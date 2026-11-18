// ==UserScript==
// @name         Surlignage + Interception PDF Passive par Site & Commune (GitHub)
// @namespace    http://tampermonkey.net/
// @version      7.0
// @description  Lit nativement l'onglet REDUCTION (sans SheetJS externe), surligne les UDB et annote le PDF Prolog
// @match        *://*/*
// @allFrames    true
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/krisscp80/prolog-reduction/main/prolog_reduction.user.js
// @downloadURL  https://raw.githubusercontent.com/krisscp80/prolog-reduction/main/prolog_reduction.user.js
// @grant        GM_addStyle
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        unsafeWindow
// @require      https://raw.githubusercontent.com/krisscp80/prolog-reduction/main/pdf-lib.min.js
// @require      https://raw.githubusercontent.com/krisscp80/prolog-reduction/main/pdf.min.js
// ==/UserScript==

(function () {
    'use strict';

    const win = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    const PDF_WORKER_LOCAL = 'file://///SERVEUR/Partage/Outils/pdf.worker.min.js';
    const DATA_KEY = 'saved_excel_udb_site_commune';

    let cachedData = GM_getValue(DATA_KEY, {});
    let isApplying = false;
    let isProcessingPdf = false;
    let observerTimer = null;

    function normaliserTexte(value) {
        if (value === null || value === undefined) return '';
        return String(value)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toUpperCase()
            .replace(/[-_'’]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function formaterQte(value) {
        const number = Number(value);
        if (!Number.isFinite(number)) return '0';
        return String(Math.round(Math.abs(number) * 1000) / 1000);
    }

    // -------------------------------------------------------------------------
    // 1. LECTEUR XLSX NATIF 100% AUTONOME (Zéro dépendance externe)
    //    Décompresse le ZIP .xlsx avec DecompressionStream (natif Chrome/Edge)
    // -------------------------------------------------------------------------
    async function lireZipXlsxNatif(arrayBuffer) {
        const view = new DataView(arrayBuffer);
        const bytes = new Uint8Array(arrayBuffer);
        const decoder = new TextDecoder('utf-8');
        const files = {};

        // Recherche de l'End of Central Directory (signature 0x06054b50)
        let eocdOffset = -1;
        for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
            if (view.getUint32(i, true) === 0x06054b50) {
                eocdOffset = i;
                break;
            }
        }
        if (eocdOffset === -1) throw new Error('Fichier .xlsx invalide (archive ZIP non reconnue).');

        const centralDirOffset = view.getUint32(eocdOffset + 16, true);
        const totalEntries = view.getUint16(eocdOffset + 10, true);
        let ptr = centralDirOffset;

        for (let e = 0; e < totalEntries; e++) {
            if (view.getUint32(ptr, true) !== 0x02014b50) break;
            const method = view.getUint16(ptr + 10, true);
            const compSize = view.getUint32(ptr + 20, true);
            const nameLen = view.getUint16(ptr + 28, true);
            const extraLen = view.getUint16(ptr + 30, true);
            const commentLen = view.getUint16(ptr + 32, true);
            const localHeaderOffset = view.getUint32(ptr + 42, true);

            const fileName = decoder.decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen));
            ptr += 46 + nameLen + extraLen + commentLen;

            if (!fileName.endsWith('.xml') && !fileName.endsWith('.rels')) continue;

            const lNameLen = view.getUint16(localHeaderOffset + 26, true);
            const lExtraLen = view.getUint16(localHeaderOffset + 28, true);
            const dataStart = localHeaderOffset + 30 + lNameLen + lExtraLen;
            const compData = bytes.subarray(dataStart, dataStart + compSize);

            if (method === 0) {
                files[fileName] = decoder.decode(compData);
            } else if (method === 8) {
                const ds = new DecompressionStream('deflate-raw');
                const writer = ds.writable.getWriter();
                writer.write(compData);
                writer.close();
                const decompBuffer = await new Response(ds.readable).arrayBuffer();
                files[fileName] = decoder.decode(decompBuffer);
            }
        }
        return files;
    }

    async function parserBufferExcelNatif(arrayBuffer) {
        if (!(arrayBuffer instanceof ArrayBuffer)) {
            throw new Error("Le fichier Excel n'a pas été chargé correctement.");
        }

        const xmlFiles = await lireZipXlsxNatif(arrayBuffer);
        const parser = new DOMParser();

        // 1. Lecture des chaînes partagées (sharedStrings.xml)
        const sharedStrings = [];
        if (xmlFiles['xl/sharedStrings.xml']) {
            const ssDoc = parser.parseFromString(xmlFiles['xl/sharedStrings.xml'], 'text/xml');
            const siNodes = ssDoc.getElementsByTagName('si');
            for (let i = 0; i < siNodes.length; i++) {
                const tNodes = siNodes[i].getElementsByTagName('t');
                let txt = '';
                for (let j = 0; j < tNodes.length; j++) {
                    txt += tNodes[j].textContent || '';
                }
                sharedStrings.push(txt);
            }
        }

        // 2. Identification du fichier XML correspondant à l'onglet REDUCTION
        const wbDoc = parser.parseFromString(xmlFiles['xl/workbook.xml'], 'text/xml');
        const relsDoc = parser.parseFromString(xmlFiles['xl/_rels/workbook.xml.rels'], 'text/xml');

        let targetRid = null;
        const sheets = wbDoc.getElementsByTagName('sheet');
        const availableSheets = [];
        for (let i = 0; i < sheets.length; i++) {
            const name = sheets[i].getAttribute('name') || '';
            availableSheets.push(name);
            if (name.trim().toUpperCase() === 'REDUCTION') {
                targetRid = sheets[i].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
                         || sheets[i].getAttribute('r:id');
            }
        }
        if (!targetRid) {
            throw new Error(`Onglet "REDUCTION" introuvable.\nOnglets disponibles : ${availableSheets.join(', ')}`);
        }

        let sheetPath = null;
        const rels = relsDoc.getElementsByTagName('Relationship');
        for (let i = 0; i < rels.length; i++) {
            if (rels[i].getAttribute('Id') === targetRid) {
                sheetPath = rels[i].getAttribute('Target').replace(/^\/+/, '');
                break;
            }
        }
        const fullSheetPath = sheetPath.startsWith('xl/') ? sheetPath : 'xl/' + sheetPath;
        const sheetXml = xmlFiles[fullSheetPath];
        if (!sheetXml) throw new Error(`Feuille XML introuvable (${fullSheetPath}).`);

        // 3. Extraction des lignes visibles 2 à 259 (Colonnes A, F, G, H)
        const sheetDoc = parser.parseFromString(sheetXml, 'text/xml');
        const rowNodes = sheetDoc.getElementsByTagName('row');
        const temporary = {};

        let currentSiteRaw = '';
        let currentSiteNorm = '';

        function lireValeurCellule(cNode) {
            if (!cNode) return '';
            const type = cNode.getAttribute('t');
            if (type === 'inlineStr') {
                const t = cNode.getElementsByTagName('t')[0];
                return t ? t.textContent : '';
            }
            const vNode = cNode.getElementsByTagName('v')[0];
            if (!vNode) return '';
            const raw = vNode.textContent || '';
            if (type === 's') {
                return sharedStrings[parseInt(raw, 10)] || '';
            }
            return raw;
        }

        for (let i = 0; i < rowNodes.length; i++) {
            const rowNode = rowNodes[i];
            const rowNum = parseInt(rowNode.getAttribute('r'), 10);
            if (isNaN(rowNum) || rowNum < 2 || rowNum > 259) continue;

            const isHidden = rowNode.getAttribute('hidden') === '1' || rowNode.getAttribute('hidden') === 'true';

            const cells = {};
            const cNodes = rowNode.getElementsByTagName('c');
            for (let j = 0; j < cNodes.length; j++) {
                const ref = cNodes[j].getAttribute('r') || '';
                const col = ref.replace(/\d+/g, '');
                cells[col] = lireValeurCellule(cNodes[j]);
            }

            // Propagation du Site (Colonne A) même si la ligne d'en-tête est masquée
            if (cells['A'] && String(cells['A']).trim() !== '') {
                currentSiteRaw = String(cells['A']).trim();
                currentSiteNorm = normaliserTexte(currentSiteRaw);
            }

            if (isHidden) continue;

            const rawUdb = String(cells['G'] || '').trim().replace(/\s+/g, '');
            if (!/^\d+$/.test(rawUdb)) continue;

            // Préserve les 7 chiffres avec zéro initial si Excel l'a stocké comme nombre (ex: 102116 -> 0102116)
            const udb = rawUdb.length === 6 ? '0' + rawUdb : rawUdb;

            const communeRaw = String(cells['F'] || '').trim();
            const communeNorm = normaliserTexte(communeRaw);
            const quantity = parseFloat(String(cells['H'] || '0').replace(',', '.').trim()) || 0;

            const groupKey = `${currentSiteNorm}|${communeNorm}`;
            if (!temporary[udb]) temporary[udb] = {};
            if (!temporary[udb][groupKey]) {
                temporary[udb][groupKey] = {
                    siteRaw: currentSiteRaw,
                    siteNorm: currentSiteNorm,
                    communeRaw,
                    communeNorm,
                    qte: quantity
                };
            } else {
                temporary[udb][groupKey].qte = Math.max(temporary[udb][groupKey].qte, quantity);
            }
        }

        const result = {};
        for (const [udb, groups] of Object.entries(temporary)) {
            result[udb] = Object.values(groups);
        }
        return result;
    }

    // -------------------------------------------------------------------------
    // 2. RÉSOLUTION SITE / COMMUNE
    // -------------------------------------------------------------------------
    function resoudreLignesBadgePourUdb(udb, pageText, compact = true) {
        const entries = cachedData[udb];
        if (!entries || entries.length === 0) return null;
        if (entries.length === 1) return [formaterQte(entries[0].qte)];

        let candidates = entries;
        const siteMatches = entries.filter(e => e.siteNorm && pageText.includes(e.siteNorm));
        if (siteMatches.length > 0) candidates = siteMatches;
        if (candidates.length === 1) return [formaterQte(candidates[0].qte)];

        const communeMatches = candidates.filter(e => e.communeNorm && pageText.includes(e.communeNorm));
        if (communeMatches.length === 1) return [formaterQte(communeMatches[0].qte)];

        const selected = communeMatches.length > 0 ? communeMatches : candidates;
        if (compact) return selected.map(e => formaterQte(e.qte));
        return [selected.map(e => `${e.communeRaw || e.siteRaw}: ${formaterQte(e.qte)}`).join(' / ')];
    }

    function setBadgeStatus(text, busy = false) {
        const badge = document.getElementById('tm-auto-badge');
        if (!badge) return;
        badge.textContent = text;
        badge.style.background = busy ? '#fff3e0' : '#e8f5e9';
        badge.style.color = busy ? '#e65100' : '#1b5e20';
        badge.style.borderColor = busy ? '#fb8c00' : '#43a047';
    }

    // -------------------------------------------------------------------------
    // 3. SURLIGNAGE ÉCRAN
    // -------------------------------------------------------------------------
    function removeHighlights() {
        document.querySelectorAll('span.tm-excel-badge').forEach(el => el.remove());
        document.querySelectorAll('mark.tm-excel-highlight').forEach(el => {
            const parent = el.parentNode;
            if (parent) {
                parent.replaceChild(document.createTextNode(el.textContent), el);
                parent.normalize();
            }
        });
        document.querySelectorAll('input.tm-excel-highlight-input').forEach(el => {
            el.classList.remove('tm-excel-highlight-input');
        });
    }

    function highlightNode(node, regex, pageText) {
        const text = node.nodeValue;
        regex.lastIndex = 0;
        if (!regex.test(text)) return;

        regex.lastIndex = 0;
        const fragment = document.createDocumentFragment();
        let lastIndex = 0;
        let match;

        while ((match = regex.exec(text)) !== null) {
            const prefix = match[1] || '';
            const udb = match[2];
            const before = text.substring(lastIndex, match.index) + prefix;
            if (before) fragment.appendChild(document.createTextNode(before));

            const mark = document.createElement('mark');
            mark.className = 'tm-excel-highlight';
            mark.textContent = udb;
            fragment.appendChild(mark);

            const lines = resoudreLignesBadgePourUdb(udb, pageText, false);
            if (lines && lines.length > 0) {
                const badge = document.createElement('span');
                badge.className = 'tm-excel-badge';
                badge.textContent = `[${lines.join(' ')}]`;
                fragment.appendChild(badge);
            }
            lastIndex = regex.lastIndex;
        }

        const after = text.substring(lastIndex);
        if (after) fragment.appendChild(document.createTextNode(after));
        if (node.parentNode) node.parentNode.replaceChild(fragment, node);
    }

    function applyHighlight() {
        if (isApplying || !document.body) return;
        cachedData = GM_getValue(DATA_KEY, cachedData);
        const keys = Object.keys(cachedData);
        if (keys.length === 0) return;

        isApplying = true;
        try {
            const pageText = normaliserTexte(document.body.innerText || document.body.textContent || '');
            const escaped = keys.sort((a, b) => b.length - a.length).map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
            const regex = new RegExp(`(^|\\D)(${escaped.join('|')})(?!\\d)`, 'g');

            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
                acceptNode(node) {
                    const parent = node.parentNode;
                    if (!parent) return NodeFilter.FILTER_REJECT;
                    const tag = parent.nodeName.toUpperCase();
                    if (['SCRIPT', 'STYLE', 'TEXTAREA', 'INPUT', 'NOSCRIPT'].includes(tag)) return NodeFilter.FILTER_REJECT;
                    if (parent.closest('#tm-excel-box, #tm-excel-mini-toggle, .tm-excel-highlight, .tm-excel-badge')) {
                        return NodeFilter.FILTER_REJECT;
                    }
                    regex.lastIndex = 0;
                    return regex.test(node.nodeValue) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
                }
            });

            const nodes = [];
            while (walker.nextNode()) nodes.push(walker.currentNode);
            nodes.forEach(node => highlightNode(node, regex, pageText));

            document.querySelectorAll('input:not(#tm-excel-file)').forEach(input => {
                const value = String(input.value || '').trim();
                if (value && cachedData[value]) input.classList.add('tm-excel-highlight-input');
            });
        } finally {
            setTimeout(() => { isApplying = false; }, 100);
        }
    }

    // -------------------------------------------------------------------------
    // 4. ANNOTATION PDF & INTERCEPTION
    // -------------------------------------------------------------------------
    function declencherTelechargement(bytes, filename) {
        const blob = new Blob([bytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    async function annoterPdfEnJs(arrayBuffer) {
        const pdfjs = win.pdfjsLib || (typeof pdfjsLib !== 'undefined' ? pdfjsLib : null);
        const pdfLib = win.PDFLib || (typeof PDFLib !== 'undefined' ? PDFLib : null);

        if (!pdfjs || !pdfLib) {
            throw new Error(
                "Les bibliothèques PDF locales (pdf.min.js / pdf-lib.min.js) sont introuvables.\n" +
                "Vérifiez que Tampermonkey a l'option « Autoriser l'accès aux URL de fichier » activée et que les fichiers sont dans le dossier partagé."
            );
        }

        // Désactive le worker externe si bloqué par le navigateur sur file://
        pdfjs.GlobalWorkerOptions.workerSrc = PDF_WORKER_LOCAL;

        const loadingTask = pdfjs.getDocument({
            data: arrayBuffer.slice(0),
            disableWorker: true
        });
        const pdfSource = await loadingTask.promise;

        const annotationsByPage = [];
        const udbKeys = Object.keys(cachedData).sort((a, b) => b.length - a.length);

        for (let pageNumber = 1; pageNumber <= pdfSource.numPages; pageNumber++) {
            const pdfPage = await pdfSource.getPage(pageNumber);
            const textContent = await pdfPage.getTextContent();
            const fullText = textContent.items.map(item => item.str || '').join(' ');
            const pageText = normaliserTexte(fullText);
            const pageAnnotations = [];

            for (const item of textContent.items) {
                const originalText = String(item.str || '').trim();
                if (!originalText) continue;
                const compactText = originalText.replace(/\s+/g, '');

                for (const udb of udbKeys) {
                    if (!compactText.includes(udb)) continue;
                    const lines = resoudreLignesBadgePourUdb(udb, pageText, true);
                    if (!lines || lines.length === 0) continue;

                    pageAnnotations.push({
                        text: originalText,
                        lines,
                        x: Number(item.transform[4]) || 0,
                        y: Number(item.transform[5]) || 0,
                        width: Number(item.width) || originalText.length * (Number(item.height) || 10) * 0.5,
                        height: Number(item.height) || 10
                    });
                    break;
                }
            }
            annotationsByPage.push(pageAnnotations);
        }

        const { PDFDocument, rgb, StandardFonts } = pdfLib;
        const pdfDocument = await PDFDocument.load(arrayBuffer);
        const font = await pdfDocument.embedFont(StandardFonts.HelveticaBold);
        const pages = pdfDocument.getPages();

        for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
            const page = pages[pageIndex];
            const annotations = annotationsByPage[pageIndex] || [];

            for (const item of annotations) {
                const udbFontSize = Math.max(item.height, 6);
                const boldUdbWidth = font.widthOfTextAtSize(item.text, udbFontSize);
                const coverWidth = Math.max(item.width, boldUdbWidth) + 2;
                const centerX = item.x + item.width / 2;

                // Masque l'UDB fine d'origine et la réécrit en gras sans bordure
                page.drawRectangle({
                    x: centerX - coverWidth / 2,
                    y: item.y - 1.5,
                    width: coverWidth,
                    height: udbFontSize + 2.5,
                    color: rgb(0.86, 0.86, 0.86)
                });

                page.drawText(item.text, {
                    x: centerX - boldUdbWidth / 2,
                    y: item.y,
                    size: udbFontSize,
                    font,
                    color: rgb(0, 0, 0)
                });

                const lineCount = item.lines.length;
                const badgeFontSize = lineCount > 1 ? Math.max(item.height * 0.55, 5.5) : Math.max(item.height * 0.7, 6.5);
                const lineHeight = badgeFontSize + 1.8;

                let maxTextWidth = 0;
                for (const line of item.lines) {
                    maxTextWidth = Math.max(maxTextWidth, font.widthOfTextAtSize(` ${line} `, badgeFontSize));
                }

                const badgeWidth = maxTextWidth + 2;
                const badgeHeight = lineCount * lineHeight + 1.5;
                const badgeX = centerX - badgeWidth / 2;
                const badgeY = item.y + item.height + 1.5;

                page.drawRectangle({
                    x: badgeX,
                    y: badgeY,
                    width: badgeWidth,
                    height: badgeHeight,
                    color: rgb(0, 0, 0)
                });

                item.lines.forEach((line, index) => {
                    const text = ` ${line} `;
                    const textWidth = font.widthOfTextAtSize(text, badgeFontSize);
                    const textX = centerX - textWidth / 2;
                    const textY = badgeY + badgeHeight - (index + 1) * lineHeight + 1;
                    page.drawText(text, { x: textX, y: textY, size: badgeFontSize, font, color: rgb(1, 1, 1) });
                });
            }
        }

        return pdfDocument.save();
    }

    async function intercepterEtAnnoter(linkId) {
        if (isProcessingPdf || Object.keys(cachedData).length === 0) return false;
        isProcessingPdf = true;
        setBadgeStatus('⏳ Annotation PDF...', true);

        try {
            const target = document.getElementById(linkId);
            const form = target ? target.closest('form') : document.getElementById('form_body');
            if (!form) throw new Error("Formulaire 'form_body' introuvable.");

            const formData = new FormData(form);
            const control = document.getElementById(linkId) || form.elements.namedItem(linkId);
            formData.set(control ? (control.name || linkId) : linkId, control ? (control.value || linkId) : linkId);

            if (target) {
                const onclick = target.getAttribute('onclick') || '';
                const parameterRegex = /['"]([^'"]+)['"]\s*:\s*['"]([^'"]+)['"]/g;
                let match;
                while ((match = parameterRegex.exec(onclick)) !== null) {
                    formData.set(match[1], match[2]);
                }
            }

            const action = form.getAttribute('action') || window.location.href;
            const response = await fetch(action, {
                method: 'POST',
                body: new URLSearchParams(formData),
                credentials: 'same-origin',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
            });

            if (!response.ok) throw new Error(`Erreur serveur HTTP ${response.status}`);

            const pdfBuffer = await response.arrayBuffer();
            const header = new TextDecoder().decode(new Uint8Array(pdfBuffer.slice(0, 5)));
            if (header !== '%PDF-') throw new Error('La réponse serveur ne contient pas un PDF valide.');

            const annotatedPdf = await annoterPdfEnJs(pdfBuffer);

            let filename = 'Document_Prolog_annote.pdf';
            const disposition = response.headers.get('content-disposition');
            if (disposition && disposition.includes('filename=')) {
                const filenameMatch = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                if (filenameMatch && filenameMatch[1]) {
                    filename = filenameMatch[1].replace(/['"]/g, '').replace(/\.pdf$/i, '_annote.pdf');
                }
            }

            declencherTelechargement(annotatedPdf, filename);
        } catch (error) {
            console.error('[Interception PDF]', error);
            alert("Erreur lors de l'annotation PDF :\n" + error.message);
        } finally {
            isProcessingPdf = false;
            setBadgeStatus('🛡️ Auto-PDF Actif', false);
        }
        return true;
    }

    function installerInterception() {
        if (typeof win.clickLink === 'function' && !win.clickLink.__tmHooked) {
            const original = win.clickLink;
            const hooked = function (id) {
                if (id && id.includes('getFile') && Object.keys(cachedData).length > 0) {
                    intercepterEtAnnoter(id);
                    return false;
                }
                return original.apply(this, arguments);
            };
            hooked.__tmHooked = true;
            win.clickLink = hooked;
        }

        const prototype = win.HTMLFormElement && win.HTMLFormElement.prototype;
        const nativeSubmit = prototype && prototype.submit;
        if (typeof nativeSubmit === 'function' && !nativeSubmit.__tmHooked) {
            const hookedSubmit = function () {
                if (this.id === 'form_body' && Object.keys(cachedData).length > 0) {
                    const input = this.querySelector('input[name*="getFile"]');
                    if (input) {
                        intercepterEtAnnoter(input.name);
                        return;
                    }
                }
                return nativeSubmit.apply(this, arguments);
            };
            hookedSubmit.__tmHooked = true;
            prototype.submit = hookedSubmit;
        }
    }

    // -------------------------------------------------------------------------
    // 5. INTERFACE UTILISATEUR (Déplaçable + Raccourci Alt+E)
    // -------------------------------------------------------------------------
    function initialiserInterface() {
        if (window.self !== window.top || !document.body || document.getElementById('tm-excel-box')) return;

        GM_addStyle(`
            #tm-excel-box {
                position: fixed !important; right: 20px; bottom: 20px; z-index: 9999999 !important;
                background: #ffffff !important; border: 2px solid #003da5 !important; border-radius: 8px !important;
                padding: 8px 12px !important; display: flex !important; align-items: center !important; gap: 8px !important;
                font: 12px Arial, sans-serif !important; box-shadow: 0 4px 15px rgba(0,0,0,.35) !important; user-select: none !important;
            }
            #tm-excel-drag { cursor: move !important; font-weight: bold !important; color: #003da5 !important; }
            #tm-excel-box button { cursor: pointer !important; padding: 4px 8px !important; border-radius: 4px !important; font-weight: bold !important; }
            #tm-excel-btn-choose { background: #e3f2fd !important; border: 1px solid #1976d2 !important; color: #0d47a1 !important; }
            #tm-excel-clear { background: #ffebee !important; border: 1px solid #d32f2f !important; color: #b71c1c !important; }
            #tm-excel-hide { background: #eceff1 !important; border: 1px solid #546e7a !important; color: #263238 !important; }
            #tm-auto-badge { background: #e8f5e9 !important; border: 1px solid #43a047 !important; color: #1b5e20 !important; padding: 4px 7px !important; border-radius: 4px !important; font-weight: bold !important; }
            .tm-excel-badge { display: inline-block !important; margin-left: 4px !important; padding: 1px 5px !important; font-size: 11px !important; font-weight: bold !important; border-radius: 3px !important; background: #e3f2fd !important; color: #1565c0 !important; border: 1px solid #90caf9 !important; }
            .tm-excel-highlight { background: #fff176 !important; color: #000000 !important; font-weight: bold !important; border: 1px solid #fbc02d !important; }
            .tm-excel-highlight-input { background: #fff176 !important; color: #000000 !important; border: 2px solid #fbc02d !important; font-weight: bold !important; }
            #tm-excel-mini-toggle {
                position: fixed !important; right: 12px; bottom: 12px; z-index: 9999999 !important; width: 32px !important; height: 32px !important;
                display: none !important; border-radius: 50% !important; cursor: move !important; background: #003da5 !important; color: #ffffff !important; border: 2px solid #ffffff !important;
            }
            @media print { #tm-excel-box, #tm-excel-mini-toggle { display: none !important; } }
        `);

        const box = document.createElement('div');
        box.id = 'tm-excel-box';
        box.innerHTML = `
            <span id="tm-excel-drag" title="Maintenir le clic pour déplacer">⋮⋮ 📊 Excel :</span>
            <span id="tm-excel-status">Aucune UDB en mémoire</span>
            <button id="tm-excel-btn-choose" type="button">🔄 Charger</button>
            <input id="tm-excel-file" type="file" accept=".xlsx,.xls" style="display:none;">
            <button id="tm-excel-clear" type="button" style="display:none;">🗑️ Vider</button>
            <span id="tm-auto-badge">🛡️ Auto-PDF Actif</span>
            <button id="tm-excel-hide" type="button" title="Masquer (Alt + E)">👁️ Masquer</button>
        `;
        document.body.appendChild(box);

        const mini = document.createElement('button');
        mini.id = 'tm-excel-mini-toggle';
        mini.type = 'button';
        mini.textContent = '📊';
        mini.title = 'Afficher (Alt + E) — Glisser pour déplacer';
        document.body.appendChild(mini);

        const fileInput = document.getElementById('tm-excel-file');
        document.getElementById('tm-excel-btn-choose').addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', async event => {
            const file = event.target.files && event.target.files[0];
            if (!file) return;
            try {
                const buffer = await file.arrayBuffer();
                const structured = await parserBufferExcelNatif(buffer);
                const count = Object.keys(structured).length;
                if (count === 0) throw new Error('Aucune UDB valide trouvée dans REDUCTION.');

                cachedData = structured;
                GM_setValue(DATA_KEY, cachedData);
                removeHighlights();
                applyHighlight();
                updateUI();
            } catch (error) {
                console.error('[Lecture Excel]', error);
                alert('Erreur de lecture Excel :\n' + error.message);
            } finally {
                fileInput.value = '';
            }
        });

        document.getElementById('tm-excel-clear').addEventListener('click', () => {
            cachedData = {};
            GM_setValue(DATA_KEY, {});
            removeHighlights();
            updateUI();
        });

        const toggleBox = () => {
            const isHidden = box.style.display === 'none';
            box.style.setProperty('display', isHidden ? 'flex' : 'none', 'important');
            mini.style.setProperty('display', isHidden ? 'none' : 'block', 'important');
        };

        document.getElementById('tm-excel-hide').addEventListener('click', toggleBox);
        mini.addEventListener('click', toggleBox);
        document.addEventListener('keydown', e => {
            if (e.altKey && !e.ctrlKey && e.key.toLowerCase() === 'e') {
                e.preventDefault();
                toggleBox();
            }
        });

        // Déplacement manuel du bandeau
        let drag = false, sx = 0, sy = 0, ol = 0, ot = 0;
        document.getElementById('tm-excel-drag').addEventListener('mousedown', e => {
            drag = true;
            const r = box.getBoundingClientRect();
            sx = e.clientX; sy = e.clientY; ol = r.left; ot = r.top;
            box.style.setProperty('left', ol + 'px', 'important');
            box.style.setProperty('top', ot + 'px', 'important');
            box.style.setProperty('right', 'auto', 'important');
            box.style.setProperty('bottom', 'auto', 'important');
        });
        document.addEventListener('mousemove', e => {
            if (!drag) return;
            box.style.setProperty('left', Math.max(0, ol + e.clientX - sx) + 'px', 'important');
            box.style.setProperty('top', Math.max(0, ot + e.clientY - sy) + 'px', 'important');
        });
        document.addEventListener('mouseup', () => { drag = false; });

        updateUI();
    }

    function updateUI() {
        const status = document.getElementById('tm-excel-status');
        const clearButton = document.getElementById('tm-excel-clear');
        const count = Object.keys(cachedData).length;
        if (status) status.textContent = count > 0 ? `${count} UDB en mémoire` : 'Aucune UDB en mémoire';
        if (clearButton) clearButton.style.display = count > 0 ? 'inline-block' : 'none';
    }

    document.addEventListener('click', event => {
        const target = event.target.closest && event.target.closest('[id*="getFile"]');
        if (target && Object.keys(cachedData).length > 0) {
            event.preventDefault();
            event.stopPropagation();
            intercepterEtAnnoter(target.id);
        }
    }, true);

    function demarrer() {
        initialiserInterface();
        updateUI();
        applyHighlight();
        setInterval(installerInterception, 400);

        const observer = new MutationObserver(() => {
            if (isApplying) return;
            if (observerTimer) clearTimeout(observerTimer);
            observerTimer = setTimeout(applyHighlight, 600);
        });
        if (document.body) observer.observe(document.body, { childList: true, subtree: true });
    }

    if (document.body) demarrer();
    else window.addEventListener('DOMContentLoaded', demarrer, { once: true });
})();