// ==UserScript==
// @name         DeepRad helpers
// @namespace    http://tsai.it/
// @version      20260529.1
// @description  Add reporting helpers to DeepRad.AI.
// @author       I-Ta Tsai
// @match        http://172.17.15.97:17000/LungRADS/
// @match        http://172.17.15.97:17000/LungRADS/*
// @match        http://172.17.15.97:17000/LungRADS*
// @run-at       document-start
// @grant        GM_registerMenuCommand
// @grant        GM_setClipboard
// ==/UserScript==

(function() {
    'use strict';

    console.info('[DeepRad helpers] loaded:', location.href);

    const SERIES_BY_REPORT_FORMAT = {
        HPA: '4',
        healthCheck: '2'
    };
    const KEEP_ALIVE_INTERVAL_MS = 5 * 60 * 1000;

    function getCleanText(element) {
        return element?.textContent?.replace(/\s+/g, ' ').trim() ?? '';
    }

    function getReportSeries() {
        const reportFormat = document.querySelector('.report-selector select')?.value;
        return SERIES_BY_REPORT_FORMAT[reportFormat] ?? SERIES_BY_REPORT_FORMAT.healthCheck;
    }

    function showDeepRadStatus(message, duration = 1800) {
        let status = document.querySelector('#deeprad-helper-status');
        if (!status) {
            status = document.createElement('div');
            status.id = 'deeprad-helper-status';
            Object.assign(status.style, {
                position: 'fixed',
                left: '50%',
                top: '33vh',
                transform: 'translateX(-50%)',
                zIndex: '2147483647',
                padding: '12px 18px',
                background: '#f3f1ec',
                color: '#35414b',
                border: '1px solid #ddd8cf',
                borderRadius: '8px',
                fontSize: '17px',
                lineHeight: '1.4',
                boxShadow: '0 12px 32px rgba(48, 56, 66, 0.13)',
                maxWidth: 'min(720px, 86vw)',
                whiteSpace: 'pre-wrap'
            });
            document.body.appendChild(status);
        }

        status.textContent = message;
        status.style.display = 'block';

        clearTimeout(status._hideTimer);
        if (duration > 0) {
            status._hideTimer = setTimeout(() => {
                status.style.display = 'none';
            }, duration);
        }
    }

    function getNoduleType(row) {
        const title = row.cells[4]?.querySelector('svg')?.getAttribute('title');
        const type = (title || getCleanText(row.cells[4])).trim().toLowerCase();
        return type === 'pure ggo' ? 'non-solid' : type;
    }

    function getNoduleRows() {
        return Array.from(document.querySelectorAll('table.nodule-table tbody tr'));
    }

    function isNoduleRowSelected(row) {
        const checkbox = row.cells[6]?.querySelector('input[type="checkbox"]');
        return Boolean(checkbox && (checkbox.checked || checkbox.matches(':checked')));
    }

    function formatNoduleRow(row, series) {
        const lobe = getCleanText(row.cells[1]);
        const image = getCleanText(row.cells[2]);
        const diameter = getCleanText(row.cells[3]);
        const type = getNoduleType(row);

        if (!lobe || !image || !diameter || !type) {
            return '';
        }

        return `A ${diameter} mm ${type} nodule in the ${lobe} of lung (Srs/Img: ${series}/${image}).`;
    }

    function copyText(text) {
        if (typeof GM_setClipboard === 'function') {
            GM_setClipboard(text, 'text');
            return Promise.resolve();
        }

        const copyWithTextarea = () => {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            textarea.remove();
        };

        if (navigator.clipboard?.writeText) {
            return navigator.clipboard.writeText(text).catch(copyWithTextarea);
        }

        copyWithTextarea();
        return Promise.resolve();
    }

    function copyNoduleTable() {
        const series = getReportSeries();
        const rows = getNoduleRows().filter(isNoduleRowSelected);
        console.info(`[DeepRad helpers] selected ${rows.length} nodule row(s).`);

        const report = rows
            .map((row) => formatNoduleRow(row, series))
            .filter(Boolean)
            .join('\r\n');

        if (report) {
            copyText(report).then(() => {
                console.info(`[DeepRad helpers] copied ${rows.length} nodule(s).`);
                showDeepRadStatus(`Copied ${rows.length} nodule(s).`);
            });
        } else {
            console.warn('[DeepRad helpers] no nodules to copy.');
            showDeepRadStatus('No selected nodules to copy.', 2400);
        }
    }

    window.addEventListener('keydown', (ev) => {
        if (ev.ctrlKey && ev.shiftKey && (ev.code === 'KeyC' || ev.key.toLowerCase() === 'c')) {
            console.info('[DeepRad helpers] Ctrl+Shift+C detected.');
            copyNoduleTable();
            ev.preventDefault();
            ev.stopPropagation();
        }
    }, true);

    if (typeof GM_registerMenuCommand === 'function') {
        GM_registerMenuCommand('Copy nodule table', copyNoduleTable);
    }

    setInterval(() => {
        document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }));
        document.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Shift', code: 'ShiftLeft' }));
        window.dispatchEvent(new Event('focus'));
        console.info('[DeepRad helpers] frontend keep alive event sent.');
    }, KEEP_ALIVE_INTERVAL_MS);
})();
