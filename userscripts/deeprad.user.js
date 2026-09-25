// ==UserScript==
// @name         DeepRad helpers
// @namespace    http://tsai.it/
// @version      20260917.1
// @description  Add reporting helpers to DeepRad.AI.
// @author       I-Ta Tsai
// @match        http://172.17.15.97:17000/
// @match        http://172.17.15.97:17000/*
// @run-at       document-start
// @grant        GM_registerMenuCommand
// @grant        GM_setClipboard
// @grant        unsafeWindow
// ==/UserScript==

/**
 * ============================================================================
 * DeepRad Helpers (胸部電腦斷層肺結節判讀輔助 Userscript)
 * ============================================================================
 *
 * 【專案簡介】
 * 本腳本專為輔助放射科醫師操作 DeepRad.AI（胸部電腦斷層肺結節 AI 輔助判讀系統）
 * 設計，主要解決報告系統文字轉換整合、連線 Session 維持、防跳轉與消除干擾通知等問題。
 *
 * 【核心功能與特色】
 *
 * 1. 結節快速擷取與格式化報告 (Nodule Reporting Helpers)：
 *    - 國健署肺癌篩檢格式 (HPA Format, 快速鍵：Ctrl + Shift + C)：
 *      - 自動統計並聚合已選取（勾選）的肺結節，依肺葉歸納 Image 編號。
 *      - 保證解剖肺葉輸出順序：RUL -> RML -> RLL -> LUL -> LLL。
 *      - 自動去重 (De-duplicate) 並依影像序號由小到大排序，輸出如：RUL:12;15;LUL:45;LLL:88。
 *    - 健檢與臨床完整格式 (Health Check Format, 快速鍵：Alt + Shift + C)：
 *      - 輸出結構化英文章句（如：A 6.5 mm (6.2 x 6.8 mm) sub-solid nodule (solid part: 3.2 mm) in the RUL of lung (Srs/Img: 2/45).）。
 *      - 智慧長短軸量測 (Axis Auto-detection)：結節分級達 Lung-RADS 3 或以上時，自動自 .key-film 讀取雙軸數據並填入。
 *      - 實質成分萃取 (Solid Part Auto-detection)：對 subsolid / part-solid 結節，自動擷取表單中的實質成分大小 (solid part: X mm)。
 *      - 純磨玻璃結節 (Pure GGO) 正規化為 non-solid。
 *      - 批次查詢最佳化：同時需要長短軸與實質成分時僅模擬切換一次點擊，讀取後自動點選復原原本畫面選取狀態。
 *    - RIS 換行相容性：報告行尾採用 CRLF (\r\n)，確保直接貼入醫院 RIS 報告系統時不會遺失換行。
 *    - 剪貼簿支援：支援 GM_setClipboard、navigator.clipboard 與 textarea fallback。
 *
 * 2. 智慧 Session 防呆與 Token 自動保活 (Session & Token Watchdog)：
 *    - 攔截 localStorage 清除：監聽並防範前端閒置定時器刪除 token，若非手動登出則自動從備份復原。
 *    - 阻擋前端跳轉至登入頁：攔截 pushState / replaceState，防止有登入憑證時被強制轉向 /login。
 *    - 記錄最後瀏覽頁面：於登入頁面載入且成功復原 Token 時，自動 replace 回原本工作路徑。
 *    - 區分主動登出：監聽使用者點擊 .logout 按鈕，手動登出時放行正常登出流程，不硬阻擋。
 *    - API 鑑權標頭修補 (Auth Header Patching)：攔截 fetch 與 XMLHttpRequest，確保所有 API 請求攜帶最新 Authorization Header。
 *    - 背景定期展延 (Token Refresh)：每 10 分鐘自動向 /auth/refresh_token 要求新 Token，並分發 storage 與自訂事件。
 *    - 前端活動模擬 (Keep-Alive)：每 5 分鐘模擬發送 mousemove、Shift 鍵與 focus 事件，防止前端純 UI 閒置超時。
 *
 * 3. 靜音式逾時通知自動關閉 (Session Expired Notification Auto-dismissal)：
 *    - 使用 MutationObserver 監控頁面 Sonner / Toast 通知。
 *    - 偵測到「session expired」或「please login again」時，透過 CSS display: none 隱藏並觸發關閉按鈕。
 *    - 避免直接操作刪除 DOM 節點，防止引發 React Virtual DOM unmount 崩潰 (removeChild 錯誤)。
 *
 * 4. 操作介面與回饋 (User Interface & Feedback)：
 *    - 畫面居中 HUD 狀態通知提示框 (Status Toast)，1.8 ~ 2.4 秒後自動淡出，不阻礙閱片視野。
 *    - 註冊 Tampermonkey / 油猴擴充腳本選單指令，提供手動觸發複製與即時 Token 刷新選項。
 *
 * 【快速鍵對照】
 * - Ctrl + Shift + C : 複製國健署篩檢格式結節 (Copy HPA nodules)
 * - Alt + Shift + C  : 複製健檢 / 完整描述格式結節 (Copy Health check nodules)
 * ============================================================================
 */

(function() {
    'use strict';

    console.info('[DeepRad helpers] loaded:', location.href);

    const SERIES_BY_REPORT_FORMAT = {
        HPA: '4',
        healthCheck: '2'
    };
    const LOBE_ORDER = ['RUL', 'RML', 'RLL', 'LUL', 'LLL'];
    const KEEP_ALIVE_INTERVAL_MS = 5 * 60 * 1000;
    const TOKEN_REFRESH_INTERVAL_MS = 10 * 60 * 1000;
    const TOKEN_BACKUP_KEY = 'deepradHelperTokenBackup';
    const LAST_PATH_KEY = 'deepradHelperLastPath';
    const MANUAL_LOGOUT_KEY = 'deepradHelperManualLogoutAt';

    function getLatestAuthHeader() {
        const token = localStorage.getItem('token');
        return token ? `Bearer ${token}` : '';
    }

    function isManualLogoutRecent() {
        const logoutAt = Number(sessionStorage.getItem(MANUAL_LOGOUT_KEY) || 0);
        return logoutAt && Date.now() - logoutAt < 60 * 1000;
    }

    function rememberToken(token) {
        if (token) {
            sessionStorage.setItem(TOKEN_BACKUP_KEY, token);
        }
    }

    function rememberCurrentPath() {
        if (location.pathname !== '/login') {
            sessionStorage.setItem(LAST_PATH_KEY, location.href);
        }
    }

    function restoreTokenIfNeeded() {
        if (localStorage.getItem('token') || isManualLogoutRecent()) {
            return false;
        }

        const backupToken = sessionStorage.getItem(TOKEN_BACKUP_KEY);
        if (!backupToken) {
            return false;
        }

        localStorage.setItem('token', backupToken);
        console.warn('[DeepRad helpers] restored token cleared by frontend timer.');
        dismissSessionExpiredNotifications();
        return true;
    }

    function isLoginUrl(url) {
        try {
            return new URL(url, location.href).pathname === '/login';
        } catch {
            return String(url).includes('/login');
        }
    }

    function patchPageAuthHeaders() {
        const pageWindow = typeof unsafeWindow === 'object' ? unsafeWindow : window;

        if (pageWindow.fetch && !pageWindow.fetch._deepradPatched) {
            const originalFetch = pageWindow.fetch.bind(pageWindow);
            const patchedFetch = (input, init = {}) => {
                const authHeader = getLatestAuthHeader();
                if (!authHeader) {
                    return originalFetch(input, init);
                }

                if (input instanceof Request) {
                    const headers = new Headers(init.headers || input.headers);
                    headers.set('Authorization', authHeader);
                    return originalFetch(input, { ...init, headers });
                }

                const headers = new Headers(init.headers || {});
                headers.set('Authorization', authHeader);
                return originalFetch(input, { ...init, headers });
            };
            patchedFetch._deepradPatched = true;
            pageWindow.fetch = patchedFetch;
        }

        const xhrPrototype = pageWindow.XMLHttpRequest?.prototype;
        if (xhrPrototype && !xhrPrototype.setRequestHeader._deepradPatched) {
            const originalSetRequestHeader = xhrPrototype.setRequestHeader;
            xhrPrototype.setRequestHeader = function(name, value) {
                if (String(name).toLowerCase() === 'authorization') {
                    const authHeader = getLatestAuthHeader();
                    return originalSetRequestHeader.call(this, name, authHeader || value);
                }

                return originalSetRequestHeader.call(this, name, value);
            };
            xhrPrototype.setRequestHeader._deepradPatched = true;
        }
    }

    function patchFrontendLogoutTimer() {
        const pageWindow = typeof unsafeWindow === 'object' ? unsafeWindow : window;
        const storagePrototype = pageWindow.Storage?.prototype;

        if (storagePrototype && !storagePrototype.removeItem._deepradPatched) {
            const originalRemoveItem = storagePrototype.removeItem;
            storagePrototype.removeItem = function(key) {
                if (this === pageWindow.localStorage && key === 'token' && restoreTokenIfNeeded()) {
                    return undefined;
                }
                return originalRemoveItem.call(this, key);
            };
            storagePrototype.removeItem._deepradPatched = true;
        }

        if (storagePrototype && !storagePrototype.clear._deepradPatched) {
            const originalClear = storagePrototype.clear;
            storagePrototype.clear = function() {
                const backupToken = this === pageWindow.localStorage ? sessionStorage.getItem(TOKEN_BACKUP_KEY) : '';
                const result = originalClear.call(this);
                if (backupToken && !isManualLogoutRecent()) {
                    pageWindow.localStorage.setItem('token', backupToken);
                    console.warn('[DeepRad helpers] restored token after localStorage.clear().');
                }
                return result;
            };
            storagePrototype.clear._deepradPatched = true;
        }

        for (const methodName of ['pushState', 'replaceState']) {
            if (!pageWindow.history?.[methodName] || pageWindow.history[methodName]._deepradPatched) {
                continue;
            }

            const originalHistoryMethod = pageWindow.history[methodName];
            pageWindow.history[methodName] = function(state, title, url) {
                if (url && isLoginUrl(url) && restoreTokenIfNeeded()) {
                    console.warn(`[DeepRad helpers] blocked frontend ${methodName} to /login.`);
                    return undefined;
                }
                return originalHistoryMethod.call(this, state, title, url);
            };
            pageWindow.history[methodName]._deepradPatched = true;
        }

        document.addEventListener('click', (ev) => {
            const logoutButton = ev.target?.closest?.('.logout, button.logout');
            if (logoutButton) {
                sessionStorage.setItem(MANUAL_LOGOUT_KEY, String(Date.now()));
            }
        }, true);
    }

    function dismissSessionExpiredNotifications() {
        const toastCandidates = document.querySelectorAll(
            '[data-sonner-toast], [data-sonner-toaster] li, section[aria-label*="Notification" i] li'
        );

        for (const toast of toastCandidates) {
            if (toast.dataset.deepradDismissed) {
                continue;
            }

            const text = (toast.textContent || '').trim();
            if (/session expired/i.test(text) || /please login again/i.test(text)) {
                toast.dataset.deepradDismissed = 'true';
                toast.style.setProperty('display', 'none', 'important');

                const closeButton = toast.querySelector(
                    'button[data-action], button[data-close-button], button[data-button], button'
                );
                if (closeButton) {
                    try {
                        closeButton.click();
                    } catch {
                        try {
                            closeButton.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                        } catch {}
                    }
                }

                console.info('[DeepRad helpers] auto-dismissed session expired notification.');
            }
        }
    }

    function startSessionExpiredNotificationObserver() {
        const observer = new MutationObserver((mutations) => {
            for (const mutation of mutations) {
                if (mutation.addedNodes.length > 0) {
                    dismissSessionExpiredNotifications();
                    break;
                }
            }
        });

        const target = document.documentElement || document;
        if (target) {
            observer.observe(target, { childList: true, subtree: true });
        }

        dismissSessionExpiredNotifications();
    }

    function getCleanText(element) {
        return element?.textContent?.replace(/\s+/g, ' ').trim() ?? '';
    }

    function getReportFormat() {
        return document.querySelector('.report-selector select')?.value;
    }

    function getReportSeries(targetFormat) {
        const reportFormat = targetFormat || getReportFormat();
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

    function getColumnIndices(table) {
        const headers = Array.from(table?.querySelectorAll('thead th') || []);
        const indices = {
            no: 0,
            lobe: 1,
            slice: 2,
            diameter: 3,
            type: 4,
            lungRads: 5,
            select: 6
        };
        headers.forEach((th, idx) => {
            const text = getCleanText(th).toLowerCase();
            if (text.includes('no.')) indices.no = idx;
            else if (text.includes('lobe')) indices.lobe = idx;
            else if (text.includes('slice')) indices.slice = idx;
            else if (text.includes('diameter')) indices.diameter = idx;
            else if (text.includes('type')) indices.type = idx;
            else if (text.includes('lung-rads') || text.includes('rads')) indices.lungRads = idx;
            else if (text.includes('select')) indices.select = idx;
        });
        return indices;
    }

    function getNoduleType(row, typeColIndex = 4) {
        const cell = row.cells[typeColIndex];
        const title = cell?.querySelector('svg')?.getAttribute('title') || cell?.querySelector('[title]')?.getAttribute('title');
        const type = (title || getCleanText(cell)).trim().toLowerCase();
        return type === 'pure ggo' ? 'non-solid' : type;
    }

    function getNoduleLungRads(row, lungRadsColIndex = 5) {
        const cell = row.cells[lungRadsColIndex];
        return getCleanText(cell);
    }

    function isLungRads3OrAbove(lungRads) {
        const match = String(lungRads || '').trim().match(/^(\d+)/);
        return match ? Number(match[1]) >= 3 : false;
    }

    function isSubsolidType(type) {
        return /sub-?solid|part-?solid/i.test(String(type || ''));
    }

    function normalizeSolidPart(str) {
        if (!str || str === '--') return '';
        const match = String(str).trim().match(/^([0-9]+(?:\.[0-9]+)?)/);
        if (!match) return '';
        const num = parseFloat(match[1]);
        if (isNaN(num) || num <= 0) return '';
        return `${match[1]} mm`;
    }

    function normalizeAxis(str) {
        if (!str || str === '--') return '';
        const match = String(str).trim().match(/([0-9.]+)\s*[*xX×,]\s*([0-9.]+)/);
        if (match) {
            return `${match[1]} x ${match[2]} mm`;
        }
        return String(str).trim();
    }

    function getAxisFromKeyFilmDOM() {
        const keyFilm = document.querySelector('.key-film');
        if (!keyFilm) return '';
        const pElements = keyFilm.querySelectorAll('.view-col p, p');
        for (const p of pElements) {
            const text = getCleanText(p);
            const match = text.match(/Axis:\s*([0-9.]+\s*[*xX×,]\s*[0-9.]+(?:\s*mm)?)/i);
            if (match) {
                return normalizeAxis(match[1]);
            }
        }
        return '';
    }

    function getSolidPartFromDOM() {
        const input = document.querySelector('.form-table input[name="solid_part"], input[name="solid_part"]');
        let val = (input?.value || input?.getAttribute?.('value') || '').trim();

        if (!val) {
            const thElements = document.querySelectorAll('.form-table th, th');
            for (const th of thElements) {
                if (/solid\s*part/i.test(getCleanText(th))) {
                    const cell = th.closest('tr')?.querySelector('td');
                    const rowInput = cell?.querySelector('input');
                    val = (rowInput?.value || rowInput?.getAttribute?.('value') || getCleanText(cell) || '').trim();
                    if (val) break;
                }
            }
        }

        return normalizeSolidPart(val);
    }

    function isRowCurrentlySelected(row) {
        return row.classList.contains('selected') || Boolean(row.querySelector('td.selected'));
    }

    async function getNoduleAxisByDOM(row) {
        if (isRowCurrentlySelected(row)) {
            return getAxisFromKeyFilmDOM();
        }

        try {
            const clickable = row.querySelector('td:not(:last-child)') || row;
            clickable.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
            await new Promise((resolve) => setTimeout(resolve, 60));
            return getAxisFromKeyFilmDOM();
        } catch {
            return '';
        }
    }

    async function getNoduleSolidPartByDOM(row) {
        if (isRowCurrentlySelected(row)) {
            return getSolidPartFromDOM();
        }

        try {
            const clickable = row.querySelector('td:not(:last-child)') || row;
            clickable.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
            await new Promise((resolve) => setTimeout(resolve, 60));
            return getSolidPartFromDOM();
        } catch {
            return '';
        }
    }

    function getNoduleRows() {
        return Array.from(document.querySelectorAll('table.nodule-table tbody tr'));
    }

    function isNoduleRowSelected(row, selectColIndex = 6) {
        const checkbox = row.cells[selectColIndex]?.querySelector('input[type="checkbox"]') ||
            row.querySelector('input[type="checkbox"]');
        return Boolean(checkbox && (checkbox.checked || checkbox.matches(':checked')));
    }

    function formatNoduleRow(row, series, axis = '', solidPart = '', cols = { lobe: 1, slice: 2, diameter: 3, type: 4, lungRads: 5 }) {
        if (typeof solidPart === 'object' && solidPart !== null) {
            cols = solidPart;
            solidPart = '';
        }

        const lobe = getCleanText(row.cells[cols.lobe ?? 1]);
        const image = getCleanText(row.cells[cols.slice ?? 2]);
        const diameter = getCleanText(row.cells[cols.diameter ?? 3]);
        const type = getNoduleType(row, cols.type ?? 4);
        const lungRads = getNoduleLungRads(row, cols.lungRads ?? 5);

        if (!lobe || !image || !diameter || !type) {
            return '';
        }

        const normAxis = isLungRads3OrAbove(lungRads) ? normalizeAxis(axis) : '';
        const axisPart = normAxis ? ` (${normAxis})` : '';
        const normSolidPart = isSubsolidType(type) ? normalizeSolidPart(solidPart) : '';
        const solidPartText = normSolidPart ? ` (solid part: ${normSolidPart})` : '';

        return `A ${diameter} mm${axisPart} ${type} nodule${solidPartText} in the ${lobe} of lung (Srs/Img: ${series}/${image}).`;
    }

    function formatHpaNoduleRows(rows, cols = { lobe: 1, slice: 2 }) {
        const imagesByLobe = new Map();

        for (const row of rows) {
            const lobe = getCleanText(row.cells[cols.lobe ?? 1]);
            const imageText = getCleanText(row.cells[cols.slice ?? 2]);
            const image = Number(imageText);
            if (!lobe || !imageText || !Number.isInteger(image)) {
                continue;
            }

            if (!imagesByLobe.has(lobe)) {
                imagesByLobe.set(lobe, new Set());
            }
            imagesByLobe.get(lobe).add(image);
        }

        const orderedLobes = LOBE_ORDER.concat(
            Array.from(imagesByLobe.keys()).filter((lobe) => !LOBE_ORDER.includes(lobe))
        );

        return orderedLobes
            .filter((lobe) => imagesByLobe.has(lobe))
            .map((lobe) => {
                const images = Array.from(imagesByLobe.get(lobe)).sort((a, b) => a - b);
                return `${lobe}:${images.join(';')}`;
            })
            .join(';');
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

    function findTokenFromRefreshResponse(data) {
        if (!data || typeof data !== 'object') {
            return '';
        }

        return data.token ||
            data.access_token ||
            data.accessToken ||
            data?.data?.token ||
            data?.data?.access_token ||
            data?.data?.accessToken ||
            '';
    }

    function getResponseKeys(data) {
        if (!data || typeof data !== 'object') {
            return [];
        }

        const keys = Object.keys(data);
        if (data.data && typeof data.data === 'object') {
            return keys.concat(Object.keys(data.data).map((key) => `data.${key}`));
        }

        return keys;
    }

    function updateStoredToken(token) {
        const oldToken = localStorage.getItem('token');
        localStorage.setItem('token', token);
        rememberToken(token);
        window.dispatchEvent(new StorageEvent('storage', {
            key: 'token',
            oldValue: oldToken,
            newValue: token,
            storageArea: localStorage,
            url: location.href
        }));
        window.dispatchEvent(new CustomEvent('deeprad-token-refreshed'));
    }

    async function refreshToken() {
        const token = localStorage.getItem('token');
        if (!token) {
            console.warn('[DeepRad helpers] token refresh skipped: localStorage token not found.');
            return;
        }

        try {
            const response = await fetch('/auth/refresh_token', {
                method: 'POST',
                credentials: 'include',
                cache: 'no-store',
                headers: {
                    Accept: 'application/json, text/plain, */*',
                    'Content-Type': 'application/x-www-form-urlencoded',
                    Authorization: `Bearer ${token}`
                }
            });

            if (!response.ok) {
                console.warn(`[DeepRad helpers] token refresh failed: HTTP ${response.status}.`);
                if (response.status === 401 || response.status === 403) {
                    sessionStorage.removeItem(TOKEN_BACKUP_KEY);
                }
                return;
            }

            const text = await response.text();
            const data = text ? JSON.parse(text) : null;
            const refreshedToken = findTokenFromRefreshResponse(data);
            if (refreshedToken) {
                updateStoredToken(refreshedToken);
                console.info('[DeepRad helpers] token refreshed.');
            } else {
                rememberToken(token);
                console.info('[DeepRad helpers] token refresh request succeeded without token field:', getResponseKeys(data));
            }
        } catch (err) {
            console.warn('[DeepRad helpers] token refresh failed:', err);
        }
    }

    async function copyNoduleTable(targetFormat = 'HPA') {
        const series = getReportSeries(targetFormat);
        const table = document.querySelector('table.nodule-table');
        const cols = getColumnIndices(table);
        const rows = getNoduleRows().filter((row) => isNoduleRowSelected(row, cols.select));
        const formatLabel = targetFormat === 'HPA' ? 'HPA' : 'Health Check';
        console.info(`[DeepRad helpers] selected ${rows.length} ${formatLabel} nodule row(s).`);

        if (targetFormat === 'HPA') {
            const report = formatHpaNoduleRows(rows, cols);
            if (report) {
                await copyText(report);
                console.info(`[DeepRad helpers] copied ${rows.length} HPA nodule(s).`);
                showDeepRadStatus(`Copied ${rows.length} HPA nodule(s).`);
            } else {
                console.warn('[DeepRad helpers] no nodules to copy.');
                showDeepRadStatus('No selected nodules to copy.', 2400);
            }
            return;
        }

        const initialSelectedRow = getNoduleRows().find(isRowCurrentlySelected);
        const formattedRows = [];
        let switchedRow = false;

        for (const row of rows) {
            const type = getNoduleType(row, cols.type);
            const lungRads = getNoduleLungRads(row, cols.lungRads);
            const needAxis = isLungRads3OrAbove(lungRads);
            const needSolidPart = isSubsolidType(type);

            let axis = '';
            let solidPart = '';

            if (needAxis || needSolidPart) {
                if (!isRowCurrentlySelected(row)) {
                    try {
                        const clickable = row.querySelector('td:not(:last-child)') || row;
                        clickable.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                        await new Promise((resolve) => setTimeout(resolve, 60));
                        switchedRow = true;
                    } catch {}
                }

                if (needAxis) {
                    axis = getAxisFromKeyFilmDOM();
                }
                if (needSolidPart) {
                    solidPart = getSolidPartFromDOM();
                }
            }

            const formatted = formatNoduleRow(row, series, axis, solidPart, cols);
            if (formatted) {
                formattedRows.push(formatted);
            }
        }

        if (switchedRow && initialSelectedRow) {
            try {
                const restoreClickable = initialSelectedRow.querySelector('td:not(:last-child)') || initialSelectedRow;
                restoreClickable.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
            } catch {}
        }

        const report = formattedRows.join('\r\n');

        if (report) {
            await copyText(report);
            console.info(`[DeepRad helpers] copied ${rows.length} Health Check nodule(s).`);
            showDeepRadStatus(`Copied ${rows.length} Health Check nodule(s).`);
        } else {
            console.warn('[DeepRad helpers] no nodules to copy.');
            showDeepRadStatus('No selected nodules to copy.', 2400);
        }
    }

    window.addEventListener('keydown', (ev) => {
        const isKeyC = ev.code === 'KeyC' || (ev.key && ev.key.toLowerCase() === 'c');
        if (ev.ctrlKey && !ev.altKey && ev.shiftKey && isKeyC) {
            console.info('[DeepRad helpers] Ctrl+Shift+C detected (HPA).');
            copyNoduleTable('HPA');
            ev.preventDefault();
            ev.stopPropagation();
        } else if (ev.altKey && !ev.ctrlKey && ev.shiftKey && isKeyC) {
            console.info('[DeepRad helpers] Alt+Shift+C detected (Health Check).');
            copyNoduleTable('healthCheck');
            ev.preventDefault();
            ev.stopPropagation();
        }
    }, true);

    if (typeof GM_registerMenuCommand === 'function') {
        GM_registerMenuCommand('Copy HPA nodules (Ctrl+Shift+C)', () => copyNoduleTable('HPA'));
        GM_registerMenuCommand('Copy Health check nodules (Alt+Shift+C)', () => copyNoduleTable('healthCheck'));
        GM_registerMenuCommand('Refresh token now', refreshToken);
    }

    rememberToken(localStorage.getItem('token'));
    rememberCurrentPath();
    patchPageAuthHeaders();
    patchFrontendLogoutTimer();
    startSessionExpiredNotificationObserver();

    if (location.pathname === '/login' && restoreTokenIfNeeded()) {
        const lastPath = sessionStorage.getItem(LAST_PATH_KEY);
        if (lastPath) {
            location.replace(lastPath);
        }
    }

    setInterval(() => {
        document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }));
        document.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Shift', code: 'ShiftLeft' }));
        window.dispatchEvent(new Event('focus'));
        console.info('[DeepRad helpers] frontend keep alive event sent.');
    }, KEEP_ALIVE_INTERVAL_MS);

    setInterval(rememberCurrentPath, 30 * 1000);
    setTimeout(refreshToken, 30 * 1000);
    setInterval(refreshToken, TOKEN_REFRESH_INTERVAL_MS);
})();
