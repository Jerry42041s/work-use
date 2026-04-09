/** @odoo-module **/

/**
 * LazyLoader — 按需載入第三方函式庫。
 * 首次呼叫時插入 <script> 標籤，後續呼叫直接返回已快取的物件。
 */

const _cache = {};

async function _loadScript(url, globalName) {
    if (_cache[globalName]) return _cache[globalName];
    if (window[globalName]) {
        _cache[globalName] = window[globalName];
        return _cache[globalName];
    }

    return new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.src = url;
        script.onload = () => {
            _cache[globalName] = window[globalName];
            resolve(_cache[globalName]);
        };
        script.onerror = () => reject(new Error(`Failed to load ${url}`));
        document.head.appendChild(script);
    });
}

/**
 * 載入 mammoth.js（DOCX → HTML）
 * 需要先將 mammoth.browser.min.js 放在 static/lib/
 */
export async function getMammoth() {
    try {
        return await _loadScript(
            "/dobtor_doc_editor/static/lib/mammoth.browser.min.js",
            "mammoth"
        );
    } catch (e) {
        throw new Error("mammoth.js 未找到。請確認 static/lib/mammoth.browser.min.js 存在。");
    }
}

/**
 * 載入 HTMLtoDOCX（HTML → DOCX 草稿級匯出）
 * 需要先將 html-to-docx.js 放在 static/lib/
 */
export async function getHTMLtoDOCX() {
    try {
        return await _loadScript(
            "/dobtor_doc_editor/static/lib/html-to-docx.js",
            "HTMLtoDOCX"
        );
    } catch (e) {
        throw new Error("html-to-docx 未找到。請確認 static/lib/html-to-docx.js 存在。");
    }
}
