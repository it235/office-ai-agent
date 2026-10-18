/**
 * utils.js - Utility functions for OfficeAI Chat
 * Common helper functions used across the application
 */

// Generate UUID v4
function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        const r = Math.random() * 16 | 0;
        const v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

// Format date time to string
function formatDateTime(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');

    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// Escape HTML special characters
function escapeHtml(unsafe) {
    if (typeof unsafe !== 'string') return '';
    return unsafe
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// Convert markdown to plain text
function markdownToPlain(md) {
    try {
        const html = sanitizeHtml(marked.parse(md || ""));
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        return tmp.innerText || tmp.textContent || "";
    } catch (err) {
        console.error('markdownToPlain error', err);
        return md;
    }
}

// ========== HTML 清洗 ==========
//
// 模型输出可以包含任意原始 HTML，marked 默认原样透传。
// 因此所有 markdown -> innerHTML 都必须经过 renderMarkdown，
// 由 sanitizeHtml 按白名单剔除脚本、事件处理器和危险 URL。

const HTML_ALLOWED_TAGS = new Set([
    'a', 'abbr', 'b', 'blockquote', 'br', 'button', 'circle', 'code', 'col',
    'colgroup', 'dd', 'del', 'details', 'div', 'dl', 'dt', 'em', 'figcaption',
    'figure', 'g', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr', 'i', 'img', 'kbd',
    'li', 'line', 'mark', 'ol', 'p', 'path', 'polygon', 'polyline', 'pre',
    'rect', 's', 'samp', 'section', 'small', 'span', 'strong', 'sub', 'summary',
    'sup', 'svg', 'table', 'tbody', 'td', 'tfoot', 'th', 'thead', 'tr', 'u',
    'ul', 'var', 'wbr'
]);

// 连同内容一起丢弃：这些标签的文本本身也不可信或无意义
const HTML_FORBIDDEN_TAGS = new Set([
    'script', 'style', 'iframe', 'frame', 'frameset', 'object', 'embed', 'form',
    'input', 'textarea', 'select', 'option', 'link', 'meta', 'base', 'template',
    'noscript', 'math', 'portal'
]);

const HTML_ALLOWED_ATTRS = new Set([
    'href', 'src', 'alt', 'title', 'class', 'colspan', 'rowspan', 'start',
    'align', 'width', 'height', 'open', 'viewbox', 'xmlns', 'fill', 'stroke',
    'stroke-width', 'stroke-linecap', 'stroke-linejoin', 'd', 'points', 'cx',
    'cy', 'r', 'rx', 'ry', 'x', 'y', 'x1', 'x2', 'y1', 'y2'
]);

const HTML_URL_ATTRS = new Set(['href', 'src']);
const HTML_SAFE_URL_SCHEMES = ['http', 'https', 'mailto', 'tel'];

function isSafeUrl(value) {
    const v = String(value).replace(/[\s\u0000-\u001f\u007f]/g, '').toLowerCase();
    if (!v) return false;
    if (v.charAt(0) === '#' || v.indexOf('//') === 0) return true;
    const schemeEnd = v.indexOf(':');
    const pathEnd = v.search(/[/?#]/);
    if (schemeEnd < 0 || (pathEnd >= 0 && pathEnd < schemeEnd)) return true;
    return HTML_SAFE_URL_SCHEMES.some((scheme) => v.indexOf(scheme + ':') === 0);
}

function sanitizeScrubChildren(parent) {
    Array.from(parent.childNodes).forEach((node) => {
        if (node.nodeType === Node.COMMENT_NODE) {
            node.remove();
        } else if (node.nodeType === Node.ELEMENT_NODE) {
            sanitizeScrubElement(node);
        }
    });
}

function sanitizeScrubElement(el) {
    const tag = el.tagName.toLowerCase();
    if (HTML_FORBIDDEN_TAGS.has(tag)) {
        el.remove();
        return;
    }
    sanitizeScrubChildren(el);
    if (!HTML_ALLOWED_TAGS.has(tag)) {
        // 未知标签只拆壳保留文本，避免整段内容消失
        const parent = el.parentNode;
        if (parent) {
            Array.from(el.childNodes).forEach((child) => parent.insertBefore(child, el));
            el.remove();
        }
        return;
    }
    Array.from(el.attributes).forEach((attr) => {
        const name = attr.name.toLowerCase();
        const isData = name.indexOf('data-') === 0 || name.indexOf('aria-') === 0;
        const isSafeAttr = HTML_ALLOWED_ATTRS.has(name) && (!HTML_URL_ATTRS.has(name) || isSafeUrl(attr.value));
        if (!isData && !isSafeAttr) el.removeAttribute(attr.name);
    });
}

// DOMParser 产出的文档是惰性的：脚本不执行、资源不加载，可以安全解析
function sanitizeHtml(html) {
    const raw = html === null || html === undefined ? '' : String(html);
    if (!raw) return '';
    const doc = new DOMParser().parseFromString(raw, 'text/html');
    sanitizeScrubChildren(doc.body);
    return doc.body.innerHTML;
}

// 渲染 markdown 的唯一入口：解析 -> 清洗 -> 写入，失败时退化为纯文本
function renderMarkdown(element, markdown) {
    if (!element) return;
    const text = markdown === null || markdown === undefined ? '' : String(markdown);
    if (typeof marked === 'undefined') {
        element.textContent = text;
        return;
    }
    try {
        element.innerHTML = sanitizeHtml(marked.parse(text));
    } catch (err) {
        console.error('[renderMarkdown] markdown解析失败，退化为纯文本:', err);
        element.textContent = text;
    }
}

// Scroll content to bottom if auto-scroll is enabled
window.contentScroll = function () {
    if (document.getElementById("settings-scroll-checked").checked) {
        window.scrollTo(0, document.body.scrollHeight);
    }
}
