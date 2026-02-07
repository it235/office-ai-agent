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

// Strip HTML tags from string
function stripHtml(html) {
    try {
        const tmp = document.createElement('div');
        tmp.innerHTML = html || '';
        return (tmp.textContent || tmp.innerText || '').trim();
    } catch (e) {
        return html || '';
    }
}

// Convert markdown to plain text
function markdownToPlain(md) {
    try {
        const html = marked.parse(md || "");
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        return tmp.innerText || tmp.textContent || "";
    } catch (err) {
        console.error('markdownToPlain error', err);
        return md;
    }
}

// Scroll content to bottom if auto-scroll is enabled
window.contentScroll = function () {
    if (document.getElementById("settings-scroll-checked").checked) {
        window.scrollTo(0, document.body.scrollHeight);
    }
}
