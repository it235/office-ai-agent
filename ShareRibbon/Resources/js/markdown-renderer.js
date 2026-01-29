/**
 * markdown-renderer.js - Markdown Stream Renderer
 * Handles incremental markdown rendering for streaming responses
 */

class MarkdownStreamRenderer {
    constructor(element) {
        console.log('[MarkdownStreamRenderer] 构造函数, element=', element);
        this.output = element instanceof HTMLElement ? element : document.getElementById(element);
        this.fullContent = '';
        console.log('[MarkdownStreamRenderer] output元素:', this.output ? 'OK (id=' + (this.output.id || 'no-id') + ')' : 'NULL');
        
        // 调试：设置可见的背景色来确认元素存在
        if (this.output) {
            console.log('[MarkdownStreamRenderer] output元素位置:', this.output.getBoundingClientRect());
            console.log('[MarkdownStreamRenderer] output父元素:', this.output.parentElement ? this.output.parentElement.id : 'no-parent');
        }
    }

    append(text) {
        console.log('[MarkdownStreamRenderer.append] text长度=' + (text ? text.length : 0) + ', text内容="' + text + '"');
        this.fullContent += text + '';
        console.log('[MarkdownStreamRenderer.append] fullContent累计="' + this.fullContent + '"');

        if (!this.output) {
            console.error('[MarkdownStreamRenderer.append] output元素为null，无法渲染');
            return;
        }

        // 验证output元素是否仍在DOM中
        if (!document.body.contains(this.output)) {
            console.error('[MarkdownStreamRenderer.append] output元素不在DOM中!');
            return;
        }

        // Check if marked is available
        if (typeof marked === 'undefined') {
            console.error('[MarkdownStreamRenderer.append] marked库未加载');
            this.output.textContent = this.fullContent;
            return;
        }

        // Use full content render
        try {
            const parsed = marked.parse(this.fullContent);
            console.log('[MarkdownStreamRenderer.append] marked.parse结果="' + parsed + '"');
            this.output.innerHTML = parsed;
            console.log('[MarkdownStreamRenderer.append] innerHTML设置完成, 实际innerHTML="' + this.output.innerHTML + '"');
        } catch (e) {
            console.error('[MarkdownStreamRenderer.append] marked.parse出错:', e);
            this.output.textContent = this.fullContent;
        }

        // Apply code highlighting
        this.output.querySelectorAll('pre code').forEach((block) => {
            hljs.highlightElement(block);
        });
    }
}
