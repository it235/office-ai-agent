/**
 * markdown-renderer.js - Markdown Stream Renderer
 * Handles incremental markdown rendering for streaming responses
 */

class MarkdownStreamRenderer {
    constructor(element) {
        this.output = element instanceof HTMLElement ? element : document.getElementById(element);
        this.fullContent = '';
    }

    append(text) {
        this.fullContent += text + '';

        // Use full content render
        this.output.innerHTML = marked.parse(this.fullContent);

        // Apply code highlighting
        this.output.querySelectorAll('pre code').forEach((block) => {
            hljs.highlightElement(block);
        });
    }
}
