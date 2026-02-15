/**
 * markdown-renderer.js - Markdown Stream Renderer
 * Handles incremental markdown rendering for streaming responses
 * Includes ReformatTemplate card rendering for AI-generated templates
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
        
        // 收到第一个内容时隐藏等待动画
        if (this.fullContent === '' && text) {
            if (typeof hideLoadingIndicator === 'function') {
                hideLoadingIndicator();
            }
        }
        
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
            
            // Process template JSON code blocks and render as cards
            this._processTemplateCards();
        } catch (e) {
            console.error('[MarkdownStreamRenderer.append] marked.parse出错:', e);
            this.output.textContent = this.fullContent;
        }

        // Apply code highlighting
        this.output.querySelectorAll('pre code').forEach((block) => {
            hljs.highlightElement(block);
        });
    }

    /**
     * 检测并处理模板JSON代码块，将其渲染为交互式卡片
     */
    _processTemplateCards() {
        // 查找所有代码块
        const codeBlocks = this.output.querySelectorAll('pre code');
        
        codeBlocks.forEach((codeBlock, index) => {
            const content = codeBlock.textContent.trim();
            
            // 尝试解析JSON
            let template = null;
            try {
                template = JSON.parse(content);
            } catch (e) {
                // 不是有效的JSON，跳过
                return;
            }
            
            // 检测是否是ReformatTemplate结构
            if (!this._isReformatTemplate(template)) {
                return;
            }
            
            console.log('[MarkdownStreamRenderer] 检测到模板JSON，渲染为卡片:', template.Name || '未命名模板');
            
            // 创建卡片HTML
            const cardHtml = this._createTemplateCard(template, index);
            
            // 替换代码块为卡片
            const preElement = codeBlock.parentElement;
            const cardContainer = document.createElement('div');
            cardContainer.innerHTML = cardHtml;
            preElement.replaceWith(cardContainer.firstElementChild);
        });
    }

    /**
     * 检测对象是否是ReformatTemplate结构
     */
    _isReformatTemplate(obj) {
        if (!obj || typeof obj !== 'object') return false;
        
        // 检查必要字段：至少有 Layout 或 BodyStyles 或 PageSettings
        const hasLayout = obj.Layout && typeof obj.Layout === 'object';
        const hasBodyStyles = Array.isArray(obj.BodyStyles);
        const hasPageSettings = obj.PageSettings && typeof obj.PageSettings === 'object';
        
        // 还需要有 Name 字段
        const hasName = typeof obj.Name === 'string';
        
        return hasName && (hasLayout || hasBodyStyles || hasPageSettings);
    }

    /**
     * 创建模板卡片HTML
     */
    _createTemplateCard(template, cardIndex) {
        const templateJson = JSON.stringify(template);
        const escapedJson = templateJson.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        
        // 构建版式元素列表
        let layoutElementsHtml = '';
        if (template.Layout && template.Layout.Elements && template.Layout.Elements.length > 0) {
            const elements = template.Layout.Elements.slice(0, 5); // 最多显示5个
            layoutElementsHtml = elements.map(el => 
                `<span class="template-card-tag">${this._escapeHtml(el.Name || el.ElementType || '元素')}</span>`
            ).join('');
            if (template.Layout.Elements.length > 5) {
                layoutElementsHtml += `<span class="template-card-tag template-card-tag-more">+${template.Layout.Elements.length - 5}</span>`;
            }
        } else {
            layoutElementsHtml = '<span class="template-card-tag template-card-tag-empty">无版式元素</span>';
        }
        
        // 构建样式规则列表
        let bodyStylesHtml = '';
        if (template.BodyStyles && template.BodyStyles.length > 0) {
            const styles = template.BodyStyles.slice(0, 4); // 最多显示4个
            bodyStylesHtml = styles.map(style => 
                `<span class="template-card-tag">${this._escapeHtml(style.RuleName || '样式')}</span>`
            ).join('');
            if (template.BodyStyles.length > 4) {
                bodyStylesHtml += `<span class="template-card-tag template-card-tag-more">+${template.BodyStyles.length - 4}</span>`;
            }
        } else {
            bodyStylesHtml = '<span class="template-card-tag template-card-tag-empty">无样式规则</span>';
        }
        
        // 构建页面设置摘要
        let pageSettingsHtml = '';
        if (template.PageSettings) {
            const ps = template.PageSettings;
            const items = [];
            if (ps.Margins) {
                items.push(`边距: ${ps.Margins.Top || 2.54}/${ps.Margins.Bottom || 2.54}/${ps.Margins.Left || 3.18}/${ps.Margins.Right || 3.18}cm`);
            }
            if (ps.PageNumber && ps.PageNumber.Enabled) {
                items.push('页码: 开启');
            }
            pageSettingsHtml = items.length > 0 
                ? items.map(item => `<span class="template-card-info">${item}</span>`).join('')
                : '<span class="template-card-info">默认设置</span>';
        }

        return `
        <div class="template-card" data-card-index="${cardIndex}">
            <div class="template-card-header">
                <div class="template-card-icon">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                        <polyline points="14 2 14 8 20 8"></polyline>
                        <line x1="16" y1="13" x2="8" y2="13"></line>
                        <line x1="16" y1="17" x2="8" y2="17"></line>
                        <polyline points="10 9 9 9 8 9"></polyline>
                    </svg>
                </div>
                <div class="template-card-title-area">
                    <div class="template-card-title">${this._escapeHtml(template.Name || '未命名模板')}</div>
                    <div class="template-card-category">${this._escapeHtml(template.Category || '通用')} · ${this._escapeHtml(template.TargetApp || 'Word')}</div>
                </div>
            </div>
            ${template.Description ? `<div class="template-card-desc">${this._escapeHtml(template.Description)}</div>` : ''}
            <div class="template-card-content">
                <div class="template-card-section">
                    <div class="template-card-section-title">版式元素</div>
                    <div class="template-card-tags">${layoutElementsHtml}</div>
                </div>
                <div class="template-card-section">
                    <div class="template-card-section-title">正文样式</div>
                    <div class="template-card-tags">${bodyStylesHtml}</div>
                </div>
                ${pageSettingsHtml ? `
                <div class="template-card-section">
                    <div class="template-card-section-title">页面设置</div>
                    <div class="template-card-infos">${pageSettingsHtml}</div>
                </div>
                ` : ''}
            </div>
            <div class="template-card-actions">
                <button class="template-card-btn template-card-btn-primary" onclick="TemplateCardActions.save('${escapedJson}')">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path>
                        <polyline points="17 21 17 13 7 13 7 21"></polyline>
                        <polyline points="7 3 7 8 15 8"></polyline>
                    </svg>
                    保存模板
                </button>
                <button class="template-card-btn" onclick="TemplateCardActions.preview('${escapedJson}')">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                        <circle cx="12" cy="12" r="3"></circle>
                    </svg>
                    预览效果
                </button>
                <button class="template-card-btn" onclick="TemplateCardActions.toggleJson(this)">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="16 18 22 12 16 6"></polyline>
                        <polyline points="8 6 2 12 8 18"></polyline>
                    </svg>
                    查看JSON
                </button>
            </div>
            <div class="template-card-json-panel" style="display: none;">
                <pre><code class="language-json">${this._escapeHtml(JSON.stringify(template, null, 2))}</code></pre>
            </div>
        </div>`;
    }

    /**
     * HTML转义
     */
    _escapeHtml(text) {
        if (!text) return '';
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }
}

/**
 * 模板卡片操作对象 - 处理卡片按钮点击事件
 */
const TemplateCardActions = {
    /**
     * 保存模板
     */
    save(templateJson) {
        try {
            const template = JSON.parse(templateJson.replace(/&quot;/g, '"'));
            console.log('[TemplateCardActions.save] 保存模板:', template.Name);
            
            // 发送消息到VB端保存
            if (window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage({
                    type: 'saveAiTemplate',
                    templateJson: JSON.stringify(template)
                });
            } else {
                console.warn('[TemplateCardActions.save] WebView2不可用');
                alert('保存功能需要在Office插件中运行');
            }
        } catch (e) {
            console.error('[TemplateCardActions.save] 解析模板失败:', e);
            alert('模板数据解析失败');
        }
    },

    /**
     * 预览模板效果
     */
    preview(templateJson) {
        try {
            const template = JSON.parse(templateJson.replace(/&quot;/g, '"'));
            console.log('[TemplateCardActions.preview] 预览模板:', template.Name);
            
            // 发送消息到VB端预览
            if (window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage({
                    type: 'previewAiTemplate',
                    templateJson: JSON.stringify(template)
                });
            } else {
                console.warn('[TemplateCardActions.preview] WebView2不可用');
                alert('预览功能需要在Office插件中运行');
            }
        } catch (e) {
            console.error('[TemplateCardActions.preview] 解析模板失败:', e);
            alert('模板数据解析失败');
        }
    },

    /**
     * 切换JSON面板显示
     */
    toggleJson(btn) {
        const card = btn.closest('.template-card');
        const jsonPanel = card.querySelector('.template-card-json-panel');
        const isHidden = jsonPanel.style.display === 'none';
        
        jsonPanel.style.display = isHidden ? 'block' : 'none';
        btn.innerHTML = isHidden 
            ? `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="16 18 22 12 16 6"></polyline>
                <polyline points="8 6 2 12 8 18"></polyline>
               </svg>
               隐藏JSON`
            : `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="16 18 22 12 16 6"></polyline>
                <polyline points="8 6 2 12 8 18"></polyline>
               </svg>
               查看JSON`;
        
        // 应用代码高亮
        if (isHidden) {
            jsonPanel.querySelectorAll('pre code').forEach((block) => {
                if (typeof hljs !== 'undefined') {
                    hljs.highlightElement(block);
                }
            });
        }
    }
};
