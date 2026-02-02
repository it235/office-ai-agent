/**
 * revision-manager.js - Revision Display and Application
 * Handles document formatting preview and revision suggestions
 */

(function () {
    if (!window._oa) window._oa = {};
    window._oa._comparisonCache = window._oa._comparisonCache || {};
    window._oa._reformatRetryCount = window._oa._reformatRetryCount || {};

    /**
     * 清理和修复常见的JSON格式错误
     */
    function cleanJsonString(str) {
        if (!str || typeof str !== 'string') return str;
        
        let cleaned = str.trim();
        
        // 移除代码块标记
        const fenceMatch = cleaned.match(/^\s*```([^\n]*)\n([\s\S]*?)\n```\s*$/);
        if (fenceMatch && fenceMatch.length >= 3) {
            cleaned = fenceMatch[2];
        } else if (cleaned.startsWith('```')) {
            cleaned = cleaned.replace(/^\s*```[^\n]*\n?/, '');
            cleaned = cleaned.replace(/\n?```\s*$/, '');
        }
        
        // 移除BOM和不可见字符
        cleaned = cleaned.replace(/^\uFEFF/, '');
        
        // 修复常见的JSON错误
        // 1. 中文引号替换为英文引号
        cleaned = cleaned.replace(/[""]/g, '"');
        cleaned = cleaned.replace(/['']/g, "'");
        
        // 2. 移除尾随逗号 (在 } 或 ] 之前的逗号)
        cleaned = cleaned.replace(/,\s*([\]}])/g, '$1');
        
        // 3. 修复未转义的换行符在字符串中
        // 先保护已转义的
        cleaned = cleaned.replace(/\\n/g, '___NEWLINE___');
        cleaned = cleaned.replace(/\\r/g, '___RETURN___');
        // 移除字符串内的实际换行
        cleaned = cleaned.replace(/"([^"]*)\n([^"]*)"/g, '"$1 $2"');
        // 恢复已转义的
        cleaned = cleaned.replace(/___NEWLINE___/g, '\\n');
        cleaned = cleaned.replace(/___RETURN___/g, '\\r');
        
        // 4. 尝试修复缺少引号的键名
        // 匹配 { 后或 , 后的非引号键名
        cleaned = cleaned.replace(/([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)\s*:/g, '$1"$2":');
        
        return cleaned.trim();
    }

    /**
     * 安全解析JSON，带有错误恢复
     */
    function safeParseJson(str) {
        if (!str) return null;
        
        // 第一次尝试：直接解析
        try {
            return JSON.parse(str);
        } catch (e1) {
            console.log('JSON parse attempt 1 failed:', e1.message);
        }
        
        // 第二次尝试：清理后解析
        const cleaned = cleanJsonString(str);
        try {
            return JSON.parse(cleaned);
        } catch (e2) {
            console.log('JSON parse attempt 2 failed:', e2.message);
        }
        
        // 第三次尝试：提取JSON部分
        try {
            // 查找第一个 { 或 [ 到最后一个 } 或 ]
            const startObj = cleaned.indexOf('{');
            const startArr = cleaned.indexOf('[');
            const start = startObj >= 0 && startArr >= 0 ? Math.min(startObj, startArr) : Math.max(startObj, startArr);
            
            if (start >= 0) {
                const isObject = cleaned[start] === '{';
                const endChar = isObject ? '}' : ']';
                const lastEnd = cleaned.lastIndexOf(endChar);
                
                if (lastEnd > start) {
                    const extracted = cleaned.substring(start, lastEnd + 1);
                    return JSON.parse(extracted);
                }
            }
        } catch (e3) {
            console.log('JSON parse attempt 3 failed:', e3.message);
        }
        
        return null;
    }

    // Format preview (支持新的rules格式和旧的array格式)
    window.showComparison = function (uuid, originalText, aiPreviewOrPlan) {
        console.log('[showComparison] uuid:', uuid);
        console.log('[showComparison] aiPreviewOrPlan:', aiPreviewOrPlan);
        
        try {
            const container = document.getElementById('content-' + uuid);
            if (!container) return;

            const footer = document.getElementById('footer-' + uuid);
            if (footer) {
                try {
                    const oldReject = footer.querySelector('.reject-btn');
                    if (oldReject) oldReject.style.display = 'none';
                    const tokenCount = footer.querySelector('.token-count');
                    if (tokenCount) tokenCount.style.display = 'none';
                    const codeButtons = footer.querySelector('.code-buttons');
                    if (codeButtons) codeButtons.style.display = 'none';
                } catch (e) { }
            }

            // 清理并解析JSON
            let raw = aiPreviewOrPlan;
            if (typeof raw === 'string') {
                raw = cleanJsonString(raw);
            }

            let parsed = null;
            let parseError = null;
            
            try {
                parsed = safeParseJson(raw);
            } catch (e) {
                parseError = e;
            }

            // 检测是否为新的rules格式
            if (parsed && parsed.rules && Array.isArray(parsed.rules)) {
                // 新格式：rules模式，直接发送给后端应用
                showRulesPreview(uuid, container, parsed);
                return;
            }

            // 旧格式：数组模式
            let planArr = [];
            if (Array.isArray(parsed)) {
                planArr = parsed;
            } else if (parsed && parsed.documentPlan) {
                planArr = parsed.documentPlan;
            }

            if (!Array.isArray(planArr) || planArr.length === 0) {
                // 解析失败，检查是否可以重试
                const retryCount = window._oa._reformatRetryCount[uuid] || 0;
                
                if (retryCount < 1 && parseError) {
                    // 第一次失败，尝试重试
                    window._oa._reformatRetryCount[uuid] = retryCount + 1;
                    
                    const wrapRetry = document.createElement('div');
                    wrapRetry.id = 'compare-' + uuid;
                    wrapRetry.className = 'reference-container';
                    wrapRetry.style.padding = '12px';
                    wrapRetry.style.marginTop = '12px';
                    wrapRetry.style.background = '#fff3cd';
                    wrapRetry.innerHTML = `
                        <div style="color:#856404;padding:8px;">
                            <strong>JSON解析失败</strong>，正在请求重试...<br>
                            <small>错误: ${parseError.message || '格式不符合规范'}</small>
                        </div>
                    `;
                    container.appendChild(wrapRetry);
                    
                    // 发送重试请求
                    sendMessageToServer({
                        type: 'retryReformat',
                        uuid: uuid,
                        error: parseError.message || '格式不符合规范'
                    });
                    return;
                }
                
                // 重试次数已用完或无错误信息
                const wrapEmpty = document.createElement('div');
                wrapEmpty.id = 'compare-' + uuid;
                wrapEmpty.className = 'reference-container';
                wrapEmpty.style.padding = '12px';
                wrapEmpty.style.marginTop = '12px';
                wrapEmpty.style.background = '#f6f8fa';
                wrapEmpty.innerHTML = '<div style="color:#666;padding:8px;">无需排版修改或大模型返回格式有问题。</div>';
                container.appendChild(wrapEmpty);
                return;
            }

            // 清除重试计数
            delete window._oa._reformatRetryCount[uuid];

            // Create preview container (旧格式的处理逻辑保持不变)
            const wrap = document.createElement('div');
            wrap.id = 'compare-' + uuid;
            wrap.className = 'reference-container';
            wrap.style.padding = '12px';
            wrap.style.marginTop = '12px';
            wrap.style.background = '#f6f8fa';

            // Header with apply all button
            const header = document.createElement('div');
            header.style.display = 'flex';
            header.style.justifyContent = 'space-between';
            header.style.alignItems = 'center';
            header.style.marginBottom = '10px';

            const title = document.createElement('div');
            title.style.fontWeight = '600';
            title.textContent = '排版预览';
            header.appendChild(title);

            const acceptAllBtn = document.createElement('button');
            acceptAllBtn.className = 'code-button';
            acceptAllBtn.style.backgroundColor = '#4CAF50';
            acceptAllBtn.textContent = '应用全部排版';
            acceptAllBtn.onclick = function () {
                wrap.querySelectorAll('.format-accept-btn:not([disabled])').forEach(b => b.click());
                acceptAllBtn.disabled = true;
                acceptAllBtn.textContent = '已全部应用';
            };
            header.appendChild(acceptAllBtn);
            wrap.appendChild(header);

            // List container
            const listWrap = document.createElement('div');
            listWrap.style.display = 'flex';
            listWrap.style.flexDirection = 'column';
            listWrap.style.gap = '8px';

            planArr.forEach((item, idx) => {
                const row = document.createElement('div');
                row.className = 'format-item';
                row.style.display = 'flex';
                row.style.justifyContent = 'space-between';
                row.style.alignItems = 'center';
                row.style.padding = '8px';
                row.style.background = '#fff';
                row.style.border = '1px solid #e6e6e6';
                row.style.borderRadius = '4px';

                const paraIdx = item.paraIndex != null ? item.paraIndex : idx;
                const previewText = item.previewText || '';
                const changes = item.changes || '';
                const formatting = item.formatting || {};

                const info = document.createElement('div');
                info.innerHTML = `<strong>[段落${paraIdx}]</strong> ${previewText.substring(0, 50)}${previewText.length > 50 ? '...' : ''}<br><em style="color:#666;font-size:12px;">${changes}</em>`;

                const acceptBtn = document.createElement('button');
                acceptBtn.className = 'format-accept-btn code-button';
                acceptBtn.textContent = '应用';
                acceptBtn.onclick = function () {
                    acceptBtn.disabled = true;
                    acceptBtn.textContent = '已应用';
                    row.style.opacity = '0.6';
                    // Send format apply request
                    sendMessageToServer({
                        type: 'applyDocumentPlanItem',
                        uuid: uuid,
                        paraIndex: paraIdx,
                        formatting: formatting
                    });
                };

                row.appendChild(info);
                row.appendChild(acceptBtn);
                listWrap.appendChild(row);
            });

            wrap.appendChild(listWrap);
            container.appendChild(wrap);

        } catch (err) {
            console.error('showComparison error', err);
        }
    };

    /**
     * 显示新格式的rules规则预览
     */
    function showRulesPreview(uuid, container, rulesData) {
        const rules = rulesData.rules || [];
        const summary = rulesData.summary || '';
        const sampleClassification = rulesData.sampleClassification || [];

        // 清除重试计数
        delete window._oa._reformatRetryCount[uuid];

        const wrap = document.createElement('div');
        wrap.id = 'compare-' + uuid;
        wrap.className = 'reference-container';
        wrap.style.padding = '12px';
        wrap.style.marginTop = '12px';
        wrap.style.background = '#f6f8fa';

        // Header
        const header = document.createElement('div');
        header.style.display = 'flex';
        header.style.justifyContent = 'space-between';
        header.style.alignItems = 'center';
        header.style.marginBottom = '10px';

        const title = document.createElement('div');
        title.style.fontWeight = '600';
        title.textContent = '排版规则预览';
        header.appendChild(title);

        const applyBtn = document.createElement('button');
        applyBtn.className = 'code-button';
        applyBtn.style.backgroundColor = '#4CAF50';
        applyBtn.textContent = '应用排版规则';
        applyBtn.onclick = function () {
            applyBtn.disabled = true;
            applyBtn.textContent = '正在应用...';
            // 发送整个rules对象给后端应用
            sendMessageToServer({
                type: 'applyDocumentPlanItem',
                uuid: uuid,
                rules: rules,
                sampleClassification: sampleClassification
            });
            setTimeout(() => {
                applyBtn.textContent = '已应用';
            }, 500);
        };
        header.appendChild(applyBtn);
        wrap.appendChild(header);

        // Summary
        if (summary) {
            const summaryDiv = document.createElement('div');
            summaryDiv.style.padding = '8px';
            summaryDiv.style.marginBottom = '10px';
            summaryDiv.style.background = '#e8f5e9';
            summaryDiv.style.borderRadius = '4px';
            summaryDiv.style.fontSize = '13px';
            summaryDiv.innerHTML = `<strong>排版策略：</strong>${summary}`;
            wrap.appendChild(summaryDiv);
        }

        // Rules list
        const listWrap = document.createElement('div');
        listWrap.style.display = 'flex';
        listWrap.style.flexDirection = 'column';
        listWrap.style.gap = '8px';

        rules.forEach((rule, idx) => {
            const row = document.createElement('div');
            row.style.padding = '8px';
            row.style.background = '#fff';
            row.style.border = '1px solid #e6e6e6';
            row.style.borderRadius = '4px';

            const ruleType = rule.type || `规则${idx + 1}`;
            const matchCondition = rule.matchCondition || '';
            const formatting = rule.formatting || {};

            // 格式化formatting为可读文本
            const formatParts = [];
            if (formatting.fontNameCN) formatParts.push(`中文字体: ${formatting.fontNameCN}`);
            if (formatting.fontNameEN) formatParts.push(`英文字体: ${formatting.fontNameEN}`);
            if (formatting.fontSize) formatParts.push(`字号: ${formatting.fontSize}pt`);
            if (formatting.bold) formatParts.push('加粗');
            if (formatting.alignment) formatParts.push(`对齐: ${formatting.alignment}`);
            if (formatting.firstLineIndent) formatParts.push(`首行缩进: ${formatting.firstLineIndent}字符`);
            if (formatting.lineSpacing) formatParts.push(`行距: ${formatting.lineSpacing}倍`);

            row.innerHTML = `
                <div style="font-weight:600;color:#1976d2;margin-bottom:4px;">${ruleType}</div>
                <div style="font-size:12px;color:#666;margin-bottom:4px;">匹配条件: ${matchCondition}</div>
                <div style="font-size:12px;color:#333;">${formatParts.join(' | ')}</div>
            `;

            listWrap.appendChild(row);
        });

        wrap.appendChild(listWrap);
        container.appendChild(wrap);
    }

    // Revision suggestions list (simplified: using paraIndex for positioning)
    window.showRevisions = function (responseUuid, revisions) {
        try {
            let revs = [];
            if (!revisions) revs = [];
            else if (typeof revisions === 'string') {
                try { revs = JSON.parse(revisions); } catch (e) { revs = []; }
            } else revs = revisions;

            const footer = document.getElementById('footer-' + responseUuid);
            if (!footer) return;

            // Hide footer buttons
            try {
                const oldReject = footer.querySelector('.reject-btn');
                if (oldReject) oldReject.style.display = 'none';
                const tokenCount = footer.querySelector('.token-count');
                if (tokenCount) tokenCount.style.display = 'none';
                const codeButtons = footer.querySelector('.code-buttons');
                if (codeButtons) codeButtons.style.display = 'none';
            } catch (e) { }

            // Remove existing container
            const exist = document.getElementById('revisions-' + responseUuid);
            if (exist) exist.remove();

            const footerWrapper = document.createElement('div');
            footerWrapper.id = 'revisions-' + responseUuid;
            footerWrapper.className = 'revisions-footer-wrapper';

            // Top controls
            const controls = document.createElement('div');
            controls.className = 'revisions-controls';

            const btnAcceptAll = document.createElement('button');
            btnAcceptAll.className = 'code-button';
            btnAcceptAll.style.backgroundColor = '#4CAF50';
            btnAcceptAll.textContent = '接受全部修改';
            btnAcceptAll.onclick = function () {
                footerWrapper.querySelectorAll('.rev-accept-btn:not([disabled])').forEach(b => b.click());
            };
            controls.appendChild(btnAcceptAll);
            footerWrapper.appendChild(controls);

            // List container
            const listInner = document.createElement('div');
            listInner.className = 'revisions-list-inner';

            revs.forEach((item, idx) => {
                const row = document.createElement('div');
                row.className = 'revision-item';
                row.id = `rev-${responseUuid}-${idx}`;

                // Display revision content
                const summary = document.createElement('div');
                summary.className = 'rev-summary';
                const paraIdx = item.paraIndex != null ? item.paraIndex : idx;
                const original = item.original || '';
                const corrected = item.corrected || '';
                const reason = item.reason || '';
                summary.innerHTML = `<strong>[段落${paraIdx}]</strong> "${original}" → "${corrected}"` + (reason ? ` <em>(${reason})</em>` : '');

                // Accept button
                const accept = document.createElement('button');
                accept.className = 'rev-accept-btn';
                accept.textContent = '接受';
                accept.setAttribute('data-idx', idx);
                accept.onclick = function () {
                    accept.disabled = true;
                    // Send simplified payload
                    const payload = {
                        type: 'applyRevisionSegment',
                        uuid: responseUuid,
                        paraIndex: paraIdx,
                        original: original,
                        corrected: corrected
                    };
                    sendMessageToServer(payload);
                    row.classList.add('rev-accepted');
                };

                row.appendChild(summary);
                row.appendChild(accept);
                listInner.appendChild(row);
            });

            footerWrapper.appendChild(listInner);
            footer.insertBefore(footerWrapper, footer.firstChild);

        } catch (err) {
            console.error('showRevisions error', err);
        }
    };

    // Fallback: trigger revision accept from frontend
    window.applyRevisionAccept = function (responseUuid, globalIndex) {
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage({ type: 'applyRevisionAccept', responseUuid: responseUuid, globalIndex: globalIndex });
        } else if (window.vsto) {
            window.vsto.sendMessage({ type: 'applyRevisionAccept', responseUuid: responseUuid, globalIndex: globalIndex });
        }
    };

    // Mark revision as handled (update UI)
    window.markRevisionHandled = function (responseUuid, globalIndex, status) {
        try {
            const container = document.getElementById('revisions-' + responseUuid);
            if (!container) return;
            const item = Array.from(container.querySelectorAll('div')).find(d => d.innerText && d.innerText.startsWith('#' + globalIndex + ' '));
            if (item) {
                item.style.opacity = '0.5';
                const badge = document.createElement('span');
                badge.className = 'token-count';
                badge.style.marginLeft = '8px';
                badge.textContent = status === 'accepted' ? '已接受' : '已拒绝';
                item.appendChild(badge);
                // Disable buttons
                item.querySelectorAll('button').forEach(b => b.disabled = true);
            }
        } catch (err) {
            console.error('markRevisionHandled error', err);
        }
    };

    // Expose for debugging
    window._oa.showComparison = window.showComparison;
    window._oa.showRevisions = window.showRevisions;
})();
