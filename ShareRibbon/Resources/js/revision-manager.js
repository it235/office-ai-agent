/**
 * revision-manager.js - Revision Display and Application
 * Handles document formatting preview and revision suggestions
 */

(function () {
    if (!window._oa) window._oa = {};
    window._oa._comparisonCache = window._oa._comparisonCache || {};

    // Format preview (simplified: using paraIndex and formatting object)
    window.showComparison = function (uuid, originalText, aiPreviewOrPlan) {
        console.log(aiPreviewOrPlan);
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

            // Parse JSON response
            let raw = aiPreviewOrPlan;
            if (typeof raw === 'string') {
                raw = raw.trim();
                // Remove code block markers
                const fenceMatch = raw.match(/^\s*```([^\n]*)\n([\s\S]*?)\n```\s*$/);
                if (fenceMatch && fenceMatch.length >= 3) {
                    raw = fenceMatch[2];
                } else if (raw.startsWith('```')) {
                    raw = raw.replace(/^\s*```[^\n]*\n?/, '');
                    raw = raw.replace(/\n?```\s*$/, '');
                }
            }

            let planArr = [];
            try {
                if (typeof raw === 'string' && raw.length > 0) {
                    planArr = JSON.parse(raw);
                } else if (Array.isArray(raw)) {
                    planArr = raw;
                }
            } catch (e) {
                planArr = [];
            }

            if (!Array.isArray(planArr) || planArr.length === 0) {
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

            // Create preview container
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
