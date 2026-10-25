/**
 * reformat-template.js - 排版模板选择模块
 * 处理模板列表显示、预览、选择和管理
 */

// 当前模板列表
let currentTemplates = [];
// 当前规范列表
let currentStyleGuides = [];
// 当前选中的模板ID（用于预览后使用）
let selectedTemplateId = null;
// 当前选中的规范ID（用于预览后使用）
let selectedStyleGuideId = null;
// 当前筛选的分类
let currentCategory = '全部';
// 是否处于管理模式
let isManageMode = false;
// 全局状态：是否处于排版模板选择模式（用于防止意外退出）
window.reformatTemplateActive = false;
// 当前资源类型 (template | styleguide | all)
let currentResourceType = 'template';

/**
 * 进入模板选择模式
 */
window.enterReformatTemplateMode = function() {
    // 设置全局状态
    window.reformatTemplateActive = true;
    
    // 隐藏聊天容器
    const chatContainer = document.getElementById('chat-container');
    if (chatContainer) {
        chatContainer.style.display = 'none';
    }
    
    // 隐藏底部输入栏
    const bottomBar = document.getElementById('chat-bottom-bar');
    if (bottomBar) {
        bottomBar.style.display = 'none';
    }
    
    // 显示模板模式容器
    const templateMode = document.getElementById('reformat-template-mode');
    if (templateMode) {
        templateMode.style.display = 'flex';
    }

    // WebView2 页面会复用 DOM；进入排版页时先清掉可能残留的旧智能排版面板。
    const wrapper = document.getElementById('template-cards-wrapper');
    if (wrapper) {
        wrapper.innerHTML = '<div class="template-empty-hint">正在加载排版模板...</div>';
    }

    // 重置状态
    isManageMode = false;
    currentResourceType = 'template';
    updateManageModeUI();
    updateResourceTabUI();
    
    // 请求规范列表（模板列表由VB端自动发送）
    sendMessageToVB({ type: 'getStyleGuides' });
    
    };

/**
 * 面板工具条上的模板库入口。E5-04 之前只有 PowerPoint 的「排版」按钮会让 VB 注入
 * enterReformatTemplateMode()，而那个按钮在 PowerPoint 自己的 Ribbon 设计器里是
 * Visible=False（Word 的「排版」走智能排版分析、Excel 的「排版」直接被隐藏），
 * 也就是说三端都没有一条能走到模板库的可见入口——编辑器与列表代码却一直存在。
 * 这里复用同一段前端逻辑，并按 PPT 路径补发一次列表请求，三端各一次点击即可到达。
 */
window.openTemplateLibrary = function() {
    window.enterReformatTemplateMode();
    sendMessageToVB({ type: 'getReformatTemplates' });
};

/**
 * 退出模板选择模式
 * @param {boolean} force - 是否强制退出（默认false）
 */
window.exitReformatTemplateMode = function(force = false) {
    // 如果不是强制退出，检查是否真的处于模板模式
    if (!force && !window.reformatTemplateActive) {
        return;
    }
    
    // 清除全局状态
    window.reformatTemplateActive = false;
    
    // 隐藏模板模式容器
    const templateMode = document.getElementById('reformat-template-mode');
    if (templateMode) {
        templateMode.style.display = 'none';
    }
    
    // 显示聊天容器
    const chatContainer = document.getElementById('chat-container');
    if (chatContainer) {
        chatContainer.style.display = 'block';
    }
    
    // 显示底部输入栏
    const bottomBar = document.getElementById('chat-bottom-bar');
    if (bottomBar) {
        bottomBar.style.display = 'flex';
    }
    
    // 关闭预览对话框
    closeTemplatePreview();
    
    };

/**
 * 加载模板列表（由VB.NET调用）
 * @param {Array} templates - 模板数组
 */
window.loadReformatTemplateList = function(templates) {
    currentTemplates = templates || [];
    renderResourceList();
    };

/**
 * 加载规范列表（由VB.NET调用）
 * @param {Array} guides - 规范数组
 */
window.loadStyleGuideList = function(guides) {
    currentStyleGuides = guides || [];
    renderResourceList();
    };

/**
 * Tab切换
 * @param {string} tabType - 资源类型 (template | styleguide | all)
 */
window.switchResourceTab = function(tabType) {
    currentResourceType = tabType;
    updateResourceTabUI();

    // 按需请求数据：如果目标类型数据为空，向VB端请求
    if (tabType === 'styleguide' && currentStyleGuides.length === 0) {
        sendMessageToVB({ type: 'getStyleGuides' });
    }
    if (tabType === 'template' && currentTemplates.length === 0) {
        sendMessageToVB({ type: 'getReformatTemplates' });
    }

    renderResourceList();
    };

/**
 * 更新资源Tab UI
 */
function updateResourceTabUI() {
    // 更新Tab按钮样式
    document.querySelectorAll('.resource-tab').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.type === currentResourceType);
    });
    
    // 更新标题
    const titleEl = document.getElementById('resource-mode-title');
    if (titleEl) {
        switch(currentResourceType) {
            case 'template':
                titleEl.textContent = '选择排版模板';
                break;
            case 'styleguide':
                titleEl.textContent = '选择排版规范';
                break;
            default:
                titleEl.textContent = '选择排版资源';
        }
    }
    
    // 更新按钮显示
    const btnNewTemplate = document.getElementById('btn-new-template');

    if (btnNewTemplate) {
        if (currentResourceType === 'styleguide') {
            btnNewTemplate.style.display = 'none';
        } else {
            btnNewTemplate.style.display = 'inline-flex';
        }
    }

    // 更新保存/导入按钮的title提示
    const btnSave = document.getElementById('btn-save-resource');
    const btnImport = document.getElementById('btn-import-resource');
    const btnManage = document.getElementById('manage-templates-btn');

    if (btnSave) btnSave.style.display = '';
    if (btnImport) btnImport.style.display = '';
    if (btnManage) btnManage.style.display = '';
    if (currentResourceType === 'styleguide') {
        if (btnSave) btnSave.title = '保存左侧内容为排版规范';
        if (btnImport) btnImport.title = '从文件导入排版规范';
    } else {
        if (btnSave) btnSave.title = '保存左侧文档内容为排版模板';
        if (btnImport) btnImport.title = '从文件导入排版模板';
    }
}

/**
 * 渲染资源列表（根据当前Tab类型）
 */
function renderResourceList() {
    if (currentResourceType === 'template') {
        renderTemplateCards(currentTemplates, currentCategory);
    } else if (currentResourceType === 'styleguide') {
        renderStyleGuideCards(currentStyleGuides);
    } else {
        renderMixedResources(currentTemplates, currentStyleGuides);
    }
}

/**
 * 渲染混合资源列表
 */
function renderMixedResources(templates, guides) {
    const wrapper = document.getElementById('template-cards-wrapper');
    if (!wrapper) return;
    
    wrapper.innerHTML = '';
    
    // 模板区域
    if (templates.length > 0) {
        const templateSection = document.createElement('div');
        templateSection.className = 'template-section';
        templateSection.innerHTML = `
            <div class="template-section-header">
                <span class="template-section-title">排版模板</span>
                <span class="template-section-count">${templates.length}个</span>
            </div>
            <div class="template-section-cards"></div>
        `;
        const cardsContainer = templateSection.querySelector('.template-section-cards');
        templates.forEach(template => {
            cardsContainer.appendChild(createTemplateCard(template));
        });
        wrapper.appendChild(templateSection);
    }
    
    // 规范区域
    if (guides.length > 0) {
        const guideSection = document.createElement('div');
        guideSection.className = 'template-section styleguide-section';
        guideSection.innerHTML = `
            <div class="template-section-header">
                <span class="template-section-title">排版规范</span>
                <span class="template-section-count">${guides.length}个</span>
            </div>
            <div class="template-section-cards"></div>
        `;
        const cardsContainer = guideSection.querySelector('.template-section-cards');
        guides.forEach(guide => {
            cardsContainer.appendChild(createStyleGuideCard(guide));
        });
        wrapper.appendChild(guideSection);
    }
    
    if (templates.length === 0 && guides.length === 0) {
        wrapper.innerHTML = '<div class="template-empty-hint">暂无资源</div>';
    }
}

/**
 * 渲染规范卡片
 * @param {Array} guides - 规范数组
 */
function renderStyleGuideCards(guides) {
    const wrapper = document.getElementById('template-cards-wrapper');
    if (!wrapper) return;
    
    wrapper.innerHTML = '';
    
    if (guides.length === 0) {
        wrapper.innerHTML = '<div class="template-empty-hint">暂无排版规范，点击"上传规范"添加</div>';
        return;
    }
    
    // 分离预置规范和自定义规范
    const presetGuides = guides.filter(g => g.IsPreset);
    const customGuides = guides.filter(g => !g.IsPreset);
    
    // 自定义规范按创建时间降序排序（新增的排在最上面）
    customGuides.sort((a, b) => {
        const timeA = a.CreatedAt ? new Date(a.CreatedAt).getTime() : 0;
        const timeB = b.CreatedAt ? new Date(b.CreatedAt).getTime() : 0;
        return timeB - timeA;
    });
    
    // 自定义规范区域（放在前面，新增的排在最上面）
    if (customGuides.length > 0) {
        const customSection = document.createElement('div');
        customSection.className = 'template-section styleguide-section';
        customSection.innerHTML = `
            <div class="template-section-header">
                <span class="template-section-title">自定义规范</span>
                <span class="template-section-count">${customGuides.length}个</span>
            </div>
            <div class="template-section-cards"></div>
        `;
        const cardsContainer = customSection.querySelector('.template-section-cards');
        customGuides.forEach(guide => {
            cardsContainer.appendChild(createStyleGuideCard(guide));
        });
        wrapper.appendChild(customSection);
    }
    
    // 预置规范区域
    if (presetGuides.length > 0) {
        const presetSection = document.createElement('div');
        presetSection.className = 'template-section styleguide-section';
        presetSection.innerHTML = `
            <div class="template-section-header">
                <span class="template-section-title">系统规范</span>
                <span class="template-section-count">${presetGuides.length}个</span>
            </div>
            <div class="template-section-cards"></div>
        `;
        const cardsContainer = presetSection.querySelector('.template-section-cards');
        presetGuides.forEach(guide => {
            cardsContainer.appendChild(createStyleGuideCard(guide));
        });
        wrapper.appendChild(presetSection);
    }
}

/**
 * 创建规范卡片
 * @param {Object} guide - 规范对象
 * @returns {HTMLElement} 卡片元素
 */
function createStyleGuideCard(guide) {
    const card = document.createElement('div');
    card.className = 'template-card styleguide-card';
    card.dataset.id = guide.Id;
    
    // 管理模式按钮
    const manageBtnsHtml = isManageMode ? `
        <div class="template-manage-btns">
            <button class="template-manage-btn copy-btn" onclick="duplicateStyleGuide('${guide.Id}')" title="复制">
                <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>
            </button>
            <button class="template-manage-btn delete-btn" onclick="deleteStyleGuide('${guide.Id}')" title="删除" ${guide.IsPreset ? 'disabled' : ''}>
                <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
            </button>
            <button class="template-manage-btn export-btn" onclick="exportStyleGuide('${guide.Id}')" title="导出">
                <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
            </button>
        </div>
    ` : '';
    
    // 预置标签
    const presetBadge = guide.IsPreset ? '<span class="template-preset-badge styleguide-badge">预置</span>' : '';
    
    // 内容摘要
    const contentSummary = guide.ContentSummary || (guide.GuideContent ? guide.GuideContent.substring(0, 80) + '...' : '暂无内容');
    
    card.innerHTML = `
        <div class="template-card-content">
            <div class="template-header">
                <div class="template-name">
                    <svg class="styleguide-icon" viewBox="0 0 24 24" width="16" height="16"><path fill="currentColor" d="M18 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zM6 4h5v8l-2.5-1.5L6 12V4z"/></svg>
                    ${escapeHtml(guide.Name)}
                    ${presetBadge}
                </div>
            </div>
            <div class="template-description styleguide-summary">${escapeHtml(contentSummary)}</div>
            ${manageBtnsHtml}
        </div>
        <div class="template-actions">
            <button class="template-btn preview-btn" onclick="previewStyleGuide('${guide.Id}')">
                <svg viewBox="0 0 24 24" width="12" height="12"><path fill="currentColor" d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
                预览
            </button>
            <button class="template-btn use-btn" onclick="useStyleGuide('${guide.Id}')">
                <svg viewBox="0 0 24 24" width="12" height="12"><path fill="currentColor" d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
                使用
            </button>
        </div>
    `;
    
    return card;
}

/**
 * 预览规范（Markdown渲染）
 * @param {string} guideId - 规范ID
 */
window.previewStyleGuide = function(guideId) {
    const guide = currentStyleGuides.find(g => g.Id === guideId);
    if (!guide) {
        console.error('[ReformatTemplate] 规范不存在:', guideId);
        return;
    }
    
    selectedStyleGuideId = guideId;
    
    // 获取预览内容区域
    const previewContent = document.getElementById('styleguide-preview-content');
    const previewTitle = document.getElementById('styleguide-preview-title');
    
    if (!previewContent) return;
    
    // 更新标题
    if (previewTitle) {
        previewTitle.textContent = guide.Name;
    }
    
    // 使用marked.js渲染Markdown
    try {
        renderMarkdown(previewContent, guide.GuideContent || '暂无内容');
        
        // 应用代码高亮
        document.querySelectorAll('#styleguide-preview-content pre code').forEach((block) => {
            if (typeof hljs !== 'undefined') {
                hljs.highlightElement(block);
            }
        });
    } catch (e) {
        console.error('[ReformatTemplate] Markdown渲染错误:', e);
        previewContent.textContent = guide.GuideContent || '暂无内容';
    }
    
    // 显示预览对话框
    const dialog = document.getElementById('styleguide-preview-dialog');
    if (dialog) {
        dialog.style.display = 'flex';
    }
    
    // 根据是否预置来显示/隐藏编辑按钮
    const editBtn = document.getElementById('styleguide-edit-btn');
    if (editBtn) {
        editBtn.style.display = guide.IsPreset ? 'none' : '';
    }
    
    };

/**
 * 关闭规范预览（同时重置编辑模式）
 */
window.closeStyleGuidePreview = function() {
    // 如果在编辑模式，先退出
    const dialogInner = document.querySelector('.styleguide-preview-dialog');
    if (dialogInner && dialogInner.dataset.mode === 'edit') {
        resetStyleGuideEditUI();
    }
    const dialog = document.getElementById('styleguide-preview-dialog');
    if (dialog) {
        dialog.style.display = 'none';
    }
};

/**
 * 从卡片编辑按钮进入编辑（先打开预览再切换到编辑模式）
 * @param {string} guideId - 规范ID
 */
/**
 * 进入规范编辑模式（预览 → 编辑）
 */
window.enterStyleGuideEditMode = function() {
    const guide = currentStyleGuides.find(g => g.Id === selectedStyleGuideId);
    if (!guide) return;
    if (guide.IsPreset) {
        alert('预置规范不可编辑');
        return;
    }

    const dialogInner = document.querySelector('.styleguide-preview-dialog');
    const previewContent = document.getElementById('styleguide-preview-content');
    const editorContainer = document.getElementById('styleguide-editor-container');
    const editorInput = document.getElementById('styleguide-editor-input');
    const editorPreview = document.getElementById('styleguide-editor-preview');

    if (!dialogInner || !editorContainer || !editorInput) return;

    // 切换模式标记
    dialogInner.dataset.mode = 'edit';

    // 隐藏只读预览，显示编辑器
    if (previewContent) previewContent.style.display = 'none';
    editorContainer.style.display = 'flex';

    // 填充Markdown源码
    editorInput.value = guide.GuideContent || '';

    // 渲染右侧实时预览
    renderEditorPreview(editorInput.value, editorPreview);

    // 绑定实时预览（输入时更新右侧）
    editorInput.oninput = function() {
        renderEditorPreview(this.value, editorPreview);
    };

    // 切换按钮可见性
    toggleEditButtons(true);

    };

/**
 * 保存规范编辑
 */
window.saveStyleGuideEdit = function() {
    const guide = currentStyleGuides.find(g => g.Id === selectedStyleGuideId);
    if (!guide || guide.IsPreset) return;

    const editorInput = document.getElementById('styleguide-editor-input');
    if (!editorInput) return;

    const newContent = editorInput.value;

    // 发送更新消息到VB
    sendMessageToVB({
        type: 'updateStyleGuide',
        guideId: selectedStyleGuideId,
        guideContent: newContent
    });

    // 本地同步更新（不等VB回调，体验更流畅）
    guide.GuideContent = newContent;

    // 退出编辑模式并刷新预览
    resetStyleGuideEditUI();

    // 刷新预览区域为新内容
    const previewContent = document.getElementById('styleguide-preview-content');
    if (previewContent) {
        renderMarkdown(previewContent, newContent || '暂无内容');
    }

    };

/**
 * 取消编辑模式（回到预览）
 */
window.cancelStyleGuideEditMode = function() {
    resetStyleGuideEditUI();
    };

/**
 * 重置编辑模式UI到预览状态
 */
function resetStyleGuideEditUI() {
    const dialogInner = document.querySelector('.styleguide-preview-dialog');
    const previewContent = document.getElementById('styleguide-preview-content');
    const editorContainer = document.getElementById('styleguide-editor-container');
    const editorInput = document.getElementById('styleguide-editor-input');

    if (dialogInner) dialogInner.dataset.mode = 'view';
    if (previewContent) previewContent.style.display = '';
    if (editorContainer) editorContainer.style.display = 'none';
    if (editorInput) editorInput.oninput = null;

    toggleEditButtons(false);
}

/**
 * 切换编辑/预览模式按钮的可见性
 * @param {boolean} editing - 是否处于编辑状态
 */
function toggleEditButtons(editing) {
    const editBtn = document.getElementById('styleguide-edit-btn');
    const useBtn = document.getElementById('styleguide-use-btn');
    const saveBtn = document.getElementById('styleguide-save-btn');
    const cancelBtn = document.getElementById('styleguide-cancel-edit-btn');

    if (editBtn) editBtn.style.display = editing ? 'none' : '';
    if (useBtn) useBtn.style.display = editing ? 'none' : '';
    if (saveBtn) saveBtn.style.display = editing ? '' : 'none';
    if (cancelBtn) cancelBtn.style.display = editing ? '' : 'none';
}

/**
 * 渲染编辑器右侧的实时预览
 * @param {string} mdText - Markdown源码
 * @param {HTMLElement} previewEl - 预览容器
 */
function renderEditorPreview(mdText, previewEl) {
    if (!previewEl) return;
    renderMarkdown(previewEl, mdText || '');
    previewEl.querySelectorAll('pre code').forEach(block => {
        if (typeof hljs !== 'undefined') hljs.highlightElement(block);
    });
}

/**
 * 使用规范
 * @param {string} guideId - 规范ID
 */
window.useStyleGuide = function(guideId) {
    const guide = currentStyleGuides.find(g => g.Id === guideId);
    if (!guide) {
        console.error('[ReformatTemplate] 规范不存在:', guideId);
        return;
    }
    
    sendMessageToVB({
        type: 'useStyleGuide',
        guideId: guideId
    });
    
    };

/**
 * 从预览对话框使用规范
 */
window.useStyleGuideFromPreview = function() {
    if (selectedStyleGuideId) {
        closeStyleGuidePreview();
        useStyleGuide(selectedStyleGuideId);
    }
};

/**
 * 删除规范
 * @param {string} guideId - 规范ID
 */
window.deleteStyleGuide = function(guideId) {
    const guide = currentStyleGuides.find(g => g.Id === guideId);
    if (!guide) return;
    
    if (guide.IsPreset) {
        alert('预置规范不可删除');
        return;
    }
    
    if (!confirm(`确定要删除规范"${guide.Name}"吗？此操作不可恢复。`)) {
        return;
    }
    
    sendMessageToVB({
        type: 'deleteStyleGuide',
        guideId: guideId
    });
    
    };

/**
 * 复制规范
 * @param {string} guideId - 规范ID
 */
window.duplicateStyleGuide = function(guideId) {
    const guide = currentStyleGuides.find(g => g.Id === guideId);
    if (!guide) return;
    
    const newName = prompt('请输入新规范名称:', guide.Name + ' (副本)');
    if (newName === null) return;
    
    sendMessageToVB({
        type: 'duplicateStyleGuide',
        guideId: guideId,
        newName: newName
    });
    
    };

/**
 * 导出规范
 * @param {string} guideId - 规范ID
 */
window.exportStyleGuide = function(guideId) {
    sendMessageToVB({
        type: 'exportStyleGuide',
        guideId: guideId
    });
    };

/**
 * 渲染模板卡片 - 分系统模板和自定义模板两组显示
 * @param {Array} templates - 模板数组
 * @param {string} filterCategory - 筛选分类
 */
function renderTemplateCards(templates, filterCategory = '全部') {
    const wrapper = document.getElementById('template-cards-wrapper');
    if (!wrapper) return;
    
    wrapper.innerHTML = '';
    
    // 筛选模板
    const filtered = filterCategory === '全部' 
        ? templates 
        : templates.filter(t => t.Category === filterCategory);
    
    if (filtered.length === 0) {
        wrapper.innerHTML = '<div class="template-empty-hint">暂无模板</div>';
        return;
    }
    
    // 分离系统模板和自定义模板
    const presetTemplates = filtered.filter(t => t.IsPreset);
    const customTemplates = filtered.filter(t => !t.IsPreset);
    
    // 自定义模板按创建时间降序排序（新增的排在最上面）
    customTemplates.sort((a, b) => {
        const timeA = a.CreatedAt ? new Date(a.CreatedAt).getTime() : 0;
        const timeB = b.CreatedAt ? new Date(b.CreatedAt).getTime() : 0;
        return timeB - timeA;
    });
    
    // 自定义模板区域（放在前面，新增的排在最上面）
    if (customTemplates.length > 0) {
        const customSection = document.createElement('div');
        customSection.className = 'template-section';
        customSection.innerHTML = `
            <div class="template-section-header">
                <span class="template-section-title">自定义模板</span>
                <span class="template-section-count">${customTemplates.length}个</span>
            </div>
            <div class="template-section-cards"></div>
        `;
        const cardsContainer = customSection.querySelector('.template-section-cards');
        customTemplates.forEach(template => {
            cardsContainer.appendChild(createTemplateCard(template));
        });
        wrapper.appendChild(customSection);
    }
    
    // 系统模板区域
    if (presetTemplates.length > 0) {
        const presetSection = document.createElement('div');
        presetSection.className = 'template-section';
        presetSection.innerHTML = `
            <div class="template-section-header">
                <span class="template-section-title">系统模板</span>
                <span class="template-section-count">${presetTemplates.length}个</span>
            </div>
            <div class="template-section-cards"></div>
        `;
        const cardsContainer = presetSection.querySelector('.template-section-cards');
        presetTemplates.forEach(template => {
            cardsContainer.appendChild(createTemplateCard(template));
        });
        wrapper.appendChild(presetSection);
    } else if (customTemplates.length > 0) {
        // 如果只有自定义模板，显示系统模板为空的提示
        const emptyPreset = document.createElement('div');
        emptyPreset.className = 'template-section';
        emptyPreset.innerHTML = `
            <div class="template-section-header">
                <span class="template-section-title">系统模板</span>
                <span class="template-section-count">0个</span>
            </div>
            <div class="template-empty-hint small">暂无系统模板</div>
        `;
        wrapper.appendChild(emptyPreset);
    }
}

/**
 * 创建单个模板卡片
 * @param {Object} template - 模板对象
 * @returns {HTMLElement} 卡片元素
 */
function createTemplateCard(template) {
    const card = document.createElement('div');
    card.className = 'template-card';
    card.dataset.id = template.Id;
    
    // 管理模式按钮（docx映射不支持编辑/复制/导出，只保留删除）
    let manageBtnsHtml = '';
    if (isManageMode) {
        if (template.IsDocxMapping) {
            manageBtnsHtml = `
                <div class="template-manage-btns">
                    <button class="template-manage-btn delete-btn" onclick="deleteDocxMapping('${template.Id}')" title="删除">
                        <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                    </button>
                </div>
            `;
        } else {
            manageBtnsHtml = `
                <div class="template-manage-btns">
                    <button class="template-manage-btn edit-btn" onclick="editTemplate('${template.Id}')" title="编辑" ${template.IsPreset ? 'disabled' : ''}>
                        <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04a1 1 0 0 0 0-1.41l-2.34-2.34a1 1 0 0 0-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                    </button>
                    <button class="template-manage-btn copy-btn" onclick="duplicateTemplate('${template.Id}')" title="复制">
                        <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M16 1H4a2 2 0 0 0-2 2v14h2V3h12V1zm3 4H8a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2zm0 16H8V7h11v14z"/></svg>
                    </button>
                    <button class="template-manage-btn delete-btn" onclick="deleteTemplate('${template.Id}')" title="删除" ${template.IsPreset ? 'disabled' : ''}>
                        <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                    </button>
                    <button class="template-manage-btn export-btn" onclick="exportTemplate('${template.Id}')" title="导出">
                        <svg viewBox="0 0 24 24" width="14" height="14"><path fill="currentColor" d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
                    </button>
                </div>
            `;
        }
    }
    
    // 预置标签 / 文档提取标签
    let badge = '';
    if (template.IsDocxMapping) {
        badge = '<span class="template-preset-badge" style="background:#38a169;color:#fff;">文档</span>';
    } else if (template.IsPreset) {
        badge = '<span class="template-preset-badge">预置</span>';
    }
    
    card.innerHTML = `
        <div class="template-card-content">
            <div class="template-header">
                <div class="template-name">
                    ${escapeHtml(template.Name)}
                    ${badge}
                </div>
            </div>
            <div class="template-description">${escapeHtml(template.Description || '暂无描述')}</div>
            ${manageBtnsHtml}
        </div>
        <div class="template-actions">
            <button class="template-btn preview-btn" onclick="previewTemplate('${template.Id}')">
                <svg viewBox="0 0 24 24" width="12" height="12"><path fill="currentColor" d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
                预览
            </button>
            <button class="template-btn use-btn" onclick="useTemplate('${template.Id}')">
                <svg viewBox="0 0 24 24" width="12" height="12"><path fill="currentColor" d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
                使用
            </button>
        </div>
    `;
    
    return card;
}

/**
 * 预览模板
 * @param {string} templateId - 模板ID
 */
window.previewTemplate = function(templateId) {
    const template = currentTemplates.find(t => t.Id === templateId);
    if (!template) {
        console.error('[ReformatTemplate] 模板不存在:', templateId);
        return;
    }
    
    selectedTemplateId = templateId;
    
    // 构建预览内容
    const previewContent = document.getElementById('template-preview-content');
    if (!previewContent) return;
    
    // docx映射卡片使用语义标签预览
    if (template.IsDocxMapping && template.SemanticTags) {
        previewContent.innerHTML = buildDocxMappingPreviewHtml(template);
        const dialog = document.getElementById('template-preview-dialog');
        if (dialog) dialog.style.display = 'flex';
        return;
    }
    
    // 版式元素HTML
    let layoutElementsHtml = '';
    if (template.Layout && template.Layout.Elements) {
        layoutElementsHtml = template.Layout.Elements.map(el => `
            <div class="preview-element">
                <span class="element-name">${escapeHtml(el.Name)}</span>
                <span class="element-font">${escapeHtml(el.Font?.FontNameCN || '默认')} ${el.Font?.FontSize || 12}pt${el.Font?.Bold ? ' 加粗' : ''}</span>
                <span class="element-align">${getAlignmentText(el.Paragraph?.Alignment)}</span>
            </div>
        `).join('');
    }
    
    // 正文样式HTML
    let bodyStylesHtml = '';
    if (template.BodyStyles) {
        bodyStylesHtml = template.BodyStyles.map(style => `
            <div class="preview-style">
                <span class="style-name">${escapeHtml(style.RuleName)}</span>
                <span class="style-condition">${escapeHtml(style.MatchCondition || '默认')}</span>
                <span class="style-font">${escapeHtml(style.Font?.FontNameCN || '默认')} ${style.Font?.FontSize || 12}pt</span>
            </div>
        `).join('');
    }
    
    // 页面设置HTML
    let pageSettingsHtml = '';
    if (template.PageSettings) {
        const ps = template.PageSettings;
        pageSettingsHtml = `
            <div class="preview-page-settings">
                <div class="page-setting-item">
                    <span class="setting-label">页边距:</span>
                    <span class="setting-value">上${ps.Margins?.Top || 2.54}cm 下${ps.Margins?.Bottom || 2.54}cm 左${ps.Margins?.Left || 3.18}cm 右${ps.Margins?.Right || 3.18}cm</span>
                </div>
                <div class="page-setting-item">
                    <span class="setting-label">页码:</span>
                    <span class="setting-value">${ps.PageNumber?.Enabled ? ps.PageNumber.Format || '第{page}页' : '不显示'}</span>
                </div>
            </div>
        `;
    }
    
    previewContent.innerHTML = `
        <div class="preview-header">
            <h3 class="preview-title">${escapeHtml(template.Name)}</h3>
            <span class="preview-category">${escapeHtml(template.Category)}</span>
        </div>
        <p class="preview-description">${escapeHtml(template.Description || '暂无描述')}</p>
        
        <div class="preview-section">
            <h4>版式配置</h4>
            <div class="preview-elements-list">
                ${layoutElementsHtml || '<div class="preview-empty">未配置版式元素</div>'}
            </div>
        </div>
        
        <div class="preview-section">
            <h4>正文样式</h4>
            <div class="preview-styles-list">
                ${bodyStylesHtml || '<div class="preview-empty">未配置正文样式</div>'}
            </div>
        </div>
        
        <div class="preview-section">
            <h4>页面设置</h4>
            ${pageSettingsHtml || '<div class="preview-empty">使用默认页面设置</div>'}
        </div>
        
        ${template.AiGuidance ? `
        <div class="preview-section">
            <h4>AI说明</h4>
            <p class="preview-ai-guidance">${escapeHtml(template.AiGuidance)}</p>
        </div>
        ` : ''}
    `;
    
    // 显示预览对话框
    const dialog = document.getElementById('template-preview-dialog');
    if (dialog) {
        dialog.style.display = 'flex';
    }
    
    };

/**
 * 获取对齐方式文本
 * @param {string} alignment - 对齐方式
 * @returns {string} 中文描述
 */
function getAlignmentText(alignment) {
    const alignMap = {
        'left': '左对齐',
        'center': '居中',
        'right': '右对齐',
        'justify': '两端对齐'
    };
    return alignMap[alignment] || '左对齐';
}

/**
 * 关闭模板预览
 */
window.closeTemplatePreview = function() {
    const dialog = document.getElementById('template-preview-dialog');
    if (dialog) {
        dialog.style.display = 'none';
    }
};

/**
 * 在Word中预览模板
 */
window.previewTemplateInWord = function() {
    if (!selectedTemplateId) return;
    
    // 获取当前应用类型 - 从全局变量获取或尝试获取
    let currentAppName = 'Word'; // 默认假设在Word中
    
    // 尝试从全局变量获取应用名称
    if (typeof window.currentOfficeAppName !== 'undefined' && window.currentOfficeAppName) {
        currentAppName = window.currentOfficeAppName;
    } else {
        // 尝试通过消息发送获取应用信息
        try {
            // 发送消息获取应用信息
            sendMessageToVB({
                type: 'getCurrentAppInfo'
            });
            // 暂时假设为Word，直到收到响应
            currentAppName = 'Word';
        } catch(e) {
            currentAppName = 'Word';
        }
    }
    
    // 检查是否为Word应用
    const isWordApp = currentAppName.toLowerCase().indexOf('word') !== -1 || 
                      currentAppName.toLowerCase().indexOf('word') !== -1 ||
                      currentAppName.toLowerCase().indexOf('word') !== -1;
    
    if (isWordApp) {
        sendMessageToVB({
            type: 'previewTemplateInWord',
            templateId: selectedTemplateId
        });
        
        } else {
        // 如果不是Word，显示提示信息
        alert(`${currentAppName}不支持模板预览功能，此功能仅适用于Word应用。`);
        }
};

/**
 * 使用模板
 * @param {string} templateId - 模板ID
 */
window.useTemplate = function(templateId) {
    const template = currentTemplates.find(t => t.Id === templateId);
    if (!template) {
        console.error('[ReformatTemplate] 模板不存在:', templateId);
        return;
    }
    
    // 注意：不在此处退出模板模式
    // 由VB后端在成功处理后调用 exitReformatTemplateMode()
    
    // 发送模板给后端
    sendMessageToVB({
        type: 'useReformatTemplate',
        templateId: templateId,
        template: template
    });
    
    };

/**
 * 从预览对话框使用模板
 */
window.useTemplateFromPreview = function() {
    if (selectedTemplateId) {
        closeTemplatePreview();
        useTemplate(selectedTemplateId);
    }
};

/**
 * 切换分类筛选
 * @param {string} category - 分类名称
 */
window.filterTemplatesByCategory = function(category) {
    currentCategory = category;
    
    // 更新按钮状态
    document.querySelectorAll('.template-category-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.category === category);
    });
    
    // 重新渲染卡片
    renderTemplateCards(currentTemplates, category);
    
    };

/**
 * 切换管理模式
 */
window.toggleManageMode = function() {
    isManageMode = !isManageMode;
    updateManageModeUI();
    renderResourceList();
    
    };

/**
 * 更新管理模式UI
 */
function updateManageModeUI() {
    const manageBtn = document.getElementById('manage-templates-btn');
    if (manageBtn) {
        const label = currentResourceType === 'styleguide' ? '规范' : '模板';
        manageBtn.textContent = isManageMode ? '完成管理' : `管理${label}`;
        manageBtn.classList.toggle('active', isManageMode);
    }
}

/**
 * 编辑模板
 * 模板编辑器（Word 走任务窗格，其余宿主走模态窗体）一直在，但 E5-04 之前没有任何按钮发
 * openTemplateEditor，所以「改一个已有模板」只有开发者知道。预置模板保持只读，先复制再改。
 * @param {string} templateId - 模板ID
 */
window.editTemplate = function(templateId) {
    const template = currentTemplates.find(t => t.Id === templateId);
    if (!template) return;

    if (template.IsPreset) {
        alert('预置模板不可直接修改，请先复制再编辑');
        return;
    }

    sendMessageToVB({
        type: 'openTemplateEditor',
        templateId: templateId
    });
};

/**
 * 复制模板
 * @param {string} templateId - 模板ID
 */
window.duplicateTemplate = function(templateId) {
    const template = currentTemplates.find(t => t.Id === templateId);
    if (!template) return;
    
    const newName = prompt('请输入新模板名称:', template.Name + ' (副本)');
    if (newName === null) return; // 用户取消
    
    sendMessageToVB({
        type: 'duplicateTemplate',
        templateId: templateId,
        newName: newName
    });
    
    };

/**
 * 删除模板
 * @param {string} templateId - 模板ID
 */
window.deleteTemplate = function(templateId) {
    const template = currentTemplates.find(t => t.Id === templateId);
    if (!template) return;
    
    if (template.IsPreset) {
        alert('预置模板不可删除');
        return;
    }
    
    if (!confirm(`确定要删除模板"${template.Name}"吗？此操作不可恢复。`)) {
        return;
    }
    
    sendMessageToVB({
        type: 'deleteTemplate',
        templateId: templateId
    });
    
    };

/**
 * 删除docx映射卡片
 * @param {string} cardId - 卡片ID（格式为 docx_mappingId）
 */
window.deleteDocxMapping = function(cardId) {
    const template = currentTemplates.find(t => t.Id === cardId);
    if (!template) return;
    
    if (!confirm(`确定要删除文档映射"${template.Name}"吗？此操作不可恢复。`)) {
        return;
    }
    
    sendMessageToVB({
        type: 'deleteDocxMapping',
        mappingId: template.MappingId || cardId.replace('docx_', '')
    });
    
    };

/**
 * 导出模板
 * @param {string} templateId - 模板ID
 */
window.exportTemplate = function(templateId) {
    sendMessageToVB({
        type: 'exportTemplate',
        templateId: templateId
    });
    
    };

/**
 * 保存当前文档为模板
 */
window.saveCurrentDocumentAsTemplate = function() {
    sendMessageToVB({
        type: 'saveCurrentDocumentAsTemplate'
    });
    
    };

/**
 * 导入资源（根据当前Tab类型自动切换导入行为）
 */
window.importTemplate = function() {
    if (currentResourceType === 'styleguide') {
        sendMessageToVB({ type: 'uploadStyleGuideDocument' });
        } else {
        sendMessageToVB({ type: 'importTemplate' });
        }
};

/**
 * 创建新模板
 */
window.createNewTemplate = function() {
    sendMessageToVB({
        type: 'openTemplateEditor',
        templateId: '' // 空ID表示新建
    });
    
    };

/**
 * 使用AI助手创建模板
 * Plan A: 直接在聊天中与AI对话，AI返回的模板JSON会自动渲染为交互式卡片
 */
/**
 * 发送消息到VB.NET后端
 * @param {Object} payload - 消息负载
 */
function sendMessageToVB(payload) {
    try {
        if (window.officeAi && typeof window.officeAi.post === 'function') {
            window.officeAi.post(payload);
        } else {
            console.warn('[ReformatTemplate] no host bridge available');
        }
    } catch (e) {
        console.error('[ReformatTemplate] 发送消息失败:', e);
    }
}

/**
 * 折叠/展开模板列表区域
 */
// 初始化事件绑定
document.addEventListener('DOMContentLoaded', function() {
    // 退出按钮（用户主动点击，强制退出）
    const exitBtn = document.getElementById('exit-template-mode-btn');
    if (exitBtn) {
        exitBtn.addEventListener('click', function() {
            exitReformatTemplateMode(true);
        });
    }
    
    // 分类筛选按钮
    document.querySelectorAll('.template-category-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            filterTemplatesByCategory(this.dataset.category);
        });
    });
    
    // 预览对话框关闭按钮
    const closePreviewBtn = document.getElementById('close-template-preview-btn');
    if (closePreviewBtn) {
        closePreviewBtn.addEventListener('click', closeTemplatePreview);
    }
    
    // 从预览使用按钮
    const useFromPreviewBtn = document.getElementById('use-template-from-preview-btn');
    if (useFromPreviewBtn) {
        useFromPreviewBtn.addEventListener('click', useTemplateFromPreview);
    }
    
    // 在Word中预览按钮
    const previewInWordBtn = document.getElementById('preview-template-in-word-btn');
    if (previewInWordBtn) {
        previewInWordBtn.addEventListener('click', previewTemplateInWord);
    }
    
    });

/**
 * 检查是否处于排版模板模式
 * @returns {boolean} 是否处于模板模式
 */
// ============================================================
// docx映射卡片预览渲染
// ============================================================

/**
 * 构建docx映射预览HTML（用于模板卡片的预览对话框）
 * @param {Object} template - 带SemanticTags的映射卡片对象
 * @returns {string} HTML内容
 */
function buildDocxMappingPreviewHtml(template) {
    const tags = template.SemanticTags || [];
    let tagsHtml = tags.map(tag => {
        const fontDesc = [];
        if (tag.Font) {
            if (tag.Font.FontNameCN) fontDesc.push(tag.Font.FontNameCN);
            if (tag.Font.FontSize > 0) fontDesc.push(tag.Font.FontSize + 'pt');
            if (tag.Font.Bold) fontDesc.push('加粗');
            if (tag.Font.Italic) fontDesc.push('斜体');
        }
        const paraDesc = [];
        if (tag.Paragraph) {
            if (tag.Paragraph.Alignment) paraDesc.push(tag.Paragraph.Alignment);
            if (tag.Paragraph.LineSpacing > 0) paraDesc.push('行距' + tag.Paragraph.LineSpacing);
            if (tag.Paragraph.FirstLineIndent > 0) paraDesc.push('缩进' + tag.Paragraph.FirstLineIndent + '字符');
        }
        return `
            <div class="preview-element" style="margin-bottom: 4px;">
                <span class="element-name" style="font-family:monospace;font-size:11px;background:#edf2f7;padding:1px 4px;border-radius:3px;">${tag.TagId || ''}</span>
                <span style="font-weight:600;margin-left:6px;">${escapeHtml(tag.DisplayName || '')}</span>
                <div style="color:#718096;font-size:11px;margin-top:2px;">
                    ${fontDesc.length ? fontDesc.join(' ') : ''}
                    ${paraDesc.length ? ' | ' + paraDesc.join(' ') : ''}
                    ${tag.MatchHint ? ' | <span style="color:#a0aec0;">提示: ' + escapeHtml(tag.MatchHint) + '</span>' : ''}
                </div>
            </div>`;
    }).join('');

    return `
        <div class="preview-header">
            <h3 class="preview-title">${escapeHtml(template.Name)}</h3>
            <span class="preview-category" style="background:#38a169;color:#fff;">文档提取</span>
        </div>
        <p class="preview-description">${escapeHtml(template.Description || '')}</p>
        <div class="preview-section">
            <h4>语义标签映射（共 ${tags.length} 个）</h4>
            <div class="preview-elements-list">
                ${tagsHtml || '<div class="preview-empty">无标签</div>'}
            </div>
        </div>
    `;
}

// ============================================================
// 语义映射预览弹窗 (.docx模板解析后展示)
// ============================================================

/**
 * 显示语义映射预览弹窗（VB解析.docx后调用）
 * @param {Object} mapping - SemanticStyleMapping对象
 */
function showMappingPreview(mapping) {
    try {
        // 创建或获取弹窗
        let overlay = document.getElementById('mapping-preview-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'mapping-preview-overlay';
            overlay.className = 'styleguide-preview-overlay';
            overlay.innerHTML = `
                <div class="styleguide-preview-dialog" style="max-width: 560px;">
                    <div class="styleguide-preview-header">
                        <h3 id="mapping-preview-title">语义映射预览</h3>
                        <button class="styleguide-preview-close-btn" onclick="closeMappingPreview()">×</button>
                    </div>
                    <div class="styleguide-preview-body" id="mapping-preview-content" style="max-height: 60vh; overflow-y: auto;">
                    </div>
                    <div class="styleguide-preview-footer" style="display: flex; gap: 8px; justify-content: flex-end; padding: 12px 16px;">
                        <button onclick="closeMappingPreview()" style="padding: 6px 16px; border: 1px solid #ccc; background: white; border-radius: 4px; cursor: pointer; font-size: 13px;">关闭</button>
                        <button onclick="useMappingFromPreview()" style="padding: 6px 16px; border: none; background: #4299e1; color: white; border-radius: 4px; cursor: pointer; font-size: 13px;">使用此映射</button>
                    </div>
                </div>
            `;
            document.body.appendChild(overlay);
        }

        // 保存当前mapping供使用按钮回调
        window._currentPreviewMapping = mapping;

        // 渲染内容
        const content = document.getElementById('mapping-preview-content');
        if (!content) return;

        let html = '';

        // 映射名称
        const title = document.getElementById('mapping-preview-title');
        if (title) title.textContent = `语义映射预览 - ${mapping.Name || '未命名'}`;

        // 语义标签列表
        const tags = mapping.SemanticTags || [];
        html += `<div style="margin-bottom: 12px;"><h4 style="font-size: 13px; color: #4a5568; margin: 0 0 8px 0;">语义标签映射（共 ${tags.length} 个）</h4>`;
        html += '<div style="display: flex; flex-direction: column; gap: 6px;">';

        for (const tag of tags) {
            const fontDesc = [];
            if (tag.Font) {
                if (tag.Font.FontNameCN) fontDesc.push(tag.Font.FontNameCN);
                if (tag.Font.FontSize > 0) fontDesc.push(tag.Font.FontSize + 'pt');
                if (tag.Font.Bold) fontDesc.push('加粗');
                if (tag.Font.Italic) fontDesc.push('斜体');
            }

            const paraDesc = [];
            if (tag.Paragraph) {
                if (tag.Paragraph.Alignment) paraDesc.push(tag.Paragraph.Alignment);
                if (tag.Paragraph.LineSpacing > 0) paraDesc.push('行距' + tag.Paragraph.LineSpacing);
                if (tag.Paragraph.FirstLineIndent > 0) paraDesc.push('缩进' + tag.Paragraph.FirstLineIndent + '字符');
            }

            html += `
                <div style="background: #f7fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 10px; font-size: 12px;">
                    <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 3px;">
                        <span style="background: #edf2f7; padding: 1px 6px; border-radius: 3px; font-family: monospace; font-size: 11px; color: #4a5568;">${tag.TagId || ''}</span>
                        <span style="font-weight: 600; color: #2d3748;">${tag.DisplayName || ''}</span>
                    </div>
                    <div style="color: #718096; font-size: 11px;">
                        ${fontDesc.length ? '<span>' + fontDesc.join(' ') + '</span>' : ''}
                        ${paraDesc.length ? ' | <span>' + paraDesc.join(' ') + '</span>' : ''}
                        ${tag.MatchHint ? ' | <span style="color: #a0aec0;">提示: ' + tag.MatchHint + '</span>' : ''}
                    </div>
                </div>
            `;
        }
        html += '</div></div>';

        // 页面设置
        if (mapping.PageConfig && mapping.PageConfig.Margins) {
            const m = mapping.PageConfig.Margins;
            html += `<div style="margin-bottom: 8px;"><h4 style="font-size: 13px; color: #4a5568; margin: 0 0 6px 0;">页面设置</h4>`;
            html += `<div style="font-size: 12px; color: #718096;">上 ${(m.Top || 0).toFixed(2)}cm  下 ${(m.Bottom || 0).toFixed(2)}cm  左 ${(m.Left || 0).toFixed(2)}cm  右 ${(m.Right || 0).toFixed(2)}cm</div>`;
            html += '</div>';
        }

        content.innerHTML = html;

        // 显示弹窗
        overlay.style.display = 'flex';
    } catch (err) {
        console.error('showMappingPreview error:', err);
    }
}

/**
 * 关闭映射预览弹窗
 */
function closeMappingPreview() {
    const overlay = document.getElementById('mapping-preview-overlay');
    if (overlay) overlay.style.display = 'none';
    window._currentPreviewMapping = null;
}

/**
 * 使用预览中的映射（关闭预览 + 刷新模板列表使其出现在卡片中）
 */
function useMappingFromPreview() {
    closeMappingPreview();
    // 映射已在VB端保存到SemanticMappingManager，刷新模板列表使其出现在卡片中
    sendMessageToVB({ type: 'getReformatTemplates' });
}

/**
 * 发送上传.docx模板消息到VB
 */
function uploadDocxTemplate() {
    try {
        const payload = JSON.stringify({ type: 'uploadDocxTemplate' });
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage(payload);
        } else if (window.vsto) {
            window.vsto.postMessage(payload);
        }
    } catch (err) {
        console.error('uploadDocxTemplate error:', err);
    }
}
