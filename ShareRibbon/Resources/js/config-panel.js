/**
 * config-panel.js - 配置面板：场景与 Skills、记忆管理（阶段二）
 * 与 VB 通过 postMessage 通信：getPromptTemplates, savePromptTemplate, deletePromptTemplate, getAtomicMemories, deleteAtomicMemory, getUserProfile, saveUserProfile
 */

window.configPanel = {
    _currentSkillsId: null,
    _currentAtomicId: null,
    _skillsList: [],
    _atomicList: [],
    /** 当前宿主应用：Excel / Word / PowerPoint，对应 scenario：excel / word / ppt */
    _appType: 'Excel',

    init: function () {
        this._appType = (typeof window.officeAppType === 'string' && window.officeAppType) ? window.officeAppType : 'Excel';
        this._buildScenarioOptions();

        const btnSkills = document.getElementById('btn-config-skills');
        const btnMemory = document.getElementById('btn-config-memory');
        const overlay = document.getElementById('config-panel-overlay');
        const panel = document.getElementById('config-panel');
        const closeBtn = document.getElementById('config-panel-close');

        if (btnSkills) btnSkills.addEventListener('click', () => this.open('skills'));
        if (btnMemory) btnMemory.addEventListener('click', () => this.open('memory'));
        if (overlay) overlay.addEventListener('click', () => this.close());
        if (closeBtn) closeBtn.addEventListener('click', () => this.close());

        document.querySelectorAll('.config-tab').forEach(tab => {
            tab.addEventListener('click', () => this.switchTab(tab.dataset.tab));
        });

        document.getElementById('config-skills-scenario')?.addEventListener('change', () => this.loadPromptTemplates());
        document.getElementById('config-skills-refresh')?.addEventListener('click', () => this.loadPromptTemplates());
        document.getElementById('config-skills-list')?.addEventListener('click', (e) => {
            const item = e.target.closest('.config-list-item');
            if (item && item.dataset.id) {
                const r = this._skillsList.find(x => String(x.id) === String(item.dataset.id));
                if (r) this.selectSkillsItem(r);
            }
        });
        document.getElementById('config-skills-save')?.addEventListener('click', () => this.savePromptTemplate());
        document.getElementById('config-skills-delete')?.addEventListener('click', () => this.deletePromptTemplate());
        document.getElementById('config-skills-new')?.addEventListener('click', () => this.newPromptTemplate());
        document.getElementById('config-skills-import-folder')?.addEventListener('click', () => this.importSkillsFromFolder());
        document.getElementById('config-skills-isskill')?.addEventListener('change', (e) => {
            document.getElementById('config-skills-supported-wrap').style.display = e.target.checked ? 'block' : 'none';
        });

        document.getElementById('config-memory-atomic-refresh')?.addEventListener('click', () => this.loadAtomicMemories());
        document.getElementById('config-memory-atomic-list')?.addEventListener('click', (e) => {
            const item = e.target.closest('.config-list-item');
            if (item && item.dataset.id) {
                const r = this._atomicList.find(x => String(x.id) === String(item.dataset.id));
                if (r) this.selectAtomicItem(r.id, r.content);
            }
        });
        document.getElementById('config-memory-atomic-delete')?.addEventListener('click', () => this.deleteAtomicMemory());
        document.getElementById('config-memory-profile-save')?.addEventListener('click', () => this.saveUserProfile());
        document.getElementById('config-memory-profile-clear')?.addEventListener('click', () => this.clearUserProfile());
    },

    open: function (tab) {
        document.getElementById('config-panel-overlay').style.display = 'block';
        document.getElementById('config-panel').style.display = 'flex';
        this.switchTab(tab || 'skills');
    },

    close: function () {
        document.getElementById('config-panel-overlay').style.display = 'none';
        document.getElementById('config-panel').style.display = 'none';
    },

    /** 按当前宿主只显示本应用 + common，默认选本应用 */
    _buildScenarioOptions: function () {
        const sel = document.getElementById('config-skills-scenario');
        if (!sel) return;
        const appToScenario = { Excel: 'excel', Word: 'word', PowerPoint: 'ppt' };
        const current = appToScenario[this._appType] || 'excel';
        const options = [
            { value: current, text: current },
            { value: 'common', text: 'common' }
        ];
        sel.innerHTML = options.map(o => `<option value="${o.value}">${o.text}</option>`).join('');
    },

    switchTab: function (tab) {
        document.querySelectorAll('.config-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tab));
        document.getElementById('config-tab-skills').style.display = tab === 'skills' ? 'block' : 'none';
        document.getElementById('config-tab-memory').style.display = tab === 'memory' ? 'block' : 'none';
        if (tab === 'skills') this.loadPromptTemplates();
        if (tab === 'memory') {
            this.loadAtomicMemories();
            this.loadUserProfile();
        }
    },

    sendToVB: function (msg) {
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage(msg);
        } else if (window.vsto && typeof window.vsto.sendMessage === 'function') {
            window.vsto.sendMessage(JSON.stringify(msg));
        }
    },

    loadPromptTemplates: function () {
        const scenario = document.getElementById('config-skills-scenario')?.value || 'excel';
        this.sendToVB({ type: 'getPromptTemplates', scenario: scenario });
    },

    selectSkillsItem: function (r) {
        if (!r) return;
        this._currentSkillsId = r.id;
        document.querySelectorAll('#config-skills-list .config-list-item').forEach(el => el.classList.remove('selected'));
        const item = document.querySelector(`#config-skills-list .config-list-item[data-id="${r.id}"]`);
        if (item) item.classList.add('selected');
        document.getElementById('config-skills-name').value = r.templateName || '';
        document.getElementById('config-skills-content').value = r.content || '';
        const chk = document.getElementById('config-skills-isskill');
        chk.checked = r.isSkill === 1;
        document.getElementById('config-skills-supported-wrap').style.display = chk.checked ? 'block' : 'none';
        let supported = '';
        if (r.extraJson) try { const o = JSON.parse(r.extraJson); supported = (o.supported_apps || o.supportedApps || []).join(','); } catch (e) {}
        document.getElementById('config-skills-supported').value = supported;
    },

    newPromptTemplate: function () {
        this._currentSkillsId = null;
        document.querySelectorAll('#config-skills-list .config-list-item').forEach(el => el.classList.remove('selected'));
        document.getElementById('config-skills-name').value = '';
        document.getElementById('config-skills-content').value = '';
        document.getElementById('config-skills-isskill').checked = false;
        document.getElementById('config-skills-supported-wrap').style.display = 'none';
        document.getElementById('config-skills-supported').value = '';
    },

    savePromptTemplate: function () {
        const scenario = document.getElementById('config-skills-scenario')?.value || 'excel';
        const name = document.getElementById('config-skills-name')?.value?.trim() || '';
        const content = document.getElementById('config-skills-content')?.value || '';
        const isSkill = document.getElementById('config-skills-isskill')?.checked ? 1 : 0;
        let extraJson = '{}';
        if (isSkill) {
            const supported = document.getElementById('config-skills-supported')?.value?.split(',').map(s => s.trim()).filter(Boolean) || [];
            extraJson = JSON.stringify({ supported_apps: supported });
        }
        this.sendToVB({
            type: 'savePromptTemplate',
            id: this._currentSkillsId ? parseInt(this._currentSkillsId, 10) : 0,
            templateName: name,
            scenario: scenario,
            content: content,
            isSkill: isSkill,
            extraJson: extraJson,
            sort: 0
        });
    },

    deletePromptTemplate: function () {
        if (!this._currentSkillsId) return;
        if (!confirm('确定删除当前项？')) return;
        const scenario = document.getElementById('config-skills-scenario')?.value || 'excel';
        this.sendToVB({ type: 'deletePromptTemplate', id: parseInt(this._currentSkillsId, 10), scenario: scenario });
    },

    /** 从文件夹批量导入 Skill（.json/.md），由 VB 弹窗选目录并解析导入 */
    importSkillsFromFolder: function () {
        const scenario = document.getElementById('config-skills-scenario')?.value || 'excel';
        this.sendToVB({ type: 'importSkillsFromFolder', scenario: scenario });
    },

    setPromptTemplates: function (list) {
        this._skillsList = list || [];
        const el = document.getElementById('config-skills-list');
        if (!el) return;
        el.innerHTML = this._skillsList.map(r => {
            const name = (r.templateName || '(未命名)').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            return `<div class="config-list-item" data-id="${r.id}">${name} ${r.isSkill === 1 ? '[Skill]' : ''}</div>`;
        }).join('');
    },

    loadAtomicMemories: function () {
        this.sendToVB({ type: 'getAtomicMemories', limit: 100, appType: this._appType });
    },

    selectAtomicItem: function (id, content) {
        this._currentAtomicId = id;
        document.querySelectorAll('#config-memory-atomic-list .config-list-item').forEach(el => el.classList.remove('selected'));
        const item = document.querySelector(`#config-memory-atomic-list .config-list-item[data-id="${id}"]`);
        if (item) item.classList.add('selected');
        document.getElementById('config-memory-atomic-content').value = content || '';
    },

    deleteAtomicMemory: function () {
        if (!this._currentAtomicId) return;
        if (!confirm('确定删除选中记忆？')) return;
        this.sendToVB({ type: 'deleteAtomicMemory', id: parseInt(this._currentAtomicId, 10), appType: this._appType });
    },

    setAtomicMemories: function (list) {
        this._atomicList = list || [];
        const el = document.getElementById('config-memory-atomic-list');
        if (!el) return;
        this._currentAtomicId = null;
        document.getElementById('config-memory-atomic-content').value = '';
        el.innerHTML = this._atomicList.map(r => {
            const preview = (r.content || '').substring(0, 60).replace(/</g, '&lt;').replace(/>/g, '&gt;');
            return `<div class="config-list-item" data-id="${r.id}">${preview}${(r.content || '').length > 60 ? '…' : ''}</div>`;
        }).join('');
    },

    loadUserProfile: function () {
        this.sendToVB({ type: 'getUserProfile' });
    },

    setUserProfile: function (content) {
        const el = document.getElementById('config-memory-profile');
        if (el) el.value = content || '';
    },

    saveUserProfile: function () {
        const content = document.getElementById('config-memory-profile')?.value || '';
        this.sendToVB({ type: 'saveUserProfile', content: content });
    },

    clearUserProfile: function () {
        if (!confirm('确定清空用户画像？')) return;
        document.getElementById('config-memory-profile').value = '';
        this.sendToVB({ type: 'saveUserProfile', content: '' });
    }
};

// VB 回调
window.setPromptTemplatesList = function (list) {
    if (window.configPanel) window.configPanel.setPromptTemplates(list);
};
window.setAtomicMemoriesList = function (list) {
    if (window.configPanel) window.configPanel.setAtomicMemories(list);
};
window.setUserProfileContent = function (content) {
    if (window.configPanel) window.configPanel.setUserProfile(content);
};
window.configSaveResult = function (ok, message) {
    if (typeof window.configPanel !== 'undefined' && message) {
        if (ok) window.configPanel.loadPromptTemplates();
        alert(message);
    }
};
