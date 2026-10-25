/**
 * check-agent-card-ui.js - E1-08 信任层 UI 契约检查
 * 用最小 DOM 桩在 node 里加载 agent-card.js，断言计划卡信任信息、重规划原位重绘、
 * 步骤状态令牌、迭代归属步骤、终态变更摘要与撤销入口。
 * 运行：node scripts/check-agent-card-ui.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repoRoot = path.resolve(__dirname, '..');
const cardPath = path.join(repoRoot, 'ShareRibbon', 'Resources', 'js', 'agent-card.js');

function assert(condition, message) {
    if (!condition) throw new Error(message);
}

// ---- 最小 DOM 桩：只支撑 agent-card.js 用到的那几件事 ----
class StubElement {
    constructor(tag, doc) {
        this.tagName = (tag || 'div').toUpperCase();
        this._doc = doc;
        this.id = '';
        this._className = '';
        this._classes = new Set();
        this.dataset = {};
        this.style = {};
        this.attributes = {};
        this.children = [];
        this.parentNode = null;
        this.disabled = false;
        this.textContent = '';
        this.title = '';
        this.type = '';
        this._html = '';
        this.open = false;
        this.listeners = {};
    }

    get parentElement() { return this.parentNode; }

    get className() { return this._className; }    set className(value) {
        this._className = value || '';
        this._classes = new Set(this._className.split(/\s+/).filter(Boolean));
    }

    get classList() {
        const self = this;
        return {
            add(...names) { names.forEach(name => self._classes.add(name)); self._sync(); },
            remove(...names) { names.forEach(name => self._classes.delete(name)); self._sync(); },
            contains(name) { return self._classes.has(name); }
        };
    }

    _sync() { this._className = [...this._classes].join(' '); }

    setAttribute(name, value) {
        this.attributes[name] = String(value);
        if (name === 'id') this.id = String(value);
    }

    get innerHTML() { return this._html; }
    set innerHTML(html) {
        // 重绘即销毁旧后代，否则测试会拿到已经不在卡片上的僵尸节点
        this.children.forEach(child => child._unregisterTree());
        this._html = String(html);
        this.children = [];
        this.textContent = '';
        const idPattern = /\sid="([^"]+)"/g;
        let match;
        while ((match = idPattern.exec(this._html)) !== null) {
            const child = new StubElement('div', this._doc);
            child.id = match[1];
            const classMatch = this._html
                .slice(Math.max(0, match.index - 200), match.index)
                .match(/class="([^"]*)"$/);
            if (classMatch) child.className = classMatch[1];
            child.parentNode = this;
            this.children.push(child);
            this._doc.register(child);
        }
    }

    _unregisterTree() {
        this._doc.unregister(this);
        this.children.forEach(child => child._unregisterTree());
        this.parentNode = null;
    }

    appendChild(node) {
        node.parentNode = this;
        this.children.push(node);
        if (node.id) this._doc.register(node);
        return node;
    }

    insertBefore(node, ref) {
        node.parentNode = this;
        const index = ref ? this.children.indexOf(ref) : -1;
        if (index < 0) this.children.push(node); else this.children.splice(index, 0, node);
        if (node.id) this._doc.register(node);
        return node;
    }

    replaceChild(next, old) {
        const index = this.children.indexOf(old);
        if (index >= 0) this.children[index] = next; else this.children.push(next);
        next.parentNode = this;
        return old;
    }

    remove() {
        if (this.parentNode) {
            const index = this.parentNode.children.indexOf(this);
            if (index >= 0) this.parentNode.children.splice(index, 1);
            this._unregisterTree();
        }
    }

    querySelector() { return null; }
    closest() { return null; }
    scrollIntoView() { }
    addEventListener(name, handler) {
        this.listeners[name] = (this.listeners[name] || []).concat(handler);
    }
    dispatch(name) {
        (this.listeners[name] || []).forEach(handler => handler());
    }

    // 递归收集可见文本，供断言使用（不含标签，避免把标记当成内容）
    text() {
        let out = this.textContent || '';
        this.children.forEach(child => { out += ' ' + child.text(); });
        return out;
    }

    hasClassInTree(name) {
        if (this._classes.has(name)) return true;
        return this.children.some(child => child.hasClassInTree(name));
    }
}

function createDocument() {
    const doc = {
        elements: new Map(),
        register(el) { if (el.id) doc.elements.set(el.id, el); },
        unregister(el) { if (el.id) doc.elements.delete(el.id); },
        getElementById(id) { return doc.elements.get(id) || null; },
        querySelector() { return null; },
        querySelectorAll(selector) {
            // 只支持 ".class" 形式：agent-card.js 用它定位撤销按钮
            const name = String(selector).replace(/^\./, '');
            const found = [];
            const walk = (el) => {
                if (el.classList.contains(name)) found.push(el);
                el.children.forEach(walk);
            };
            walk(chatRoot);
            return found;
        },
        createElement(tag) { return new StubElement(tag, doc); }
    };
    return doc;
}

const document = createDocument();
const chatRoot = new StubElement('div', document);
chatRoot.id = 'chat-container';
document.register(chatRoot);

const sandbox = {
    window: {},
    document,
    console,
    setTimeout: () => 0,
    prompt: () => null,
    escapeHtml: (value) => String(value == null ? '' : value)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;'),
    formatDateTime: () => '2026-10-24 10:00:00',
    generateUUID: () => 'fallback-uuid',
    requestApprove: () => { sandbox.__approved = (sandbox.__approved || 0) + 1; },
    requestApproveToolSession: () => { },
    requestReject: () => { },
    requestRefinePlan: () => { },
    requestAbortAgent: () => { },
    sendMessageToServer: (message) => { sandbox.__posted = (sandbox.__posted || []).concat(message); }
};

const source = fs.readFileSync(cardPath, 'utf8');
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: cardPath });

const card = sandbox.window;
assert(typeof card.showAgentPlanCard === 'function', 'showAgentPlanCard is not exported');
assert(typeof card.completeAgent === 'function', 'completeAgent is not exported');

// agent-card.js 不得再用同名 getStepIcon 覆盖 code-handler.js 的图标字典
assert(!/function\s+getStepIcon\s*\(/.test(source),
    'agent-card.js declares getStepIcon again, which shadows code-handler.js icon map');

// ---- 1. 计划卡显示范围/数量/风险 ----
const sessionId = 'run-1';
card.showAgentPlanCard({
    sessionId,
    understanding: '把选中的三段正文改成宋体小四',
    planVersion: 1,
    scope: 'Word《季度报告.docx》 当前选区（3 项）',
    stepCount: 2,
    toolCount: 2,
    riskLevel: 'risky',
    riskyTools: ['FormatRange'],
    summary: '两段格式调整',
    steps: [
        { description: '设置字体', code: '{}', language: 'json' },
        { description: '设置字号', code: '{}', language: 'json' }
    ],
    replaceThinkingUuid: ''
});

const planContainer = document.getElementById('agent-plan-' + sessionId);
assert(planContainer, 'plan card container missing');
assert(/agent-trust-scope/.test(planContainer.innerHTML), 'plan card renders no scope chip');
assert(/当前选区（3 项）/.test(planContainer.innerHTML), 'plan card lost the scope text');
assert(/数量 2 步 · 2 个工具/.test(planContainer.innerHTML), 'plan card renders no count chip');
assert(/agent-trust-risk-risky/.test(planContainer.innerHTML), 'plan card renders no risk chip');
assert(/需确认 1 项/.test(planContainer.innerHTML), 'plan card risk chip has no risky tool count');
assert(/开始执行/.test(document.getElementById('agent-actions-' + sessionId).innerHTML),
    'first plan card lost its confirm-execution button');

// 步骤状态令牌：正在执行 / 失败 / 用户拒绝（跳过）互不覆盖
card.updateAgentStep(sessionId, 0, 'running', '');
const step0 = document.getElementById(`agent-step-${sessionId}-0`);
assert(step0.classList.contains('step-running'), 'running step lost its class');

card.updateAgentStep(sessionId, 0, 'failed', '宿主超时');
assert(step0.classList.contains('step-failed'), 'failed step lost its class');

card.updateAgentStep(sessionId, 0, 'skipped', '重新规划后跳过这一步');
assert(step0.attributes['data-status'] === 'skipped', 'skipped step status attribute missing');
assert(!step0.classList.contains('step-failed'), 'skipped step kept the failed styling');
assert(document.getElementById(`step-icon-${sessionId}-0`).textContent === '⏭', 'skipped step icon wrong');

// 迭代详情按所属步骤落位：重规划后全局迭代号和步骤号必然错开
card.updateAgentIteration(sessionId, {
    index: 7,
    stepIndex: 0,
    thought: '需要先确认选区',
    action: 'FormatRange',
    observation: '已设置 3 段'
});
assert(/已设置 3 段/.test(document.getElementById(`iteration-area-${sessionId}-0`).innerHTML),
    'iteration detail did not land in its owning step area');

// ---- 2. 重规划按同一会话原位重绘，且不退回“等待确认执行” ----
const childrenBefore = chatRoot.children.length;
card.showAgentPlanCard({
    sessionId,
    understanding: '继续未完成部分',
    planVersion: 2,
    scope: 'Word《季度报告.docx》 整篇文档',
    stepCount: 1,
    toolCount: 1,
    riskLevel: 'safe',
    riskyTools: [],
    summary: '补写标题',
    steps: [{ description: '插入标题', code: '{}', language: 'json' }],
    replaceThinkingUuid: ''
});
assert(chatRoot.children.length === childrenBefore, 'replan appended a second plan card');
assert(/agent-plan-version/.test(planContainer.innerHTML), 'replanned card shows no version marker');
assert(/第 2 版/.test(planContainer.innerHTML), 'replanned card version marker lost');
const replanActions = document.getElementById('agent-actions-' + sessionId);
assert(/终止执行/.test(replanActions.innerHTML), 'replanned card fell back to the confirm-execution state');
assert(!/confirmAgentExecution/.test(replanActions.innerHTML), 'replanned card still offers start execution');
assert(/按新计划继续执行/.test(document.getElementById('agent-status-' + sessionId).innerHTML),
    'replanned card status bar did not switch to running');
assert(!document.getElementById(`agent-step-${sessionId}-1`),
    'stale step node from the old plan survived the re-render');

// 终止态单独一个令牌，不能和失败混在一起
card.updateAgentStep(sessionId, 0, 'aborted', '审批已放弃');
assert(document.getElementById(`step-icon-${sessionId}-0`).textContent === '⏹', 'aborted step icon wrong');

// ---- 3. 终态变更摘要 + 不可撤销时给出原因 ----
card.completeAgent(sessionId, true, '已完成 2 个操作', '', {
    writtenCount: 2,
    rolledBackCount: 1,
    changeTruncated: false,
    changes: [
        { seq: 1, toolId: 'FormatRange', summary: '将 3 段设为宋体', partialApply: false },
        { seq: 2, toolId: 'InsertText', summary: '<img src=x onerror=alert(1)>', partialApply: true }
    ],
    canUndo: false,
    runId: '',
    units: 0,
    undoReason: '有改动缺少可靠的撤销单元或文档指纹'
});

const summaryBox = document.getElementById('agent-change-summary-' + sessionId);
assert(summaryBox, 'terminal change summary missing');
assert(/本轮已写入 2 处/.test(summaryBox.text()), 'change summary lost the written count');
assert(/将 3 段设为宋体/.test(summaryBox.text()), 'change summary lost an entry');
assert(/部分生效/.test(summaryBox.text()), 'change summary did not flag a partial write');
assert(/已自动回滚 1 处/.test(summaryBox.text()), 'change summary lost the rollback note');
assert(!/onerror/.test(summaryBox.innerHTML), 'change summary wrote host text as markup');
const cardText = summaryBox.parentNode.text();
assert(/有改动缺少可靠的撤销单元/.test(cardText), 'card hid the undo entry without a reason');

// ---- 4. 可撤销时给按钮，并且不追加原因 ----
const second = 'run-2';
card.showAgentPlanCard({
    sessionId: second, understanding: 'u', planVersion: 1, scope: '', stepCount: 1,
    toolCount: 1, riskLevel: 'safe', riskyTools: [], summary: 's',
    steps: [{ description: 'd', code: '{}', language: 'json' }],
    replaceThinkingUuid: ''
});
card.completeAgent(second, true, 'done', '', {
    writtenCount: 1,
    changes: [{ seq: 1, toolId: 'InsertText', summary: '写入一段', partialApply: false }],
    canUndo: true, runId: 'run-xyz', units: 3, undoReason: ''
});

const secondActions = document.getElementById('agent-actions-' + second);
const undoButton = secondActions.children.find(el => el.classList.contains('agent-btn-undo'));
assert(undoButton, 'undo button missing when the run is fully undoable');
assert(undoButton.dataset.runId === 'run-xyz', 'undo button carries the wrong run id');
assert(/撤销本轮（3 个撤销单元）/.test(undoButton.textContent), 'undo button lost its unit count');
assert(!secondActions.hasClassInTree('agent-undo-reason'), 'undo reason shown next to an enabled undo button');

undoButton.dispatch('click');
assert(sandbox.__posted && sandbox.__posted.some(m => m.type === 'agent:undoRun' && m.runId === 'run-xyz'),
    'undo click did not post agent:undoRun with the run id');

// 撤销结果回写：按 runId 精确替换按钮
card.setAgentUndoResult('run-xyz', true, '已核对文档指纹', true);
assert(/已撤销本轮/.test(secondActions.text()), 'undo result did not replace the button');

// 没有计划卡/没有终态数据时终态仍要收敛，不能抛异常
card.completeAgent('run-none', false, '已终止', '', null);

console.log('agent-card UI contract checks passed');
