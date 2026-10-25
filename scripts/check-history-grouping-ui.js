/**
 * check-history-grouping-ui.js - E2-05 历史列表分组与空状态契约检查
 * 用最小 DOM 桩在 node 里加载 history-manager.js，断言：分组只按「首次出现」切桶、
 * 绝不重排后端顺序；按时间分组用的是排序键 lastActiveAt；文档名交给组头承载；
 * 空状态给出可点击的引导而不是空白列表。
 * 运行：node scripts/check-history-grouping-ui.js
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repoRoot = path.resolve(__dirname, '..');
const historyPath = path.join(repoRoot, 'ShareRibbon', 'Resources', 'js', 'history-manager.js');

function assert(condition, message) {
    if (!condition) throw new Error(message);
}

// ---- 最小 DOM 桩：只支撑 history-manager.js 用到的那几件事 ----
function escapeText(value) {
    return String(value == null ? '' : value)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

class StubElement {
    constructor(doc, tag) {
        this._doc = doc;
        this.tagName = (tag || 'div').toUpperCase();
        this.id = '';
        this._classes = new Set();
        this.dataset = {};
        this.attributes = {};
        this.children = [];
        this._html = '';
        this._text = '';
        this.focused = 0;
        this.listeners = {};
    }

    get className() { return [...this._classes].join(' '); }
    set className(value) { this._classes = new Set(String(value || '').split(/\s+/).filter(Boolean)); }

    get classList() {
        const self = this;
        return {
            add(...names) { names.forEach(name => self._classes.add(name)); },
            remove(...names) { names.forEach(name => self._classes.delete(name)); },
            contains(name) { return self._classes.has(name); }
        };
    }

    setAttribute(name, value) {
        this.attributes[name] = String(value);
        if (name === 'id') this.id = String(value);
    }

    get textContent() { return this._text; }
    // escapeHtml 的实现就是「写 textContent、读 innerHTML」，桩必须同样转义，否则测不到 XSS 口径
    set textContent(value) { this._text = String(value == null ? '' : value); this._html = ''; }
    get innerHTML() { return this._html || escapeText(this._text); }
    set innerHTML(html) { this._html = String(html); this._text = ''; }

    set innerText(value) { this._text = String(value); }
    get innerText() { return this._text; }

    focus() { this.focused += 1; }
    select() { }
    addEventListener(name, handler) { this.listeners[name] = (this.listeners[name] || []).concat(handler); }
    dispatch(name, event) { (this.listeners[name] || []).forEach(handler => handler(event)); }
    querySelector() { return null; }
    closest() { return null; }

    hasClassInTree(name) {
        if (this._classes.has(name)) return true;
        return this.children.some(child => child.hasClassInTree(name));
    }
}

function createDocument() {
    const doc = {
        elements: new Map(),
        register(el) { if (el.id) doc.elements.set(el.id, el); },
        getElementById(id) { return doc.elements.get(id) || null; },
        querySelector() { return null; },
        createElement(tag) { return new StubElement(doc, tag); },
        addEventListener() { }
    };
    return doc;
}

const document = createDocument();
const sidebar = new StubElement(document, 'div');
sidebar.id = 'history-sidebar';
document.register(sidebar);
const overlay = new StubElement(document, 'div');
overlay.id = 'sidebar-overlay';
document.register(overlay);
const historyList = new StubElement(document, 'div');
historyList.id = 'history-list';
document.register(historyList);
const smartInput = new StubElement(document, 'div');
smartInput.id = 'smart-input';
document.register(smartInput);
const toggleBtn = new StubElement(document, 'button');
toggleBtn.id = 'history-toggle-btn';
document.register(toggleBtn);
const closeBtn = new StubElement(document, 'button');
closeBtn.id = 'close-sidebar-btn';
document.register(closeBtn);
const newSessionBtn = new StubElement(document, 'button');
newSessionBtn.id = 'new-session-btn';
document.register(newSessionBtn);

const posted = [];
const started = [];
const ctx = {
    document,
    console: { log: () => { }, warn: () => { }, error: (...args) => { started.push('console.error:' + args.join(' ')); } },
    setTimeout: () => 0,
    confirm: () => ctx.__confirmAnswer,
    officeAiPost: null
};
ctx.window = ctx;
vm.createContext(ctx);

const source = fs.readFileSync(historyPath, 'utf8');
vm.runInContext(source, ctx, { filename: historyPath });

ctx.window.officeAi = { post: (message) => posted.push(message) };
ctx.window.showOnboarding = () => started.push('onboarding');

const manager = ctx.window.historyManager;
assert(manager && typeof manager.displayHistoryFiles === 'function', 'historyManager is not exported');
manager.init();
manager.init();
assert(historyList.listeners.click.length === 1,
    'bindListEvents attached the delegated click handler more than once');
assert(historyList.dataset.bound === '1', 'delegated binding guard lost its dataset marker');

// ---- 1. scope=all 按文档分组：当前文档最前、同名不同文件不并桶、通用会话单独一组 ----
const day = (offsetDays) => {
    const now = new Date();
    const stamp = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) + offsetDays * 86400000);
    return `${stamp.getUTCFullYear()}-${String(stamp.getUTCMonth() + 1).padStart(2, '0')}-` +
        `${String(stamp.getUTCDate()).padStart(2, '0')} 09:30:00`;
};

// 后端给的顺序就是「当前文档整块在前，其余按最近活跃倒序」；同名不同文件、同文件被别的应用
// 会话插在中间都是真实形态，分组必须扛住这两种情况。
const allItems = [
    { sessionId: 'cur-1', title: '当前文档最近', createdAt: day(-9), lastActiveAt: day(0), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
    { sessionId: 'cur-2', title: '当前文档稍早', createdAt: day(-7), lastActiveAt: day(-2), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
    { sessionId: 'docB-1', title: '预算甲', createdAt: day(-8), lastActiveAt: day(-1), documentId: 'bbb', documentName: '预算.xlsx', isCurrentDocument: false },
    { sessionId: 'docC-1', title: '同名不同文件', createdAt: day(-5), lastActiveAt: day(-3), documentId: 'ccc', documentName: '预算.xlsx', isCurrentDocument: false },
    { sessionId: 'docB-2', title: '预算乙', createdAt: day(-6), lastActiveAt: day(-4), documentId: 'bbb', documentName: '预算.xlsx', isCurrentDocument: false },
    { sessionId: 'gen-1', title: '通用会话一条', createdAt: day(-4), lastActiveAt: day(-5), documentId: '', documentName: '', isCurrentDocument: false }
];

manager.displayHistoryFiles({
    items: allItems, scope: 'all', documentName: '报告.docx', hasDocument: true
});

const groups = manager.buildGroups();
assert(groups.length === 4, `scope=all must produce 4 document buckets, got ${groups.length}`);
assert(groups.map(g => g.label).join('|') === '报告.docx|预算.xlsx|预算.xlsx|通用会话',
    'document bucket labels wrong: ' + groups.map(g => g.label).join('|'));
assert(groups[0].isCurrent === true && !groups[1].isCurrent,
    'only the current-document bucket may carry the current marker');
assert(groups.map(g => g.items.length).join(',') === '2,2,1,1',
    'bucket counts wrong: ' + groups.map(g => g.items.length).join(','));
assert(groups[1].key !== groups[2].key,
    'two different files that share a name must not merge into one bucket');
assert(groups[2].items.map(s => s.sessionId).join(',') === 'docC-1',
    'the same-name bucket split lost an item: ' + groups.map(g => g.items.map(s => s.sessionId).join('+')).join(' | '));

const positionOf = new Map(allItems.map((s, index) => [s.sessionId, index]));
const flattened = groups.reduce((acc, g) => acc.concat(g.items), []);
assert(flattened.length === allItems.length, 'grouping dropped or duplicated a session');
assert(new Set(flattened.map(s => s.sessionId)).size === allItems.length,
    'grouping reported one session twice');
groups.forEach(g => {
    const positions = g.items.map(s => positionOf.get(s.sessionId));
    assert(positions.every((value, index) => index === 0 || value > positions[index - 1]),
        `bucket ${g.key} reordered rows the backend already sorted`);
});
const firstSeen = groups.map(g => Math.min(...g.items.map(s => positionOf.get(s.sessionId))));
assert(firstSeen.every((value, index) => index === 0 || value > firstSeen[index - 1]),
    'buckets must follow their first appearance in the backend list');

const listHtml = historyList.innerHTML;
assert(listHtml.split('class="history-group"').length - 1 === 4, 'rendered group wrappers mismatch');
assert(/history-group-current/.test(listHtml), 'current-document group header lost its marker class');
assert((listHtml.match(/history-group-current/g) || []).length === 1,
    'only one group header may carry the current-document marker');
const metaParts = (listHtml.match(/<span class="history-meta-part">([^<]*)<\/span>/g) || [])
    .map(part => part.replace(/<[^>]*>/g, ''));
assert(metaParts.length === allItems.length, 'every item must keep one meta line');
assert(metaParts.every(value => /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(value)),
    'item meta must be a minute-precision timestamp only: ' + metaParts.join(' / '));
assert(!/报告\.docx<\/span><span class="history-meta-part">/.test(listHtml),
    'the document name now lives in the group header, not in every item meta');

// ---- 2. 分组标题只走 escapeHtml：文档名/标题里的标签不能变成 markup ----
manager.displayHistoryFiles({
    items: [{
        sessionId: 'x-1', title: '<img src=x onerror=alert(1)>', createdAt: day(0), lastActiveAt: day(0),
        documentId: 'ddd', documentName: '<b>evil.docx</b>', isCurrentDocument: false
    }],
    scope: 'all', documentName: '报告.docx', hasDocument: true
});
const unsafeHtml = historyList.innerHTML;
assert(!/<img src=x/.test(unsafeHtml), 'session title reached the sidebar as markup');
assert(/&lt;img src=x/.test(unsafeHtml), 'session title was not escaped through textContent');
assert(!/<b>evil\.docx<\/b>/.test(unsafeHtml), 'document name reached the group header as markup');

// ---- 3. scope=document 按时间分组：桶必须跟着排序键 lastActiveAt，而不是首条消息时间 ----
manager.displayHistoryFiles({
    items: [
        { sessionId: 't-1', title: '很久以前开、今天还在聊', createdAt: '2020-01-01 08:00:00', lastActiveAt: day(0), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
        { sessionId: 't-2', title: '今天', createdAt: day(0), lastActiveAt: day(0), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
        { sessionId: 't-3', title: '昨天', createdAt: day(-2), lastActiveAt: day(-1), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
        { sessionId: 't-4', title: '这周', createdAt: day(-6), lastActiveAt: day(-5), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
        { sessionId: 't-5', title: '这个月', createdAt: day(-21), lastActiveAt: day(-20), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
        { sessionId: 't-6', title: '很久以前', createdAt: day(-201), lastActiveAt: day(-200), documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true },
        { sessionId: 't-7', title: '时间读不出来', createdAt: '', lastActiveAt: '', documentId: 'aaa', documentName: '报告.docx', isCurrentDocument: true }
    ],
    scope: 'document', documentName: '报告.docx', hasDocument: true
});
const timeGroups = manager.buildGroups();
assert(timeGroups.map(g => g.label).join('|') === '今天|昨天|最近 7 天|最近 30 天|更早',
    'time buckets wrong: ' + timeGroups.map(g => g.label).join('|'));
assert(timeGroups[0].items.map(s => s.sessionId).join(',') === 't-1,t-2',
    'a session active today must bucket by lastActiveAt even when its first message is years old');
assert(timeGroups[timeGroups.length - 1].items.map(s => s.sessionId).join(',') === 't-6,t-7',
    'unparsable timestamps must fall into the oldest bucket instead of breaking the list');

// ---- 4. 空状态要有引导，而不是一句「没有」 ----
manager.displayHistoryFiles({ items: [], scope: 'document', documentName: '报告.docx', hasDocument: true });
let emptyHtml = historyList.innerHTML;
assert(/《报告\.docx》还没有会话/.test(emptyHtml),
    'document scope empty state lost the document name');
assert(/data-empty="start"/.test(emptyHtml) && /data-scope="all"/.test(emptyHtml),
    'document scope empty state must offer both "start here" and "show all"');

manager.displayHistoryFiles({ items: [], scope: 'document', documentName: '', hasDocument: false });
emptyHtml = historyList.innerHTML;
assert(/还没保存/.test(emptyHtml) && /通用会话/.test(emptyHtml),
    'unsaved-document scope must explain why nothing is bound to it');

manager.displayHistoryFiles({ items: [], scope: 'all', documentName: '', hasDocument: false });
emptyHtml = historyList.innerHTML;
assert(/还没有任何历史会话/.test(emptyHtml), 'global empty state lost its headline');
assert(/empty-state-hint/.test(emptyHtml) && /把销售表按金额排序/.test(emptyHtml),
    'global empty state must suggest a concrete first task');
assert(/data-empty="onboarding"/.test(emptyHtml), 'global empty state lost the onboarding entry');

delete ctx.window.showOnboarding;
manager.displayHistoryFiles({ items: [], scope: 'all', documentName: '', hasDocument: false });
assert(!/data-empty="onboarding"/.test(historyList.innerHTML),
    'onboarding entry must disappear when the wizard is not on the page');
ctx.window.showOnboarding = () => started.push('onboarding');

// ---- 5. 空状态与列表的动作都走委托绑定，且不另开发送路径 ----
function targetFor(attrs) {
    return {
        classList: { contains: () => false },
        closest(selector) {
            if (!Object.prototype.hasOwnProperty.call(attrs, selector)) return null;
            const node = attrs[selector];
            return node === true ? { getAttribute: () => null } : node;
        }
    };
}

function attrNode(attrs) {
    return { getAttribute: (name) => (Object.prototype.hasOwnProperty.call(attrs, name) ? attrs[name] : null) };
}

posted.length = 0;
manager.isOpen = true;
historyList.dispatch('click', { target: targetFor({ '[data-empty]': attrNode({ 'data-empty': 'start' }) }) });
assert(manager.isOpen === false, 'the start-guided action must close the sidebar');
assert(smartInput.focused === 1, 'the start-guided action must hand focus back to the composer');
assert(posted.length === 0, 'the start-guided action must not post a second send path');

historyList.dispatch('click', { target: targetFor({ '[data-empty]': attrNode({ 'data-empty': 'onboarding' }) }) });
assert(started.indexOf('onboarding') >= 0, 'the onboarding action must reuse window.showOnboarding');

posted.length = 0;
historyList.dispatch('click', { target: targetFor({ '[data-scope]': attrNode({ 'data-scope': 'document' }) }) });
assert(posted.some(m => m.type === 'getSessionList' && m.scope === 'document'),
    'the scope tab must reload the list through the existing message contract');

posted.length = 0;
historyList.dispatch('click', {
    target: targetFor({ '.history-item': attrNode({ 'data-session-id': 'docB-1' }) })
});
assert(posted.some(m => m.type === 'loadSession' && m.sessionId === 'docB-1'),
    'clicking a grouped item must still load that session');

posted.length = 0;
ctx.__confirmAnswer = false;
historyList.dispatch('click', {
    target: targetFor({
        '.history-item': attrNode({ 'data-session-id': 'docB-1' }),
        '.history-item-delete': attrNode({})
    })
});
assert(posted.length === 0, 'a cancelled delete confirmation must not reach the backend');

ctx.__confirmAnswer = true;
historyList.dispatch('click', {
    target: targetFor({
        '.history-item': attrNode({ 'data-session-id': 'docB-1' }),
        '.history-item-delete': attrNode({})
    })
});
assert(posted.some(m => m.type === 'deleteSession' && m.sessionId === 'docB-1'),
    'a confirmed delete must reach the backend');

console.log('history grouping UI contract checks passed');
