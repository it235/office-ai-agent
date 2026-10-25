/**
 * command-palette.js - E4-03 命令面板
 *
 * 一个输入框同时接受「命令」（QuickActionCatalog 同步下来的宿主动作）和「自然语言」：
 * 两类选择都只发一条 startAgent 给后端，与面板输入、Ribbon 按钮、右键菜单同一个闸门、
 * 同一条 Agent 主链。这里不判断模型有没有配好，也不复制任何话术。
 * 唤起方式只有两条：宿主快捷键由 VB 调 showCommandPalette()，面板按钮直接调同一个函数。
 */

window.commandPaletteState = window.commandPaletteState || { actions: [], packages: [] };

/** 设置面板里那几个原生入口开关的通用赋值：元素不存在就什么都不做，免得新前端还没加载时把整段注入脚本打断。 */
function setNativeEntrySetting(elementId, value, title) {
    const element = document.getElementById(elementId);
    if (!element) return;
    if (typeof value === 'boolean') element.checked = value;
    else if (typeof value === 'string') element.value = value;
    if (typeof title === 'string') element.title = title;
}

function commandPaletteActions() {
    const list = window.commandPaletteState && window.commandPaletteState.actions;
    return Array.isArray(list) ? list : [];
}

window.syncCommandPalette = function (state) {
    if (state && typeof state === 'object') {
        window.commandPaletteState = state;
        if (!Array.isArray(state.actions)) window.commandPaletteState.actions = [];
        // 旧后端没下发 packages 时按"全部未归组"渲染，而不是让分组代码去猜包名
        if (!Array.isArray(state.packages)) window.commandPaletteState.packages = [];
    }
};

window.showCommandPalette = function () {
    const overlay = document.getElementById('command-palette-overlay');
    const palette = document.getElementById('command-palette');
    if (overlay && palette && palette.style.display === 'block') {
        commandPaletteClose();
        return;
    }

    if (!commandPaletteEnsureDom()) return;
    overlay.style.display = 'block';
    palette.style.display = 'block';

    const input = document.getElementById('command-palette-input');
    input.value = '';
    // 每次打开都从头开始：上一次停在第几条和这次要做什么没有关系
    commandPaletteCursor = 0;
    commandPaletteRender('');
    input.focus();
};

function commandPaletteClose() {
    const overlay = document.getElementById('command-palette-overlay');
    const palette = document.getElementById('command-palette');
    if (overlay) overlay.style.display = 'none';
    if (palette) palette.style.display = 'none';
}

/** 浮层只建一次：反复 innerHTML 重建会把输入焦点和候选列表一起丢掉。 */
function commandPaletteEnsureDom() {
    if (document.getElementById('command-palette')) return true;
    if (!document.body) return false;

    const overlay = document.createElement('div');
    overlay.id = 'command-palette-overlay';
    overlay.className = 'command-palette-overlay';
    overlay.onclick = commandPaletteClose;

    const palette = document.createElement('div');
    palette.id = 'command-palette';
    palette.className = 'command-palette';

    const input = document.createElement('input');
    input.id = 'command-palette-input';
    input.className = 'command-palette-input';
    input.type = 'text';
    input.placeholder = '输入命令，或直接用自然语言描述要做什么';
    input.setAttribute('aria-label', '命令面板');
    input.oninput = function () { commandPaletteRender(input.value); };
    input.onkeydown = commandPaletteKeydown;

    const list = document.createElement('div');
    list.id = 'command-palette-list';
    list.className = 'command-palette-list';

    const hint = document.createElement('div');
    hint.className = 'command-palette-hint';
    hint.textContent = '↑↓ 选择 · Enter 执行 · Esc 关闭';

    palette.appendChild(input);
    palette.appendChild(list);
    palette.appendChild(hint);
    document.body.appendChild(overlay);
    document.body.appendChild(palette);
    return true;
}

/**
 * 候选 = 命中的目录动作 + 一条「按原话交给 Agent」。
 * 空输入也给出动作清单，用户不用先想好措辞；自然语言那条始终排在最后，
 * 免得它抢在动词前面把「命令面板」变成第二个输入框。
 * E6-04：后端把动作按任务包顺序排好再下发，这里只在包边界上插一行标题，
 * 光标、Enter 仍然数同一条扁平列表——分组是显示层的事，不该长出第二套选择逻辑。
 */
function commandPaletteMatches(query) {
    const text = (query || '').trim();
    const lowered = text.toLowerCase();
    const actions = commandPaletteActions().filter(function (action) {
        if (!lowered) return true;
        return (String(action.label || '') + ' ' + String(action.id || '')).toLowerCase().indexOf(lowered) >= 0;
    });
    if (text) {
        actions.push({ id: '__natural_language__', label: '按原话执行：' + text, request: text });
    }
    return actions;
}

/** 任务包标题只从后端下发的 packages 里取；缺了就当没有分组，绝不在前端另立一份包清单。 */
function commandPalettePackageTitle(packageId) {
    if (!packageId) return '';
    const packages = window.commandPaletteState && window.commandPaletteState.packages;
    if (!Array.isArray(packages)) return '';
    const hit = packages.filter(function (item) { return item && item.id === packageId; })[0];
    return hit ? String(hit.title || '') : '';
}

let commandPaletteCursor = 0;

function commandPaletteRender(query) {
    const list = document.getElementById('command-palette-list');
    if (!list) return;
    const items = commandPaletteMatches(query);
    commandPaletteCursor = Math.max(0, Math.min(commandPaletteCursor, items.length - 1));
    if (items.length === 0) commandPaletteCursor = 0;
    list.innerHTML = '';

    if (items.length === 0) {
        const empty = document.createElement('div');
        empty.className = 'command-palette-empty';
        empty.textContent = '没有匹配的快捷动作；继续输入这句话，按 Enter 交给 AI。';
        list.appendChild(empty);
        return;
    }

    let lastPackage = null;
    items.forEach(function (item, rank) {
        const packageId = String(item.packageId || '');
        if (packageId !== lastPackage) {
            lastPackage = packageId;
            const title = commandPalettePackageTitle(packageId);
            if (title) {
                const group = document.createElement('div');
                group.className = 'command-palette-group';
                group.textContent = title;
                list.appendChild(group);
            }
        }

        const row = document.createElement('div');
        row.className = 'command-palette-item' + (rank === commandPaletteCursor ? ' active' : '');
        row.textContent = item.label;
        row.title = item.request;
        row.onclick = function () { commandPaletteRun(item); };
        list.appendChild(row);
    });
}

function commandPaletteCurrentItem() {
    const items = commandPaletteMatches(
        (document.getElementById('command-palette-input') || {}).value || '');
    return items[commandPaletteCursor];
}

function commandPaletteKeydown(event) {
    const key = event.key || '';
    if (key === 'Escape') {
        event.preventDefault();
        commandPaletteClose();
        return;
    }
    if (key === 'ArrowDown' || key === 'ArrowUp') {
        event.preventDefault();
        const total = commandPaletteMatches(event.target.value).length;
        if (total === 0) return;
        commandPaletteCursor = key === 'ArrowDown'
            ? (commandPaletteCursor + 1) % total
            : (commandPaletteCursor - 1 + total) % total;
        commandPaletteRender(event.target.value);
        commandPaletteScrollCursorIntoView();
        return;
    }
    if (key === 'Enter') {
        event.preventDefault();
        const item = commandPaletteCurrentItem();
        if (item) commandPaletteRun(item);
    }
}

function commandPaletteScrollCursorIntoView() {
    const active = document.querySelector('#command-palette-list .command-palette-item.active');
    if (active && active.scrollIntoView) active.scrollIntoView({ block: 'nearest' });
}

/**
 * 唯一出口：一条 startAgent。命令面板自己不执行任何东西，也不判断风险——
 * 预览、风险确认、变更摘要和「撤销本轮」都在这条主链的下游。
 */
function commandPaletteRun(item) {
    const request = String(item.request || '').trim();
    commandPaletteClose();
    if (!request) return;
    // 入口身份跟着动作走，任务包断言才知道这一轮该按哪份成功标准收尾。
    // 「按原话执行」那条是自由文本，没有入口动作，必须留空而不是编一个 id。
    const entryActionId = item.id === '__natural_language__' ? '' : String(item.id || '');
    sendMessageToServer({ type: 'startAgent', request: request, entryActionId: entryActionId });
}
