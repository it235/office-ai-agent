// E5-04 入口可达性门禁。
//
// 面板里「点了没反应」从来不是异常：ChatCommandRouter 对未注册的 type 只 Debug.WriteLine，
// 被注释掉的按钮、没人调用的 window.* 函数、查不到的 getElementById 同理全部静默失效。
// 所以这里把「入口」当成合同来核，十个方向都要闭合：
//   A 面板发出的 type 必须注册   B 注册的 type 必须有发送方
//   C HTML 注释里不留控件        D 内联事件的函数必须有定义
//   E getElementById 的 id 必须有节点   F vsto 桥方法必须有调用点
//   G Ribbon 上可见的按钮必须有非空 Click 处理
//   H VB 注入脚本里的字面量 id 必须有节点   I VB 注入脚本调用的 JS 函数必须有定义
//   J 导航文档（AGENTS.md/CLAUDE.md/README）引用的本地路径必须存在于磁盘
//
// 覆盖范围是被踩出来的：VB 生成的标记（SmartProofreadFocusMode 里的 id="proofread-summary"）和
// VB 注入的 vsto 桥（GetVstoBridgeScript 里的 type: 'checkedChange'）都算发送方/声明方，
// 只看 Resources 会得到相反的结论。
//
// H/I 是 E5-05 从 audit-p0-guardrails.ps1 的 P4-C 片段并轨过来的：那条只扫一份手写文件清单、
// 只看载荷第一个调用、并且把带 VB 插值的载荷整条跳过——而注入载荷绝大多数都带插值。
// 并轨后语料来自磁盘上四个模块的全部 .vb，载荷由 VB 词法决定而不是正则猜。

const fs = require('fs');
const path = require('path');

const repoRoot = path.resolve(__dirname, '..');
const failures = [];
const notes = [];

function read(...parts) {
    return fs.readFileSync(path.join(repoRoot, ...parts), 'utf8');
}

function walk(dir, pattern, out) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        if (entry.name === 'bin' || entry.name === 'obj' || entry.name === 'packages') continue;
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) walk(full, pattern, out);
        else if (pattern.test(entry.name)) out.push(full);
    }
    return out;
}

const rel = (p) => path.relative(repoRoot, p).replace(/\\/g, '/');
const lineOf = (src, index) => src.slice(0, index).split('\n').length;

// ---------- corpus ----------
const htmlPath = path.join(repoRoot, 'ShareRibbon', 'Resources', 'chat-template-refactored.html');
const html = read('ShareRibbon', 'Resources', 'chat-template-refactored.html');
const jsDir = path.join(repoRoot, 'ShareRibbon', 'Resources', 'js');
const jsFiles = fs.readdirSync(jsDir).filter(f => f.endsWith('.js')).sort()
    .map(f => ({ file: path.join(jsDir, f), src: fs.readFileSync(path.join(jsDir, f), 'utf8') }));
const jsCorpus = jsFiles.map(j => j.src).join('\n');

// VB 侧：router 注册 + 注入脚本 + 由 VB 拼出来的 HTML 标记
const vbFiles = walk(path.join(repoRoot, 'ShareRibbon'), /\.vb$/, [])
    .concat(walk(path.join(repoRoot, 'WordAi'), /\.vb$/, []))
    .concat(walk(path.join(repoRoot, 'ExcelAi'), /\.vb$/, []))
    .concat(walk(path.join(repoRoot, 'PowerPointAi'), /\.vb$/, []))
    .filter(f => !/\.Designer\.vb$/.test(f) && !/\\My Project\\/.test(f));
const vbSources = vbFiles.map(f => ({ file: f, src: fs.readFileSync(f, 'utf8') }));
const vbCorpus = vbSources.map(v => v.src).join('\n');

// ---------- 注入载荷扫描（规则 I）----------
// 注入 API 的实参在 VB 里常常是好几段字面量用 & 拼起来的，而且每段都带 {VB 表达式} 插值。
// 所以先按 VB 词法把实参表达式里的字符串字面量逐段取出来，再把插值区域挖掉，
// 剩下的才是真正会跑在页面里的那段 JS。用正则猜会在插值处断在半句上，把 VB 关键字当成 JS 调用。
// 已知边界：插值空洞里的字符串不需要 doubled quote，所以 `""` 判定偶尔把载荷读偏一格；
// 规则 I 因此按「解析得到的载荷」取下限，字面量 id 交给规则 H 全文扫，不受这个边界影响。
const injectApiRe = /\b(?:_executeScript|ExecuteJavaScriptAsyncJS|ExecuteJavaScriptAsync|ExecuteScriptAsync|ExecuteScript)\s*\(\s*/g;

// 从实参第一个字符开始读，字符串内部视为不透明（"" 是一个引号），读到该实参表达式的右括号为止。
function readVbArguments(src, start) {
    const literals = [];
    let depth = 1, i = start;
    while (i < src.length && depth > 0 && i - start < 6000) {
        const c = src[i];
        if (c === '"') {
            let j = i + 1, buf = '';
            while (j < src.length) {
                if (src[j] === '"') {
                    if (src[j + 1] === '"') { buf += '"'; j += 2; continue; }
                    j++;
                    break;
                }
                buf += src[j];
                j++;
            }
            literals.push(buf);
            i = j;
            continue;
        }
        if (c === '(') depth++;
        else if (c === ')') { depth--; if (depth === 0) break; }
        i++;
    }
    return literals;
}

// {VB 表达式} 可以嵌套，反复剥掉最内层直到没有成对花括号为止。
function stripVbInterpolation(text) {
    let prev;
    let cur = text;
    do {
        prev = cur;
        cur = cur.replace(/\{[^{}]*\}/g, ' ');
    } while (cur !== prev);
    return cur;
}

const injections = [];
for (const v of vbSources) {
    for (const m of v.src.matchAll(injectApiRe)) {
        const literals = readVbArguments(v.src, m.index + m[0].length);
        if (!literals.length) continue;
        injections.push({
            where: `${rel(v.file)}:${lineOf(v.src, m.index)}`,
            js: literals.map(stripVbInterpolation).join('\n'),
        });
    }
}
if (injections.length < 100) {
    failures.push(`只解析到 ${injections.length} 段注入载荷，是这条扫描自己失效了，不是桥变干净了。`);
}

// ---------- A: JS/HTML 发出的 type 必须被注册 ----------
const routerSrc = read('ShareRibbon', 'Controls', 'BaseChatControl.vb');
const registered = new Map();
for (const m of routerSrc.matchAll(/router\.Register\("([^"]+)"/g)) {
    registered.set(m[1], `BaseChatControl.vb:${lineOf(routerSrc, m.index)}`);
}

function collectLiteralTypes(src, out) {
    for (const m of src.matchAll(/\btype\s*:\s*['"]([A-Za-z][\w:]*)['"]/g)) {
        if (!out.has(m[1])) out.set(m[1], []);
        if (out.get(m[1]).length < 2) out.get(m[1]).push(lineOf(src, m.index));
    }
}

// VB 注入的 vsto 桥：函数体里写死 type，只有该函数真的被面板调用时才算一个发送方。
const bridgeSrc = (() => {
    const src = read('ShareRibbon', 'Controls', 'Services', 'UtilsService.vb');
    const start = src.indexOf('GetVstoBridgeScript');
    return start < 0 ? '' : src.slice(start, src.indexOf('End Function', start));
})();
const bridgeSenders = [];
for (const m of bridgeSrc.matchAll(/(\w+)\s*:\s*function[\s\S]{0,400}?type:\s*'([A-Za-z][\w:]*)'/g)) {
    if (new RegExp('\\bvsto\\.' + m[1] + '\\s*\\(').test(jsCorpus)) bridgeSenders.push({ type: m[2], fn: 'vsto.' + m[1] });
}

const sentFromPanel = new Map();
collectLiteralTypes(jsCorpus, sentFromPanel);
collectLiteralTypes(html, sentFromPanel);
for (const b of bridgeSenders) {
    if (!sentFromPanel.has(b.type)) sentFromPanel.set(b.type, []);
    sentFromPanel.get(b.type).push(`vsto 桥 ${b.fn}`);
}

for (const [type, lines] of [...sentFromPanel].sort()) {
    if (!registered.has(type)) {
        failures.push(`面板发送了 type "${type}"，但 router 没有注册，消息会被静默丢弃：` +
            lines.map(l => `Resources ${l}`).join(', '));
    }
}
notes.push(`消息 type：面板发送 ${sentFromPanel.size} 个，router 注册 ${registered.size} 个`);

// ---------- B: 注册的 type 必须有发送方 ----------
for (const [type, where] of [...registered].sort()) {
    const hasLiteral = new RegExp("['\"]" + type.replace(/[$()*+.?[\\\]^{|}]/g, '\\$&') + "['\"]")
        .test(jsCorpus + html);
    if (hasLiteral) continue;
    failures.push(`router 注册了 "${type}"（${where}），但面板里没有任何发送方——这是一条走不到的入口。`);
}

// ---------- C: HTML 注释里不得留可交互控件 ----------
for (const m of html.matchAll(/<!--[\s\S]*?-->/g)) {
    if (/<\s*(button|input|select|textarea|a)\b/i.test(m[0])) {
        failures.push(`chat-template-refactored.html:${lineOf(html, m.index)} 注释里藏着可交互控件 ` +
            `(${m[0].match(/<\s*(button|input|select|textarea|a)\b/i)[1]}): 入口要么接通要么删掉，不留「只有开发者知道」的按钮。`);
    }
}

// ---------- D: 内联事件引用的函数必须有定义 ----------
const handlerRefs = [];
for (const [src, label] of [[html, rel(htmlPath)], ...jsFiles.map(j => [j.src, rel(j.file)])]) {
    for (const m of src.matchAll(/\son[a-z]+\s*=\s*["'`]?\s*([A-Za-z_$][\w$]*)\s*\(/g)) {
        handlerRefs.push({ fn: m[1], where: `${label}:${lineOf(src, m.index)}` });
    }
}
const definedFns = new Set();
for (const name of ['window', 'globalThis', 'self']) {
    for (const m of jsCorpus.matchAll(new RegExp(name + '\\.([A-Za-z_$][\\w$]*)\\s*=', 'g'))) definedFns.add(m[1]);
}
for (const m of jsCorpus.matchAll(/(?:function\s+|const\s+|let\s+|var\s+)([A-Za-z_$][\w$]*)\s*(?:\(|=)/g)) definedFns.add(m[1]);
let missingHandlers = 0;
for (const ref of handlerRefs) {
    if (definedFns.has(ref.fn)) continue;
    missingHandlers++;
    failures.push(`${ref.where} 内联事件调用 ${ref.fn}()，但 JS 里没有这个函数——点了不会有任何反应。`);
}
notes.push(`内联事件：${handlerRefs.length} 处引用、${new Set(handlerRefs.map(r => r.fn)).size} 个函数名，${handlerRefs.length - missingHandlers} 处解析到定义`);

// ---------- E: getElementById 的字面量 id 必须有声明 ----------
const declaredIds = new Set();
for (const m of html.matchAll(/\sid="([^"]+)"/g)) declaredIds.add(m[1]);
for (const m of jsCorpus.matchAll(/\bid\s*=\s*['"`]([A-Za-z][\w-]*)['"`]/g)) declaredIds.add(m[1]);
for (const m of jsCorpus.matchAll(/\.id\s*=\s*['"]([A-Za-z][\w-]*)['"]/g)) declaredIds.add(m[1]);
// VB 也会拼 HTML 片段（例如校对面板的摘要块），这些 id 同样是真节点
for (const m of vbCorpus.matchAll(/id\s*=\s*""([A-Za-z][\w-]*)""/g)) declaredIds.add(m[1]);
for (const m of vbCorpus.matchAll(/id="([A-Za-z][\w-]*)"/g)) declaredIds.add(m[1]);

const lookedUp = new Map();
for (const j of jsFiles) {
    for (const m of j.src.matchAll(/getElementById\(\s*['"]([^'"$]{2,})['"]\s*\)/g)) {
        if (!lookedUp.has(m[1])) lookedUp.set(m[1], `${rel(j.file)}:${lineOf(j.src, m.index)}`);
    }
}
for (const m of html.matchAll(/getElementById\(\s*['"]([^'"$]{2,})['"]\s*\)/g)) {
    if (!lookedUp.has(m[1])) lookedUp.set(m[1], `${rel(htmlPath)}:${lineOf(html, m.index)}`);
}
let orphanIds = 0;
for (const [id, where] of [...lookedUp].sort()) {
    if (declaredIds.has(id)) continue;
    orphanIds++;
    failures.push(`${where} 查找 id "${id}"，但 HTML/JS/VB 标记里都没有这个节点，代码会拿到 null 并静默跳过。`);
}
notes.push(`DOM id：${lookedUp.size} 个字面量查询，${lookedUp.size - orphanIds} 个能解析到节点`);

// ---------- F: 注入的 vsto 桥方法必须真的被调用 ----------
// 桥里的每个方法都往一个消息 type 上写；方法没人调，等于 type 也没人发，是同一类静默失效。
const vstoCalls = new Set();
for (const [src] of [[jsCorpus], [html]]) {
    for (const m of src.matchAll(/(?:window\s*\.\s*)?\bvsto\s*\.\s*([A-Za-z_$][\w$]*)\s*\(/g)) vstoCalls.add(m[1]);
}
for (const m of bridgeSrc.matchAll(/([A-Za-z_$][\w$]*)\s*:\s*function/g)) {
    if (!vstoCalls.has(m[1])) {
        failures.push(`vsto 桥注入了 ${m[1]}()，面板里没有任何调用点；它 posts 的 type 因此也永远发不出去。`);
    }
}

// ---------- G: Ribbon 可见按钮必须有非空 Click 处理 ----------
// 每个宿主都从同一个 Designer 继承整条 Ribbon，靠宿主自己的 Designer 把不适用项设成
// Visible=False。这条规则钉的是「看不见的按钮不该再留着能点的实现」与「看得见的按钮必须真的会动」：
// 空处理体（只有注释）在编译和运行期都不报错，是典型「点了没反应」。
const hosts = ['WordAi', 'ExcelAi', 'PowerPointAi'];
const baseRibbonDesigner = read('ShareRibbon', 'Ribbon', 'BaseOfficeRibbon.Designer.vb');
const baseRibbon = read('ShareRibbon', 'Ribbon', 'BaseOfficeRibbon.vb');
const labels = new Map();
for (const m of baseRibbonDesigner.matchAll(/Me\.([A-Za-z0-9_]+Button)\.Label = "([^"]+)"/g)) labels.set(m[1], m[2]);
const baseHidden = new Set();
for (const m of baseRibbonDesigner.matchAll(/Me\.([A-Za-z0-9_]+Button)\.Visible = (False)/g)) baseHidden.add(m[1]);

function emptyClickHandlers(src) {
    const empty = [];
    for (const m of src.matchAll(/Sub\s+([A-Za-z0-9_]+)_Click[\s\S]*?\r?\n {4}End Sub/g)) {
        const stmts = m[0].split('\n').slice(1, -1).map(l => l.trim()).filter(l => l && !l.startsWith("'"));
        if (stmts.length === 0) empty.push(m[1]);
    }
    return empty;
}

const handledControls = new Set();
for (const m of (baseRibbon + '\n' + hosts.map(h => read(h, 'Ribbon1.vb')).join('\n'))
    .matchAll(/Handles\s+([A-Za-z0-9_]+)\.Click/g)) handledControls.add(m[1]);

for (const host of hosts) {
    const designer = read(host, 'Ribbon1.Designer.vb');
    const visible = new Map();
    for (const [btn, label] of labels) {
        const own = new RegExp('Me\\.' + btn + '\\.Visible = (True|False)').exec(designer);
        if (!(own ? own[1] === 'False' : baseHidden.has(btn))) visible.set(btn, label);
    }
    for (const [btn, label] of visible) {
        if (!handledControls.has(btn)) {
            failures.push(`${host} 的 Ribbon 上「${label}」(${btn}) 可见，但没有任何 Click 处理——点了不会有任何反应。`);
        }
    }
    const hostSrc = read(host, 'Ribbon1.vb');
    // 宿主覆盖 + 共享基类里的空实现都算，因为基类空实现等于三端都点不动
    for (const name of [...emptyClickHandlers(hostSrc), ...emptyClickHandlers(baseRibbon)]) {
        if (visible.has(name)) {
            failures.push(`${host} 的 Ribbon 上「${visible.get(name)}」(${name}) 可见，但 ${name}_Click 只剩注释——点了不会有任何反应。`);
        }
    }
    notes.push(`Ribbon 入口：${host} 可见 ${visible.size}/${labels.size}`);
}

// ---------- H: VB 侧的字面量 id 必须有节点 ----------
// 注入块是顺序执行的一整段：getElementById 拿到 null 后取属性就抛，其后每一行都不再执行，
// 症状是「设置项同步静默失效」而不是报错。这条按全文扫（去掉整行注释），因为大部分注入载荷
// 是先 `Dim js As String = ...` 再传给注入 API 的，锚在调用点上会整个漏掉；而 .vb 里出现
// getElementById('x') 只有一个理由：那是一段要跑在页面里的 JS。
// id 只在字面量整段时才谈存在不存在，'chat-' + uuid 这类拼接查找属于动态 id，直接跳过。
const idLookupPatterns = [
    /getElementById\(\s*'([^']*)'\s*\)/g,
    /querySelector\(\s*'#([^'\s]*)'/g,
    /setNativeEntrySetting\(\s*'([^']*)'/g,
];
const vbIdLookups = new Map();
for (const v of vbSources) {
    const code = v.src.split('\n').filter(l => !/^\s*'/.test(l)).join('\n');
    for (const pattern of idLookupPatterns) {
        for (const m of code.matchAll(pattern)) {
            if (!/^[A-Za-z][\w-]*$/.test(m[1])) continue;
            if (!vbIdLookups.has(m[1])) vbIdLookups.set(m[1], `${rel(v.file)}:${lineOf(code, m.index)}`);
        }
    }
}
for (const [id, where] of [...vbIdLookups].sort()) {
    if (declaredIds.has(id)) continue;
    failures.push(`${where} 注入的脚本查找 id "${id}"，但面板 HTML/JS/VB 标记里都没有这个节点：` +
        `取属性会抛 TypeError，这段注入的后半句从此不再执行。`);
}
notes.push(`注入 id：VB 语料 ${vbIdLookups.size} 个字面量查询，${[...vbIdLookups.keys()].filter(i => declaredIds.has(i)).length} 个能解析到节点`);

// ---------- I: VB 注入脚本调用的 JS 函数必须有定义 ----------
// 未定义的调用在页面控制台里是一句 ReferenceError，用户只看到「点了没反应」。
// 面板自己调用同名函数说明名字是真的，所以只看 VB 侧载荷；typeof 守卫过的可选钩子除外。
const jsGlobals = new Set();
for (const src of [jsCorpus, html]) {
    for (const m of src.matchAll(/function\s+([A-Za-z_$][\w$]*)\s*\(/g)) jsGlobals.add(m[1]);
    for (const m of src.matchAll(/(?:window|globalThis|self)\.([A-Za-z_$][\w$]*)\s*=/g)) jsGlobals.add(m[1]);
    for (const m of src.matchAll(/^\s*([A-Za-z_$][\w$]*)\s*:\s*function/gm)) jsGlobals.add(m[1]);
    for (const m of src.matchAll(/(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:function|\(|async)/g)) jsGlobals.add(m[1]);
}
if (jsGlobals.size < 100) {
    failures.push(`面板语料只解析到 ${jsGlobals.size} 个 JS 函数名，是这条规则自己失效了，不是桥没问题。`);
}
// 浏览器 API 与 VB 关键字：后者只在插值没剥干净的残句里冒出来，留着比让规则瞎掉便宜。
const callSkip = new Set([
    'document', 'window', 'console', 'localStorage', 'sessionStorage', 'JSON', 'Object', 'Array', 'String',
    'Number', 'Boolean', 'Date', 'Math', 'Promise', 'RegExp', 'Error', 'Event', 'CustomEvent', 'URL',
    'setTimeout', 'setInterval', 'clearTimeout', 'clearInterval', 'requestAnimationFrame', 'getComputedStyle',
    'alert', 'confirm', 'parseInt', 'parseFloat', 'isNaN', 'encodeURIComponent', 'decodeURIComponent', 'url',
    'eval', 'atob', 'btoa', 'Blob', 'Image', 'Audio', 'Worker', 'navigator', 'location', 'history',
    'if', 'for', 'while', 'switch', 'return', 'function', 'new', 'typeof', 'catch', 'else', 'then',
    'var', 'let', 'const', 'do', 'try', 'in', 'of', 'await',
    'If', 'Then', 'Else', 'ElseIf', 'End', 'Not', 'And', 'Or', 'AndAlso', 'OrElse', 'Select', 'Case',
    'Each', 'With', 'Sub', 'True', 'False', 'Nothing', 'As', 'Is', 'IsNot',
]);
const vbBridgeCalls = new Map();
const vbCallNames = new Set();
for (const inj of injections) {
    const guarded = new Set();
    for (const m of inj.js.matchAll(/typeof\s+([A-Za-z_$][\w$]*)\s*[!=]==?\s*['"](?:function|undefined)/g)) guarded.add(m[1]);
    for (const m of inj.js.matchAll(/(?:^|[\s;})(,])([A-Za-z_$][\w$]*)\s*\(/g)) {
        const name = m[1];
        vbCallNames.add(name);
        if (callSkip.has(name) || jsGlobals.has(name) || guarded.has(name)) continue;
        if (!vbBridgeCalls.has(name)) vbBridgeCalls.set(name, inj.where);
    }
}
for (const [name, where] of [...vbBridgeCalls].sort()) {
    failures.push(`${where} 注入的脚本调用 ${name}()，但面板语料里没有这个函数——ReferenceError 只会出现在页面控制台里。`);
}
notes.push(`注入调用：${injections.length} 段载荷、${vbCallNames.size} 个调用名，${vbCallNames.size - vbBridgeCalls.size} 个解析到定义或被 typeof 守卫`);

// ---------- 规则 J：导航文档引用的路径必须存在 ----------
// AGENTS.md 是 Agent 的地图，地图上的路径坏掉时没有任何检查会响：本轮实测根 AGENTS.md 指向
// docs/build-and-installer.md 与 docs/signing-and-certificates.md，两个文件从未存在，
// CLAUDE.md 的 docs/VisualStudio调试问题诊断.md 同样是空的——症状不是报错，是下次每个人
// （或每个 Agent）重新考古一遍。历史文档（openspec 记录、归档、.omc 研究笔记）允许引用当时的路径，
// 所以语料只取导航层：根三份 + 每个目录最近的 AGENTS.md + Storage 的 README。
// 只核两种能唯一解析的写法（markdown 链接、以顶层目录开头的反引号路径）；正文里随手引用的裸文件名
// （`agent-card.js`）和模块内相对写法（`Storage/AppTypeContract.vb`）不参与判定——把它们算成引用
// 会得到 13 条假阳性，而一条 noisy 的门禁最后只会被人关掉。
const navIgnoreDirs = new Set(['.git', 'bin', 'obj', 'packages', 'node_modules', '.vs', 'openspec', '.omc', '.qoder', 'docs']);
const navExt = /\.(md|ps1|js|vb|yml|sql|html|css|bat|vbproj|sln)$/;
function collectNavDocs(dir, out) {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
        if (navIgnoreDirs.has(e.name) && dir === repoRoot) continue;
        const abs = path.join(dir, e.name);
        if (e.isDirectory()) collectNavDocs(abs, out);
        else if (e.name === 'AGENTS.md') out.push(abs);
    }
    return out;
}
const navDocs = collectNavDocs(repoRoot, []);
for (const extra of ['AGENTS.md', 'CLAUDE.md', 'README.md', 'ShareRibbon/Storage/README_DATABASE.md']) {
    const abs = path.join(repoRoot, ...extra.split('/'));
    if (fs.existsSync(abs) && !navDocs.includes(abs)) navDocs.push(abs);
}
const navRefs = new Map();
for (const abs of navDocs) {
    const rel = path.relative(repoRoot, abs).split(path.sep).join('/');
    const text = fs.readFileSync(abs, 'utf8');
    const add = (target, baseDir) => {
        if (!target || target.startsWith('/') || /^[a-zA-Z]:/.test(target)) return;
        if (/[*<>]|NNN|\.\.\./.test(target)) return;
        if (!navExt.test(target)) return;
        const key = rel + ' -> ' + target;
        if (navRefs.has(key)) return;
        // markdown 链接按文档目录解析、反引号路径按仓库根解析，两种都再试一次另一个基准：
        // 导航文档里这两种写法混用，把它们统一成一种是一次没有必要的重写。
        const hits = [path.resolve(baseDir, target), path.resolve(repoRoot, target)].filter(p => fs.existsSync(p));
        navRefs.set(key, hits.length > 0);
    };
    for (const m of text.matchAll(/\]\((?!https?:|mailto:|#)([^)#\s]+)(#[^)]*)?\)/g)) {
        add(m[1].replace(/^file:\/\/\//, ''), path.dirname(abs));
    }
    for (const m of text.matchAll(/`((?:docs|scripts|build|openspec|ShareRibbon|WordAi|ExcelAi|PowerPointAi|\.github)\/[A-Za-z0-9_.\/+-]+)`/g)) {
        add(m[1], repoRoot);
    }
}
if (navRefs.size < 30) {
    failures.push(`导航合同只解析到 ${navRefs.size} 条引用，是这条规则自己失效了，不是地图变干净了。`);
}
for (const [key, ok] of [...navRefs].sort()) {
    if (!ok) failures.push(`${key}：导航文档指向一个不存在的文件。`);
}
notes.push(`导航路径：${navDocs.length} 份导航文档、${navRefs.size} 条本地引用，${[...navRefs.values()].filter(Boolean).length} 条能解析到磁盘`);

// ---------- 输出 ----------
for (const n of notes) console.log('  ' + n);
console.log(`  vsto 桥计入的发送方：${bridgeSenders.map(b => b.type + '(' + b.fn + ')').join(', ') || '无'}`);
if (failures.length) {
    console.error(`\n入口可达性检查失败（${failures.length} 项）：`);
    for (const f of failures) console.error('  - ' + f);
    process.exit(1);
}
console.log('入口可达性检查通过：消息 type、内联事件、DOM id、vsto 桥方法、Ribbon 可见按钮、注入载荷的 id 与调用、导航文档路径十个方向都闭合。');
