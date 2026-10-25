/**
 * onboarding.js - 首次使用引导（E3-01 / E3-02）
 *
 * 只剩两步：隐私确认 -> 示例任务。选服务商、填 Key、测连接都用既有配置窗口（openApiConfigForm ->
 * ConfigApiForm），这里不再实现第二套配置逻辑，也不在前端保存步骤状态：向导走到哪一步只由 VB 推来的状态决定。
 * 示例任务的话术全部来自 QuickActionCatalog：带 request 的那条是一键示例，点了就交给 startAgent 主链。
 */

const ONBOARDING_STEPS = ['privacy', 'example'];
const ONBOARDING_STEP_LABELS = {
    privacy: '隐私确认',
    example: '试一个任务'
};

window.onboardingState = null;

function onboardingCard(create) {
    const existing = document.getElementById('onboarding-card');
    if (existing || !create) return existing;
    const container = document.getElementById('chat-container');
    if (!container) return null;
    const card = document.createElement('div');
    card.id = 'onboarding-card';
    card.className = 'onboarding-card';
    container.appendChild(card);
    return card;
}

function onboardingStep(state) {
    return ONBOARDING_STEPS.indexOf(state.step) >= 0 ? state.step : 'example';
}

function onboardingClose() {
    const card = onboardingCard(false);
    if (card) card.remove();
}

window.renderOnboarding = function (state) {
    if (state) window.onboardingState = state;
    state = window.onboardingState || {};
    // 设置面板里那一行和向导用的是同一份状态：同意状态、边界文案都只从宿主这一份 payload 来。
    syncPrivacyConsentRow(state);

    const existing = onboardingCard(false);
    const isAuto = existing && existing.dataset.pinned !== 'true';
    // 隐私确认过、模型也配好了，就没有引导可讲；用户手动点开的卡片照常刷新。
    if (onboardingStep(state) === 'example' && !state.needsSetup && isAuto) {
        existing.remove();
        return;
    }

    const view = onboardingCard(true);
    if (!view) return;
    view.innerHTML = '';
    view.appendChild(onboardingHeader(state));
    view.appendChild(onboardingStep(state) === 'privacy' ? onboardingPrivacy(state) : onboardingExample(state));
};

function onboardingHeader(state) {
    const header = document.createElement('div');
    header.className = 'onboarding-header';

    const title = document.createElement('div');
    title.className = 'onboarding-title';
    title.textContent = '开始使用 Office AI';
    header.appendChild(title);

    const dots = document.createElement('div');
    dots.className = 'onboarding-dots';
    const current = ONBOARDING_STEPS.indexOf(onboardingStep(state));
    ONBOARDING_STEPS.forEach(function (step, rank) {
        const dot = document.createElement('span');
        dot.className = 'onboarding-dot' + (rank === current ? ' active' : '') + (rank < current ? ' done' : '');
        dot.title = ONBOARDING_STEP_LABELS[step];
        dots.appendChild(dot);
    });
    header.appendChild(dots);

    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'onboarding-close';
    close.setAttribute('aria-label', '关闭引导');
    close.textContent = '×';
    close.addEventListener('click', onboardingClose);
    header.appendChild(close);
    return header;
}

function onboardingButton(label, cls, handler) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'agent-btn ' + (cls || 'agent-btn-execute');
    btn.textContent = label;
    if (handler) btn.addEventListener('click', handler);
    return btn;
}

function onboardingConfigButton() {
    return onboardingButton('打开配置窗口', 'agent-btn-approve-session', function () {
        if (typeof openModelConfig === 'function') openModelConfig();
    });
}

function onboardingActions(extra) {
    const row = document.createElement('div');
    row.className = 'onboarding-actions';
    (extra || []).forEach(function (btn) { row.appendChild(btn); });

    const skip = document.createElement('button');
    skip.type = 'button';
    skip.className = 'onboarding-skip';
    skip.textContent = '稍后再说';
    skip.addEventListener('click', onboardingClose);
    row.appendChild(skip);
    return row;
}

function onboardingStepNode(title, bodyNodes, actions) {
    const wrap = document.createElement('div');
    wrap.className = 'onboarding-step';

    const heading = document.createElement('div');
    heading.className = 'onboarding-step-title';
    heading.textContent = title;
    wrap.appendChild(heading);

    (bodyNodes || []).forEach(function (node) { wrap.appendChild(node); });
    if (actions) wrap.appendChild(actions);
    return wrap;
}

function onboardingHint(text) {
    const hint = document.createElement('div');
    hint.className = 'onboarding-hint';
    hint.textContent = text;
    return hint;
}

function onboardingPrivacy(state) {
    const list = onboardingPrivacyList(state, 'onboarding-privacy');
    return onboardingStepNode('1. 隐私确认', [list], onboardingActions([
        onboardingButton('同意并继续', 'agent-btn-execute', function () {
            setPrivacyConsent(true);
        })
    ]));
}

/**
 * 数据边界三条说明：文案由宿主（PrivacyConsentGate）下发，前端只渲染。
 * 在这里抄一份是最坏的那种副本——用户看到的承诺会和实际发送的内容各自漂移。
 */
function onboardingPrivacyList(state, className) {
    const list = document.createElement('ul');
    list.className = className;
    ((state && state.privacyBoundary) || []).forEach(function (text) {
        const li = document.createElement('li');
        li.textContent = text;
        list.appendChild(li);
    });
    return list;
}

function setPrivacyConsent(accepted) {
    sendMessageToServer({ type: 'onboardingSetPrivacy', accepted: accepted ? 'true' : 'false' });
}

/**
 * 设置面板中的「数据边界同意」行（E3-05）：勾选状态与边界文案都只由宿主推来的状态决定，
 * 点一下就把选择送回去，宿主写完立刻回推真实状态——用户看到的永远是盘上的那一份。
 * 宿主还没给过状态时保持 disabled，不显示成一个能点的默认值。
 */
function syncPrivacyConsentRow(state) {
    const box = document.getElementById('settings-privacy-consent');
    if (!box) return;
    const known = !!(state && typeof state.privacyAccepted === 'boolean');
    box.disabled = !known;
    box.checked = !!(state && state.privacyAccepted === true);

    const list = document.getElementById('settings-privacy-boundary');
    if (!list) return;
    while (list.firstChild) list.removeChild(list.firstChild);
    ((state && state.privacyBoundary) || []).forEach(function (text) {
        const li = document.createElement('li');
        li.textContent = text;
        list.appendChild(li);
    });
}

function onboardingExample(state) {
    const rows = [];
    if (state.needsSetup) {
        rows.push(onboardingHint('还没有可用的模型配置，先配好模型服务，示例才能真的跑起来。'));
        rows.push(onboardingConfigButton());
    } else {
        rows.push(onboardingHint('一键示例直接把任务交给 Agent 跑一次；其余例子填进输入框，也可以自己写需求。'));
    }
    (state.examples || []).forEach(function (sample) {
        const text = (sample && sample.text) || '';
        if (!text) return;
        const request = (sample && sample.request) || '';
        const item = request
            ? onboardingButton(text, 'onboarding-demo', function () { onboardingRunSample(request); })
            : onboardingButton(text, 'onboarding-example', function () {
                onboardingFillInput(text);
                sendMessageToServer({ type: 'onboardingFinish' });
            });
        item.classList.add('onboarding-example-row');
        rows.push(item);
    });
    return onboardingStepNode('2. 试一个任务', rows, onboardingActions([
        onboardingButton('完成', 'agent-btn-execute', function () {
            sendMessageToServer({ type: 'onboardingFinish' });
            onboardingClose();
        })
    ]));
}

/**
 * 一键示例（E3-02）：把目录里带来的请求原话交给 startAgent，与面板输入同一个闸门、
 * 同一条主链。这里不判断模型是否配好——发送闸门是唯一权威，没配好就回落到配置 CTA，
 * 免得向导里再养一份配置判断。
 */
function onboardingRunSample(request) {
    sendMessageToServer({ type: 'startAgent', request: request });
    sendMessageToServer({ type: 'onboardingFinish' });
    onboardingClose();
}

function onboardingFillInput(text) {
    const smartInput = document.getElementById('smart-input');
    const chatInput = document.getElementById('chat-input');
    if (smartInput) {
        smartInput.innerText = text;
        if (typeof syncToHiddenTextarea === 'function') syncToHiddenTextarea();
        smartInput.focus();
    } else if (chatInput) {
        chatInput.value = text;
        chatInput.focus();
    }
}

window.showOnboarding = function () {
    const card = onboardingCard(true);
    if (!card) return;
    card.dataset.pinned = 'true';
    sendMessageToServer({ type: 'getOnboardingState' });
};

(function initOnboarding() {
    // 启动时先问一次配置状态：没配好才画向导，配好的老用户不会被打扰
    setTimeout(function () {
        if (typeof sendMessageToServer === 'function') {
            sendMessageToServer({ type: 'getOnboardingState' });
        }
    }, 800);

    // E3-05：设置面板那一行的勾选直接写回宿主。绑定只在页面装载时做一次，
    // 值本身永远由宿主回推的 renderOnboarding 决定，所以这里不保存任何本地状态。
    const consentBox = document.getElementById('settings-privacy-consent');
    if (consentBox) {
        consentBox.addEventListener('change', function () {
            setPrivacyConsent(consentBox.checked === true);
        });
    }
})();
