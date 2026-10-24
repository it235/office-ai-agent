/**
 * onboarding.js - 首次使用引导（E3-01）
 *
 * 只剩两步：隐私确认 -> 示例任务。选服务商、填 Key、测连接都用既有配置窗口（openApiConfigForm ->
 * ConfigApiForm），这里不再实现第二套配置逻辑，也不在前端保存步骤状态：向导走到哪一步只由 VB 推来的状态决定。
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
    view.appendChild(onboardingStep(state) === 'privacy' ? onboardingPrivacy() : onboardingExample(state));
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

function onboardingPrivacy() {
    const list = document.createElement('ul');
    list.className = 'onboarding-privacy';
    [
        '文档内容、选区和会话会发送给你所选的模型接口，用于生成结果。',
        '会话历史、变更日志和记忆只保存在本机 SQLite，不上传统计服务。',
        'API Key 只写入当前用户的加密存储，配置文件里只保留占位符。'
    ].forEach(function (text) {
        const li = document.createElement('li');
        li.textContent = text;
        list.appendChild(li);
    });
    return onboardingStepNode('1. 隐私确认', [list], onboardingActions([
        onboardingButton('同意并继续', 'agent-btn-execute', function () {
            sendMessageToServer({ type: 'onboardingSetPrivacy', accepted: 'true' });
        })
    ]));
}

function onboardingExample(state) {
    const rows = [];
    if (state.needsSetup) {
        rows.push(onboardingHint('还没有可用的模型配置，先配好模型服务，示例才能真的跑起来。'));
        rows.push(onboardingConfigButton());
    } else {
        rows.push(onboardingHint('点一个例子填进输入框，或直接写你自己的需求。'));
    }
    (state.examples || []).forEach(function (text) {
        const item = onboardingButton(text, 'onboarding-example', function () {
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
})();
