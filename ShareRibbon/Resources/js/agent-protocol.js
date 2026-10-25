/**
 * agent-protocol.js - 统一 Agent 前后端类型化消息协议
 * 替代松散 JSON，提供结构化消息常量与发送辅助函数
 */

const AgentMessage = {
    // 后端 → 前端
    PLAN_GENERATED: 'agent:planGenerated',
    ITERATION_UPDATE: 'agent:iterationUpdate',
    STEP_COMPLETED: 'agent:stepCompleted',
    STEP_FAILED: 'agent:stepFailed',
    STATUS_CHANGED: 'agent:statusChanged',
    APPROVAL_REQUEST: 'agent:approvalRequest',
    AGENT_COMPLETED: 'agent:completed',

    // 前端 → 后端
    START_AGENT: 'startAgent',
    ABORT_AGENT: 'abortAgent',
    APPROVE_PLAN: 'agent:approvePlan',
    REJECT_PLAN: 'agent:rejectPlan',
    APPROVE_STEP: 'agent:approveStep',
    APPROVE_TOOL_SESSION: 'agent:approveToolSession',
    REFINE_PLAN: 'agent:refinePlan',
    STEP_ACTION: 'agent:stepAction',
    EDIT_STEP: 'agent:editStep',
};

/**
 * 发送 Agent 消息到后端
 * @param {string} type - 消息类型（使用 AgentMessage 常量）
 * @param {object} payload - 消息载荷
 */
function sendAgentMessage(type, payload) {
    if (typeof sendMessageToServer === 'function') {
        sendMessageToServer({
            type: type,
            payload: payload || {}
        });
    } else {
        console.error('[AgentProtocol] sendMessageToServer 不可用');
    }
}

/**
 * 向后端请求启动统一 Agent
 * @param {string} request - 用户请求内容
 * @param {string[]} filePaths - 引用的文件路径（可选）
 * @param {object[]} selectedContent - 引用的选中内容（可选）
 */
function requestStartAgent(request, filePaths, selectedContent) {
    sendAgentMessage(AgentMessage.START_AGENT, {
        request: request,
        filePaths: filePaths || [],
        selectedContent: selectedContent || []
    });
}

/**
 * 向后端请求终止 Agent
 */
function requestAbortAgent() {
    sendAgentMessage(AgentMessage.ABORT_AGENT, {});
}

/**
 * 用户批准当前计划或步骤
 * @param {string} sessionId - 会话 ID
 */
function requestApprove(sessionId) {
    sendAgentMessage(AgentMessage.APPROVE_PLAN, { sessionId: sessionId });
}

/**
 * 用户拒绝当前计划或步骤
 * @param {string} sessionId - 会话 ID
 */
function requestReject(sessionId) {
    sendAgentMessage(AgentMessage.REJECT_PLAN, { sessionId: sessionId });
}

/**
 * 用户允许本会话内该工具的所有调用
 * @param {string} sessionId - 会话 ID
 * @param {string} toolId - 工具 ID（用于 UI 回显，后端按当前待决审批裁决）
 */
function requestApproveToolSession(sessionId, toolId) {
    sendAgentMessage(AgentMessage.APPROVE_TOOL_SESSION, {
        sessionId: sessionId,
        toolId: toolId || ''
    });
}

/**
 * 用户请求修改计划
 * @param {string} sessionId - 会话 ID
 * @param {string} feedback - 修改意见
 */
function requestRefinePlan(sessionId, feedback) {
    sendAgentMessage(AgentMessage.REFINE_PLAN, {
        sessionId: sessionId,
        feedback: feedback
    });
}

/**
 * 用户对单个步骤下达纠偏动作。
 * 走独立通道而不是 agent:approveStep：后者承载的是安全审批的超时/放弃语义，
 * 纠偏没有"待决"状态，也不允许被审批超时判成中止。
 * @param {string} sessionId - 会话 ID
 * @param {number} stepIndex - 步骤序号（0 基，与卡片渲染顺序一致）
 * @param {string} action - retry / skip，其余取值由宿主判为无法识别
 */
function requestStepAction(sessionId, stepIndex, action) {
    sendAgentMessage(AgentMessage.STEP_ACTION, {
        sessionId: sessionId,
        stepIndex: stepIndex,
        action: action
    });
}

/**
 * 用户改写单个步骤的说明与参数（E10-02）。
 * 与纠偏分开一条通道：载荷不同，回执也必须是"保存成功/失败 + 原因"，而不是按钮的启用状态。
 * @param {string} sessionId - 会话 ID
 * @param {number} stepIndex - 步骤序号（0 基）
 * @param {string} description - 改写后的步骤说明
 * @param {string} code - 改写后的 {"command":"工具ID","params":{...}} JSON 文本
 */
function requestStepEdit(sessionId, stepIndex, description, code) {
    sendAgentMessage(AgentMessage.EDIT_STEP, {
        sessionId: sessionId,
        stepIndex: stepIndex,
        description: description,
        code: code
    });
}

// 导出到全局（兼容现有代码风格）
window.AgentMessage = AgentMessage;
window.sendAgentMessage = sendAgentMessage;
window.requestStartAgent = requestStartAgent;
window.requestAbortAgent = requestAbortAgent;
window.requestApprove = requestApprove;
window.requestApproveToolSession = requestApproveToolSession;
window.requestReject = requestReject;
window.requestRefinePlan = requestRefinePlan;
window.requestStepAction = requestStepAction;
window.requestStepEdit = requestStepEdit;
