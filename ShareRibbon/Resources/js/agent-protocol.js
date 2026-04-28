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
    REFINE_PLAN: 'agent:refinePlan',
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

// 导出到全局（兼容现有代码风格）
window.AgentMessage = AgentMessage;
window.sendAgentMessage = sendAgentMessage;
window.requestStartAgent = requestStartAgent;
window.requestAbortAgent = requestAbortAgent;
window.requestApprove = requestApprove;
window.requestReject = requestReject;
window.requestRefinePlan = requestRefinePlan;
