/**
 * mcp-manager.js - MCP Connection Management
 * Handles MCP (Model Context Protocol) connection display and configuration
 */

// MCP state variables
let mcpConnections = [];
let enabledMcpList = [];
let mcpSupported = false;

// Toggle MCP dialog visibility
function toggleMcpDialog() {
    requestMcpConnections();
    document.getElementById('mcp-dialog').style.display = 'block';
    document.getElementById('mcp-overlay').style.display = 'block';
}

// Close MCP dialog
function closeMcpDialog() {
    document.getElementById('mcp-dialog').style.display = 'none';
    document.getElementById('mcp-overlay').style.display = 'none';
}

// Request MCP connections from backend
function requestMcpConnections() {
    sendMessageToServer({
        type: 'getMcpConnections'
    });
}

// Render MCP connections list
function renderMcpConnections(connections, enabledList, supported) {
    mcpSupported = supported;

    // Show or hide MCP button
    const mcpButton = document.getElementById('mcp-toggle-btn');
    mcpButton.style.display = supported ? 'flex' : 'none';

    const mcpList = document.getElementById('mcp-list');
    if (!mcpList) return;

    mcpList.innerHTML = '';

    // Show warning if model doesn't support MCP
    if (!mcpSupported) {
        mcpList.innerHTML = '<div class="mcp-warning">当前选择的模型不支持MCP功能。</div>';
        return;
    }

    // Show message if no connections available
    if (!connections || connections.length === 0) {
        mcpList.innerHTML = '<div class="mcp-warning">没有可用的MCP连接。请先配置MCP连接。</div>';
        return;
    }

    // Create item for each connection
    connections.forEach(connection => {
        const item = document.createElement('div');
        item.className = 'mcp-item';

        // Create header (title and toggle)
        const header = document.createElement('div');
        header.className = 'mcp-item-header';

        // Title
        const title = document.createElement('div');
        title.className = 'mcp-item-title';
        title.textContent = connection.name;
        header.appendChild(title);

        // Toggle switch
        const toggleLabel = document.createElement('label');
        toggleLabel.className = 'mcp-toggle';

        const toggleInput = document.createElement('input');
        toggleInput.type = 'checkbox';
        toggleInput.checked = enabledList && enabledList.includes(connection.name);
        toggleInput.setAttribute('data-mcp-name', connection.name);

        const toggleSlider = document.createElement('span');
        toggleSlider.className = 'mcp-toggle-slider';

        toggleLabel.appendChild(toggleInput);
        toggleLabel.appendChild(toggleSlider);
        header.appendChild(toggleLabel);

        // Add description if available
        if (connection.description) {
            const desc = document.createElement('div');
            desc.className = 'mcp-item-description';
            desc.textContent = connection.description;
            item.appendChild(desc);
        }

        // Add connection type info
        const typeInfo = document.createElement('div');
        typeInfo.className = 'mcp-item-description';
        const connectionType = connection.command ? "Stdio" : "HTTP";
        const connectionUrl = connection.baseUrl || (connection.command ? `${connection.command} ${connection.args.join(' ')}` : "未知URL");
        typeInfo.textContent = `${connectionType}: ${connectionUrl}`;

        // Assemble item
        item.appendChild(header);
        if (connection.description) {
            const desc = document.createElement('div');
            desc.className = 'mcp-item-description';
            desc.textContent = connection.description;
            item.appendChild(desc);
        }
        item.appendChild(typeInfo);

        mcpList.appendChild(item);
    });
}

// Save MCP settings
function saveMcpSettings() {
    const enabledMcps = [];

    // Get all enabled MCPs
    document.querySelectorAll('#mcp-list input[type="checkbox"]:checked').forEach(checkbox => {
        enabledMcps.push(checkbox.getAttribute('data-mcp-name'));
    });

    // Send to backend
    sendMessageToServer({
        type: 'saveMcpSettings',
        enabledList: enabledMcps
    });

    // Update local cache
    enabledMcpList = enabledMcps;

    closeMcpDialog();
}

// Set MCP support status
function setMcpSupport(supported, connections, enabledList) {
    mcpSupported = supported;

    // Show or hide MCP button
    const mcpButton = document.getElementById('mcp-toggle-btn');
    mcpButton.style.display = supported ? 'flex' : 'none';

    // Map backend properties to frontend format
    if (connections) {
        mcpConnections = connections.map(conn => ({
            name: conn.name,
            description: conn.description || "",
            connectionType: conn.command ? "Stdio" : "HTTP",
            isActive: conn.isActive,
            baseUrl: conn.baseUrl || "",
            command: conn.command,
            args: conn.args || [],
            env: conn.env || {}
        }));

        renderMcpConnections(mcpConnections, enabledList, supported);
    }

    // Update enabled list cache
    if (enabledList) {
        enabledMcpList = enabledList;
    }
}
