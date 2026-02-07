# MCP PROTOCOL IMPLEMENTATION

## OVERVIEW
Model Context Protocol client implementation enabling communication with AI model servers. Provides configuration management, connection handling, and JSON-RPC communication.

## STRUCTURE
```
Mcp/
├── StreamJsonRpcMCPClient.vb  # Core MCP client
├── MCPConnectionConfig.vb     # Connection configuration
├── MCPConfigForm.vb           # Configuration UI
└── MCPEntities.vb             # Data structures
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| MCP client implementation | StreamJsonRpcMCPClient.vb | Core communication logic |
| Connection configuration | MCPConnectionConfig.vb | Server settings and auth |
| Configuration UI | MCPConfigForm.vb | User interface for MCP settings |
| Data structures | MCPEntities.vb | Protocol message definitions |

## CONVENTIONS
- Uses StreamJsonRpc for communication
- Follows MCP specification exactly
- Configuration managed through ConfigManager
- Supports multiple connection types (stdio, HTTP, etc.)

## ANTI-PATTERNS
- Never bypass MCP protocol specification
- Never hardcode server addresses
- Never expose API keys in client code
- Never block main thread during MCP operations