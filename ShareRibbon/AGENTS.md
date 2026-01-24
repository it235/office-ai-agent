# SHARE RIBBON COMPONENTS

## OVERVIEW
Shared components library used by all Office add-ins. Contains core services, UI controls, configuration management, and MCP protocol implementation.

## STRUCTURE
```
ShareRibbon/
├── Config/           # Configuration management
├── Controls/         # UI components and services
├── Mcp/              # Model Context Protocol implementation
└── Resources/        # Embedded resources
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Configuration management | Config/ | API keys, settings, prompts |
| Chat UI components | Controls/BaseChatControl.vb | Core chat interface |
| HTTP communication | Controls/Services/HttpStreamService.vb | AI API communication |
| MCP protocol | Mcp/ | Model Context Protocol client |
| Resource management | Resources/ | Embedded JS/CSS assets |

## CONVENTIONS
- All shared functionality must be in ShareRibbon
- Services follow dependency injection pattern
- Configuration managed through ConfigManager
- UI components use WebView2 for rendering

## ANTI-PATTERNS
- Never duplicate shared functionality in individual add-ins
- Never access Office interop directly outside ShareRibbon
- Never hardcode API keys or configuration values