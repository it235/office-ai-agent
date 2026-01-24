# CONFIGURATION MANAGEMENT

## OVERVIEW
Centralized configuration management for all Office add-ins. Handles API keys, chat settings, prompts, and global application configuration.

## STRUCTURE
```
Config/
├── ConfigManager.vb      # Core configuration manager
├── ChatSettings.vb       # Chat-specific settings
├── ConfigSettings.vb     # Global configuration
└── ConfigApiForm.vb      # API configuration UI
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Configuration access | ConfigManager.vb | Main configuration entry point |
| Chat settings | ChatSettings.vb | Chat-specific preferences |
| Global settings | ConfigSettings.vb | Application-wide configuration |
| API configuration UI | ConfigApiForm.vb | User interface for API keys |

## CONVENTIONS
- All configuration accessed through ConfigManager
- Settings stored in user's AppData directory
- API keys encrypted at rest
- Configuration changes trigger events

## ANTI-PATTERNS
- Never access configuration files directly
- Never store API keys in plain text
- Never hardcode configuration values
- Never bypass ConfigManager for settings access