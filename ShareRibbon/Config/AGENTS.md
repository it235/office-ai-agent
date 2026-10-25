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
- `ChatSettings.LoadSettings` 只在 `ContainsKey(key)` 时覆盖静态值：默认值代表「配置文件还没表态」，
  因此把某个开关翻成默认开启（如 `selectedCellChecked`，E4-04）不会把显式存过 `false` 的用户再打开。
- `ChatSettings.SaveSettings` 必须读改写同一个 `JObject`，只覆盖自己认识的键；用固定字典重建文件会静默
  删掉表外的键（`enableMcpList` 曾经就是这样被吃掉的）。
- 默认开启的开关必须同时存在面板设置项与载荷回读（`HandleSaveSettings` 从载荷读，不从静态字段写回），
  否则它只能通过重新编译来关闭。
- 在**读开关之前**先确保配置已被读过一次：`ChatSettings.EnsureLoaded(applicationInfo)` 首次调用才构造（构造即
  `LoadSettings`），重复调用返回同一实例，宿主入口方法（面板可见性事件里挂右键菜单/选区工具条）必须先调它，
  否则进程内第一次拿到的是代码默认值。缓存一份是安全的，因为聊天设置文件与宿主无关
  （`ApplicationInfo.GetChatSettingsFilePath` 三端都指向 `office_ai_chat_settings.json`）。
- `EnsureLoaded` 不等于热重载：面板自身装载路径仍每次 `New ChatSettings(...)`，外部编辑文件在下一次装载时
  生效。不要把它扩展成配置订阅机制。

## ANTI-PATTERNS
- Never access configuration files directly
- Never store API keys in plain text
- Never hardcode configuration values
- Never bypass ConfigManager for settings access