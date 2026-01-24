# SHARED UI CONTROLS

## OVERVIEW
Core user interface components shared across all Office add-ins. Provides chat interfaces, data capture panes, and deepseek chat functionality using WebView2.

## STRUCTURE
```
Controls/
├── Services/              # UI-related services
├── BaseChatControl.vb     # Core chat component
├── BaseDataCapturePane.vb # Data capture interface
└── BaseDeepseekChat.vb    # DeepSeek-specific chat
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Chat UI implementation | BaseChatControl.vb | Core chat functionality |
| Data capture | BaseDataCapturePane.vb | Web content extraction |
| DeepSeek integration | BaseDeepseekChat.vb | DeepSeek-specific features |
| UI services | Services/ | WebView, message handling |

## CONVENTIONS
- All UI components inherit from UserControl
- WebView2 used for modern HTML rendering
- JavaScript/CSS resources embedded in ShareRibbon/Resources
- Services handle backend communication

## ANTI-PATTERNS
- Never access Office interop directly in UI controls
- Never block UI thread with long operations
- Never hardcode UI strings (use resources)