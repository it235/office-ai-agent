# POWERPOINT AI ADD-IN

## OVERVIEW
PowerPoint-specific add-in providing AI-driven presentation creation and slide manipulation capabilities. Integrates with PowerPoint through VSTO and uses shared components from ShareRibbon.

## STRUCTURE
```
PowerPointAi/
├── ChatControl.vb    # PowerPoint-specific chat interface
├── Ribbon1.vb        # PowerPoint ribbon implementation
└── ThisAddIn.vb      # PowerPoint add-in entry point
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| PowerPoint ribbon customization | Ribbon1.vb | PowerPoint tab and button definitions |
| Presentation operations | ChatControl.vb | Slide and presentation manipulation |
| Chat interface | ChatControl.vb | PowerPoint-specific chat UI |
| Add-in initialization | ThisAddIn.vb | Startup and shutdown logic |

## CONVENTIONS
- PowerPoint-specific functionality only
- Inherits from shared BaseChatControl
- Uses PowerPoint interop through ShareRibbon services
- Follows PowerPoint VSTO patterns

## ANTI-PATTERNS
- Never access Excel or Word objects
- Never duplicate shared functionality from ShareRibbon
- Never bypass PowerPoint slide protection