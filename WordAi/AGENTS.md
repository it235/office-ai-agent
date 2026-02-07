# WORD AI ADD-IN

## OVERVIEW
Word-specific add-in providing AI-driven document processing and text manipulation capabilities. Integrates with Word through VSTO and uses shared components from ShareRibbon.

## STRUCTURE
```
WordAi/
├── ChatControl.vb         # Word-specific chat interface
├── Ribbon1.vb             # Word ribbon implementation
├── ThisAddIn.vb           # Word add-in entry point
└── WebDataCapturePane.vb  # Web content capture
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Word ribbon customization | Ribbon1.vb | Word tab and button definitions |
| Document processing | WebDataCapturePane.vb | Text extraction and manipulation |
| Chat interface | ChatControl.vb | Word-specific chat UI |
| Add-in initialization | ThisAddIn.vb | Startup and shutdown logic |

## CONVENTIONS
- Word-specific functionality only
- Inherits from shared BaseChatControl
- Uses Word interop through ShareRibbon services
- Follows Word VSTO patterns

## ANTI-PATTERNS
- Never access Excel or PowerPoint objects
- Never duplicate shared functionality from ShareRibbon
- Never bypass Word document protection