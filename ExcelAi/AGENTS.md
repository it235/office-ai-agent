# EXCEL AI ADD-IN

## OVERVIEW
Excel-specific add-in providing AI-driven data analysis and automation capabilities. Integrates with Excel through VSTO and uses shared components from ShareRibbon.

## STRUCTURE
```
ExcelAi/
├── ChatControl.vb    # Excel-specific chat interface
├── Ribbon1.vb        # Excel ribbon implementation
├── ThisAddIn.vb      # Excel add-in entry point
└── ExcelFunctions.vb # Excel-specific functionality
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Excel ribbon customization | Ribbon1.vb | Excel tab and button definitions |
| Excel data analysis | ExcelFunctions.vb | Cell and range operations |
| Chat interface | ChatControl.vb | Excel-specific chat UI |
| Add-in initialization | ThisAddIn.vb | Startup and shutdown logic |

## CONVENTIONS
- Excel-specific functionality only
- Inherits from shared BaseChatControl
- Uses Excel interop through ShareRibbon services
- Follows Excel VSTO patterns

## ANTI-PATTERNS
- Never access Word or PowerPoint objects
- Never duplicate shared functionality from ShareRibbon
- Never bypass Excel security model