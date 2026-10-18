---
name: word-document-agent
description: Use for Word tasks that need document context, proofreading, formatting, heading and numbering restructuring, style application, translation, continuation, or multi-step document editing.
application: Word
default_for_application: true
tags: word, docx, document, proofread, format, heading, numbering, style, translate, writing
allowed-tools: ListParagraphs, GetParagraphInfo, InsertText, FormatText, ReplaceText, DeleteText, ApplyStyle, SetParagraphFormat, InsertTable, GenerateTOC, BeautifyDocument
intent_types: proofread, formatting, numbering, heading, writing, translation
---

# Word Document Agent

Use this skill when the user asks Word to modify the current document rather than merely chat about it.

## Operating Rules

1. Read the current selection, paragraph structure, headings, styles, and nearby context before planning.
2. If the user selected content, operate on the selection by default; otherwise infer whether the request applies to the whole document.
3. Do not repeatedly ask for range, style, numbering, or visible text when Word context can provide it.
4. Prefer structured document tools and Word-specific harnesses over free-form VBA.
5. For formatting, proofreading, numbering, and heading work, plan, preview or explain, execute, observe, and repair.
6. Keep `ShareRibbon` generic; Word COM execution belongs in `WordAi`.

## Common Tasks

- Rebuild headings and numbering
- Normalize fonts, spacing, indentation, and title hierarchy
- Proofread typos, punctuation, and formal expression
- Apply official document, report, or paper styles
- Generate or update table of contents
- Continue or rewrite selected content
