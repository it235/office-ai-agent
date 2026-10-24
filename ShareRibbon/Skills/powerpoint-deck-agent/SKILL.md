---
name: powerpoint-deck-agent
description: Use for PowerPoint tasks that need slide generation, deck structure, layout cleanup, speaker notes, visual consistency, translation, review, or chart/table insertion.
application: PowerPoint
default_for_application: true
tags: powerpoint, ppt, slide, deck, presentation, layout, theme, chart, notes, review
allowed-tools: CreateSlides, InsertSlide, DeleteSlide, DuplicateSlide, MoveSlide, SetSlideLayout, InsertShape, InsertText, InsertTable, FormatSlide, ApplyTheme, ApplyTransition, AddAnimation, AddSpeakerNotes, BeautifySlides, DiscoverOfficeCapability, OfficeObjectOperation
intent_types: slide_generation, formatting, review, translation, presentation
---

# PowerPoint Deck Agent

Use this skill when the user asks PowerPoint to create or improve slides.

## Operating Rules

1. Read current slide index, total slide count, selected shapes, text boxes, slide titles, and visible content before planning.
2. Treat the selected slide or selected shapes as the primary working area unless the user says the whole deck.
3. Use registered slide tools. If no registered tool can perform an operation, report a capability gap instead of generating VBA implicitly.
4. For design work, preserve the user's deck intent and improve readability, hierarchy, consistency, and presentation flow.
5. After execution, observe slide count, modified slide index, inserted shapes, notes, and text changes.
6. If the result is not correct, repair the plan using observed slide state.
7. For a long-tail object capability not covered by a high-level tool, call `DiscoverOfficeCapability` first and use only returned executable `MemberId` values in `OfficeObjectOperation`.
8. For deck creation, prefer the professional `CreateSlides` Scene contract with one coherent `designSystem` and explicit page archetypes. Do not default every page to title-and-bullets.

## Declarative Object Operations

- Use only canonical refs such as `PowerPoint:presentations/active/slides/2/shapes`; never invent a COM object path.
- Build `OfficeObjectOperation.batch` with `schemaVersion=1.0`, `appType=PowerPoint`, unique operation IDs, and actions limited to `get/set/invoke/create/delete/collection_item`.
- Copy `MemberId` exactly from the latest `DiscoverOfficeCapability` result. Do not derive or shorten it.
- For SmartArt, discover and create on the target slide's `shapes` collection. Read the returned `resultRef` from Observation/Data before addressing the created shape.
- SmartArt node text can be addressed beneath the returned shape as `/smartart/nodes/{1-based-index}/textframe2/textrange`; discover the writable text member and use `action=set` with `arguments.value`.
- Include `expectedEffects` such as `hasSmartArt`, `nodeCount`, `text`, or `nodeTexts` when the expected state is known. Treat `VERIFY_FAILED` as a signal to observe and repair parameters or refs.

## Common Tasks

- Generate slides from an outline
- Insert, delete, duplicate, or reorder slides
- Apply a named layout from the current deck, or a standard title/content/blank layout
- Insert shapes and text boxes
- Beautify the current slide or whole deck
- Rewrite slide text for business reporting
- Add speaker notes
- Insert charts or tables
- Apply transitions or shape animations
- Apply a theme file (`.thmx`); built-in theme names cannot be applied through COM, so ask for a theme file instead of claiming success
