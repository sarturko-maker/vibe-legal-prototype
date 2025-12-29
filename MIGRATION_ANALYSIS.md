# Vibe Legal - Migration Analysis

## Source File
**File:** `scriptlab/vibe-legal-beta-0.2.yaml`  
**Lines:** 17,571  
**Size:** 742KB

---

## Code Inventory

### Core Services (Extraction Priority)

| Section | Lines | Description | Target Module |
|---------|-------|-------------|---------------|
| Provider Configuration | 20-52 | API key, model selection for Gemini/Claude | `services/providers/` |
| Track Changes Author | 54-100 | Author detection and mode selection | `services/document/authorService.ts` |
| Cache System | 104-125 | TTL, context caching config | `services/cache/` |
| OOXML Validation | 126-270 | Pre-change OOXML structure validation | `services/ooxml/validator.ts` |
| Helper Functions | 274-600 | Path building, element utilities | `utils/` |
| Gemini API | 604-730 | API calls, response parsing | `services/gemini/` |
| Contract Map Builder | 730-900 | Clause detection, paragraph indexing | `services/document/contractMap.ts` |
| Redline Engine | 900-1250 | Surgical diff, track changes generation | `services/diff/` |
| Markdown Parser | 1250-1400 | **bold**, *italic*, __underline__ handling | `services/formatting/` |
| Operations Router | 1400-2000 | INSERT, AMEND, DELETE processing | `services/operations/` |
| OOXML Track Changes | 2000-10000 | `<w:ins>`, `<w:del>` creation | `services/ooxml/trackChanges.ts` |

### UI Components

| Component | Approx Lines | Description |
|-----------|--------------|-------------|
| Header | 14888-15000 | Settings, refresh, help buttons |
| Chat | 15100-15400 | Message display, loading states |
| InputArea | 15400-15700 | Text input, send button |
| Settings | 15700-15860 | API key, model, author config |
| HelpGuide | Modal inside components | Usage instructions |
| App | 15861-17500 | Root component, state management |

---

## Dependency Graph

```
User Input
  → handleAction()
    → getCurrentApiKey() / getCurrentModel()
    → Word.run()
      → buildContractMap() / ensureContractMap()
      → callGeminiRouter() 
        → System prompt construction
        → Gemini API call
        → JSON response parsing
      → parseOperations()
      → For each operation:
        → AMEND_SIMPLE: applyRedlineToOxmlWithFormatting()
          → extractFormattingMap()
          → surgical diff
          → createTrackChangeRuns()
        → INSERT: insertBlocks()
          → cloneOxml()
          → parseMarkdownToSegments()
          → createFormattedRuns()
        → AMEND (tree): processParaOperations()
          → paragraph-level diff
    → Update UI state
```

---

## Tight Couplings (Risks)

1. **Global State Variables**
   - `providerConfig`, `authorSettings`, `contractMapState` - need proper React context
   - `styleCache` - document styles cached globally

2. **handleAction() Monolith**
   - Lines 15900-17500: Single 1600-line async function
   - Handles all operation types inline
   - Mixed Word API calls with AI calls

3. **OOXML Functions with UI State**
   - Some OOXML functions call `setMessages()` directly for error reporting
   - Need separation: OOXML returns result, UI layer handles messaging

---

## Reusable Patterns

1. **WORD_NS constant** - XML namespace, use everywhere
2. **createTextRun()** - Creates `<w:r>` elements with text
3. **createTrackChange()** - Wraps runs in `<w:ins>` or `<w:del>`
4. **parseFormattingMarkers()** - Markdown to segments
5. **Clause number regex** - `/^(\d+(?:\.\d+)+)/` pattern used in multiple places

---

## Migration Notes

### High Value / Low Risk
- Types and interfaces (extract first)
- Markdown parser (standalone)
- Utility functions (pure)

### High Value / Medium Risk
- OOXML track changes (complex but critical)
- Surgical diff engine (the "crown jewel")

### Medium Value / High Risk
- handleAction() decomposition (1600 lines, many dependencies)
- Contract map builder (Word API dependent)

---

## Next Steps

1. Create folder structure per ARCHITECTURE.md
2. Extract types first (no dependencies)
3. Extract utilities (pure functions)
4. Extract services bottom-up (formatting → diff → ooxml → operations)
5. Wire up React components to new services
