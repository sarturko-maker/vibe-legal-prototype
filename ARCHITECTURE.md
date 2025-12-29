# Vibe Legal - Architecture Design

## Extraction Source

**Primary:** `vibe-legal-beta-0.2.yaml` (stable, 17,571 lines)  
**Post-extraction:** Add 0.3 Pro Mode features as a layer

---

## Folder Structure

```
vibe-legal-react/
├── scriptlab/                      # PRESERVED
│   ├── vibe-legal-beta-0.2.yaml
│   └── vibe-legal-beta-0.3.yaml
│
├── src/
│   ├── taskpane/
│   ├── components/
│   │   ├── App.tsx
│   │   ├── ModeToggle/
│   │   ├── ProModeSelector/
│   │   ├── FocusIndicator/         # 0.3 feature
│   │   ├── ChatArea/
│   │   ├── PreviewPanel/           # 0.3 feature
│   │   ├── QueuePanel/             # 0.3 feature
│   │   ├── InputArea/
│   │   ├── Settings/
│   │   └── index.ts
│   │
│   ├── services/
│   │   ├── gemini/
│   │   ├── ooxml/                  # 8K lines → 11 files (see below)
│   │   ├── diff/
│   │   ├── formatting/
│   │   ├── document/
│   │   └── operations/             # 5 separate handlers
│   │
│   ├── prompts/
│   ├── state/
│   ├── types/
│   └── utils/
```

---

## OOXML Sub-Module Breakdown (8K lines → 11 files)

| File | Source Lines | Est. Lines | Purpose |
|------|--------------|------------|---------|
| `namespaces.ts` | L116 | ~50 | `WORD_NS` constants |
| `validator.ts` | L126-270 | ~150 | Structure validation |
| `types.ts` | L140-180 | ~50 | TextSegment, StructuralMap interfaces |
| `runBuilder.ts` | L2100-2160 | ~200 | `cloneRunWithText()`, `cloneRunWithDelText()` |
| `paragraphBuilder.ts` | L2700-3000 | ~300 | `<w:p>` creation, numPr injection |
| `segmentMap.ts` | L2045-2100 | ~150 | `findSegmentsForRange()`, `findInsertionPoint()` |
| `runSplitter.ts` | L2199-2300 | ~200 | `splitRunForDeletion()`, `splitRunForInsertion()` |
| `trackChanges.ts` | L2304-2600 | ~400 | `applyStructuredDeletion()`, `applyStructuredInsertion()` |
| `styleInheritance.ts` | L3000-3500 | ~500 | Style copying, font/size inheritance |
| `numberingPreserver.ts` | L3500-4000 | ~500 | List numbering, numPr handling |
| `processor.ts` | L4000-6000 | ~2000 | Main orchestration, `applyRedlineToOxmlWithFormatting()` |

**Total: ~4,500 lines** (remainder is helper functions and edge cases scattered throughout)

---

## Operations Handler Breakdown (L1400-2000)

Each operation type becomes a separate file:

| File | Est. Lines | Key Function |
|------|------------|--------------|
| `router.ts` | ~50 | `executeOperations()` switch |
| `amendSimple.ts` | ~150 | `handleAmendSimple()` |
| `amendTree.ts` | ~200 | `handleAmendTree()` |
| `insert.ts` | ~150 | `handleInsert()` |
| `insertBlock.ts` | ~100 | `handleInsertBlock()` |
| `delete.ts` | ~50 | `handleDelete()` |

---

## handleAction() Decomposition

```typescript
// Target: 5 focused functions
async function handleAction(input: string) {
    const context = await prepareDocumentContext();  // 100 lines
    const response = await callAI(context, input);   // 200 lines
    const operations = parseResponse(response);      // 100 lines
    const results = await executeOperations(ops);    // 50 lines (router)
    updateUIWithResults(results);                    // 100 lines
}
```

---

## State Migration Table (Including 0.3)

### SettingsContext
| Variable | Type |
|----------|------|
| `provider` | `'gemini' \| 'claude'` |
| `geminiApiKey` | `string` |
| `claudeApiKey` | `string` |
| `geminiModel` | `string` |
| `claudeModel` | `string` |
| `authorMode` | `'auto' \| 'vibe' \| 'custom'` |
| `customAuthor` | `string` |
| `detectedAuthor` | `string \| null` |

### DocumentContext
| Variable | Type |
|----------|------|
| `contractMap` | `ContractMap \| null` |
| `styleCache` | `Map<string, StyleInfo>` |
| `documentHash` | `string \| null` |
| `focusedClause` | `FocusedClause \| null` |

### ChatContext
| Variable | Type |
|----------|------|
| `messages` | `Message[]` |
| `isProcessing` | `boolean` |
| `processingStage` | `string` |
| `appMode` | `'chat' \| 'pro'` |
| `proSubMode` | `'ask' \| 'draft' \| 'auto'` |
| `pendingPreview` | `AmendmentPreview \| null` |
| `amendmentQueue` | `Amendment[]` |
| `currentAmendmentIndex` | `number` |

---

## Error Handling

```typescript
type Result<T> = 
    | { success: true; data: T }
    | { success: false; error: string; code?: string };
```

| Category | Action |
|----------|--------|
| Critical | Error boundary, stop |
| Recoverable | Return Result, continue |
| Warning | Log only |

---

## Console Logging

```typescript
// src/utils/logger.ts
const DEBUG = process.env.NODE_ENV === 'development';

export const log = (...args: any[]) => {
    if (DEBUG) console.log('[Vibe]', ...args);
};

export const logWarn = (...args: any[]) => {
    if (DEBUG) console.warn('[Vibe]', ...args);
};

export const logError = (...args: any[]) => {
    console.error('[Vibe Error]', ...args);  // Always log errors
};
```

---

## Testing Approach

1. **Unit tests** - Pure functions (markdown parser, diff engine)
2. **Browser testing** - Use `src/utils/officeMock.ts`
3. **Comparison** - Verify output matches Script Lab version

### Verification Checkpoints
```bash
# After each extraction step
npm run build  # Must pass

# After services complete
npm start      # UI renders

# Final validation
# Sideload in Word Online
```

---

## Extraction Order

1. `types/` ← No deps
2. `utils/` ← Types only
3. `services/formatting/` ← Pure
4. `services/diff/` ← Uses formatting
5. `services/ooxml/` ← Uses diff, formatting
6. `services/document/` ← Word API
7. `services/operations/` ← Uses all above
8. `services/gemini/` ← AI calls
9. `state/` ← React contexts
10. `components/` ← Wire up UI
11. Add 0.3 features last
