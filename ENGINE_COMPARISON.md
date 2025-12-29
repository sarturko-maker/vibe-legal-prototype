# Engine Comparison: Legacy vs New Modular Code

## Summary

| Function | Status | Critical? | Notes |
|----------|--------|-----------|-------|
| Paragraph Loading | ⚠️ PARTIAL MATCH | HIGH | Legacy loads paragraphs ONCE and stores in `lookup` object |
| Contract Map Building | ✅ SIMILAR | MEDIUM | Pattern matching is similar |
| Get Paragraph OOXML | ❌ DIFFERS | **CRITICAL** | Legacy doesn't use individual getParagraphOoxml |
| Insert OOXML | ❌ DIFFERS | **CRITICAL** | Legacy uses `insertBlocks` with styleCache |
| AMEND_SIMPLE Handler | ❌ DIFFERS | **CRITICAL** | Legacy uses `sectionRange.expandTo` + `applyRedlineToOxml` |
| INSERT Handler | ❌ DIFFERS | **CRITICAL** | Legacy uses `insertBlocks`, new uses `insertOoxml` |
| Surgical Diff | ❌ DIFFERS | **CRITICAL** | Legacy uses `applyRedlineToOxml`, new uses `computeSurgicalDiff` |
| Track Change Wrappers | ⚠️ PARTIAL | MEDIUM | Structure similar but invocation differs |

## 🔴 LIKELY ROOT CAUSE

**The new modular code completely diverges from the legacy architecture:**

1. **Legacy Pattern:**
   - Loads paragraphs ONCE into `lookup[id]` object at start of `handleAction`
   - AMEND uses `sectionRange.expandTo()` spanning start/end paragraphs
   - Calls `applyRedlineToOxml(rangeOxml.value, sectionRange.text, amended_text)`
   - Uses `sectionRange.insertOoxml(result.oxml, Word.InsertLocation.replace)`

2. **New Pattern:**
   - Re-fetches paragraphs for each operation in separate `Word.run()`
   - AMEND calls `getParagraphOoxml(targetId)` for individual paragraph
   - Calls `computeSurgicalDiff()` + `applyDiffsToOoxml()`
   - Uses `insertOoxmlAtParagraph(targetId, oxml, 'Replace')`

**The fundamental mismatch:** Legacy operates on **live paragraph references within a single Word.run context**, while new code **re-fetches paragraphs in separate Word.run calls**.

---

## Detailed Comparisons

### 1. Paragraph Loading & Lookup

#### Legacy (L7662-7694)
```typescript
const mapLines = [];
const lookup = {};

for (let i = 0; i < paragraphs.items.length; i++) {
    const p = paragraphs.items[i];
    if (p.isListItem) {
        p.listItem.load("level");
    }
}
await context.sync();

for (let i = 0; i < paragraphs.items.length; i++) {
    const p = paragraphs.items[i];
    const id = i + 1;              // 1-indexed!
    lookup[id] = p;                // STORES LIVE PARAGRAPH REFERENCE
    
    context.trackedObjects.add(p); // TRACKS FOR LATER USE
    trackedItems.push(p);
    
    const text = p.text.substring(0, 150).replace(/\n/g, " ");
    // ... build document map
}
```

#### New (wordApi.ts)
```typescript
export async function loadParagraphs(context: any): Promise<ParagraphInfo[]> {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    const result: ParagraphInfo[] = [];
    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load('text,style,isListItem');
        await context.sync();

        result.push({
            id: i + 1,  // 1-indexed - MATCHES
            text: para.text || '',
            style: para.style || 'Normal',
            isListItem: para.isListItem || false
        });
    }
    return result;  // Returns DATA only, not live references!
}
```

**❌ CRITICAL DIFFERENCE:**
- Legacy stores **live paragraph references** in `lookup[id]`
- New code stores **data snapshots** with no live references
- Legacy uses `context.trackedObjects.add(p)` to keep references valid
- New code loses references when `loadParagraphs` returns

---

### 2. AMEND_SIMPLE Handler

#### Legacy (L8008-8039)
```typescript
else if (op.amended_text) {
    const startPara = lookup[range.startId];  // USES STORED REFERENCE
    const endPara = lookup[range.endId];
    if (startPara && endPara) {
        const sectionRange = startPara.getRange("Start").expandTo(endPara.getRange("End"));
        sectionRange.load("text");
        const rangeOxml = sectionRange.getOoxml();
        await context.sync();
        
        const redlineResult = applyRedlineToOxml(rangeOxml.value, sectionRange.text, op.amended_text);
        
        if (redlineResult.hasChanges) {
            sectionRange.insertOoxml(redlineResult.oxml, Word.InsertLocation.replace);
            await context.sync();
        }
    }
}
```

#### New (amendSimple.ts)
```typescript
export async function handleAmendSimple(
    context: any,
    operation: AmendSimpleOperation,
    author: string
): Promise<ExecutionResult> {
    const targetId = operation.target_id;

    // Get current OOXML - RE-FETCHES PARAGRAPHS!
    const currentOoxml = await getParagraphOoxml(context, targetId);
    const currentText = extractTextFromOoxml(currentOoxml);

    // Compute diff - DIFFERENT ALGORITHM!
    const diffs = computeSurgicalDiff(currentText, operation.amended_text);

    // Apply diffs to OOXML
    const result = applyDiffsToOoxml(currentOoxml, diffs, author);

    // Insert modified OOXML
    await insertOoxmlAtParagraph(context, targetId, result.oxml, 'Replace');
}
```

**❌ CRITICAL DIFFERENCES:**
1. Legacy uses `lookup[id]` (stored reference), new uses `getParagraphOoxml` (re-fetches)
2. Legacy uses `sectionRange.expandTo()` for multi-paragraph, new only handles single
3. Legacy uses `applyRedlineToOxml(oxml, originalText, modifiedText)`, new uses `computeSurgicalDiff` + `applyDiffsToOoxml`
4. Legacy applies to `sectionRange.insertOoxml()`, new applies to individual paragraph

---

### 3. INSERT Handler

#### Legacy (L8077-8118)
```typescript
if (resolvedId && op.content) {
    const refP = lookup[resolvedId];  // USES STORED REFERENCE
    if (refP) {
        // Smart styling logic...
        let blockStyleId = "T_Normal";
        
        const addedCount = await insertBlocks(
            context,
            [{ content: op.content, styleId: blockStyleId }],
            styleCache,         // HAS STYLE CACHE
            refP,               // USES LIVE REFERENCE
            routerResponse.analysis,
            null
        );
        await context.sync();
    }
}
```

#### New (insert.ts)
```typescript
export async function handleInsert(
    context: any,
    operation: InsertOperation,
    author: string
): Promise<ExecutionResult> {
    const insertAfter = operation.insert_after;

    // Get reference OOXML - RE-FETCHES!
    const refOoxml = await getParagraphOoxml(context, insertAfter);

    // Build new paragraph
    const newOoxml = buildInsertionOoxml(refOoxml, operation.content, author);

    // Insert after target - RE-FETCHES AGAIN!
    await insertOoxmlAtParagraph(context, insertAfter, newOoxml, 'After');
}
```

**❌ CRITICAL DIFFERENCES:**
1. Legacy uses `lookup[id]` (stored reference), new re-fetches twice
2. Legacy uses `insertBlocks()` with `styleCache`, new uses simple `buildInsertionOoxml`
3. Legacy has smart styling with `routerResponse.analysis`, new doesn't
4. Legacy works within same `Word.run`, new may be in separate context

---

### 4. applyRedlineToOxml vs computeSurgicalDiff + applyDiffsToOoxml

#### Legacy (L1903-1976)
```typescript
function applyRedlineToOxml(
    oxml: string,
    originalText: string,
    modifiedText: string
): { oxml: string; hasChanges: boolean } {
    const parser = new DOMParser();
    let xmlDoc = parser.parseFromString(oxml, "text/xml");

    // ENHANCED: Clean up AI response
    const cleanModifiedText = hasNumbering 
        ? sanitizeForNumberedList(originalText, modifiedText, true)
        : sanitizeAiResponse(modifiedText);

    if (cleanModifiedText.trim() === originalText.trim()) {
        return { oxml, hasChanges: false };
    }

    // PARAGRAPH-AWARE MODE for multi-paragraph changes
    if (isMultiParagraph && structureChanging && hasNumbering) {
        return applyParagraphAwareMode(xmlDoc, originalText, cleanModifiedText, serializer);
    }

    // Complex structure detection
    if (needsSurgicalMode) {
        return applySurgicalMode(xmlDoc, originalText, cleanModifiedText, serializer);
    } else {
        return applyReconstructionMode(xmlDoc, originalText, cleanModifiedText, serializer);
    }
}
```

#### New (processor.ts)
```typescript
export function applyDiffsToOoxml(
    oxml: string,
    diffs: StructuredDiff[],  // PRE-COMPUTED DIFFS
    author: string,
    date?: string
): OoxmlModificationResult {
    // Parse OXML
    const xmlDoc = parser.parseFromString(oxml, 'application/xml');

    // Find paragraph, build segment map
    const paragraph = paragraphs[0] as Element;
    const map = buildStructuralMap(paragraph);

    // Apply diffs in reverse order
    const reversedDiffs = [...diffs].sort((a, b) => b.start - a.start);
    for (const diff of reversedDiffs) {
        applyDiff(xmlDoc, map, diff, author, changeDate, sessionRsid);
    }
}
```

**❌ CRITICAL DIFFERENCES:**
1. Legacy takes `originalText` + `modifiedText`, computes diff internally
2. New takes pre-computed `StructuredDiff[]` array
3. Legacy has `sanitizeAiResponse()` and `sanitizeForNumberedList()` - new does not
4. Legacy detects paragraph structure and chooses mode (surgical/reconstruction)
5. Legacy handles multi-paragraph scenarios - new only handles single paragraph

---

## 🔧 RECOMMENDED FIXES

### Option A: Minimal Fix (Keep new architecture, fix targeting)
1. Ensure `executeOperations` runs in SAME `Word.run` as paragraph loading
2. Store paragraph references properly for reuse
3. Use 1-indexed IDs consistently

### Option B: Port Legacy Engine (Higher fidelity)
1. Port `applyRedlineToOxml` to new module
2. Port `insertBlocks` with styleCache
3. Keep `lookup` pattern with live paragraph references
4. Ensure all operations happen in single `Word.run`

### Option C: Hybrid Approach (Recommended)
1. Keep new modular structure
2. Port `applyRedlineToOxml` as the core diff engine
3. Modify `handleAction` to:
   - Load paragraphs ONCE
   - Store in lookup object
   - Pass lookup to `executeOperations`
   - Execute all in single `Word.run`

---

## Immediate Action Items

1. **Fix handleAction to keep single Word.run context:**
```typescript
await Word.run(async (context) => {
    // 1. Load paragraphs ONCE
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items,text,style,isListItem');
    await context.sync();
    
    // 2. Build lookup
    const lookup: Record<number, any> = {};
    for (let i = 0; i < paragraphs.items.length; i++) {
        lookup[i + 1] = paragraphs.items[i];
        context.trackedObjects.add(paragraphs.items[i]);
    }
    
    // 3. Call AI (outside Word.run or with progress updates)
    // 4. Execute operations with lookup reference
    for (const op of operations) {
        if (op.type === 'AMEND_SIMPLE') {
            const para = lookup[op.target_id];
            // Use para directly...
        }
    }
});
```

2. **Port applyRedlineToOxml to processor.ts** - it handles edge cases the new code doesn't

3. **Port insertBlocks for INSERT operations** - it handles styling correctly
