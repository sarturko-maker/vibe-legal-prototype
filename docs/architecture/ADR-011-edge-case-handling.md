# ADR-011: Edge Case Handling in Redline Engine

**Status:** Accepted  
**Date:** 2025-12-28  
**Authors:** Artur (Vibe Legal)  
**Supersedes:** None

---

## Context

The Vibe Legal redline engine performs surgical OOXML manipulation to insert track changes into Microsoft Word documents. Legal documents have complex structures including numbered lists, nested clauses, mixed formatting, and cross-references. 

During development, numerous edge cases were discovered that caused incorrect targeting, formatting loss, or document corruption. This ADR catalogues all identified edge cases, their solutions, and implementation locations.

---

## Decision

We will maintain this living document as the authoritative reference for edge case handling. All edge cases follow a consistent pattern:

1. **Detection** — How to identify when the edge case applies
2. **Solution** — The approach taken to handle it
3. **Implementation** — Where in the codebase the solution lives
4. **Test Case** — How to verify correct handling

---

## Edge Case Categories

| Category | Count | Description |
|----------|-------|-------------|
| Numbering | 4 | Auto vs manual numbering, preservation, renumbering |
| Structure | 5 | Paragraph targeting, clause hierarchy, document layout |
| Formatting | 5 | Styles, fonts, markdown, run-level formatting |
| Lists | 4 | Bullets, numbered lists, nesting, numId handling |
| Track Changes | 4 | OOXML wrappers, RSID, attribution |
| Map & Targeting | 4 | Contract map, paragraph resolution, context |

---

## Numbering Edge Cases

### EC-01: Auto-Numbered Lists (Word's w:numPr)

**Problem:** Word's automatic numbering is stored in `<w:numPr>` with `numId` and `ilvl` (indent level). When inserting new paragraphs, the numbering must be preserved or Word breaks the list sequence.

**Detection:**
```typescript
if (paragraph.isListItem) {
  const level = paragraph.listItem.level;
}
```

**Solution:** Clone the `<w:numPr>` element from a reference paragraph at the same level. Never construct `numId` values manually — extract from existing document.

**Implementation:** [insertBlocks.ts](file:///home/sarturko/vibe-legal-react/src/services/ooxml/insertBlocks.ts) — OOXML construction phase

**Test Case:** Insert a new clause "2.5" between 2.4 and 2.6. Verify Word renumbers automatically.

---

### EC-02: Manually Typed Numbers

**Problem:** Some documents use manually typed numbers ("3.1", "3.2") instead of Word's numbering. These paragraphs have `isListItem: false` but contain number prefixes in the text.

**Detection:**
```typescript
const manualNumberPattern = /^[\(\[]?\d+(\.\d+)*[\)\]]?\.?\s/;
const hasManualNumber = manualNumberPattern.test(paragraph.text);
const isAutoNumbered = paragraph.isListItem;
```

**Solution:** Detect via regex on paragraph text. Include `isListItem` flag in contract map so AI knows the difference.

**Implementation:** [contractMap.ts](file:///home/sarturko/vibe-legal-react/src/services/document/contractMap.ts) — contract map builder

**Test Case:** Document with "3.1" typed manually. Map should show `{List: false}` with number detected in text.

---

### EC-21: Manual Number Preservation in AMEND

**Problem:** When AI amends a paragraph with a manually typed number (e.g., "3.2. Delivery shall..."), it often strips the number and returns only the body text.

**Detection:**
```typescript
// AI returns "Delivery shall occur within 5 days..."
// Missing the original "3.2. " prefix
```

**Solution:** Post-processing function `restoreManualNumber()` that checks if original had manual number and restores if AI stripped it.

**Implementation:** [applyRedline.ts](file:///home/sarturko/vibe-legal-react/src/services/ooxml/applyRedline.ts#L936-L962) — `restoreManualNumber()` function

**Test Case:** Amend clause "3.2. Delivery shall..." → result should still start with "3.2."

---

### EC-22: Renumbering After Mid-List Insertion

**Problem:** When inserting a new clause in the middle of a manually-numbered list, subsequent clauses need renumbering.

**Solution:** **NOT IMPLEMENTED** — Requires user confirmation before renumbering.

**Recommendation:** Display warning to user: "Inserting here may require manual renumbering."

---

## Structure Edge Cases

### EC-03: Heading vs Body Paragraph Targeting

**Problem:** Many clauses have TWO paragraphs: a heading ("5. LIMITATION OF LIABILITY") and a body ("Neither party shall be liable..."). AI must target the body, not the heading, when amending content.

**Detection:**
```typescript
// Heading: { id: 22, style: "Heading 3", charCount: 26 }
// Body:    { id: 23, style: "List Paragraph", charCount: 245 }
```

**Solution:** Include ALL paragraphs in contract map with style metadata. AI uses style to distinguish headings from body paragraphs.

**Implementation:** [contractMap.ts](file:///home/sarturko/vibe-legal-react/src/services/document/contractMap.ts) — `buildContractMapContext()` outputs all paragraphs with `{Style: X} {List: Lvl Y}`

**Test Case:** "Amend clause 5" should target P23 (body), not P22 (heading).

---

### EC-04: Multi-Paragraph Clauses

**Problem:** Some clauses span multiple paragraphs (e.g., a clause with sub-points a, b, c). AMEND with scope CLAUSE must capture all related paragraphs.

**Solution:** When `scope: "CLAUSE"`, expand target range to include all subsequent paragraphs at higher indent levels.

**Implementation:** Operation handlers with `end_id` support

**Test Case:** "Delete clause 4.1 entirely" should remove main clause and all sub-points.

---

### EC-05: Block vs Inline Clause Structure

**Problem:** Documents use different patterns:
- **Block:** Heading on own line, body on next paragraph
- **Inline:** Heading and body combined

**Solution:** Detect pattern from document snapshot. Match existing pattern when inserting.

**Implementation:** Document analysis in AI prompt, INSERT handlers

---

### EC-06: Signature Block Detection

**Problem:** New clauses should be inserted BEFORE the signature block.

**Solution:** System prompt instructs AI about standard contract order. Contract map flags signature paragraphs.

**Implementation:** System prompt, contract map builder

---

### EC-07: Table Handling

**Problem:** Tables have different OOXML structure. Standard paragraph operations don't work.

**Solution:** Use "surgical mode" for table cells. Preserve `<w:tc>` wrapper structure.

**Implementation:** Mode router, OOXML processor

---

## Formatting Edge Cases

### EC-08: Style Inheritance

**Problem:** Word API's `styleId` often returns `undefined`. Use `p.style` (display name) instead.

**Implementation:** [handleAction.ts](file:///home/sarturko/vibe-legal-react/src/services/handleAction.ts) — style cache and resolution

---

### EC-09: Font Inheritance

**Problem:** Inserted text may use wrong font due to document defaults.

**Solution:** "Sandwich" architecture — insert via API, capture OOXML, add track changes.

**Implementation:** [insertBlocks.ts](file:///home/sarturko/vibe-legal-react/src/services/ooxml/insertBlocks.ts)

---

### EC-10: Markdown Parsing

**Problem:** AI returns `**bold**`, `*italic*` markers that must convert to OOXML.

**Implementation:** [markdownParser.ts](file:///home/sarturko/vibe-legal-react/src/services/formatting/markdownParser.ts)

---

### EC-11: Run-Level Formatting Preservation

**Problem:** Legal documents have character-level formatting. Unchanged portions must retain original formatting.

**Solution:** Character-level diff with property mapping.

**Implementation:** [applyRedline.ts](file:///home/sarturko/vibe-legal-react/src/services/ooxml/applyRedline.ts)

---

### EC-12: Mixed Formatting in Headings

**Problem:** Some headings mix bold/regular formatting.

**Solution:** Preserve existing runs when amending.

---

## List Edge Cases

### EC-13: Bullet Point Insertion

**Problem:** New bullets require correct `numId` and `ilvl`.

**Solution:** Clone `<w:numPr>` from reference paragraph.

**Implementation:** [insertBlocks.ts](file:///home/sarturko/vibe-legal-react/src/services/ooxml/insertBlocks.ts)

---

### EC-14: List Level Preservation

**Problem:** Nested lists have different `ilvl` values.

**Solution:** AI specifies `list_level`, code finds reference at that level.

---

### EC-15: List Continuation

**Problem:** After non-list paragraph, continuing previous list requires finding correct `numId`.

**Solution:** Search backwards for matching list paragraph.

---

### EC-16: Multiple Independent Lists

**Problem:** Document may have multiple lists with different `numId` values.

**Solution:** Always find reference paragraph NEAR insertion point.

---

## Track Changes Edge Cases

### EC-17: Insertion Wrapper (`<w:ins>`)

**Implementation:** [trackChanges.ts](file:///home/sarturko/vibe-legal-react/src/services/ooxml/trackChanges.ts)

---

### EC-18: Deletion Wrapper (`<w:del>`)

**Implementation:** [trackChanges.ts](file:///home/sarturko/vibe-legal-react/src/services/ooxml/trackChanges.ts)

---

### EC-19: RSID Management

**Problem:** Word uses RSID to track editing sessions.

**Solution:** Generate consistent RSID per session.

---

### EC-20: Nested Track Changes

**Problem:** Avoid nesting `<w:ins>` inside `<w:ins>`.

**Solution:** Detect existing track changes, create new revision.

---

## Map & Targeting Edge Cases

### EC-23: Contract Map Completeness

**Problem:** Contract map only contained heading paragraphs, missing body paragraphs.

**Solution:** Build map with ALL paragraphs including style and list level:
```typescript
for (const p of contractMap.paragraphs) {
    let meta = `{Style: ${p.style || 'Normal'}}`;
    if (p.isListItem && p.listLevel !== undefined) {
        meta += ` {List: Lvl ${p.listLevel}}`;
    }
    paragraphMapLines.push(`[${p.id}] ${meta} ${text}`);
}
```

**Implementation:** [contractMap.ts](file:///home/sarturko/vibe-legal-react/src/services/document/contractMap.ts) — `buildContractMapContext()`

**Test Case:** Document with 36 paragraphs → map should show all 36 with full metadata.

---

### EC-24: Paragraph ID Consistency

**Problem:** Paragraph IDs must be consistent between map and execution.

**Solution:** Use `context.trackedObjects.add(p)` to keep references alive in single `Word.run()`.

**Implementation:** [handleAction.ts](file:///home/sarturko/vibe-legal-react/src/services/handleAction.ts) — tracked objects pattern

---

### EC-25: Content vs Map Separation

**Decision:** Keep content and map separate:
- Document content: Full text for AI understanding
- Structure map: Truncated text (120 chars) for targeting

---

### EC-26: Cross-Reference Awareness

**Problem:** Amending clause numbers may break cross-references.

**Solution:** Detect cross-references, warn user when editing referenced clauses.

**Implementation:** Partial — cross-reference detection regex

---

## Implementation Status

| EC | Status | Location |
|----|--------|----------|
| EC-01 | ✅ | insertBlocks.ts |
| EC-02 | ✅ | contractMap.ts |
| EC-03 | ✅ | contractMap.ts - buildContractMapContext |
| EC-04 | ✅ | Operation handlers |
| EC-05 | ✅ | AI prompt, INSERT handlers |
| EC-06 | ✅ | System prompt |
| EC-07 | ✅ | Mode router |
| EC-08 | ✅ | handleAction.ts |
| EC-09 | ✅ | insertBlocks.ts |
| EC-10 | ✅ | markdownParser.ts |
| EC-11 | ✅ | applyRedline.ts |
| EC-12 | ⚠️ Partial | Heading handlers |
| EC-13 | ✅ | insertBlocks.ts |
| EC-14 | ✅ | INSERT handler |
| EC-15 | ⚠️ Partial | Reference finding |
| EC-16 | ✅ | Proximity-based finding |
| EC-17 | ✅ | trackChanges.ts |
| EC-18 | ✅ | trackChanges.ts |
| EC-19 | ⚠️ Partial | RSID utilities |
| EC-20 | ❌ | Not implemented |
| EC-21 | ✅ | applyRedline.ts - restoreManualNumber |
| EC-22 | ❌ Deferred | Renumbering system |
| EC-23 | ✅ | contractMap.ts - buildContractMapContext |
| EC-24 | ✅ | handleAction.ts - tracked objects |
| EC-25 | ✅ | Architecture decision |
| EC-26 | ⚠️ Partial | Cross-reference detection |

---

## Consequences

**Positive:**
- Comprehensive documentation prevents regression
- New contributors understand edge case handling
- Test cases provide verification path

**Negative:**
- Maintenance burden to keep ADR updated
- Some edge cases may be implementation-specific

**Mitigations:**
- Update ADR as part of PR process for edge case changes
