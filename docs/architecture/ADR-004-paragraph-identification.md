# ADR-004: Paragraph Identification

**Status:** Accepted  
**Date:** 2025-12-28  
**Authors:** Artur (Vibe Legal)  
**Depends On:** ADR-002, ADR-003

---

## Context

For AI to suggest changes to a contract, it needs to know:

1. **What's in the document** — The structure and content of each clause
2. **How to target specific parts** — A way to say "change paragraph 23" unambiguously

Without a reliable identification system, the AI might suggest changing "clause 5" but we wouldn't know which exact paragraph to modify.

### The Challenge

Word documents don't have built-in paragraph IDs. Paragraphs are just a sequence — they have positions, but those positions change as content is added or removed.

---

## Decision

Implement a **Contract Map** system that:

1. Assigns every paragraph a sequential ID (`P1`, `P2`, `P3`...)
2. Captures metadata about each paragraph (style, list status, text preview)
3. Sends this map to the AI as context
4. AI uses `[P#]` notation to target specific paragraphs in its response

### The Contract Map

When we read the document, we build a map like this:

```
PARAGRAPH MAP:
[P1] {Style: Title} CONFIDENTIALITY AGREEMENT
[P2] {Style: Normal} This Confidentiality Agreement is entered into...
[P3] {Style: Heading 2} 1. DEFINITIONS
[P4] {Style: List Paragraph} {List: Lvl 0} "Confidential Information" means...
[P5] {Style: List Paragraph} {List: Lvl 0} "Disclosing Party" means...
[P6] {Style: Heading 2} 2. OBLIGATIONS
[P7] {Style: List Paragraph} {List: Lvl 0} The Receiving Party shall...
...
```

Each line includes:
- **`[P#]`** — The paragraph ID (1-indexed)
- **`{Style: X}`** — The Word style applied (helps AI distinguish headings from body)
- **`{List: Lvl Y}`** — If it's a list item and what indent level
- **Text preview** — First 120 characters (truncated for brevity)

### How AI Uses It

The AI receives this map and uses `[P#]` notation in its operations:

```json
{
  "action": "AMEND",
  "target_id": 7,
  "original_text": "The Receiving Party shall protect...",
  "amended_text": "The Receiving Party shall use reasonable efforts to protect..."
}
```

The `target_id: 7` tells us exactly which paragraph to modify.

---

## Key Design Decisions

### 1. All Paragraphs Included

Early versions only included "important" paragraphs (headings, clause starts). This caused problems — the AI couldn't target body paragraphs.

**Solution:** Include ALL paragraphs in the map, even empty ones. The map is truncated (120 chars per paragraph) to keep token usage reasonable.

### 2. 1-Indexed IDs

Paragraph IDs start at 1, not 0. This matches how humans think ("paragraph 1") and avoids off-by-one confusion when lawyers review AI output.

Internally, we convert to 0-indexed when accessing the array:
```
arrayIndex = paragraphId - 1
```

### 3. Style Metadata

Including the style (`Heading 2`, `List Paragraph`, `Normal`) helps the AI understand document structure:

- **Headings** — Usually clause titles, shouldn't be amended for content
- **Body paragraphs** — Where the substantive text lives
- **List items** — May be part of a numbered list, formatting matters

### 4. Separate Map from Content

We send two things to the AI:

| Data | Purpose | Size |
|------|---------|------|
| **Document text** | Full content for understanding | Up to 15,000 chars |
| **Paragraph map** | Truncated structure for targeting | ~100 chars per paragraph |

This separation keeps the map scannable while giving AI full context.

---

## ID Echo Protocol

For operations that return modified text (AMEND across multiple paragraphs), we use the **ID Echo Protocol**:

### Sending to AI

We wrap each paragraph with ID tags:

```
<P7>The Receiving Party shall protect all Confidential Information.</P7>
<P8>The Receiving Party shall not disclose Confidential Information to third parties.</P8>
```

### Receiving from AI

The AI returns modified content with the same tags:

```
<P7>The Receiving Party shall use reasonable efforts to protect all Confidential Information.</P7>
<P8>The Receiving Party shall not disclose Confidential Information to any third party without prior written consent.</P8>
```

### Why This Matters

With tagged input and output:
- We know **exactly** which paragraph maps to which
- No fuzzy matching required
- If AI removes a tag entirely, we know that paragraph should be deleted
- If AI adds untagged text, we know it's a new insertion

This makes the system **deterministic** — same input always produces reliably targetable output.

---

## Handling ID Drift ("The Earthquake Effect")

When paragraphs are inserted or deleted, all subsequent IDs shift. This is called the "Earthquake Effect."

**Example:**
- Original: P1, P2, P3, P4, P5
- Insert new paragraph after P2
- Now: P1, P2, **P3(new)**, P4(was P3), P5(was P4), P6(was P5)

### Solution: Fresh Lookup Per Operation

We rebuild the paragraph lookup before each operation:

1. Load current paragraphs from Word
2. Build fresh ID → paragraph mapping
3. Execute the operation
4. Sync with Word

For sequential operations (multiple inserts), we refresh between each one.

### Tracked Objects Pattern

Word's API can "lose" paragraph references across sync boundaries. We use `trackedObjects`:

```typescript
context.trackedObjects.add(paragraph);
// ... do work ...
context.trackedObjects.remove(paragraph);
```

This keeps paragraph references alive during complex operations.

---

## Key Files

| File | Purpose |
|------|---------|
| `src/services/document/contractMap.ts` | `buildContractMapContext()` — Creates the map |
| `src/services/document/wordApi.ts` | `loadParagraphs()` — Reads paragraphs from Word |
| `src/services/handleAction.ts` | Builds lookup table, handles tracked objects |

---

## Consequences

### Benefits

- **Precise targeting** — AI can specify exact paragraphs, no ambiguity
- **Deterministic matching** — ID Echo Protocol eliminates fuzzy matching errors
- **Full coverage** — Every paragraph is addressable, not just headings
- **Debugging** — Clear IDs make it easy to trace what went wrong

### Limitations

- **Token overhead** — Map takes space in the AI prompt
- **Stale IDs** — If document changes between map build and execution, IDs may be wrong
- **Large documents** — Documents with 200+ paragraphs may hit token limits

### Mitigations

- Truncate text to 120 chars per paragraph
- Rebuild map before each operation sequence
- Cache map with document hash to detect staleness

---

## Map Staleness Detection

We cache the contract map with:

- `documentHash` — Hash of the full document text
- `totalParagraphCount` — Number of paragraphs

If either changes, we invalidate the cache and rebuild:

```typescript
if (contractMapState.documentHash !== currentHash ||
    contractMapState.totalParagraphCount !== currentParagraphCount) {
    // Rebuild map
}
```

This catches both content changes and structural changes.

---

## Verification Points for Code Review

1. `contractMap.ts` — All paragraphs included, not just headings
2. `wordApi.ts` — Paragraph IDs are 1-indexed
3. `handleAction.ts` — Lookup rebuilt before operations
4. `handleAction.ts` — `trackedObjects` used for paragraph references
5. Prompt includes `PARAGRAPH MAP:` section with `[P#]` notation

---

## Related

- ADR-003: Track Changes Implementation (uses paragraph IDs for targeting)
- ADR-011: Edge Case Handling (EC-23, EC-24, EC-25 cover map issues)
