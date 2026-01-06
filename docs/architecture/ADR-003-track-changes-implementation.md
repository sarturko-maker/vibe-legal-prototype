# ADR-003: Track Changes Implementation

**Status:** Accepted  
**Date:** 2025-12-28 (Updated: 2026-01-05)  
**Authors:** Artur (Vibe Legal)  
**Depends On:** ADR-002 (Track Changes Strategy)

---

## Context

ADR-002 established that we use Word's native track changes for the add-in. This ADR documents **how** each operation type is implemented.

The core challenge: When the AI suggests a change like "replace 'thirty days' with 'sixty days'", we need to:

1. Find exactly where that text is in the paragraph
2. Delete the old text and insert the new text
3. Have Word show it as proper track changes (strikethrough + underline)

### Why This Matters

If we replace the entire paragraph, Word shows the whole thing as deleted and re-inserted — not useful for a lawyer reviewing changes. We need **surgical precision**: only the changed words should be marked.

---

## Decision

Implement three operation types, each with its own approach:

| Operation | What It Does | Implementation |
|-----------|--------------|----------------|
| **AMEND** | Change text within a paragraph | Diff algorithm + DELETE_AFTER/INSERT_AFTER pairs |
| **INSERT** | Add a new paragraph | Insert empty paragraph, then add text |
| **DELETE** | Remove a paragraph | Delete with tracking on |

---

## AMEND Implementation

AMEND is the most complex operation. The AI provides the amended version of a paragraph, and we need to apply only the changes.

### The Diff Algorithm

We use Google's **diff-match-patch** library to compare the original text with the AI's amended version. This finds the minimal set of changes.

**Example:**

Original: `"The Buyer shall pay within thirty (30) days of invoice."`

Amended: `"The Buyer shall pay within sixty (60) days of invoice."`

The diff algorithm identifies:
- Delete: `"thirty (30)"`
- Insert: `"sixty (60)"`

Everything else stays untouched.

### Why diff-match-patch?

| Alternative | Problem |
|-------------|---------|
| Word-level diff | Splits "30" and "days" incorrectly; loses punctuation context |
| AI-generated find/replace | Non-deterministic; AI might hallucinate different text |
| Full paragraph replacement | Shows entire paragraph as changed; poor UX |

diff-match-patch uses character-level comparison with semantic cleanup, so "4.1" stays together as a unit, and punctuation is handled correctly.

### The AMEND Flow (Anchor-Based DELETE + INSERT)

```
1. Get original paragraph text via getReviewedText()
2. Get amended text from AI response
3. Run diff-match-patch to find all changes
4. Merge adjacent micro-changes into larger operations
5. For each change, find a unique anchor (3+ chars, appears exactly once)
6. Create DELETE_AFTER + INSERT_AFTER operation pairs
7. Skip operations that cannot find unique anchors
8. Turn on track changes
9. Apply operations in reverse order (last to first)
10. Turn off track changes (restore original state)
```

**Key principles:**
- All operations use DELETE + INSERT (never REPLACE)
- All operations share the same baseline text snapshot
- Anchors must be unique or operation is skipped
- Automatic spacing added to INSERTs

### Why Reverse Order?

Changes are applied from end to start. If we applied them start to end, each replacement would shift the position of subsequent text, causing later searches to fail.

**Example:** Paragraph has changes at position 10 and position 50.
- If we change position 10 first, and the new text is longer, position 50 is now at position 55.
- By going in reverse (50 first, then 10), earlier positions aren't affected.

---

## Anchor Uniqueness Requirements

Every text change requires an **anchor** — a piece of text that tells Word exactly where to apply the change.

**The anchor must be:**
- Unique (appears exactly 1 time in the paragraph)
- At least 3 characters long (excluding whitespace)
- Located before the change position

**Finding unique anchors:**

```typescript
1. Start with 5 words before the change
2. Check if those 5 words appear exactly once
3. If not unique, expand to 7 words, then 9, then 11...
4. Continue expanding up to 50 words
5. If still not unique after 50 words, skip the operation
```

**Example:**

```
Paragraph: "The Seller shall deliver Products within thirty days"
Change: "thirty" → "sixty"

Anchor search:
- Try "deliver Products within thirty" (5 words) → Unique! ✓
- Use this as anchor for DELETE + INSERT
```

**Why this matters:**

Single-character anchors like "(" can appear multiple times, especially after previous operations add text with parentheses. By requiring 3+ character anchors and testing for uniqueness, we prevent corruption from operations matching the wrong location.

---

## Automatic INSERT Spacing

When inserting text after an anchor, we automatically add a leading space if needed to prevent text running together.

**The logic:**

```typescript
if (!anchor.endsWith(' ') && !textToInsert.startsWith(' ')) {
  textToInsert = ' ' + textToInsert;  // Add leading space
}
```

**Examples:**

| Anchor | Text | Result | Reason |
|--------|------|--------|--------|
| `"made"` | `"DAP"` | `" DAP"` | Neither has space, add one |
| `"made "` | `"DAP"` | `"DAP"` | Anchor ends with space, don't add |
| `"made"` | `" DAP"` | `" DAP"` | Text starts with space, don't add |

This prevents corruption like:
- "madeDAP" (should be "made DAP")
- "followingthe" (should be "following the")

---

## Diff Merging

Adjacent micro-changes are automatically merged into larger operations.

**Problem without merging:**
```
Original: "thirty (30) Days"
Amended:  "sixty (60) days"

Raw diff produces 6+ operations:
- Delete "t", Insert "s"
- Delete "h", Insert "i"
- Delete "irty", Insert "xty"
- Delete "3", Insert "6"
- Delete "D", Insert "d"
```

**With merging:**
```
Merged diff produces 2 operations:
- Delete "thirty (30) Days"
- Insert "sixty (60) days"
```

Merging creates larger, more anchorable chunks and reduces operation count.

---

## Operation Skipping Strategy

When a change cannot be anchored uniquely, we **skip that operation** rather than risk corruption.

**Operations are skipped when:**
- No unique anchor can be found (after trying up to 50 words)
- Anchor would be shorter than 3 characters
- Text to change appears multiple times in paragraph

**Logging:**
```
[textDiff] SKIPPING operation - could not find unique anchor
[ANCHOR DEBUG] Could not find unique anchor after 50 words
```

**Impact:**
- Most skipped operations are micro-changes (1-2 characters)
- Important changes use longer phrases that anchor reliably
- Better to skip a cosmetic change than corrupt the document

**Example:**

If changing "t" → "a" or "5" → "10", and these characters appear multiple times, we skip the operation. The user sees 6 of 7 changes applied, which is better than corrupted text.

---

## INSERT Implementation

INSERT adds a new paragraph after a specified location.

### The "Empty First" Strategy

We discovered that inserting a paragraph with content directly sometimes caused formatting issues. The solution:

1. Insert an **empty** paragraph after the reference
2. Apply the correct style to the empty paragraph
3. **Then** turn on track changes
4. Insert the text into the empty paragraph
5. Turn off track changes

This way, Word correctly marks the text as an insertion, and the paragraph inherits the right formatting.

### List Item Handling

If the reference paragraph is a list item (numbered or bulleted), we need the new paragraph to continue the list. We copy the list formatting from the reference:

- Check if reference `isListItem`
- If yes, the new paragraph inherits the list formatting
- Word handles renumbering automatically

---

## DELETE Implementation

DELETE is the simplest operation.

### Flow

1. Find the target paragraph by ID
2. Turn on track changes
3. Call `paragraph.delete()`
4. Turn off track changes

Word automatically shows the deleted paragraph as struck-through red text.

### Fallback: Content Matching

If the paragraph ID isn't found (can happen if document structure changed), we fall back to searching by content — looking for a paragraph that matches the expected text.

---

## State Preservation Pattern

All three operations follow the same pattern for track changes state:

```typescript
1. Load current tracking mode
2. Remember if user already had tracking on
3. Force tracking on
4. ... do the operation ...
5. If user didn't have tracking on, turn it off
6. If user had tracking on, leave it on
```

This respects the user's preference. If they're working with track changes enabled, we don't interfere.

---

## Spacing and Punctuation Cleanup

The diff algorithm can produce artifacts like:

- `"losses ."` instead of `"losses."`
- `"A ll"` instead of `"All"` (split words)

We normalise these before applying:

```
- Remove spaces before punctuation: " ." → "."
- Fix split words: "A ll" → "All"  
- Collapse multiple spaces
```

---

## Key Files

| File | Purpose |
|------|---------|
| `src/utils/textDiff.ts` | diff-match-patch wrapper, anchor finding, operation creation, diff merging |
| `src/services/handleAction.ts` | Operation handlers: `handleAmendOperation()`, `handleInsertOperation()`, `handleDeleteOperation()` |
| `src/services/handleAction.ts:applyTextChanges()` | Applies DELETE_AFTER and INSERT_AFTER operations with automatic spacing |

---

## Consequences

### Benefits

- **Surgical changes** — Only changed words are marked, making review easy
- **Deterministic** — Same input always produces same output (unlike AI-generated find/replace)
- **State respectful** — Doesn't interfere with user's track changes preference
- **Corruption-resistant** — Operations skip rather than corrupt when anchors aren't unique

### Limitations

- **Formatting markers** — If the AI returns markdown (`**bold**`), we strip it before diffing. Formatting is inherited, not created from markdown.
- **Complex restructuring** — If AI completely rewrites a paragraph, the diff may show many small changes rather than one clean replacement.
- **Numbered lists** — Inserting between list items sometimes causes Word to renumber unexpectedly.
- **Micro-change skipping** — Very small changes (1-2 characters) may be skipped if they cannot be anchored uniquely. This is **intentional** to prevent corruption.
- **Operation count** — A single paragraph change may generate 2-6 operations (DELETE + INSERT pairs). Complex changes create more operations.
- **Sequential operation dependency** — Operations applied to the same paragraph can affect each other. Example: First operation adds "(INDEMNIFICATION)", later operation using "(" as anchor may match wrong location. The 3-character minimum rule mitigates this.
- **Anchor expansion overhead** — Finding unique anchors requires testing up to 50 word combinations. This is fast but adds computational cost.

---

## Verification Points for Code Review

1. `textDiff.ts` — Confirm diff-match-patch is used to find changes
2. `textDiff.ts` — Confirm anchors are tested for uniqueness (appear exactly once)
3. `textDiff.ts` — Confirm minimum 3-character anchor length enforced
4. `textDiff.ts` — Confirm operations are skipped (not applied) when no unique anchor found
5. `textDiff.ts` — Confirm adjacent micro-changes are merged
6. `handleAmendOperation()` — Changes applied in reverse order (last to first)
7. `applyTextChanges()` — INSERT operations automatically add leading space when needed
8. All operations — State preservation pattern (detect → enable → operate → restore)

---

## Related

- ADR-002: Track Changes Strategy (why native approach)
- ADR-004: Paragraph Identification (how we target the right paragraphs)
- ADR-011: Edge Case Handling (detailed edge cases)
