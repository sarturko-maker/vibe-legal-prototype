# ADR-003: Track Changes Implementation

**Status:** Accepted  
**Date:** 2025-12-28  
**Authors:** Artur (Vibe Legal)  
**Depends On:** ADR-002 (Track Changes Strategy)

---

## Context

ADR-002 established that we use Word's native track changes for the add-in. This ADR documents **how** each operation type is implemented.

The core challenge: When the AI suggests a change like "replace 'thirty days' with 'sixty days'", we need to:

1. Find exactly where that text is in the paragraph
2. Replace only those words (not the whole paragraph)
3. Have Word show it as a proper track change

### Why This Matters

If we replace the entire paragraph, Word shows the whole thing as deleted and re-inserted — not useful for a lawyer reviewing changes. We need **surgical precision**: only the changed words should be marked.

---

## Decision

Implement three operation types, each with its own approach:

| Operation | What It Does | Implementation |
|-----------|--------------|----------------|
| **AMEND** | Change text within a paragraph | Diff algorithm + search/replace |
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

### The AMEND Flow

```
1. Get original paragraph text from Word
2. Get amended text from AI response
3. Run diff-match-patch to find all changes
4. Turn on track changes
5. For each change (in reverse order):
   - Search for the "find" text in the paragraph
   - Replace with the "replace" text
6. Turn off track changes (restore original state)
```

### Why Reverse Order?

Changes are applied from end to start. If we applied them start to end, each replacement would shift the position of subsequent text, causing later searches to fail.

**Example:** Paragraph has changes at position 10 and position 50.
- If we change position 10 first, and the new text is longer, position 50 is now at position 55.
- By going in reverse (50 first, then 10), earlier positions aren't affected.

### Context Wrapping

The diff algorithm adds surrounding context to each find/replace pair. This prevents false matches.

**Example:**

If the AI changes "the" to "a" somewhere, we don't want to replace every "the" in the paragraph. By including context:

- Find: `"pay the Buyer"` 
- Replace: `"pay a Buyer"`

We only match the specific instance.

### Fallback: Full Paragraph Replacement

If the diff approach fails (e.g., text not found due to formatting differences), we fall back to replacing the entire paragraph. This is less elegant but ensures the change is applied.

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

## Key Files

| File | Purpose |
|------|---------|
| `src/utils/textDiff.ts` | diff-match-patch wrapper, `findMinimalChanges()` |
| `src/services/handleAction.ts` | Operation handlers: `handleAmendOperation()`, `handleInsertOperation()`, `handleDeleteOperation()` |

---

## Consequences

### Benefits

- **Surgical changes** — Only changed words are marked, making review easy
- **Deterministic** — Same input always produces same output (unlike AI-generated find/replace)
- **State respectful** — Doesn't interfere with user's track changes preference

### Limitations

- **Formatting markers** — If the AI returns markdown (`**bold**`), we strip it before diffing. Formatting is inherited, not created from markdown.
- **Complex restructuring** — If AI completely rewrites a paragraph, the diff may show many small changes rather than one clean replacement.
- **Numbered lists** — Inserting between list items sometimes causes Word to renumber unexpectedly.

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

## Verification Points for Code Review

1. `textDiff.ts` — Confirm diff-match-patch is used, not AI-generated find/replace
2. `handleAmendOperation()` — Changes applied in reverse order
3. All operations — State preservation pattern (detect → enable → operate → restore)
4. `normalizeSpacing()` — Punctuation cleanup before applying

---

## Related

- ADR-002: Track Changes Strategy (why native approach)
- ADR-004: Paragraph Identification (how we target the right paragraphs)
- ADR-011: Edge Case Handling (detailed edge cases)
