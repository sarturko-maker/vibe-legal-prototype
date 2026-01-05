# ADR-005: Style Preservation

**Status:** Accepted  
**Date:** 2025-12-28  
**Authors:** Artur (Vibe Legal)  
**Depends On:** ADR-002, ADR-003

---

## Context

Legal documents have precise formatting requirements. A contract might use:

- **Heading 2** for clause titles (bold, 12pt)
- **List Paragraph** for numbered clauses (indented, numbered)
- **Normal** for body text (11pt, justified)
- Custom fonts like Times New Roman or specific firm styles

When AI inserts or amends text, it must match the surrounding formatting. If a new clause appears in Comic Sans when the rest of the document is Times New Roman, the lawyer will reject it immediately.

### The Challenge

Word documents have multiple layers of formatting:

| Layer | Examples |
|-------|----------|
| **Paragraph styles** | Heading 1, Normal, List Paragraph |
| **Character formatting** | Bold, italic, underline |
| **Font properties** | Font name, size, colour |
| **List properties** | Numbering, bullets, indent level |

All must be preserved or correctly applied for professional output.

---

## Decision

Implement a multi-layered style preservation system:

1. **Style Menu** — Scan document styles on load, build a reference list
2. **Style Normalisation** — Map AI style names to document-specific names
3. **Empty-First Insertion** — Apply formatting before adding tracked content
4. **List Inheritance** — Copy numbering properties from reference paragraphs

---

## Style Menu

When we load a document, we build a "style menu" by scanning all paragraphs:

```
Style Menu:
- "Heading 2" (used 12 times, bold, 14pt, example: P3)
- "List Paragraph" (used 45 times, 11pt, example: P7)
- "Normal" (used 8 times, 11pt, example: P2)
```

This tells us:
- What styles exist in this document
- How each style looks (font, size, bold/italic)
- Example paragraphs we can reference

### Why Scan Rather Than Assume?

Every firm has different templates. "Heading 2" in one document might be 14pt Arial bold. In another, it's 12pt Times New Roman underlined. We can't assume — we must observe.

---

## Style Normalisation

The AI might say "use heading style" but the document has "Heading 2". We normalise AI responses to match document reality.

### The Normalisation Function

```
AI says: "h3" → We look for: "Heading 3"
AI says: "body" → We look for: "Normal" or "Body Text"
AI says: "clause heading" → We look for: "Heading 3"
```

Built-in aliases:

| AI Might Say | Maps To |
|--------------|---------|
| h1, title, heading1 | Heading 1 |
| h2, heading2 | Heading 2 |
| h3, clause heading | Heading 3 |
| body, body text, paragraph, text, standard | Normal |

If no match is found, we return `null` and fall back to inheriting from the reference paragraph.

---

## Empty-First Insertion Strategy

When inserting a new paragraph, we discovered that adding content with formatting in one step caused problems. The solution:

### The Three-Step Pattern

```
STEP 1: Insert empty paragraph (structure only)
        → No track changes yet
        → Just creates the paragraph container

STEP 2: Apply formatting to empty paragraph
        → Set paragraph style
        → Set font properties
        → Set list properties
        → Still no track changes

STEP 3: Turn on track changes, add content
        → Now insert the actual text
        → Word sees this as "new content" (tracked)
        → Formatting was already applied, so it's inherited
```

### Why This Works

Word tracks *content* changes, not *formatting* changes. By setting up the formatting first on an empty paragraph, then adding content with tracking on, we get:

- Content marked as an insertion (blue underline)
- Formatting that matches the document
- No formatting artefacts in the track changes

---

## List Formatting Inheritance

Numbered lists are the trickiest part. Word stores list properties (`numPr`) separately from paragraph properties.

### When Inserting Into a List

1. Check if reference paragraph `isListItem`
2. If yes, and AI specifies `list_level`:
   - Load the new paragraph's `listItem`
   - Set `listItem.level` to match
3. Word automatically continues the numbering

### The AI's Role

The AI is instructed to include `list_level` when inserting into lists:

```json
{
  "type": "INSERT",
  "insert_after": 20,
  "content": "The Goods shall be free from defects.",
  "list_level": 0
}
```

If `list_level` is omitted, the paragraph inserts as plain text — not a list item.

### Manual Number Preservation

Some documents use manually typed numbers ("3.1", "3.2") instead of Word's automatic numbering. For these:

- We detect the pattern in the original text
- If AI strips the number in its response, we restore it
- `restoreManualNumber()` handles this post-processing

---

## AMEND Formatting Preservation

When amending existing text, formatting preservation is simpler because we're modifying in place:

1. **Unchanged text keeps its formatting** — The diff algorithm only changes specific words
2. **New text inherits from context** — Word applies the surrounding run's formatting
3. **Paragraph style stays intact** — We're not replacing the paragraph, just its content

The surgical approach (ADR-003) ensures we don't accidentally destroy formatting by replacing entire paragraphs.

---

## Handling AI Markdown

AI sometimes returns markdown formatting:

- `**bold text**` for defined terms
- `*italic text*` for emphasis

### Current Approach

We strip markdown before applying to Word:

```
"The **Buyer** shall pay" → "The Buyer shall pay"
```

Formatting is inherited from the document, not created from markdown.

### Why Not Convert Markdown to Word Formatting?

Complexity vs. value. Converting `**text**` to actual bold would require:
- Parsing markdown
- Splitting text into runs
- Applying character formatting to each run

For MVP, we decided inherited formatting is sufficient. Future enhancement could add markdown conversion.

---

## Key Files

| File | Purpose |
|------|---------|
| `src/services/handleAction.ts` | Style menu building, `normalizeStyleName()`, `handleInsertOperation()` |
| `src/services/formatting/markdownParser.ts` | `stripFormattingMarkers()` — removes markdown before Word operations |
| `vibe-style-poc/styleManager.js` | Proof-of-concept style extraction |

---

## Consequences

### Benefits

- **Professional output** — Inserted text matches document style
- **No manual cleanup** — Lawyers don't need to fix fonts after AI edits
- **List continuity** — Numbered lists stay numbered
- **Document integrity** — Styles aren't corrupted by AI operations

### Limitations

- **No markdown to formatting** — AI can't make text bold via `**text**`
- **Style must exist** — Can't create new styles, only use existing ones
- **List edge cases** — Some complex nested lists may not inherit correctly

### Known Issues

- **Numbered list insertion** — Inserting between list items sometimes causes Word to renumber unexpectedly. This is a Word API limitation, not fully solved.
- **Custom styles** — If a document uses unusual style names, AI might not recognise them. Aliases help but don't cover everything.

---

## Style Override Pattern

For operations that need specific formatting, the AI can specify explicit properties:

```json
{
  "type": "INSERT",
  "insert_after": 10,
  "content": "INDEMNIFICATION",
  "style": "Heading 2",
  "bold": true,
  "font": "Times New Roman",
  "fontSize": 14
}
```

The handler applies these in order:
1. Paragraph style (if specified)
2. Font name (if specified)
3. Font size (if specified)
4. Bold/italic/underline (if specified)

This allows precise control when needed while defaulting to inheritance when not.

---

## Verification Points for Code Review

1. `handleAction.ts` — Style menu built from document scan
2. `normalizeStyleName()` — Aliases map AI terms to document styles
3. `handleInsertOperation()` — Empty paragraph created first, formatted, then content added
4. List operations — Check `isListItem` and set `list_level`
5. `stripFormattingMarkers()` — Markdown removed before Word operations

---

## Related

- ADR-003: Track Changes Implementation (formatting context for AMEND)
- ADR-004: Paragraph Identification (style metadata in contract map)
- ADR-011: Edge Case Handling (EC-08 through EC-11 cover formatting edge cases)
