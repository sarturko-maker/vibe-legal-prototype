# ADR-002: Track Changes Strategy

**Status:** Accepted  
**Date:** 2025-12-27  
**Authors:** Artur (Vibe Legal)  
**Supersedes:** Earlier OOXML-only approach

---

## Context

Vibe Legal needs to apply AI-suggested changes to Word documents in a way that lawyers can review — meaning track changes (redlines) must be visible. The user sees deleted text struck through in red and new text underlined in blue, just like any legal redline.

The fundamental question: **How do we create track changes in a Word document?**

### Two Possible Approaches

| Approach | How It Works |
|----------|--------------|
| **OOXML Manipulation** | Open the document's internal structure, manually insert `<w:ins>` and `<w:del>` tags around changed text |
| **Native Track Changes** | Tell Word to turn on track changes, make edits through Word's own tools, let Word handle the markup |

---

## The Journey

### Phase 1: OOXML Manipulation (Original Approach)

We started by manipulating the document's internal structure directly. A Word document is actually a zip file containing XML files. Track changes are represented as:

- `<w:ins>` tags — wrapping inserted text
- `<w:del>` tags — wrapping deleted text  
- `<w:delText>` — the actual deleted content

This seemed like the "proper" way — it's how commercial legal tech tools work.

**What we built:**

- Surgical mode: Change individual text runs without touching paragraph structure
- Reconstruction mode: Rebuild paragraphs with track change markup
- Paragraph-aware mode: Match paragraphs between original and modified versions
- Extensive edge case handling (see ADR-011)

**Why it was difficult:**

| Problem | Description |
|---------|-------------|
| Numbering breaks | Word's automatic numbering (1, 2, 3...) is stored separately from paragraph text. Touching paragraphs can break the link. |
| Formatting loss | Character-level formatting (bold, italic) lives in "runs" — splitting runs incorrectly loses formatting |
| Style inheritance | Paragraphs inherit styles in complex ways. New paragraphs might get wrong fonts or spacing. |
| Edge cases everywhere | Tables, footnotes, bookmarks, comments, hyperlinks — each has special handling |

The complexity was enormous. Each fix created new edge cases. "Vibe coding" this was extraordinarily difficult.

### Phase 2: Discovery — Native Track Changes Work

The breakthrough came when we discovered that **Word's JavaScript API supports turning track changes on and off**:

```
context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
```

This is available in WordApi 1.4 and works in Word Online.

**The realisation:** Instead of building track changes ourselves, we could:

1. Turn on Word's native track changes
2. Make edits using Word's standard API
3. Turn track changes back off
4. Let Word handle all the markup automatically

### Phase 3: Current Approach — Native Track Changes

We now use native track changes for all operations in the Word add-in:

| Operation | How It Works |
|-----------|--------------|
| **AMEND** | Turn on tracking → find/replace text in paragraph → turn off tracking |
| **INSERT** | Turn on tracking → insert new paragraph → turn off tracking |
| **DELETE** | Turn on tracking → delete paragraph → turn off tracking |

Word automatically:
- Wraps deleted text in strikethrough formatting
- Marks inserted text as additions
- Preserves the author name and timestamp
- Handles all the complex numbering and formatting

---

## Decision

**For the Word Add-In:** Use native track changes via Word's API.

**For the Automation Product:** Use OOXML manipulation (the work wasn't wasted).

### Why Two Approaches?

| Product | Word Running? | Approach | Reason |
|---------|---------------|----------|--------|
| **Word Add-In** | Yes — user is in Word | Native track changes | Simpler, more reliable, Word handles edge cases |
| **Automation Server** | No — batch processing | OOXML manipulation | No Word application to delegate to |

The automation product processes documents without Word open (e.g., review 50 NDAs overnight). Since there's no Word application, we must build the track change markup ourselves — exactly what our OOXML code does.

---

## Key Files

### Native Track Changes (Current Add-In)

| File | Purpose |
|------|---------|
| `src/services/handleAction.ts` | INSERT, DELETE, AMEND operations using native tracking |

### OOXML Engine (For Automation Product)

| File | Purpose |
|------|---------|
| `src/utils/OxmlEngine.ts` | OOXML engine with surgical & reconstruction modes |
| `src/utils/OxmlValidator.ts` | Validates OOXML structure before applying |
| `src/utils/OxmlEngineValidation.ts` | Test cases for OOXML engine |
| `src/taskpane/taskpane.legacy.tsx` | Original OOXML engine (preserved for reference) |

---

## Consequences

### Benefits of Native Approach

- **Dramatically simpler** — hundreds of lines of edge case handling eliminated
- **More reliable** — Word handles its own complexity
- **Better formatting preservation** — Word knows how to maintain its own structure
- **Faster development** — new features don't require OOXML expertise

### Remaining Challenges

- **Numbered lists** — Still tricky even with native approach. Word sometimes renumbers unexpectedly when inserting between list items.
- **State preservation** — Must detect user's original tracking state and restore it after operations

### OOXML Work Not Wasted

The OOXML manipulation code becomes the foundation for the automation product:

- Batch document processing without Word
- Playbook-driven automated reviews
- API for workflow integrations

---

## Alternatives Considered

| Alternative | Reason for Rejection |
|-------------|---------------------|
| OOXML only | Too complex for vibe coding; endless edge cases |
| Native only (no automation product) | Leaves batch processing use case unserved |
| Third-party library | None found that handle legal document complexity well |

---

## Implementation Notes

### Enabling Track Changes

```typescript
// Detect current state
context.document.load('changeTrackingMode');
await context.sync();
const wasAlreadyTracking = context.document.changeTrackingMode === Word.ChangeTrackingMode.trackAll;

// Enable tracking
context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
await context.sync();

// ... make changes ...

// Restore original state
if (!wasAlreadyTracking) {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await context.sync();
}
```

### Verification Points for Code Review

1. `handleAction.ts` — Confirm all operations use native tracking pattern
2. No direct OOXML insertion in the add-in flow
3. State preservation: tracking mode restored after operations
4. Legacy OOXML code clearly marked and separated

---

## Related

- ADR-001: Serverless BYOK Architecture (why processing happens client-side)
- ADR-003: Paragraph Identification (how we target the right paragraphs)
- ADR-011: Edge Case Handling (documents remaining challenges)
