# ADR-008: Definitions Feature with Combined Document Analysis

## Status
Accepted

## Date
2025-12-27

## Context
Users need to see all defined terms in a contract at a glance. Legal contracts define terms throughout the document - not just in a "Definitions" clause, but also:

- Inline definitions: `ACME Corporation (the "Buyer")`
- Schedule definitions: Terms defined in appendices
- Contextual definitions: `"Confidential Information" means...`

Additionally, we already send the document to AI on load for party detection. Sending it again for definitions would be wasteful and slow.

### Requirements
1. Extract defined terms from anywhere in the document
2. Display in a simple, searchable list
3. Don't duplicate AI calls - combine with party detection
4. Keep UI simple (no document navigation for MVP)

## Decision

### 1. Combined Document Analysis
Replace `partyDetection.ts` with `documentAnalysis.ts` that returns both parties AND definitions in a single AI call:
```typescript
interface DocumentAnalysis {
  parties: DetectedParties;
  definitions: DefinedTerm[];
}
```

### 2. AI-Powered Extraction
Use AI (not regex) to find defined terms because:
- Definitions can appear anywhere in the document
- Patterns vary widely between contracts
- AI understands context (distinguishes defined terms from regular capitalized words)
- Single prompt handles all variations

### 3. Simple UI
- "Definitions" button in toolbar (next to Sides and Context)
- Popover with alphabetized, searchable list
- Shows term name and first 150 chars of definition
- No "Go to" navigation (keep MVP simple)

### Data Flow
```
Document Load
    → analyzeDocument(apiKey, docText)
    → AI returns { parties, definitions }
    → Store in documentAnalysis state
    → Sides feature uses documentAnalysis.parties
    → Definitions feature uses documentAnalysis.definitions
```

## Consequences

### Benefits
- Single AI call for all document analysis
- Users can quickly look up any defined term
- Extensible for future analysis features

### Drawbacks
- Longer AI prompt (more tokens)
- No "Go to definition" navigation in MVP

## Alternatives Considered

| Alternative | Reason for Rejection |
|-------------|---------------------|
| Regex extraction | Too fragile; misses non-standard patterns |
| Separate AI calls | Wasteful; increases latency |
| Full document search | Complex; deferred to future version |

## Implementation Notes

### Files
- `src/services/documentAnalysis.ts` — Combined analysis service
- `src/components/Definitions.tsx` — Popover UI component
- `src/components/Definitions.css` — Styling
- `src/components/Toolbar.tsx` — Contains Definitions button
- `src/components/App.tsx` — State management

## Related
- ADR-006: Sides feature
- ADR-007: Context feature
