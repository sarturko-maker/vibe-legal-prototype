# ADR-007: Deal Context Feature

## Status
Accepted

## Date
2025-12-27

## Context
Users need to provide persistent context about their deal that shapes ALL AI interactions. Currently, users must re-explain deal specifics with each query:

- "We're the buyer, budget is £50k"
- "Critical that we own all custom IP"
- "Supplier is a large US vendor"

This context must:
1. Combine with existing features (Sides) and future features (Playbook)
2. Be immediately active once saved
3. Be visible to users (green indicator when active)
4. Follow serverless architecture (no backend storage)

## Decision
Implement a "Context" button in the toolbar (next to Sides) that opens a popover:

1. **UI Components:**
   - Button with green dot indicator when active
   - Popover with textarea for free-form deal description
   - "Done" button to save and activate
   - "Clear" button to remove context

2. **State Management:**
   ```typescript
   interface DealContextState {
     description: string;
     isActive: boolean;
   }
   ```

3. **Prompt Injection:**
   Context combines additively with other enhancements:
   ```typescript
   const combinedEnhancements = [
     buildDealContextSection(dealContext),
     buildSideInstruction(side, parties),
     buildChatHistoryContext(chatHistory)
   ].filter(Boolean).join('\n\n');
   ```

4. **Data Flow:**
   ```
   App.tsx (state) → Toolbar.tsx (button + popover)
     → InputArea.tsx → handleAction.ts → systemPrompt.ts
   ```

## Consequences

### Benefits
- Users describe deal once, not repeatedly
- AI responses tailored to specific situation
- Combines naturally with Sides feature
- Extensible pattern for future features

### Drawbacks
- Longer system prompts (increased token usage)
- Users may forget to update stale context
- Session-only persistence

## Alternatives Considered

| Alternative | Reason for Rejection |
|-------------|---------------------|
| Auto-detect from document | Unreliable; may miss user priorities |
| Per-query context field | Too repetitive; poor UX |
| Server-side storage | Breaks serverless architecture |

## Implementation Notes

### Files
- `src/components/Context.tsx` — Popover UI component
- `src/components/Toolbar.tsx` — Contains Context button
- `src/components/App.tsx` — State management
- `src/components/InputArea.tsx` — Pass to handleAction
- `src/services/handleAction.ts` — Accept parameter, combine in prompt
- `src/prompts/systemPrompt.ts` — buildDealContextSection()

### Key Function
```typescript
export function buildDealContextSection(context: DealContextState | null): string {
  if (!context?.description?.trim()) return '';
  
  return `
=== DEAL CONTEXT ===
${context.description.trim()}

Consider this context when providing ALL advice.
`;
}
```

## Related
- Sides feature (ADR-006)
- Future: Playbook, Tone features
