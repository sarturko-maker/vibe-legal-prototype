# ADR-010: Mind Map Feature

## Status
Accepted

## Date
2025-12-27

## Context
Users need a quick way to understand what a contract covers without reading it clause-by-clause. Legal contracts are often dense and structured by legal convention rather than readability.

A visual, topic-based overview helps users:
1. Quickly grasp what the contract is about
2. Identify areas of interest or concern
3. Navigate directly to topics that matter to them
4. Get AI-powered explanations of specific areas

## Decision

### Combined Document Analysis
Extend the existing `documentAnalysis.ts` to generate mind map topics alongside parties and definitions in a **single AI call**. This is efficient because:
- One API call on document load
- All analysis available immediately
- No separate loading for Mind Map feature
```typescript
interface DocumentAnalysis {
  parties: DetectedParties;
  definitions: DefinedTerm[];
  mindMap: MindMapTopic[];  // 5 topics with sub-topics
}
```

### Structure
- **5 topics** (exactly 5, AI decides themes based on contract content)
- **2-5 sub-topics** per topic
- Each topic/sub-topic has a 2-sentence summary
- Key figures extracted where relevant (amounts, dates, percentages)

### User Interface
Panel-based UI consistent with other features (Definitions, Negotiate):

1. **Overview**: 5 clickable topic cards
2. **Expanded**: Topic detail with sub-topic list
3. **AI View**: AI-generated advice with follow-up capability

### AI Integration
Users can ask AI about any topic or sub-topic. The AI prompt includes:
- Contract text (for context)
- Deal context (from Context feature, if set)
- Selected side (from Sides feature, if set)
- Specific topic/sub-topic being asked about

### No Complex Caching
Mind map is generated once on document load as part of combined analysis. When document changes (redlines applied), the analysis naturally refreshes on next load. No separate caching layer needed.

## Consequences

### Benefits
- Quick contract comprehension without reading every clause
- AI explanations tailored to user's side and deal context
- Consistent UI pattern with other features
- Efficient single API call for all document analysis

### Drawbacks
- Initial document load slightly slower (more analysis)
- AI-generated topics may occasionally miss nuances
- Limited to 5 topics (may oversimplify complex contracts)

## Alternatives Considered

| Alternative | Reason for Rejection |
|-------------|---------------------|
| Fixed categories (Money, Risk, etc.) | Doesn't reflect actual contract content |
| Clause-by-clause view | Duplicates Word's outline; not user-friendly |
| Separate API call for mind map | Inefficient; adds latency |
| Complex caching system | Unnecessary; document state manages freshness |

## Implementation Notes

### Files
- `services/documentAnalysis.ts` - Extended with mindMap generation
- `types/mindmap.ts` - View state types
- `components/MindMap.tsx` - Main container
- `components/TopicOverview.tsx` - Cards view
- `components/TopicExpanded.tsx` - Expanded topic view
- `components/TopicAIView.tsx` - AI response view
- `services/mindmapAI.ts` - AI query generation

### View States
```typescript
type MindMapView = 'overview' | 'expanded' | 'ai';
```

### AI Prompt Pattern
Topic advice prompt includes:
- Contract excerpt
- Deal context (if available)
- Selected side (if set)
- Topic and sub-topic context
- User's specific question (if any)

## Related
- [ADR-007](ADR-007-context-feature.md) - Deal Context (used in AI prompts)
- [ADR-006](ADR-006-sides-feature.md) - Sides (used in AI prompts)
- [ADR-008](ADR-008-definitions-feature.md) - Combined document analysis pattern
