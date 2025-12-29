# ADR-009: Negotiate Feature

## Status
Accepted

## Date
2025-12-27

## Context
Lawyers need to prepare for negotiations by exploring arguments from both perspectives. This requires:

1. Understanding counterparty's likely objections
2. Strengthening their own position
3. Finding weaknesses in their arguments before the other side does

Current solutions involve manual brainstorming or asking colleagues, which is time-consuming and may miss perspectives.

## Decision
Implement a Negotiate feature with two modes:

### Mode 1: Auto-Debate
- AI generates arguments for both sides automatically
- Up to 5 arguments per side (10 total)
- Each argument RESPONDS to the previous one (not batch-generated)
- Arguments appear sequentially with animation

### Mode 2: Interactive Brainstorm
- User picks a side (FOR or AGAINST their position)
- User types arguments, AI plays opponent
- "Let AI help me" option generates arguments for user's side
- Turn-by-turn conversation

### Key Design Decisions

1. **Sequential Generation**: Each argument responds to the previous one. This creates realistic debate flow rather than parallel lists of pre-generated points.

2. **Context Integration**: Uses Deal Context and Sides settings to tailor arguments to the specific deal situation.

3. **Two Modes**: Auto-Debate for quick exploration, Interactive for deep preparation on specific points.

4. **Chat Bubble UI**: Visual distinction between sides (green FOR, amber AGAINST) makes debate easy to follow.

### Data Flow
```
User enters position
    → Select mode (Auto/Interactive)
    → [Auto] AI generates FOR argument #1
    → [Auto] AI generates AGAINST argument #1 (responding to FOR #1)
    → [Auto] Continue until 5 per side or stopped
    
    → [Interactive] User types argument
    → [Interactive] AI responds with counter
    → [Interactive] Continue turn-by-turn
```

## File Structure
```
types/negotiate.ts           - DebateMessage, NegotiateState interfaces
services/negotiate.ts        - AI generation: buildNegotiatePrompt, generateDebateArgument, runAutoDebate
components/
├── Negotiate.tsx           - Main modal container
├── Negotiate.css           - All negotiate styles
├── NegotiateSetup.tsx      - Position input + mode selection
├── NegotiateActions.tsx    - Copy/Continue in chat buttons
├── DebateView.tsx          - Message list with auto-scroll
├── DebateBubble.tsx        - Individual message bubble
└── InteractiveInput.tsx    - User input for interactive mode
```

## Consequences

### Positive
- Lawyers can quickly explore both sides of a negotiation point
- Arguments respond to each other naturally, surfacing realistic objections
- Integrates with existing Context and Sides for tailored arguments
- Copy functionality for use in emails or documents

### Negative
- Additional bundle size (~3KB gzipped)
- Requires API calls for each argument (vs batch)
- No persistence of debate history between sessions

## Alternatives Considered

1. **Batch Generation**: Generate all FOR arguments, then all AGAINST
   - Rejected: Arguments wouldn't respond to each other

2. **Single Mode Only**: Just auto-debate or just interactive
   - Rejected: Different use cases need different approaches

3. **Integrated in Main Chat**: No separate modal
   - Rejected: Debate UI benefits from dedicated space with side-by-side visual
