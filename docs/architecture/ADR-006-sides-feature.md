# ADR-006: Sides Feature

**Status:** Accepted  
**Date:** 2025-12-27  
**Authors:** Artur (Vibe Legal)  
**Foundation for:** ADR-012 (Risk Tolerance)

---

## Context

In legal negotiations, a lawyer always represents one party. When reviewing a contract:

- A **Buyer's lawyer** looks for risks to the Buyer and suggests Buyer-friendly amendments
- A **Seller's lawyer** does the opposite
- A **neutral reviewer** considers both perspectives equally

For AI-powered contract review to be useful, it must understand **whose side it's on**. Generic advice like "this clause could be risky" isn't helpful — risky for whom?

### The Problem

Without knowing the user's side, AI gives balanced advice:

> "The limitation of liability clause caps damages at the contract value. This benefits the Seller by limiting exposure, but the Buyer may want higher protection."

With side awareness, AI gives actionable advice:

> "This limitation of liability clause is unfavourable to you as the Buyer. Consider requesting a cap of 2x contract value and carving out liability for gross negligence."

---

## Decision

Implement a **Sides Feature** that:

1. **Auto-detects parties** from the document using AI
2. **Lets users select** which party they represent
3. **Injects side context** into all AI prompts
4. **Shows visual indicator** when a side is active

### Three States

| State | Meaning | AI Behaviour |
|-------|---------|--------------|
| **Neutral** | No side selected | Balanced advice for both parties |
| **Party A** | User represents first party | Advice favours Party A |
| **Party B** | User represents second party | Advice favours Party B |

---

## Party Detection

When the user saves their settings (triggering document load), we analyse the contract to identify parties.

### AI Prompt for Detection

```
Analyze this contract and identify the two main parties.

Return JSON only:
{
  "partyA": {
    "shortName": "ABC",
    "fullName": "ABC Holdings Limited",
    "role": "Buyer"
  },
  "partyB": {
    "shortName": "XYZ",
    "fullName": "XYZ International B.V.",
    "role": "Seller"
  }
}
```

### What We Extract

| Field | Purpose | Example |
|-------|---------|---------|
| `shortName` | Button label (max 15 chars) | "Acme" |
| `fullName` | Full legal name for prompts | "Acme Corporation Ltd" |
| `role` | Contract role for context | "Buyer", "Licensor", "Landlord" |

### Common Role Pairs

The AI recognises standard contract relationships:

- Buyer / Seller
- Licensor / Licensee
- Landlord / Tenant
- Discloser / Recipient (NDAs)
- Employer / Employee
- Service Provider / Client

---

## User Interface

### Side Selector Button

The toolbar shows a "Side" button that:

- Displays "Side" when neutral
- Displays party short name when selected (e.g., "Acme")
- Shows green dot indicator when active
- Opens dropdown on click

### Dropdown Options

```
┌─────────────────────────┐
│  Acme                   │
│  Buyer               ✓  │
├─────────────────────────┤
│  GlobalCorp             │
│  Seller                 │
├─────────────────────────┤
│  Neutral                │
│  Balanced advice        │
└─────────────────────────┘
```

Users can switch sides at any time. The change takes effect on the next AI interaction.

---

## Prompt Injection

When a side is selected, we inject context into every AI prompt:

### The Side Instruction

```
=== CLIENT REPRESENTATION ===
You are advising Acme Corporation Ltd (the "Buyer") in this transaction.

All your advice must:
- FAVOUR Acme's interests
- IDENTIFY risks TO Acme
- SUGGEST amendments that BENEFIT Acme
- FLAG provisions that are UNFAVOURABLE to Acme

When drafting or amending, always ask: "Is this good for Acme?"
```

### How It Combines

The side instruction is added to every prompt alongside other context:

```
[System Prompt]
[Side Instruction]      ← "You are advising the Buyer..."
[Deal Context]          ← From Context feature (ADR-007)
[Risk Tolerance]        ← From Risk feature (ADR-012)
[Document Content]
[User Question]
```

---

## Data Flow

```
Document Load
    │
    ▼
detectParties(docText, apiKey)
    │
    ▼
Store in state: detectedParties = { partyA, partyB }
    │
    ▼
User clicks Side button → sees detected parties
    │
    ▼
User selects "Acme (Buyer)"
    │
    ▼
selectedSide = { selected: 'partyA' }
    │
    ▼
All subsequent AI calls include side instruction
```

---

## Key Files

| File | Purpose |
|------|---------|
| `src/services/partyDetection.ts` | AI call to detect parties |
| `src/components/SideSelector.tsx` | UI component for side selection |
| `src/components/SideSelector.css` | Styling |
| `src/prompts/systemPrompt.ts` | `buildSideInstruction()` function |
| `src/components/App.tsx` | State management for `selectedSide` and `detectedParties` |

---

## Consequences

### Benefits

- **Actionable advice** — AI suggestions are specific to user's position
- **Natural workflow** — Matches how lawyers actually work
- **Quick setup** — Auto-detection means users don't type party names
- **Flexible** — Can switch sides or go neutral at any time

### Limitations

- **Two parties only** — Doesn't handle multi-party agreements well
- **Detection accuracy** — AI might misidentify parties in unusual contracts
- **Session only** — Side selection not persisted (follows serverless pattern)

### Dependencies

Other features build on Sides:

| Feature | Dependency |
|---------|------------|
| Risk Tolerance (ADR-012) | Requires side selected before enabling |
| Negotiate (ADR-009) | Uses sides for debate simulation |

---

## Alternatives Considered

| Alternative | Reason for Rejection |
|-------------|---------------------|
| Manual party entry | Extra friction; auto-detection is more user-friendly |
| Persist side selection | Breaks serverless pattern; would need storage |
| Regex party detection | Too fragile; misses variations in contract language |
| No sides (always neutral) | Reduces usefulness for actual legal work |

---

## Implementation Notes

### State Structure

```typescript
interface SideState {
  selected: 'neutral' | 'partyA' | 'partyB';
}

interface DetectedParties {
  partyA: { shortName: string; fullName: string; role: string };
  partyB: { shortName: string; fullName: string; role: string };
}
```

### Null Handling

- If party detection fails, `detectedParties` is `null`
- Side selector shows "Save Settings to detect parties"
- User can still use the tool in neutral mode

### Short Name Generation

AI is instructed to create short names (max 15 chars) that fit in the button:

- "ABC Holdings Ltd" → "ABC"
- "John Smith" → "Smith"
- Unknown company → Use role ("Buyer")

---

## Verification Points for Code Review

1. `partyDetection.ts` — AI prompt returns valid JSON with both parties
2. `SideSelector.tsx` — Three options shown (partyA, partyB, neutral)
3. `systemPrompt.ts` — `buildSideInstruction()` generates correct prompt text
4. Side instruction included in AI calls when side is selected
5. Green indicator shown when side is active (not neutral)

---

## Related

- ADR-007: Deal Context Feature (combines with sides in prompts)
- ADR-008: Definitions Feature (shares party detection call)
- ADR-012: Risk Tolerance (requires sides to be enabled)
