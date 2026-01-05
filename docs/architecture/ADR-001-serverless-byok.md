# ADR-001: Serverless BYOK Architecture

**Status:** Accepted  
**Date:** 2025-12-27  
**Authors:** Artur (Vibe Legal)  
**Foundation ADR** — This decision shapes the entire project architecture.

---

## Context

Vibe Legal is an open-source Word add-in for AI-powered contract redlining. Legal documents are sensitive, and law firms have strict rules about where client data can go.

The project needed to answer a fundamental question: **Where does the AI processing happen?**

Three options existed:

1. **Hosted service** — We run servers, users connect to us
2. **Self-hosted server** — Users run their own backend
3. **Serverless browser-only** — Everything runs in the browser, users bring their own AI keys

### Constraints

- **Open source project** — No budget for server infrastructure
- **Solo maintainer** — Cannot support server operations long-term
- **Legal industry** — Extreme sensitivity about document privacy
- **Vibe-coded** — Built without traditional programming knowledge, simpler is better

---

## Decision

Vibe Legal uses a **Serverless BYOK (Bring Your Own Key)** architecture:

1. **No backend servers** — The add-in runs entirely in the browser
2. **User provides AI key** — Entered in Settings, stored in browser session only
3. **Direct API calls** — Browser calls AI provider (Gemini/Claude) directly
4. **Document never leaves Word** — AI sees text content, but document stays local

### How It Works

```
┌─────────────────────────────────────────────────────────────┐
│                     USER'S COMPUTER                         │
│  ┌─────────────┐    ┌─────────────────────────────────────┐ │
│  │   MS Word   │◄──►│         Vibe Legal Add-in           │ │
│  │  (Document) │    │  • Reads document text              │ │
│  └─────────────┘    │  • Applies track changes            │ │
│                     │  • Stores API key in session        │ │
│                     └──────────────┬──────────────────────┘ │
└────────────────────────────────────┼────────────────────────┘
                                     │ HTTPS (document text only)
                                     ▼
                        ┌────────────────────────┐
                        │    AI Provider API     │
                        │  (Gemini / Claude /    │
                        │   Future: Local LLM)   │
                        └────────────────────────┘
```

### Key Files

| File | Purpose |
|------|---------|
| `src/components/Settings.tsx` | API key entry and storage |
| `src/state/SettingsContext.tsx` | React state management for API keys |
| `src/services/gemini/client.ts` | Direct browser-to-Gemini calls |
| `src/services/handleAction.ts` | Orchestrates AI calls from browser |

---

## Future Vision: Local AI Models

A key reason for this architecture is **forward compatibility with open-weight AI models**.

As local AI improves (Llama, Mistral, etc.), users will be able to:

- Run AI entirely on their own machine or office server
- Keep absolutely nothing in the cloud
- Eliminate per-token costs after initial setup

The BYOK pattern supports this by simply pointing to a different endpoint:

```
Current:  api.gemini.google.com
Future:   localhost:8080 (local Ollama server)
          office-server:8080 (firm's private AI)
```

This makes Vibe Legal suitable for:
- **Regulated industries** — Banking, healthcare, government
- **Cost-conscious firms** — No ongoing API fees
- **Privacy-maximalist users** — Full air-gap possible

---

## Consequences

### Benefits

- **Maximum privacy** — Document text goes only to user's chosen AI provider
- **Zero infrastructure** — No servers to maintain, no hosting costs
- **User control** — Users manage their own API keys and costs
- **Open source friendly** — Anyone can fork and run without dependencies
- **Future-proof** — Ready for local AI models when they mature

### Drawbacks

- **Feature limitations** — Some features are harder without a backend:
  - Shared playbooks and prompt libraries
  - Team collaboration features
  - Usage analytics
  - Automatic updates to prompts
- **User responsibility** — Users must manage their own API keys
- **No fallback** — If user's API key is invalid, nothing works

### Mitigations

- Clear error messages when API key is missing or invalid
- Settings panel validates key on save
- Future: Local storage for user's personal playbooks

---

## Alternatives Considered

| Alternative | Reason for Rejection |
|-------------|---------------------|
| Hosted SaaS backend | Cannot maintain servers; privacy concerns for legal docs |
| Proxy server (keys stay with us) | Still requires server infrastructure; trust issue |
| Electron desktop app | More complex; harder to distribute; Word add-in is simpler |
| Server-side document processing | Privacy concerns; server maintenance burden |

---

## Implementation Notes

### API Key Storage

Keys are stored in React state (session only). They are NOT:
- Saved to localStorage (persists after browser close)
- Sent to any server we control
- Logged anywhere

### Supported Providers

Currently:
- **Gemini** — Primary provider, implemented in `gemini/client.ts`
- **Claude** — Planned, same pattern

Future:
- **OpenAI-compatible endpoints** — For local models (Ollama, LM Studio)

### Verification Points for Code Review

1. `Settings.tsx` — Confirm API key is stored in React state only
2. `SettingsContext.tsx` — Confirm state management uses `useState`, not localStorage
3. `gemini/client.ts` — Confirm calls go direct to Google, no proxy
4. No `fetch()` calls to any domain we control
5. No analytics or telemetry sending document content

---

## Related

- ADR-002: OOXML Manipulation Strategy (how document editing works client-side)
- ADR-007: Deal Context Feature (example of session-only state)
