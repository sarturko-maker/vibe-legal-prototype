# Vibe Legal

**AI-powered contract redlining for Microsoft Word**

---

> ⚠️ **Please Read Before Using**
>
> This is a **vibe coded** project. Built by a lawyer who doesn't know how to code (and with a help of a community that does), using AI development tools.
>
> It exists to generate ideas, spark conversation, and contribute to the open source legal tech community. It is **not** a commercial product. It is **not** production-ready software.
>
> **No liability is accepted.** Use this at your own risk. Always review AI suggestions carefully. Never rely on this tool for legal advice or critical work.

---

Vibe Legal is a Word add-in that helps legal professionals review and negotiate contracts using AI. It works directly inside Microsoft Word Online — just ask for changes in plain English and watch them appear as tracked changes.

If you're a lawyer with ideas for features, you can build them too. That's the point.

---

## What It Does

Open a contract in Word Online. Tell the add-in what you want. Get proper tracked changes back.

**Examples:**
- "Make the indemnity mutual"
- "Reduce the liability cap to £100,000"  
- "Add a 30-day termination for convenience"
- "What risks does this clause create for us?"

The AI reads your document, understands the context, and applies changes surgically — preserving your formatting, numbering, and defined terms.

---

## Features

### Pick Your Side
Tell the add-in which party you represent. AI detects the parties automatically — you just choose one. All advice and drafting then favours your position.

### Add Deal Context
Explain what matters to your client. Working on a time-sensitive deal? Conservative client? Specific red lines? The AI remembers this throughout your session.

### Set Risk Appetite
Show the AI what "aggressive" and "conservative" look like using your own examples. It calibrates its suggestions accordingly.

### See Defined Terms
View all defined terms instantly. No more scrolling to the definitions section or hunting through schedules.

### Debate Prep
Let the AI argue both sides of a position. Useful for anticipating counterarguments or stress-testing your approach.

### Mind Map
See the top 5 topics in your contract at a glance. Drill into any area and discuss it with the AI.

### Preview Before Applying
Review what the AI wants to change before it touches your document.

### Choose Your AI Provider
Switch between Google Gemini, Groq, and Mistral. Bring your own API key — nothing gets stored on a server.

### Chat History
Start new conversations, auto-generate names for old ones, and access your history throughout the session.

---

## Coming Soon (Preview Mode)

These features are visible in Preview Mode to show where the project is heading:

- **Local deployment** — Run everything on your own machine with open-weight models. No data leaves your laptop.
- **Community tools** — A library of templates and playbooks created by the legal community.

---

## Security & Privacy

**This is a "Bring Your Own Key" tool.** You provide your own API key from your chosen AI provider. Your documents go directly from Word to the AI — there's no Vibe Legal server in between.

**What this means:**
- Your API key is stored locally in your browser
- Your documents are sent directly to your chosen AI provider (Google, Groq, or Mistral)
- Nothing passes through Vibe Legal infrastructure

**What you should know:**
- If your computer is compromised, someone could access your stored API key
- Make sure sending documents to your chosen AI provider complies with your firm's policies and any confidentiality obligations
- This is research software — always review AI suggestions carefully

---

## Getting Started

### The Honest Version

Right now, running this add-in requires some technical setup. You'll need to run a local development server on your computer before the add-in will work in Word.

If you're comfortable with that (or know someone who is), you'll need:
- Node.js installed on your computer
- A Microsoft 365 account (Word Online)
- An API key from one of the supported providers (Gemini, Groq, or Mistral)

Clone this repository, run the development server, then sideload the manifest into Word Online.

### If That Sounds Like Too Much

That's completely fair. This is an open source project — the code and ideas are here for anyone to learn from, adapt, or build upon.

If there's interest, I may host a live version in the future. In the meantime, feel free to reach out if you'd like help getting it running.

---

## Reminder

AI makes mistakes. It might remove a liability cap while telling you it fixed a typo. **Always review every change.** You are responsible for the final document.

---

## The Bigger Picture

This Word add-in is Part 1 of a larger project:

- **Part 2** — A batch processing tool that redlines entire documents against your playbook (coming soon)
- **Part 3** — A vision for what legal workflows could look like (coming soon)

---

## Contributing

This project is open source under the **GPL-3.0 licence**.

If you modify and distribute this software, you must open-source your changes. If you're interested in building something proprietary using these ideas, get in touch.

---

## Acknowledgements

Built using:
- [diff-match-patch](https://github.com/google/diff-match-patch) (Google) — for computing text differences
- [React](https://react.dev/) (Meta) — for the interface
- [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/) (Microsoft) — for Word integration

---

## Get In Touch

Questions? Ideas? Found a bug?

Open an issue on GitHub or connect on LinkedIn.
