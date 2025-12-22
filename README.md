# Vibe Legal (v0.2 Alpha)
**Stateful AI Contract Redlining Engine - Research Prototype**

> **⚠️ EXPERIMENTAL SOFTWARE (SCRIPT LAB):**
> This is a **research prototype** provided as a YAML snippet for **Script Lab** (a Microsoft Word prototyping add-in). It is **not** a compiled, standalone plugin. It is designed for developers and legal engineers to experiment with "Surgical OXML" strategies.

---

> ### 🛑 SECURITY & LEGAL WARNING - READ BEFORE USING
>
> **1. API KEY SECURITY RISK (CRITICAL):**
> Vibe Legal is a **Client-Side Only** application. It creates a direct connection between your specific instance of Word and the AI Provider (Google/Anthropic).
> * **NO BACKEND SECURITY:** Your API keys are stored in your browser's **Local Storage** (`localStorage`) and in active memory during the session.
> * **VULNERABILITY:** If your machine is compromised, or if a malicious add-in accesses your local storage, **your API keys can be stolen.**
> * **DO NOT DEPLOY** this version in a production environment without wrapping it in a secure proxy backend.
>
> **2. NO LEGAL ADVICE:**
> This software uses Large Language Models (LLMs) which hallucinate. It may delete a liability cap while telling you it "fixed a typo." **NEVER** rely on this tool for legal advice. You assume full liability for all output.
>
> **3. DATA PRIVACY:**
> When you use features like "Context Caching" (Gemini) or deep analysis, the **full text of your document** is sent to the AI provider. Ensuring this complies with applicable laws, including data protection laws, is your responsibility.

---

## What is Vibe Legal?

Vibe Legal is a TypeScript engine that runs inside Microsoft Word to perform "Surgical Redlining." Unlike standard chatbots that rewrite entire documents (destroying formatting), Vibe Legal uses a **Stateful Holistic Architecture** to inject changes precisely into the existing XML structure of the document.

### New in v0.2 (Alpha):
* **The ID Echo Protocol:** To prevent AI "drift" (where the model loses track of which clause it is editing), this version wraps paragraphs in unique ID tags (e.g., `<P38>`) before sending them to the LLM. The AI must "echo" these IDs back, ensuring precise placement of edits.
* **Delta Manager:** For larger documents, a singleton state manager that tracks changes across the session (Insert/Modify/Delete), ensuring the AI remembers what it just changed.  This feature (like all others) is experimental.
* **Structured Reasoning Router:** The AI outputs a JSON "Change Plan" (`intent`, `reasoning`, `actions`), allowing complex logic (like "Make this mutual") to be broken down into specific atomic actions.

## Quick Start (Script Lab)

Vibe Legal is designed to run inside **Script Lab**, a Microsoft prototyping tool for Word.

1.  **Install Script Lab** from the **Insert > Add-ins > Store** menu in Microsoft Word.
2.  **Import the Code:**
    * Open the `vibe-legal.yaml` file in this repository.
    * Copy the raw text content.
    * Open the **Script Lab** tab in Word, click **Import**, and paste the YAML.
3.  **Run:** Click the "Run" button in the Script Lab panel.
4.  **Configuration:**
    * Click the **Settings** (gear icon) in the task pane.
    * Select your provider (Gemini or Claude).
    * Enter your API Key. **(Stored locally on your machine—see Security Warning above).**

## Architecture & Features

* **Surgical OXML Injection:** Attempts to manipulate the Open XML layer directly to preserve tables, numbering, and styles.
* **Paragraph-Aware Mode:** Intelligently handles multi-paragraph clause insertions without breaking Word's auto-numbering lists.
* **Native Track Changes:** AI modifications appear as standard Word Track Changes (Insertions/Deletions), attributable to "Vibe AI" or a custom name.

## Contributing & License

This project is open source under the **GNU General Public License v3** (GPLv3).

**Note for Enterprise Developers:** This license requires that if you modify and distribute this software, you must open-source your changes. If you are looking to build a proprietary backend using this frontend logic, please contact the maintainers.

## Acknowledgments

Vibe Legal stands on the shoulders of open-source giants. We gratefully acknowledge:

* **[diff-match-patch](https://github.com/google/diff-match-patch)** (Google): The core algorithm used for computing semantic text differences.
* **[React](https://react.dev/)** (Meta): Powering the reactive task pane interface.
* **[Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)** (Microsoft): The API enabling direct manipulation of the Word Object Model.
