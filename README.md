# Vibe Legal (MVP 1.1 Prototype)
**Serverless AI Contract Redlining Engine - Experimental Build**

> **⚠️ WORK IN PROGRESS:**
> This project is an early-stage **prototype**. While the core engine works, it is experimental software. There is a lot of work to do, and bugs should be expected. We are building this in public to explore what is possible with Client-Side AI.

---

> ### 🛑 CRITICAL LEGAL DISCLAIMER - READ BEFORE USING
>
> **THIS SOFTWARE IS AN ENGINEERING PROTOTYPE, NOT A LAWYER.**
>
> 1.  **NO LEGAL ADVICE:** Vibe Legal uses Large Language Models (LLMs) which can hallucinate, omit critical legal terms, or misinterpret instructions. **NEVER** rely on this tool for legal advice.
> 2.  **USER RESPONSIBILITY:** You are solely responsible for verifying every single change made by this software. By using this tool, you agree that you are competent to review the output and accept full liability for its use.
> 3.  **NO WARRANTY:** This software is provided "AS IS", without warranty of any kind, express or implied. The Vibe Legal Project & Contributors are **NOT** liable for any damages, malpractice claims, or contract disputes arising from its use.
> 4.  **ALWAYS BACKUP:** This tool manipulates the XML structure of your document. Always keep a clean backup copy of your contract before running Vibe Legal.

---

## What is Vibe Legal?
Vibe Legal is a client-side TypeScript engine for **Microsoft Word (Script Lab)** designed to experiment with using Large Language Models (LLM) for surgical document redlining.

Unlike standard text-replacement tools, Vibe Legal attempts to manipulate the **Open XML (OXML)** layer directly. Our goal is to:

* **Experiment with AI Forensics:** We are testing algorithms that scan documents to detect fonts, numbering styles, and heading formats, attempting to force the AI to match your specific style ("Document DNA").
* **Test Hybrid Intents:** We are building an "Orchestrator" that can handle complex requests like *"Make this buyer friendly"* by attempting to simultaneously **Modify** existing clauses and **Insert** new ones.
* **Prioritize Privacy:** The architecture is designed so your contract text goes directly from your Word instance to the LLM provider (Google Gemini). It **never** passes through a Vibe Legal server.

## Quick Start (Script Lab)

Vibe Legal is designed to run inside **Script Lab**, a Microsoft prototyping tool for Word.

1.  **Install Script Lab** from the Insert > Add-ins > Store in Microsoft Word.
2.  **Import the Code:**
    * Open the `vibe-legal.yaml` file in this repository.
    * Copy the raw content.
    * Open Script Lab in Word, click **Import**, and paste the YAML.
3.  **Run:** Click the "Run" button.
4.  **API Key:** You will need a Google Gemini API Key (free tier available at aistudio.google.com).

## Current Features (Experimental)

* **⚡ Model Freedom:** A dynamic settings menu that fetches available models from Google.
* **🧠 Smart Caching:** Attempts to cache long contracts to save tokens and speed up Q&A.
* **👀 True Vision Focus:** An experimental fix for the "Ghost Text" bug in Track Changes, ensuring the AI sees the accepted view of the text.
* **📝 Native Track Changes:** AI modifications appear as standard Word Track Changes.

## Contributing & License
This project is open source under the **GNU General Public License v3**.
We welcome contributions! There is a lot of work to be done to make this a reliable tool for legal professionals.
