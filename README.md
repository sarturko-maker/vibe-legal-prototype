Vibe Legal (Open Source Edition)
A Surgical OXML Redlining Engine for Microsoft Word.

Vibe Legal is a client-side TypeScript engine that uses Large Language Models (LLM) to "surgically" redline Word documents. Unlike standard text-replacement tools, Vibe Legal manipulates the Open XML (OXML) layer directly, preserving:

Table structures

Paragraph properties (indentation, alignment)

Run properties (bold, italic, color) via "Nearest Neighbor" inheritance.

⚠️ CRITICAL DISCLAIMER OF LIABILITY
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.

Not Legal Advice: This tool is an engineering prototype, not a lawyer. It uses AI (Gemini), which can hallucinate, omit text, or misinterpret instructions.

No Warranty: The authors (Vibe Legal & Contributors) make no representations or warranties of any kind concerning the safety, suitability, lack of viruses, inaccuracies, typographical errors, or other harmful components of this software.

Limitation of Liability: In no event shall the authors be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software.

User Responsibility: You are solely responsible for verifying every change made by this tool before accepting it. Always keep a backup of your original document.

Features
Zero Server: Runs entirely in the browser (Client-Side). Your API keys and data go directly to Google; they never pass through a middleman server.

Zero Trust Architecture: Assumes AI output is "dirty" and sanitizes it (strips LaTeX, Markdown) before injection.

Re-Flow Engine: Supports multi-paragraph selection by rebuilding the XML structure dynamically.

How to Run (Script Lab)
Install the Script Lab add-in for Word.

Copy the code from vibe-legal-v2.tsx.

Paste it into the Script tab in Script Lab.

Add the following libraries in the Libraries tab:

Plaintext

https://appsforoffice.microsoft.com/lib/1/hosted/office.js
https://unpkg.com/react@17/umd/react.development.js
https://unpkg.com/react-dom@17/umd/react-dom.development.js
https://cdnjs.cloudflare.com/ajax/libs/diff_match_patch/20121119/diff_match_patch.js
Run.
