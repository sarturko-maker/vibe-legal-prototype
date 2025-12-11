VIBE LEGAL (OPEN SOURCE EDITION) A Surgical OXML Redlining Engine for Microsoft Word.

Vibe Legal is a client-side TypeScript engine that uses Large Language Models (LLM) to "surgically" redline Word documents. Unlike standard text-replacement tools, Vibe Legal manipulates the Open XML (OXML) layer directly, preserving:

Table structures

Paragraph properties (indentation, alignment)

Run properties (bold, italic, color) via "Nearest Neighbor" inheritance.

======================================================================== CRITICAL DISCLAIMER OF LIABILITY THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
Not Legal Advice: This tool is an engineering prototype, not a lawyer. It uses AI (Gemini), which can hallucinate, omit text, or misinterpret instructions.

No Warranty: The authors (Arturs & Contributors) make no representations or warranties of any kind concerning the safety, suitability, lack of viruses, inaccuracies, typographical errors, or other harmful components of this software.

Limitation of Liability: In no event shall the authors be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software.

User Responsibility: You are solely responsible for verifying every change made by this tool before accepting it. ALWAYS KEEP A BACKUP OF YOUR ORIGINAL DOCUMENT.

FEATURES

Zero Server: Runs entirely in the browser (Client-Side). Your API keys and data go directly to Google; they never pass through a middleman server.

Zero Trust Architecture: Assumes AI output is "dirty" and sanitizes it (strips LaTeX, Markdown) before injection.

Re-Flow Engine: Supports multi-paragraph selection by rebuilding the XML structure dynamically.

HOW TO RUN (SCRIPT LAB)

To run Vibe Legal inside Microsoft Word, follow these steps:

INSTALL SCRIPT LAB Search for "Script Lab" in the Microsoft Office Add-in Store (Insert > Add-ins) and add it to Word.

CONFIGURE THE "SCRIPT" TAB Copy the entire code from the file "vibe-legal-v2.tsx" in this repository. Paste it into the Script tab in Script Lab.

CONFIGURE THE "HTML" TAB Paste this single line into the HTML tab:

<div id="root"></div>

CONFIGURE THE "CSS" TAB Paste these styles into the CSS tab to ensure the UI renders correctly:

body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #fff; }

{ box-sizing: border-box; }

CONFIGURE THE "LIBRARIES" TAB Paste these URLs into the Libraries tab (one per line):

https://appsforoffice.microsoft.com/lib/1/hosted/office.js

https://unpkg.com/react@17/umd/react.development.js

https://unpkg.com/react-dom@17/umd/react-dom.development.js

https://cdnjs.cloudflare.com/ajax/libs/diff_match_patch/20121119/diff_match_patch.js

RUN Click the "Run" button in the Script Lab ribbon.

You will need a Google Gemini API Key (free tier available at aistudio.google.com).

Select text in your document and click "Redline".

## Third-Party Licenses

This project uses the following open-source software:

* **diff-match-patch**: Copyright (c) 2018 The diff-match-patch Authors. Licensed under Apache 2.0.
* **React**: Copyright (c) Meta Platforms, Inc. Licensed under MIT.
