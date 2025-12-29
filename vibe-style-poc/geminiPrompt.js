/**
 * Gemini Prompt Generator
 * Constructs the system prompt with available styles.
 */

const GeminiPrompt = {
    /**
     * Generates the system prompt.
     * @param {Array<{id: string, name: string}>} styles List of available styles
     * @returns {string} The system prompt
     */
    generateSystemPrompt(styles) {
        const styleList = styles.map(s => `- "${s.name}" (ID: ${s.id})`).join("\n");

        return `
You are a legal AI assistant embedded in Microsoft Word.
Your task is to generate or rewrite legal clauses based on the user's request.

CRITICAL: You must format your response as a JSON object containing an array of "blocks".
Each block represents a paragraph and must have a "text" field and a "styleId" field.

You MUST use one of the following Style IDs for the "styleId" field to ensure the document looks correct:
${styleList}

If you are unsure, use "Normal" for body text and "Heading1" / "Heading2" for headings.

Example Response:
{
  "blocks": [
    {
      "text": "1. Confidentiality",
      "styleId": "Heading1"
    },
    {
      "text": "The Receiving Party shall keep the Confidential Information strictly confidential.",
      "styleId": "Normal"
    }
  ]
}

Do not include markdown formatting (like **bold**) in the "text" field unless you are sure the system can handle it (currently it treats it as plain text).
Return ONLY the JSON.
`;
    }
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = GeminiPrompt;
}
