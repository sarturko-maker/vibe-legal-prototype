import { callGemini } from './geminiApi';
import { mergeIdenticalRuns, stripRsidAttributes } from './OxmlPreProcessor';
import { correctDoubleEncodedEntities, cloneBulletNumbering, remapIds, removeGhostComments } from './OxmlPostProcessor';
import { validateOxml } from './OxmlValidator';

export interface AgentResult {
    oxml: string;
    originalOoxml: string; // For Rollback
    error?: string;
}

/**
 * Constructs the prompt for the OOXML Redline Agent.
 * @param oxml The raw OOXML string.
 * @param instruction The user's editing instruction.
 * @returns The formatted prompt string.
 */
/**
 * Constructs the prompt for the OOXML Redline Agent.
 * @param oxml The raw OOXML string.
 * @param instruction The user's editing instruction.
 * @param visualContext High-level visual metadata (styles, formatting) from Word.
 * @returns The formatted prompt string.
 */
export function constructOxmlPrompt(oxml: string, instruction: string, visualContext: string = ""): string {
    const date = new Date().toISOString();
    const prompt = `
""" SYSTEM ROLE: You are an expert OpenXML (OOXML) Engine. You do not speak conversational English; you speak strict XML.

CORE DIRECTIVES:

NO SUMMARIZATION: You must return the FULL, VALID XML for the requested scope. Never use placeholders like \`\`. If the input has 50 rows, the output must have 50 rows.

NATIVE TRACK CHANGES:

Wrap deletions in <w:del w:id="[Unique]" w:author="Vibe Legal" w:date="${date}"><w:r><w:delText>old</w:delText></w:r></w:del>.

Wrap insertions in <w:ins w:id="[Unique]" w:author="Vibe Legal" w:date="${date}"><w:r><w:t>new</w:t></w:r></w:ins>.

SENTINEL PROTECTION:

Never modify text inside <w:instrText> (Field Codes).

Never remove <w:drawing> or <w:pict> tags unless explicitly asked to delete the image.

STRUCTURAL INTEGRITY:

Do not break <w:tbl> grids. If you add a cell to one row, you must add it to all rows.

Preserve <w:sectPr> (Section Breaks) at all costs.

OUTPUT FORMAT: JSON

You must return a JSON object with the following structure:
{
  "analysis": "Brief explanation of your plan.",
  "changes": ["List of specific changes made"],
  "oxml": "The modified OOXML string"
}

INPUT CONTEXT: The user will provide a snippet of OOXML and a Visual Context summary.

YOUR TASK: Analyze the Visual Context to understand the user's intent (e.g., preserving styles). Then, apply the edits to the OOXML. Return ONLY the JSON object. """

<VISUAL_CONTEXT>
${visualContext}
</VISUAL_CONTEXT>

<INPUT_OOXML>
${oxml}
</INPUT_OOXML>

<INSTRUCTION>
${instruction}
</INSTRUCTION>
`;
    return prompt;
}

/**
 * Runs the OOXML Redline Agent with Retry Protocol and Validation.
 * @param apiKey The Gemini API Key.
 * @param model The model to use (e.g., "gemini-1.5-pro").
 * @param oxml The raw OOXML string.
 * @param instruction The user's editing instruction.
 * @param visualContext High-level visual metadata.
 * @returns The modified OOXML string.
 */
export async function runOxmlAgent(apiKey: string, model: string, oxml: string, instruction: string, visualContext: string = ""): Promise<AgentResult> {
    // 1. Pre-Processing
    let processedOxml = mergeIdenticalRuns(oxml);
    processedOxml = stripRsidAttributes(processedOxml);

    const MAX_ATTEMPTS = 3;
    let currentInstruction = instruction;
    let lastError = "";

    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
        console.log(`[Agent] Attempt ${attempt}/${MAX_ATTEMPTS}`);

        // Construct prompt (append error if retry)
        let prompt = constructOxmlPrompt(processedOxml, currentInstruction, visualContext);
        if (attempt > 1) {
            prompt += `\n\nPREVIOUS ATTEMPT FAILED. ERROR: ${lastError}\nFIX THE XML AND TRY AGAIN.`;
        }

        try {
            const rawResponse = await callGemini(apiKey, model, prompt, "OOXML");

            // Parse JSON Response
            let resultOxml = "";
            try {
                // Clean markdown code blocks if present
                const cleanJson = rawResponse.replace(/```json/g, "").replace(/```/g, "").trim();
                const parsed = JSON.parse(cleanJson);

                console.log("[Agent] Analysis:", parsed.analysis);
                console.log("[Agent] Changes:", parsed.changes);

                if (!parsed.oxml) throw new Error("JSON missing 'oxml' field");
                resultOxml = parsed.oxml;
            } catch (e) {
                console.warn("[Agent] JSON Parsing Failed. Falling back to raw text (risky).", e);
                // Fallback: If parsing fails, maybe the model returned raw XML?
                // Or we can treat it as an error and retry.
                // Let's try to extract XML if it looks like XML
                if (rawResponse.trim().startsWith("<")) {
                    resultOxml = rawResponse;
                } else {
                    throw new Error("Invalid Output Format: Not JSON and not XML.");
                }
            }

            // 2. Post-Processing
            let finalOxml = correctDoubleEncodedEntities(resultOxml);
            finalOxml = cloneBulletNumbering(finalOxml);
            finalOxml = remapIds(finalOxml);
            finalOxml = removeGhostComments(finalOxml);

            // 3. Validation
            const validation = validateOxml(finalOxml);
            if (validation.isValid) {
                return { oxml: finalOxml, originalOoxml: oxml };
            } else {
                console.warn(`[Agent] Validation Failed: ${validation.error}`);
                lastError = validation.error || "Unknown Validation Error";
            }
        } catch (error: any) {
            console.error(`[Agent] API Error: ${error.message}`);
            lastError = error.message;
        }
    }

    // Attempt 3 Failed: Abort
    return {
        oxml: oxml, // Return original on failure (Safe default)
        originalOoxml: oxml,
        error: `Complex edit failed after ${MAX_ATTEMPTS} attempts. Try a simpler instruction. Last error: ${lastError}`
    };
}
