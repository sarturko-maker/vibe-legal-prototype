/**
 * Document Analysis Service
 * Combined extraction of parties AND defined terms in a single AI call
 */

import { callAIForText } from './gemini/client';
import { AIProvider } from '../types';

export interface DefinedTerm {
    term: string;
    definition: string;  // First 400 chars
}

export interface DetectedParties {
    partyA: { shortName: string; fullName: string; role: string };
    partyB: { shortName: string; fullName: string; role: string };
}

export interface MindMapSubTopic {
    id: string;
    name: string;
    summary: string;  // 2 sentences
}

export interface MindMapTopic {
    id: string;
    title: string;
    icon: string;     // Emoji
    summary: string;  // 2 sentences
    keyFigures: string[];
    subTopics: MindMapSubTopic[];
}

export interface DocumentAnalysis {
    parties: DetectedParties;
    definitions: DefinedTerm[];
    mindMap: MindMapTopic[];  // NEW: Mind map topics
}

/**
 * Analyze document to extract parties AND defined terms in one AI call
 */
export async function analyzeDocument(
    documentText: string,
    apiKey: string,
    model: string = 'gemini-2.0-flash',
    provider: AIProvider = 'gemini'
): Promise<DocumentAnalysis | null> {
    console.log('[analyzeDocument] Starting combined analysis...');
    console.log('[analyzeDocument] API key present:', !!apiKey);
    console.log('[analyzeDocument] Document length:', documentText?.length);

    if (!apiKey) {
        console.error('[analyzeDocument] No API key provided');
        return null;
    }

    if (!documentText || documentText.length < 50) {
        console.error('[analyzeDocument] Document text too short or missing');
        return null;
    }

    const systemPrompt = `You are a legal document parser. Extract party information, defined terms, and mind map topics from contracts. Return valid JSON only.`;

    const userPrompt = `Analyze this contract and extract three things:

1. PARTIES: Identify the two main parties to this agreement.

2. DEFINED TERMS: Find ALL defined terms anywhere in the document.
   Look for patterns like:
   - "Term" means/shall mean...
   - "Term": [definition]
   - (the "Defined Term")
   - (hereinafter "Term")
   - Any capitalized term that is clearly defined
   
   Search the ENTIRE document including:
   - Definitions sections/clauses
   - Schedules and appendices
   - Inline definitions within clauses
   - Recitals and preamble

3. MIND MAP: Identify EXACTLY 5 major themes/topics this contract covers.
   These should be plain-language topics (not clause titles).
   Think: "What are the 5 main things this contract is about?"

CONTRACT TEXT:
${documentText.substring(0, 20000)}

Return JSON only, no markdown code blocks, no explanation:
{
  "parties": {
    "partyA": { "shortName": "ACME", "fullName": "ACME Corporation Ltd", "role": "Buyer" },
    "partyB": { "shortName": "TechCo", "fullName": "TechCo Inc", "role": "Seller" }
  },
  "definitions": [
    { "term": "Agreement", "definition": "This Master Services Agreement including all Schedules..." },
    { "term": "Confidential Information", "definition": "Any information disclosed by either party that is marked..." }
  ],
  "mindMap": [
    {
      "id": "payments",
      "title": "Payments & Pricing",
      "icon": "💰",
      "summary": "Two sentence summary of payment arrangements in this contract.",
      "keyFigures": ["£500k total", "Net 30"],
      "subTopics": [
        { "id": "pricing", "name": "Pricing Structure", "summary": "Two sentences about pricing." },
        { "id": "terms", "name": "Payment Terms", "summary": "Two sentences about payment terms." }
      ]
    }
  ]
}

RULES:
- shortName: max 15 characters, use abbreviated company name or role
- definitions: return up to 50 terms, sorted alphabetically
- definition text: max 400 characters each
- Return EXACTLY 5 mind map topics
- Each topic has 2-5 sub-topics
- Mind map summaries must be specific to THIS contract
- Use plain language, not legal jargon
- Choose appropriate emoji icons for each topic`;

    try {
        console.log('[analyzeDocument] Calling AI with provider:', provider, 'model:', model);
        const response = await callAIForText(provider, apiKey, model, systemPrompt, userPrompt);
        console.log('[analyzeDocument] Raw response length:', response?.length);

        // Extract JSON from response (handle potential markdown code blocks)
        let jsonStr = response;

        // Remove markdown code blocks if present
        const codeBlockMatch = response.match(/```(?:json)?\s*([\s\S]*?)```/);
        if (codeBlockMatch) {
            jsonStr = codeBlockMatch[1];
        }

        // Extract JSON object
        const jsonMatch = jsonStr.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            const parsed = JSON.parse(jsonMatch[0]);

            // Validate structure
            if (parsed.parties?.partyA?.shortName && parsed.parties?.partyB?.shortName) {
                // Ensure definitions array exists
                if (!Array.isArray(parsed.definitions)) {
                    parsed.definitions = [];
                }

                // Ensure mindMap array exists
                if (!Array.isArray(parsed.mindMap)) {
                    parsed.mindMap = [];
                }

                console.log('[analyzeDocument] SUCCESS - Parties:',
                    parsed.parties.partyA.shortName, 'vs', parsed.parties.partyB.shortName);
                console.log('[analyzeDocument] SUCCESS - Definitions:',
                    parsed.definitions.length, 'terms found');
                console.log('[analyzeDocument] SUCCESS - Mind Map:',
                    parsed.mindMap.length, 'topics found');

                return parsed as DocumentAnalysis;
            } else {
                console.warn('[analyzeDocument] Invalid structure:', parsed);
            }
        } else {
            console.warn('[analyzeDocument] No JSON found in response');
        }

        return null;
    } catch (e) {
        console.error('[analyzeDocument] Error:', e);
        return null;
    }
}
