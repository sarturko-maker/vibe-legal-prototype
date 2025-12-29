
export interface SelectionState {
    isEmpty: boolean;
    text: string;
}

export interface DocumentContext {
    headers: string[]; // List of section headers in the doc
    cursorSection: string; // The section the cursor is currently in
}

export type IntentType = 'QUESTION' | 'COMMAND' | 'CLARIFY' | 'BLOCK';
export type ScopeType = 'SELECTION' | 'WHOLE_DOC' | 'AUTO_PLACE' | 'NONE';

export interface RouterResult {
    intent: IntentType;
    scope: ScopeType;
    targetSection?: string;
    message?: string;
}

/**
 * Determines the user's intent and the scope of the operation.
 */
/**
 * Determines the user's intent and the scope of the operation.
 */
export function determineIntent(prompt: string, selection: SelectionState, context: DocumentContext, visualContext: string = ""): RouterResult {
    const lowerPrompt = prompt.toLowerCase().trim();

    // R1: Nuclear Safety
    // "Delete everything", "Remove all", "Clear document" + No Selection
    if (selection.isEmpty) {
        if (lowerPrompt.match(/^(delete|remove|clear)\s+(everything|all|document|entire doc)/)) {
            return {
                intent: 'BLOCK',
                scope: 'NONE',
                message: "Nuclear Safety Protocol: Bulk deletion without selection is blocked. Please select the text you wish to delete."
            };
        }
    }

    // R5: Read vs Write (Questions)
    // Starts with "What", "Who", "Why", "Summarise", "Explain", "List"
    // Unless it says "List the changes" (which might be a command? No, that's read)
    // "Summarise the indemnity" -> QUESTION
    if (lowerPrompt.match(/^(what|who|why|how|when|where|summarise|summarize|explain|list|describe|analyze|check)\b/)) {
        return {
            intent: 'QUESTION',
            scope: 'NONE' // Questions don't modify the doc
        };
    }

    // R4: Ambiguity
    // "Missing definitions" -> Could be "Find them" (Q) or "Add them" (C).
    // Catch phrases that start with "missing" or just "definitions" (noun phrases).
    if (lowerPrompt.match(/^(missing\s+\w+|definitions|clerical errors|formatting issues)/)) {
        return {
            intent: 'CLARIFY',
            scope: 'NONE',
            message: `Ambiguous request: "${prompt}". Do you want me to FIND them (Question) or FIX them (Command)?`
        };
    }

    // R3: Auto-Place
    // "Add X clause" + Cursor in middle of unrelated paragraph (simulated by context)
    // If command is "Add" or "Insert" AND selection is empty.
    if (selection.isEmpty && lowerPrompt.match(/^(add|insert|include)\b/)) {
        // Check if we should auto-place.
        // If the prompt mentions a section that exists in headers, target it.
        // Example: "Add a Force Majeure clause"
        // We need to map "Force Majeure" to a section.
        // Simple heuristic: Look for keywords in headers.

        // Common sections to look for if not specified
        const targetKeywords = ['general', 'miscellaneous', 'definitions', 'indemnity', 'termination'];

        for (const header of context.headers) {
            const lowerHeader = header.toLowerCase();
            // If prompt mentions the header explicitly
            if (lowerPrompt.includes(lowerHeader)) {
                return {
                    intent: 'COMMAND',
                    scope: 'AUTO_PLACE',
                    targetSection: header
                };
            }
        }

        // Visual Context Heuristics for Auto-Place
        // If visual context suggests we are near a "Definition List" or "Signature Block", use that.
        if (visualContext.toLowerCase().includes("definition") || visualContext.toLowerCase().includes("defined terms")) {
            if (lowerPrompt.includes("definition")) {
                // We are likely in the definitions section, so insert here?
                // Or if we are NOT, but the prompt asks for definitions, we should find the definitions section.
                // This logic is tricky without full doc scan.
                // Let's assume if Visual Context says we are looking at definitions, and user says "Add definition", we insert at selection (which is empty, so cursor).
                return {
                    intent: 'COMMAND',
                    scope: 'SELECTION' // Insert at cursor in definitions section
                };
            }
        }

        // If prompt implies a type of clause, try to find a home.
        // "Force Majeure" -> "General Provisions" or "Miscellaneous"
        if (lowerPrompt.includes('force majeure') || lowerPrompt.includes('notice') || lowerPrompt.includes('severability')) {
            const miscHeader = context.headers.find(h => h.toLowerCase().includes('miscellaneous') || h.toLowerCase().includes('general'));
            if (miscHeader) {
                return {
                    intent: 'COMMAND',
                    scope: 'AUTO_PLACE',
                    targetSection: miscHeader
                };
            }
        }

        // If no target found, default to WHOLE_DOC (or insert at cursor?)
        // R2 says "Make it..." + No Selection -> COMMAND + WHOLE_DOC
        // But "Add..." might imply insertion.
        // Let's default "Add" to AUTO_PLACE if we can't find a target? No, that's risky.
        // If we can't find a target, maybe just insert at cursor (SELECTION with empty text)?
        // Or WHOLE_DOC implies "Add this to the document structure".
        // Let's stick to R2 for generic commands.
    }

    // R2: Context Switch (Command)
    // If we are here, it's likely a command.
    if (!selection.isEmpty) {
        return {
            intent: 'COMMAND',
            scope: 'SELECTION'
        };
    } else {
        // No selection.
        return {
            intent: 'COMMAND',
            scope: 'WHOLE_DOC'
        };
    }
}
