/**
 * Validation Layer for AI-Generated Operations
 * 
 * Two-stage validation:
 * 1. Structural validation (code-based, strict allowlist)
 * 2. Semantic validation (AI-based, for complex operations)
 */

import { callGeminiApi } from './gemini/client';

// === CONSTANTS ===
const ALLOWED_OPERATION_TYPES = new Set([
    'AMEND_SIMPLE',
    'INSERT',
    'DELETE'
]);

// === TYPES ===

export interface ValidationError {
    code: string;
    message: string;
    operationIndex?: number;
    field?: string;
}

export interface ValidationResult {
    valid: boolean;
    errors: ValidationError[];
    warnings: string[];
}

export interface AIResponse {
    intent?: string;
    answer?: string;
    explanation?: string;
    operations?: Operation[];
}

export interface Operation {
    type: string;
    target_id?: number;
    end_id?: number;
    insert_after?: number;
    // AMEND_SIMPLE fields (new targeted replacement)
    find_text?: string;       // Exact text to find in paragraph
    replace_text?: string;    // Replacement text
    amended_text?: string;    // Legacy: full paragraph replacement (deprecated)
    content?: string;
    block?: string; // Legacy/forbidden support detection
    styleToken?: string;
    style?: string; // Style name for INSERT
    list_level?: number;
    description?: string;
    original_text?: string; // Runtime addition
    // Font info for INSERT operations (ADR-011)
    font?: string;
    fontSize?: number;
    // Extended formatting for INSERT operations
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string; // Hex color like "#1a73e8" or null for auto
}

export interface DocumentContext {
    totalParagraphs: number;
    clauses: any[];
    paragraphs?: any[]; // Rich paragraph info
}

// ============================================================================
// STAGE 1: Code-based structural validation (Strict & Fast)
// ============================================================================

export function validateOperations(
    response: AIResponse,
    documentContext: DocumentContext,
    styleTokens?: Set<string>
): ValidationResult {
    console.log('[Validation] Starting validation pipeline...');

    const errors: ValidationError[] = [];
    const warnings: string[] = [];

    // 1. Check intent exists
    if (!response.intent) {
        errors.push({
            code: 'MISSING_INTENT',
            message: 'Response missing required "intent" field'
        });
    }

    // 2. Check operations array exists
    // If no operations but intent is MODIFY, that's an issue. If intent is ANSWER, ops can be empty/undefined.
    // However, the prompt ensures operations array is present.
    if (response.intent === 'MODIFY' || response.intent === 'HYBRID') {
        if (!response.operations || !Array.isArray(response.operations)) {
            errors.push({
                code: 'MISSING_OPERATIONS',
                message: 'Response missing "operations" array'
            });
            return { valid: false, errors, warnings };
        }
    }

    // 3. Validate each operation
    if (response.operations && Array.isArray(response.operations)) {
        response.operations.forEach((op, index) => {

            // CRITICAL: Operation type allowlist check
            if (!op.type) {
                errors.push({
                    code: 'MISSING_TYPE',
                    message: `Operation ${index} missing "type" field`,
                    operationIndex: index
                });
            } else if (!ALLOWED_OPERATION_TYPES.has(op.type)) {
                errors.push({
                    code: 'INVALID_OPERATION_TYPE',
                    message: `Operation ${index} has unsupported type "${op.type}". Allowed: ${[...ALLOWED_OPERATION_TYPES].join(', ')}`,
                    operationIndex: index,
                    field: 'type'
                });
            }

            // Target ID bounds check
            const targetId = op.target_id || op.insert_after;
            if (targetId !== undefined) {
                if (typeof targetId !== 'number' || !Number.isInteger(targetId)) {
                    errors.push({
                        code: 'INVALID_TARGET_ID',
                        message: `Operation ${index}: target_id must be an integer, got ${typeof targetId}`,
                        operationIndex: index,
                        field: 'target_id'
                    });
                } else if (targetId < 1 || targetId > documentContext.totalParagraphs) {
                    errors.push({
                        code: 'TARGET_OUT_OF_RANGE',
                        message: `Operation ${index}: target_id ${targetId} out of range (1-${documentContext.totalParagraphs})`,
                        operationIndex: index,
                        field: 'target_id'
                    });
                }
            }

            // End_id validation
            if (op.end_id !== undefined && (op.target_id || 0)) {
                if (op.end_id < (op.target_id || 0)) {
                    errors.push({
                        code: 'INVALID_RANGE',
                        message: `Operation ${index}: end_id (${op.end_id}) cannot be less than target_id (${op.target_id})`,
                        operationIndex: index,
                        field: 'end_id'
                    });
                }
            }

            // Style token check
            if (op.styleToken && styleTokens && !styleTokens.has(op.styleToken)) {
                // Note: Changed to error in user spec 'UNKNOWN_STYLE', but kept as warning in legacy code.
                // User spec says: Blocking Error: UNKNOWN_STYLE
                errors.push({
                    code: 'UNKNOWN_STYLE',
                    message: `Operation ${index}: unknown styleToken "${op.styleToken}"`,
                    operationIndex: index,
                    field: 'styleToken'
                });
            }

            // List level validation
            if (op.list_level !== undefined && (op.list_level < 0 || op.list_level > 9)) {
                errors.push({
                    code: 'INVALID_LIST_LEVEL',
                    message: `Operation ${index}: list_level must be 0-9, got ${op.list_level}`,
                    operationIndex: index,
                    field: 'list_level'
                });
            }

            // Content presence checks
            if (op.type === 'INSERT' && !op.content) {
                errors.push({
                    code: 'MISSING_CONTENT',
                    message: `Operation ${index}: INSERT requires "content" field`,
                    operationIndex: index,
                    field: 'content'
                });
            }

            if (op.type === 'AMEND_SIMPLE') {
                // Prefer find_text/replace_text (new targeted approach)
                if (op.find_text && !op.replace_text) {
                    errors.push({
                        code: 'MISSING_REPLACE_TEXT',
                        message: `Operation ${index}: find_text provided without replace_text`,
                        operationIndex: index,
                        field: 'replace_text'
                    });
                }
                if (op.replace_text && !op.find_text) {
                    errors.push({
                        code: 'MISSING_FIND_TEXT',
                        message: `Operation ${index}: replace_text provided without find_text`,
                        operationIndex: index,
                        field: 'find_text'
                    });
                }
                // Fallback: amended_text still allowed for backwards compatibility
                if (!op.find_text && !op.amended_text) {
                    errors.push({
                        code: 'MISSING_AMEND_DATA',
                        message: `Operation ${index}: AMEND_SIMPLE requires find_text/replace_text`,
                        operationIndex: index
                    });
                }
            }

            // DELETE requires target_id
            if (op.type === 'DELETE' && !op.target_id) {
                errors.push({
                    code: 'MISSING_TARGET_ID',
                    message: `Operation ${index}: DELETE requires "target_id" field`,
                    operationIndex: index,
                    field: 'target_id'
                });
            }

            // === WARNINGS (non-blocking) ===
            const content = op.content || op.amended_text || '';

            if (op.list_level !== undefined && /^\d+\./.test(content)) {
                warnings.push(`Operation ${index}: Content starts with number but list_level=${op.list_level} — may duplicate numbering`);
            }

            if (/^[\u2022\u2023\u25E6\u2043\u2219•]\s/.test(content)) {
                warnings.push(`Operation ${index}: Content starts with bullet character — remove it, let Word handle formatting`);
            }

            if (/\*\*[^*]+\*\*/.test(content)) {
                warnings.push(`Operation ${index}: Content contains markdown bold (**text**) — Word won't render this`);
            }

            if (op.type === 'INSERT' && op.insert_after && documentContext.paragraphs) {
                const targetP = documentContext.paragraphs.find((p: any) => p.id === op.insert_after);
                if (targetP?.isListItem && op.list_level === undefined) {
                    warnings.push(`Operation ${index}: INSERT after list item P${op.insert_after} without list_level — will be plain text`);
                }
            }
        });

    }

    // Validating progressive IDs in INSERTs
    const inserts = (response.operations || []).filter(op => op.type === 'INSERT' && op.insert_after !== undefined);
    if (inserts.length > 1) {
        const targets = inserts.map(op => op.insert_after as number);
        const uniqueTargets = new Set(targets);

        // If multiple unique targets exist, check if they are sequential (progressive)
        // e.g. 27, 28, 29
        if (uniqueTargets.size > 1) {
            const sorted = [...targets].sort((a, b) => a - b);

            // Check if every target is exactly 1 greater than previous
            // Only strictly creating a sequence: 27, 28, 29...
            let isProgressive = true;
            for (let i = 1; i < sorted.length; i++) {
                if (sorted[i] !== sorted[i - 1] + 1) {
                    isProgressive = false;
                    break;
                }
            }

            if (isProgressive && uniqueTargets.size === targets.length) {
                warnings.push(
                    `INSERTs have progressive targets (${sorted.join(', ')}). Use same anchor for sequential inserts — tool handles offset calculation.`
                );
            }
        }
    }

    const valid = errors.length === 0;
    console.log(`[Validation] ${valid ? 'Passed' : 'Failed'} - ${errors.length} errors, ${warnings.length} warnings`);

    if (!valid) console.log('[Validation] Errors:', errors);
    if (warnings.length > 0) console.warn('[Validation] Warnings:', warnings);

    return { valid, errors, warnings };
}

// ============================================================================
// STAGE 2: AI semantic validation (for complex operations)
// ============================================================================

const VALIDATION_SYSTEM_PROMPT = `
You are validating proposed contract modifications for correctness.

## CRITICAL: INSERT ORDER HANDLING

The tool processes same-anchor INSERTs using REVERSE STACKING:
- Array order [A, B, C] with same anchor → Document order [A, B, C]
- The tool internally reverses and processes C→B→A to achieve correct order

**VALID EXAMPLE - This is CORRECT, do NOT flag:**
\`\`\`json
[
  { "insert_after": 25, "content": "7. ENTIRE AGREEMENT", "style": "heading 3" },
  { "insert_after": 25, "content": "7.1 This Agreement constitutes...", "style": "Normal" },
  { "insert_after": 25, "content": "8. FORCE MAJEURE", "style": "heading 3" },
  { "insert_after": 25, "content": "8.1 Neither party shall be liable...", "style": "Normal" }
]
\`\`\`
This produces document order: 7. heading → 7.1 content → 8. heading → 8.1 content ✓

**DO NOT FLAG any of these:**
- Multiple INSERTs sharing the same anchor
- Heading appearing before its sub-clause content in the array
- Sequential clause numbers (7, 7.1, 8, 8.1) in ascending order

**ONLY flag these as errors:**
- Sub-clause content (7.1) appearing BEFORE its heading (7.) in the array
- Clause 8 appearing BEFORE Clause 7 in the array
- Content that makes no legal sense

DOCUMENT INFO:
- Total paragraphs: {totalParagraphs}
- Clause structure: {clauseSummary}

PROPOSED CHANGES:
{operations}

RETURN JSON ONLY - no markdown, no explanation:
{
  "valid": true,
  "errors": [],
  "warnings": []
}
`;

export async function validateSemantics(
    response: AIResponse,
    documentContext: DocumentContext,
    apiKey: string,
    model: string
): Promise<ValidationResult> {
    try {
        const clauseSummary = documentContext.clauses?.slice(0, 10).map(c =>
            `P${c.paragraphId}: ${c.number}. ${c.title}`
        ).join('\n') || 'No clause structure available';

        const prompt = VALIDATION_SYSTEM_PROMPT
            .replace('{totalParagraphs}', String(documentContext.totalParagraphs))
            .replace('{clauseSummary}', clauseSummary)
            .replace('{operations}', JSON.stringify(response.operations, null, 2));

        console.log('[Validation] Running semantic validation...');

        const result = await callGeminiApi(
            apiKey,
            model,
            'You validate contract changes. Return JSON only.',
            prompt
        );

        // Parse JSON response
        const jsonMatch = result.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            const parsed = JSON.parse(jsonMatch[0]);

            // Map strings to ValidationError objects if necessary, or just treat as messages
            // The semantic validation prompt returns strings. We need to adapt to ValidationError interface.
            const errors: ValidationError[] = (parsed.errors || []).map((msg: string) => ({
                code: 'SEMANTIC_ERROR',
                message: msg
            }));

            return {
                valid: parsed.valid ?? true,
                errors: errors,
                warnings: parsed.warnings || []
            };
        }

        // If no JSON found, assume valid (fail safe)
        console.warn('[Validation] Could not parse semantic validation response');
        return { valid: true, errors: [], warnings: ['Semantic validation skipped'] };

    } catch (e) {
        console.error('[Validation] Semantic validation error:', e);
        // On error, allow operation to proceed (fail safe)
        return { valid: true, errors: [], warnings: ['Semantic validation failed: ' + String(e)] };
    }
}

// ============================================================================
// HELPERS
// ============================================================================

export function formatValidationError(result: ValidationResult): string {
    if (result.valid) return '';

    // Handle various error formats safely
    const messages = result.errors.map(e => {
        if (typeof e === 'string') return e;
        if (e && typeof e.message === 'string') return e.message;
        if (e && typeof e === 'object') return JSON.stringify(e);
        return String(e);
    });

    if (messages.length === 1) {
        return messages[0];
    }

    return `${messages.length} validation errors:\n• ${messages.join('\n• ')}`;
}

// ============================================================================
// ORCHESTRATOR
// ============================================================================

export async function validateAIResponse(
    response: AIResponse,
    documentContext: DocumentContext,
    apiKey: string,
    model: string,
    styleTokens?: Set<string>
): Promise<ValidationResult> {

    // Stage 1: Structural validation (Strict)
    const structural = validateOperations(response, documentContext, styleTokens);

    if (!structural.valid) {
        console.error('[Validation] Structural errors:', structural.errors);
        return structural;
    }

    // Stage 2: Semantic validation (AI) — only for complex batches
    const operations = response.operations || [];
    const needsSemanticValidation =
        operations.length > 2 ||
        operations.some(op => op.type === 'INSERT' || op.type === 'DELETE'); // Strict check on major structural changes

    if (needsSemanticValidation && apiKey) {
        console.log('[Validation] Running semantic check for complex operation...');
        const semantic = await validateSemantics(response, documentContext, apiKey, model);

        if (!semantic.valid) {
            console.error('[Validation] Semantic errors:', semantic.errors);
            return {
                valid: false,
                errors: [...structural.errors, ...semantic.errors],
                warnings: [...structural.warnings, ...semantic.warnings]
            };
        }

        // Merge warnings
        return {
            valid: true,
            errors: [],
            warnings: [...structural.warnings, ...semantic.warnings]
        };
    }

    return structural;
}
