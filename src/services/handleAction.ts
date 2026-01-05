/**
 * Handle Action - Primary document operation coordinator
 * 
 * ARCHITECTURE (Native Track Changes + Deterministic Diff):
 * - All operations use Word.ChangeTrackingMode.trackAll
 * - AMEND uses diff algorithm to find minimal change
 * - Preserves user's existing track changes state
 */

import { log, logError, logWarn } from '../utils/logger';
import { buildRouterSystemPrompt, buildSideInstruction, buildDealContextSection, SideState, DetectedParties, DealContextState } from '../prompts/systemPrompt';
import { callGeminiRouter } from './gemini/client';
import { ContractMap, ParagraphInfo, Message } from '../types';
import { buildSimpleContractMap } from './document/contractMap';
import { stripFormattingMarkers } from './formatting/markdownParser';
import { validateAIResponse, ValidationResult, formatValidationError } from './validation';
import { buildRiskToleranceInstruction } from '../prompts/riskTolerancePrompt';
import { RiskTolerance } from '../types/state';
import { findMinimalChanges, findTextChanges, TextChange } from '../utils/textDiff';
import { toMarkdown } from '../utils/markdownNormalizer';
import { formatDiffForAI, convertDmpToTextChanges, TextChange as DiffTextChange } from '../utils/diffFormatter';
import { executeAmendChange } from './amendExecutor';
import { validateChangesWithAI } from './aiSelfValidator';
import { diff_match_patch } from 'diff-match-patch';



/**
 * Build conversation context from recent chat history
 * Helps AI understand prior discussion when user says "do it" or "proceed"
 */
function buildChatHistoryContext(history: Message[] | null | undefined): string {
    if (!history || history.length === 0) return '';

    // Take last 6 messages (3 exchanges) to avoid token bloat
    const recent = history.slice(-6);

    const formatted = recent.map(m => {
        const role = m.role === 'user' ? 'USER' : 'ASSISTANT';
        const content = m.content.length > 500
            ? m.content.substring(0, 500) + '... [truncated]'
            : m.content;
        return `${role}: ${content}`;
    }).join('\n\n');

    return `
=== CONVERSATION HISTORY ===
Use this to understand what was previously discussed.
If user says "do it", "proceed", "amend now", "OK", refer to this history.
ONLY modify clauses that were discussed - do NOT add unrelated changes.

${formatted}

=== END CONVERSATION ===
`;
}

// Declare Word (provided by Office.js)
declare var Word: any;

// === INTERFACES ===

export interface StyleInfo {
    name: string;
    font: {
        name: string;
        size: number;
        bold: boolean;
        italic: boolean;
        underline: boolean;
        color: string | null;
    };
    paragraphFormat: {
        alignment: string;
        lineSpacing: number;
        spaceBefore: number;
        spaceAfter: number;
    };
    usageCount: number;
    exampleParagraphId: number;
    usedForHeadings: boolean;
    usedForBody: boolean;
    hasDirectFormattingOverrides?: boolean;
}

export interface ActionResult {
    success: boolean;
    intent: 'ANSWER' | 'MODIFY' | 'HYBRID' | 'ERROR';
    answer?: string;
    operationsExecuted?: number;
    error?: string;
    isDraft?: boolean;
    preview?: {
        intent: string;
        operations: any[];
        message: string;
        originalTexts?: { [key: number]: string };
    };
}

/**
 * Main entry point - MATCHES LEGACY handleAction PATTERN
 * @param selectedSide - Optional party selection for client-aware advice
 * @param detectedParties - Optional detected parties from document
 * @param chatHistory - Optional recent chat messages for context
 * @param dealContext - Optional deal context for persistent context injection
 * @param riskTolerance - Optional risk tolerance for negotiation calibration (ADR-012)
 */
export async function handleAction(
    message: string,
    apiKey: string,
    model: string,
    author: string,
    mode: 'chat' | 'draft' = 'chat',
    onProgress?: (stage: string) => void,
    selectedSide?: SideState | null,
    detectedParties?: DetectedParties | null,
    chatHistory?: Message[] | null,
    dealContext?: DealContextState | null,
    riskTolerance?: RiskTolerance | null
): Promise<ActionResult> {

    if (typeof Word === 'undefined') {
        return { success: false, intent: 'ERROR', error: 'Please open in Microsoft Word.' };
    }

    try {
        // ═══════════════════════════════════════════════════════════
        // PHASE 1: Build document context (read-only Word.run)
        // ═══════════════════════════════════════════════════════════
        onProgress?.('Reading document...');

        const documentContext = await buildDocumentContext();

        console.log('=== DOCUMENT CONTEXT ===');
        console.log('Paragraphs:', documentContext.paragraphCount);
        console.log('Clauses:', documentContext.contractMap?.clauses?.length || 0);
        console.log('=== CONTRACT MAP ===');
        console.log(JSON.stringify(documentContext.contractMap?.clauses, null, 2));

        // ═══════════════════════════════════════════════════════════
        // PHASE 2: Call AI (outside Word.run - can take time)
        // ═══════════════════════════════════════════════════════════
        onProgress?.('Thinking...');

        // Build all enhancement sections
        const contextSection = buildDealContextSection(dealContext || null);
        const sideInstruction = buildSideInstruction(selectedSide || null, detectedParties || null);
        const chatContext = buildChatHistoryContext(chatHistory);
        const riskInstruction = (selectedSide && riskTolerance)
            ? buildRiskToleranceInstruction(selectedSide, riskTolerance)
            : '';

        // Combine all enhancements
        const combinedEnhancements = [contextSection, sideInstruction, riskInstruction, chatContext]
            .filter(Boolean)
            .join('\n\n');

        console.log('[handleAction] Deal context:', dealContext?.isActive ? 'Active' : 'None');
        console.log('[handleAction] Side instruction:', sideInstruction ? 'Active' : 'Neutral');
        console.log('[handleAction] Risk tolerance:', riskTolerance?.enabled ? `${riskTolerance.level}%` : 'None');
        console.log('[handleAction] Chat history:', chatHistory?.length || 0, 'messages');

        const systemPrompt = buildRouterSystemPrompt(
            documentContext.contractMap,
            documentContext.documentText,
            combinedEnhancements,
            documentContext.styleMenu // Pass style menu
        );

        const response = await callGeminiRouter(apiKey, model, systemPrompt, message);

        console.log('=== AI RESPONSE ===');
        console.log('Intent:', response.intent);
        console.log('Operations:', JSON.stringify(response.operations, null, 2));

        const intent = response.intent ||
            (response.operations?.length > 0 ? 'MODIFY' : 'ANSWER');

        // ═══════════════════════════════════════════════════════════
        // If ANSWER only, return without document changes
        // ═══════════════════════════════════════════════════════════
        if (intent === 'ANSWER' || !response.operations?.length) {
            return {
                success: true,
                intent: 'ANSWER',
                answer: response.answer || response.explanation || 'No answer provided.'
            };
        }

        // ═══════════════════════════════════════════════════════════
        // PHASE 2.5: VALIDATE AI RESPONSE (before applying)
        // ═══════════════════════════════════════════════════════════
        onProgress?.('Validating...');

        const validation = await validateAIResponse(
            response,
            {
                totalParagraphs: documentContext.paragraphCount,
                clauses: documentContext.contractMap?.clauses || [],
                paragraphs: documentContext.contractMap?.paragraphs || []
            },
            apiKey,
            model
        );

        if (!validation.valid) {
            console.error('[handleAction] Validation failed:', validation.errors);
            return {
                success: false,
                intent: 'ERROR',
                error: formatValidationError(validation)
            };
        }

        if (validation.warnings.length > 0) {
            console.warn('[handleAction] Validation warnings:', validation.warnings);
        }

        // ═══════════════════════════════════════════════════════════
        // DRAFT MODE: Return preview instead of applying
        // ═══════════════════════════════════════════════════════════
        if (mode === 'draft' && response.operations && response.operations.length > 0) {
            console.log('=== DRAFT MODE: Returning preview ===');

            // Collect original texts and STORE on each operation for content matching
            const originalTexts: { [key: number]: string } = {};
            if (documentContext.contractMap?.paragraphs) {
                for (const op of response.operations) {
                    if (op.target_id && documentContext.contractMap.paragraphs[op.target_id - 1]) {
                        const origText = documentContext.contractMap.paragraphs[op.target_id - 1].text;
                        originalTexts[op.target_id] = origText;
                        // CRITICAL: Store original_text on the operation itself for content matching
                        op.original_text = origText;
                    }
                }
            }

            return {
                success: true,
                intent: intent as 'MODIFY' | 'HYBRID',
                isDraft: true,
                preview: {
                    intent: intent,
                    operations: response.operations,
                    message: response.explanation || response.answer || 'Review proposed changes below',
                    originalTexts: originalTexts
                }
            };
        }

        // ═══════════════════════════════════════════════════════════
        // CHAT MODE: Execute operations (SINGLE Word.run with TRACKED REFS)
        // THIS IS THE KEY FIX - matches legacy pattern exactly
        // ═══════════════════════════════════════════════════════════
        onProgress?.('Applying changes...');

        const result = await executeOperations(
            response.operations || [],
            author,
            documentContext.styleMenu
        );

        console.log('=== OPERATION RESULTS ===');
        console.log('Success:', result.successCount, 'Errors:', result.errorCount);

        let answer = response.answer || response.explanation || '';
        if (result.successCount > 0) {
            answer += `\n\n✅ Applied ${result.successCount} change(s) to the document.`;
        }
        if (result.errorCount > 0) {
            answer += `\n⚠️ ${result.errorCount} change(s) failed.`;
        }

        return {
            success: result.errorCount === 0,
            intent: intent as 'MODIFY' | 'HYBRID',
            answer,
            operationsExecuted: result.successCount
        };

    } catch (error: any) {
        logError('handleAction failed:', error);
        console.error('=== HANDLEACTION ERROR ===', error);
        return { success: false, intent: 'ERROR', error: error.message };
    }
}

/**
 * Build document context - reads paragraphs for AI prompt
 * Enhanced with rich paragraph data for style cloning and auto-numbering awareness
 */
async function buildDocumentContext(): Promise<{
    contractMap: ContractMap | null;
    documentText: string;
    paragraphCount: number;
    styleMenu: StyleInfo[];
}> {
    return Word.run(async (context: any) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        // Load rich paragraph properties in a single consolidated call to avoid overwrite
        paragraphs.load(
            'items/style, items/text, items/isListItem, items/alignment, items/lineSpacing, items/spaceBefore, items/spaceAfter, ' +
            'items/font/name, items/font/size, items/font/bold, items/font/italic, items/font/underline, items/font/color'
        );

        await context.sync();

        // Load list levels for list items (requires separate sync)
        for (let i = 0; i < paragraphs.items.length; i++) {
            const p = paragraphs.items[i];
            if (p.isListItem) {
                try {
                    p.listItem.load('level');
                } catch (e) {
                    // Some list items may not have list property accessible
                }
            }
        }
        await context.sync();

        // --- STYLE ANALYSIS ---
        // Track font sizes per style to find MOST COMMON (mode) rather than first occurrence
        const styleMap = new Map<string, {
            count: number;
            exampleIndex: number;
            font: any;
            paragraphFormat: any;
            fontSizes: Map<number, number>; // fontSize -> count
        }>();
        const styleFormats = new Map<string, Set<string>>();

        for (let i = 0; i < paragraphs.items.length; i++) {
            const p = paragraphs.items[i];
            const styleName = p.style || 'Normal';

            // Track formatting key
            const formatKey = `${p.font.name}-${p.font.size}-${p.font.bold}`;
            if (!styleFormats.has(styleName)) styleFormats.set(styleName, new Set());
            styleFormats.get(styleName)!.add(formatKey);

            if (!styleMap.has(styleName)) {
                styleMap.set(styleName, {
                    count: 1,
                    exampleIndex: i + 1,
                    font: {
                        name: p.font.name,
                        size: p.font.size, // Will be updated to mode later
                        bold: p.font.bold,
                        italic: p.font.italic,
                        underline: p.font.underline !== 'None',
                        color: p.font.color !== 'automatic' ? p.font.color : null
                    },
                    paragraphFormat: {
                        alignment: p.alignment,
                        lineSpacing: p.lineSpacing,
                        spaceBefore: p.spaceBefore,
                        spaceAfter: p.spaceAfter
                    },
                    fontSizes: new Map([[p.font.size, 1]])
                });
            } else {
                const style = styleMap.get(styleName)!;
                style.count++;
                // Track font size occurrences
                const currentCount = style.fontSizes.get(p.font.size) || 0;
                style.fontSizes.set(p.font.size, currentCount + 1);
            }
        }

        // Update each style's font.size to the MOST COMMON (mode) font size
        for (const [name, data] of styleMap) {
            let maxCount = 0;
            let modeSize = data.font.size; // Default to first occurrence
            for (const [size, count] of data.fontSizes) {
                if (count > maxCount) {
                    maxCount = count;
                    modeSize = size;
                }
            }
            if (modeSize !== data.font.size) {
                console.log(`[Style Detection] "${name}" font size: first=${data.font.size}pt, mode=${modeSize}pt (using mode)`);
            }
            data.font.size = modeSize;
        }

        const styleMenu: StyleInfo[] = [];
        for (const [name, data] of styleMap) {
            const isHeading = /heading/i.test(name) || (data.font.bold && data.font.size >= 12);
            const isBody = /normal|body/i.test(name) || (!data.font.bold && data.count > 5);
            const formats = styleFormats.get(name);
            const hasOverrides = formats ? formats.size > 1 : false;

            styleMenu.push({
                name,
                font: data.font,
                paragraphFormat: data.paragraphFormat,
                usageCount: data.count,
                exampleParagraphId: data.exampleIndex,
                usedForHeadings: isHeading,
                usedForBody: isBody,
                hasDirectFormattingOverrides: hasOverrides
            });
        }
        styleMenu.sort((a, b) => {
            if (a.usedForHeadings && !b.usedForHeadings) return -1;
            if (!a.usedForHeadings && b.usedForHeadings) return 1;
            return b.usageCount - a.usageCount;
        });

        // --- TEXT & CONTRACT MAP ---
        const paragraphInfos: ParagraphInfo[] = [];
        const textParts: string[] = [];

        for (let i = 0; i < paragraphs.items.length; i++) {
            const p = paragraphs.items[i];
            const id = i + 1;
            const text = p.text || '';

            const isAllCaps = text.length > 0 && text === text.toUpperCase() && /[A-Z]/.test(text);
            const hasBoldPattern = /^\d+\.\s+[A-Z]/.test(text.trim());

            let listLevel = 0;
            if (p.isListItem) {
                try {
                    listLevel = p.listItem?.level || 0; // Access loaded level
                } catch (e) { listLevel = 0; }
            }

            paragraphInfos.push({
                id,
                text: text,
                style: p.style || 'Normal',
                isListItem: p.isListItem || false,
                listLevel: listLevel,
                alignment: p.alignment || 'Left',
                isAllCaps: isAllCaps,
                isHeadingPattern: hasBoldPattern || text.length < 60,
                // Font info for contract map enrichment
                font: p.font?.name || undefined,
                fontSize: p.font?.size || undefined
            });
            textParts.push(text);
        }

        const contractMap = buildSimpleContractMap(paragraphInfos);
        const documentText = textParts.join('\n').substring(0, 15000);

        console.log('[buildDocumentContext] Success. Styles:', styleMenu.length, 'Paragraphs:', paragraphInfos.length);

        return {
            contractMap,
            documentText,
            paragraphCount: paragraphs.items.length,
            styleMenu
        };
    });
}

/**
 * Find paragraph by content match (not index)
 * This handles paragraph ID shifts when content is inserted/deleted
 */
async function findParagraphByContent(
    context: any,
    expectedText: string
): Promise<any | null> {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    for (const para of paragraphs.items) {
        para.load('text');
    }
    await context.sync();

    // Use first 50 chars for matching
    const searchText = expectedText.trim().substring(0, 50);

    for (const para of paragraphs.items) {
        const paraText = para.text.trim();
        if (paraText.startsWith(searchText) || paraText.includes(searchText)) {
            context.trackedObjects.add(para);
            return para;
        }
    }
    return null;
}

/**
 * Execute all operations - SINGLE Word.run with TRACKED REFERENCES
 * THIS IS THE KEY FIX - matches legacy pattern exactly
 * EXPORTED for use by PreviewPanel Accept button in Draft mode
 */
export async function executeOperations(
    operations: any[],
    author: string,
    styleMenu: StyleInfo[] = [], // Default to empty array
    apiKey?: string,
    model?: string
): Promise<{ successCount: number; errorCount: number }> {

    return Word.run(async (context: any) => {
        let successCount = 0;
        let errorCount = 0;

        // ════════════════════════════════════════════
        // STEP 1: Load paragraphs and build lookup
        // ════════════════════════════════════════════
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        // Load all paragraph properties
        // Split into chunks if too large? For now assume reasonable size.
        for (let i = 0; i < paragraphs.items.length; i++) {
            const p = paragraphs.items[i];
            p.load('text,style,isListItem');
        }
        await context.sync();

        // Build lookup with TRACKED OBJECTS
        const lookup: { [id: number]: any } = {};
        const trackedItems: any[] = [];

        for (let i = 0; i < paragraphs.items.length; i++) {
            const p = paragraphs.items[i];
            const id = i + 1;  // 1-indexed
            lookup[id] = p;
            context.trackedObjects.add(p);  // CRITICAL: Keep reference alive
            trackedItems.push(p);
        }

        const lookupKeys = Object.keys(lookup).map(Number);
        console.log('[executeOperations] Lookup built:', lookupKeys.length, 'paragraphs tracked');
        console.log('[executeOperations] Lookup range: P' + Math.min(...lookupKeys) + ' to P' + Math.max(...lookupKeys));

        // Legacy style cache not needed - using handleInsertOperation with styleMenu
        // const styleCache = await buildStyleCache(context);

        // ════════════════════════════════════════════
        // STEP 2: Separate & Sort Operations
        // ════════════════════════════════════════════

        const deletes = operations.filter(op => op.type === 'DELETE');
        const amends = operations.filter(op => op.type === 'AMEND_SIMPLE');
        const inserts = operations.filter(op => op.type === 'INSERT');

        // Defensive check for unsupported types
        const unknown = operations.filter(op => !['DELETE', 'AMEND_SIMPLE', 'INSERT'].includes(op.type));
        if (unknown.length > 0) {
            throw new Error(`Unsupported operation type(s): ${unknown.map((o: any) => o.type).join(', ')}`);
        }

        // 2.1 DELETEs - reverse order (highest ID first)
        const sortedDeletes = [...deletes].sort((a, b) =>
            (b.target_id || 0) - (a.target_id || 0)
        );

        // 2.2 AMENDs - reverse order by target_id, but preserve array order for same target_id
        // This ensures same-paragraph operations execute in the order AI specified
        const sortedAmends = [...amends].map((op, idx) => ({ op, originalIndex: idx }))
            .sort((a, b) => {
                if (a.op.target_id !== b.op.target_id) {
                    return (b.op.target_id || 0) - (a.op.target_id || 0); // Different paragraphs: bottom to top
                }
                return a.originalIndex - b.originalIndex; // Same paragraph: preserve original order
            })
            .map(item => item.op);

        // 2.3 INSERTs - Group by Anchor
        const insertsByAnchor = new Map<number, any[]>();
        for (const op of inserts) {
            const anchor = op.insert_after || 0;
            if (!insertsByAnchor.has(anchor)) {
                insertsByAnchor.set(anchor, []);
            }
            insertsByAnchor.get(anchor)!.push(op);
        }

        // Sort anchors High -> Low (so earlier inserts don't affect later anchor lookups)
        const sortedAnchors = [...insertsByAnchor.keys()].sort((a, b) => b - a);

        // ════════════════════════════════════════════
        // STEP 3: Execute Operations
        // ════════════════════════════════════════════

        // 3.1 Execute DELETEs
        for (const op of sortedDeletes) {
            try {
                console.log('[executeOperations] DELETE target:', op.target_id);
                const result = await handleDeleteOperation(context, op, lookup);
                if (result) successCount++; else errorCount++;
            } catch (e: any) {
                console.error('DELETE failed:', e);
                errorCount++;
            }
        }

        // 3.2 Execute AMENDs
        for (const op of sortedAmends) {
            try {
                const targetId = op.target_id;
                const lookupHasTarget = lookup.hasOwnProperty(targetId);
                console.log('[executeOperations] AMEND target:', targetId, '| Exists in lookup:', lookupHasTarget);
                if (!lookupHasTarget) {
                    console.error('[executeOperations] AMEND TARGET MISSING! Available IDs:', Object.keys(lookup).slice(0, 10).join(', '), '...');
                }
                const result = await handleAmendOperation(context, op, lookup, author, apiKey, model);
                if (result.success) {
                    successCount++;
                } else {
                    console.error('[executeOperations] AMEND failed:', result.error || 'Unknown error');
                    errorCount++;
                }
            } catch (e: any) {
                console.error('[executeOperations] AMEND exception:', e?.message || e);
                errorCount++;
            }
        }

        // 3.3 Execute INSERT Groups
        for (const anchor of sortedAnchors) {
            const group = insertsByAnchor.get(anchor)!;

            console.log(`[executeOperations] Processing ${group.length} inserts for anchor P${anchor}`);

            // PROCESSING STRATEGY: REVERSE STACKING
            // To achieve Order [A, B, C] after Anchor P:
            // 1. Insert C after P. (Doc: P, C)
            // 2. Insert B after P. (Doc: P, B, C)
            // 3. Insert A after P. (Doc: P, A, B, C)
            // This allows us to use the static 'lookup[anchor]' for every insert.

            const reversedGroup = [...group].reverse();

            for (const op of reversedGroup) {
                try {
                    console.log('[executeOperations] INSERT at P' + anchor + ' (Reverse Stack)');

                    // RE-FETCH ANCHOR for robustness
                    // In sequential inserts, the anchor object from initial lookup gets stale after doc mutation.
                    // We must fetch fresh reference by index (anchor is 1-indexed, items are 0-indexed)

                    // 1. Sync to get latest state
                    // context.document.body.paragraphs.items[anchor - 1] retrieval requires items loaded?
                    // Safer to just re-load items for the range or entire body if needed, but simple index access on re-loaded paragraphs collection is best.

                    const freshParagraphs = context.document.body.paragraphs;
                    freshParagraphs.load('items');
                    await context.sync();

                    // 2. Mock lookup with fresh object
                    // We construct a temporary lookup just for this operation with the fresh anchor
                    const freshAnchor = freshParagraphs.items[anchor - 1];
                    if (!freshAnchor) throw new Error(`Anchor paragraph P${anchor} lost during sequential insert`);

                    const tempLookup = { [anchor]: freshAnchor };

                    const result = await handleInsertOperation(context, op, tempLookup, styleMenu, author);
                    if (result) successCount++; else errorCount++;

                    // 3. Sync after insert ensures next iteration sees valid state
                    await context.sync();

                } catch (e: any) {
                    console.error('INSERT failed:', e);
                    errorCount++;
                }
            }
        }

        // ════════════════════════════════════════════
        // STEP 4: Cleanup
        // ════════════════════════════════════════════
        for (const item of trackedItems) {
            try {
                context.trackedObjects.remove(item);
            } catch (e) {
                // Ignore cleanup errors
            }
        }

        return { successCount, errorCount };
    });
}

/**
 * Apply TextChange array to a paragraph with proper native handling.
 * - insert_after: Uses native Word insertion at end of anchor range
 * - delete_after: Searches for anchor+textToDelete, replaces with just anchor
 * - replace: Standard search/replace
 */
async function applyTextChanges(
    context: any,
    paragraph: any,
    changes: TextChange[]
): Promise<{ successCount: number; failedCount: number }> {
    let successCount = 0;
    let failedCount = 0;

    // Apply in reverse order for stability (changes don't affect earlier positions)
    for (let i = changes.length - 1; i >= 0; i--) {
        const change = changes[i];
        console.log(`[applyTextChanges] Applying change ${i + 1}/${changes.length}`);

        try {
            if (change.type === 'insert_after') {
                // NATIVE INSERTION
                console.log(`[applyTextChanges] → Insert "${change.text}" after anchor "${change.anchor}"`);

                let searchResults = paragraph.search(change.anchor, { matchCase: false });
                searchResults.load('items');
                await context.sync();

                if (searchResults.items.length === 0) {
                    console.warn(`[applyTextChanges] Anchor not found: "${change.anchor}"`);
                    failedCount++;
                    continue;
                }

                // Get end of anchor range and insert text there
                const anchorRange = searchResults.items[0];
                anchorRange.insertText(change.text, Word.InsertLocation.end);
                await context.sync();

                console.log(`[applyTextChanges] ✓ Insert applied successfully`);
                successCount++;

            } else if (change.type === 'delete_after') {
                // NATIVE DELETION - search for ONLY the text to delete (not anchor + text)
                // Anchor stays untouched = won't show in track changes
                console.log(`[applyTextChanges] → Delete "${change.textToDelete.substring(0, 30)}..."`);

                // Search for the text to delete directly
                let searchResults = paragraph.search(change.textToDelete, { matchCase: false });
                searchResults.load('items');
                await context.sync();

                if (searchResults.items.length === 0) {
                    console.warn(`[applyTextChanges] Text to delete not found: "${change.textToDelete.substring(0, 50)}"`);
                    failedCount++;
                    continue;
                }

                if (searchResults.items.length > 1) {
                    console.warn(`[applyTextChanges] Multiple matches (${searchResults.items.length}), using first`);
                }

                // Delete it (replace with empty string to respect track changes)
                searchResults.items[0].insertText('', Word.InsertLocation.replace);
                await context.sync();

                console.log(`[applyTextChanges] ✓ Delete applied successfully`);
                successCount++;

            } else if (change.type === 'replace') {
                // STANDARD SEARCH/REPLACE
                console.log(`[applyTextChanges] → Replace: "${change.find.substring(0, 30)}..." → "${change.replace.substring(0, 30)}..."`);

                let searchResults = paragraph.search(change.find, { matchCase: false });
                searchResults.load('items');
                await context.sync();

                if (searchResults.items.length === 0) {
                    console.warn(`[applyTextChanges] Find text not found: "${change.find.substring(0, 50)}"`);
                    failedCount++;
                    continue;
                }

                searchResults.items[0].insertText(change.replace, Word.InsertLocation.replace);
                await context.sync();

                console.log(`[applyTextChanges] ✓ Replace applied successfully`);
                successCount++;
            }
        } catch (error: any) {
            console.error(`[applyTextChanges] Error applying change ${i}:`, error?.message || error);
            failedCount++;
        }
    }

    return { successCount, failedCount };
}

/**
 * Handle AMEND operation - Deterministic Diff for Word-Level Track Changes
 * 
 * Uses diff algorithm to find minimal change between original and amended text.
 * Returns { success: boolean, error?: string } for better error reporting.
 */
async function handleAmendOperation(
    context: any,
    op: any,
    lookup: { [id: number]: any },
    author: string,
    apiKey?: string,
    model?: string
): Promise<{ success: boolean; error?: string }> {
    try {
        const targetId = op.target_id;
        console.log('[handleAmendOperation] ══════════════════════════════════════');
        console.log('[handleAmendOperation] FOUR-LAYER APPROACH for target_id:', targetId);

        // Step 1: Lookup paragraph
        console.log('[handleAmendOperation] Step 1: Looking up paragraph...');
        let startPara = lookup[targetId];

        if (!startPara && op.original_text) {
            console.log('[handleAmendOperation] ID lookup failed, trying content match...');
            startPara = await findParagraphByContent(context, op.original_text);
        }

        if (!startPara) {
            console.error('[handleAmendOperation] FAILED: Paragraph not found:', targetId);
            return { success: false, error: `Paragraph ${targetId} not found` };
        }
        console.log('[handleAmendOperation] Step 1 DONE: Paragraph found');

        // Step 2: Load paragraph text with track changes applied (virtually)
        // Always use getReviewedText - works whether or not track changes exist
        // Word API cannot detect track changes created in same session, so don't check
        console.log('[handleAmendOperation] Step 2: Loading paragraph text...');

        const paragraphRange = startPara.getRange(Word.RangeLocation.whole);
        let rawOriginalText: string;

        try {
            // Get "reviewed" text - returns current state with any track changes applied
            const reviewedText = paragraphRange.getReviewedText(Word.ChangeTrackingVersion.current);
            await context.sync();
            rawOriginalText = (reviewedText.value || '').trim();
            console.log('[handleAmendOperation] Step 2: Got text via getReviewedText');
        } catch (reviewError: any) {
            console.warn('[handleAmendOperation] getReviewedText failed, falling back to .text:', reviewError?.message);
            // Fallback to normal text load (may not include track changes)
            startPara.load('text');
            await context.sync();
            rawOriginalText = startPara.text.trim();
        }

        console.log(`[handleAmendOperation] Step 2 DONE: Text loaded (${rawOriginalText.length} chars)`);
        console.log('[handleAmendOperation] Text preview:', rawOriginalText.substring(0, 100));

        const originalMarkdown = toMarkdown(rawOriginalText);

        // Step 3: Execute change (Two-Step AI OR legacy amended_text)
        console.log('[handleAmendOperation] Step 3: Executing change...');
        let amendedMarkdown: string;

        // Check if we have API credentials and change_description for two-step AI
        if (apiKey && model && op.change_description) {
            console.log('[handleAmendOperation] Using two-step AI execution');
            amendedMarkdown = await executeAmendChange({
                originalMarkdown,
                changeDescription: op.change_description,
                userInstruction: op.user_instruction || op.change_description,
                apiKey,
                model
            });
            amendedMarkdown = toMarkdown(amendedMarkdown);
        } else if (op.amended_text) {
            // Fallback to legacy amended_text
            console.log('[handleAmendOperation] Using legacy amended_text');
            amendedMarkdown = toMarkdown(stripFormattingMarkers(op.amended_text).trim());
        } else {
            console.error('[handleAmendOperation] FAILED: No change_description or amended_text');
            return { success: false, error: 'Missing change_description or amended_text' };
        }

        console.log('[handleAmendOperation] Step 3 DONE: Amended text ready');
        console.log('[handleAmendOperation] Amended:', amendedMarkdown.substring(0, 100));

        // Step 4: Calculate diff
        console.log('[handleAmendOperation] Step 4: Calculating diff...');
        const dmp = new diff_match_patch();
        let dmpDiffs = dmp.diff_main(originalMarkdown, amendedMarkdown);
        dmp.diff_cleanupSemantic(dmpDiffs);
        let textChanges = convertDmpToTextChanges(dmpDiffs);

        if (textChanges.every(c => c.type === 'equal')) {
            console.log('[handleAmendOperation] No changes detected (texts are identical)');
            return { success: true };
        }

        console.log('[handleAmendOperation] Step 4 DONE: Diff calculated');
        console.log('[handleAmendOperation] Diff preview:', formatDiffForAI(textChanges));

        // Step 5: AI Self-Validation (if we have API credentials)
        if (apiKey && model) {
            console.log('[handleAmendOperation] Step 5: AI Self-Validation...');
            const validation = await validateChangesWithAI({
                originalText: originalMarkdown,
                userInstruction: op.user_instruction || op.change_description || 'Amend the text',
                changeDescription: op.change_description || 'Amendment',
                amendedText: amendedMarkdown,
                calculatedDiff: textChanges,
                apiKey,
                model
            });

            if (!validation.approved && validation.corrected_text) {
                console.log('[handleAmendOperation] AI self-correction triggered');
                console.log('[handleAmendOperation] Feedback:', validation.feedback);

                // Use corrected text and recalculate diff
                const correctedMarkdown = toMarkdown(validation.corrected_text);
                dmpDiffs = dmp.diff_main(originalMarkdown, correctedMarkdown);
                dmp.diff_cleanupSemantic(dmpDiffs);
                textChanges = convertDmpToTextChanges(dmpDiffs);
                amendedMarkdown = correctedMarkdown;

                console.log('[handleAmendOperation] Corrected diff:', formatDiffForAI(textChanges));
            } else {
                console.log('[handleAmendOperation] ✓ AI approved all changes');
            }
        } else {
            console.log('[handleAmendOperation] Step 5: Skipping AI validation (no API credentials)');
        }

        // Step 6: Convert to TextChange format for Word (with proper insert_after/delete_after)
        console.log('[handleAmendOperation] Step 6: Converting to Word operations...');
        const textChangeList = findTextChanges(originalMarkdown, amendedMarkdown);

        if (textChangeList.length === 0) {
            console.log('[handleAmendOperation] No changes to apply');
            return { success: true };
        }

        console.log('[handleAmendOperation] Step 6 DONE: Found', textChangeList.length, 'changes');
        textChangeList.forEach((c, i) => {
            if (c.type === 'replace') {
                console.log(`[handleAmendOperation]   [${i}] REPLACE: "${c.find.substring(0, 40)}..." → "${c.replace.substring(0, 40)}..."`);
            } else if (c.type === 'insert_after') {
                console.log(`[handleAmendOperation]   [${i}] INSERT_AFTER: anchor="${c.anchor.substring(0, 40)}" text="${c.text}"`);
            } else if (c.type === 'delete_after') {
                console.log(`[handleAmendOperation]   [${i}] DELETE_AFTER: anchor="${c.anchor.substring(0, 40)}" delete="${c.textToDelete.substring(0, 30)}..."`);
            }
        });

        // Step 7: Load track changes state
        console.log('[handleAmendOperation] Step 7: Loading track changes state...');
        context.document.load('changeTrackingMode');
        await context.sync();
        const originalMode = context.document.changeTrackingMode;
        const wasAlreadyTracking = originalMode === Word.ChangeTrackingMode.trackAll
            || originalMode === Word.ChangeTrackingMode.trackMineOnly
            || originalMode === 'TrackAll'
            || originalMode === 'TrackMineOnly';
        console.log('[handleAmendOperation] Step 7 DONE: Mode =', originalMode);

        // Step 8: Enable track changes
        console.log('[handleAmendOperation] Step 8: Enabling track changes...');
        if (!wasAlreadyTracking) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
            console.log('[handleAmendOperation] Track changes ENABLED');
        }

        // Step 9: Apply changes using new TextChange handler
        console.log('[handleAmendOperation] Step 9: Applying', textChangeList.length, 'changes...');
        const { successCount, failedCount } = await applyTextChanges(context, startPara, textChangeList);

        // Step 10: Restore state
        console.log('[handleAmendOperation] Step 10: Restoring track changes state...');
        if (!wasAlreadyTracking) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }

        console.log('[handleAmendOperation] ══════════════════════════════════════');
        console.log(`[handleAmendOperation] SUCCESS - ${successCount}/${textChangeList.length} changes applied`);

        if (failedCount > 0) {
            console.warn(`[handleAmendOperation] ${failedCount} changes could not be applied`);
        }

        return { success: true };

    } catch (error: any) {
        console.error('[handleAmendOperation] ══════════════════════════════════════');
        console.error('[handleAmendOperation] EXCEPTION:', error?.message || error);
        console.error('[handleAmendOperation] Stack:', error?.stack);
        console.error('[handleAmendOperation] ══════════════════════════════════════');
        return { success: false, error: error?.message || 'Unknown exception' };
    }
}

/**
 * Fallback: Apply full paragraph replacement when diff fails
 */
async function applyFullParagraphReplacement(
    context: any,
    startPara: any,
    lookup: { [id: number]: any },
    targetId: number,
    amendedText: string,
    op: any
): Promise<{ success: boolean; error?: string }> {
    console.warn('[handleAmendOperation] FALLBACK: Using full paragraph replacement');

    try {
        const endId = op.end_id || targetId;
        const endPara = lookup[endId] || startPara;
        const range = startPara.getRange('Start').expandTo(endPara.getRange('End'));

        context.document.load('changeTrackingMode');
        await context.sync();
        const originalMode = context.document.changeTrackingMode;
        const wasAlreadyTracking = originalMode === Word.ChangeTrackingMode.trackAll
            || originalMode === Word.ChangeTrackingMode.trackMineOnly
            || originalMode === 'TrackAll'
            || originalMode === 'TrackMineOnly';

        if (!wasAlreadyTracking) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
        }

        range.insertText(amendedText, Word.InsertLocation.replace);
        await context.sync();

        if (!wasAlreadyTracking) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }

        console.log('[handleAmendOperation] Fallback replacement complete');
        return { success: true };

    } catch (error: any) {
        console.error('[handleAmendOperation] Fallback FAILED:', error?.message);
        return { success: false, error: error?.message || 'Fallback failed' };
    }
}




// === STYLE NORMALIZATION ===

export function normalizeStyleName(
    aiStyle: string,
    documentStyles: StyleInfo[]
): string | null {

    // Case-insensitive match
    const lowerStyle = aiStyle.toLowerCase().trim();

    for (const style of documentStyles) {
        if (style.name.toLowerCase() === lowerStyle) {
            return style.name; // Return exact name from document
        }
    }

    // Common aliases
    const aliases: Record<string, string[]> = {
        'heading 1': ['h1', 'title', 'heading1'],
        'heading 2': ['h2', 'heading2'],
        'heading 3': ['h3', 'clause heading', 'heading3'],
        'heading 4': ['h4', 'heading4'],
        'normal': ['body', 'body text', 'paragraph', 'text', 'standard']
    };

    for (const [styleName, aliasList] of Object.entries(aliases)) {
        if (aliasList.includes(lowerStyle)) {
            // Find if this aliased style exists in document
            // e.g. "h3" maps to "heading 3", we check if doc has "Heading 3"
            const match = documentStyles.find(s => s.name.toLowerCase() === styleName);
            if (match) return match.name;
        }
    }

    return null;
}

/**
 * Handle INSERT operation - Uses Native Track Changes with State Preservation
 * 
 * Approach:
 * 1. Detect user's current track changes state
 * 2. Enable track changes if not already on
 * 3. Insert paragraph (Word auto-tracks as insertion)
 * 4. Apply list formatting if reference is a list item
 * 5. Restore original state
 */
async function handleInsertOperation(
    context: any,
    op: any,
    lookup: { [id: number]: any },
    documentStyles: StyleInfo[],
    author: string
): Promise<boolean> {
    const insertAfterId = op.insert_after;
    let refParagraph = lookup[insertAfterId];

    if (!refParagraph) {
        console.error('[handleInsertOperation] Anchor paragraph not found:', insertAfterId);
        return false;
    }

    const content = op.content || '';
    if (!content) return false;

    // Clean bullet characters if present (Word will handle formatting)
    let cleanContent = content;
    if (/^[\u2022\u2023\u25E6\u2043\u2219•]\s/.test(content)) {
        cleanContent = content.replace(/^[\u2022\u2023\u25E6\u2043\u2219•]\s/, '');
    }

    console.log(`[handleInsertOperation] Inserting after P${insertAfterId}:`, cleanContent.substring(0, 50));
    console.log('[INSERT DEBUG] ========== START ==========');
    console.log('[INSERT DEBUG] Content to insert:', cleanContent.substring(0, 50));

    try {
        // ═══════════════════════════════════════════════════════════
        // STEP 1: Detect user's current track changes state
        // ═══════════════════════════════════════════════════════════
        context.document.load('changeTrackingMode');
        console.log('[INSERT DEBUG] Loading document tracking state...');
        refParagraph.load('isListItem');
        await context.sync();
        console.log('[INSERT DEBUG] Current changeTrackingMode:', context.document.changeTrackingMode);

        const originalMode = context.document.changeTrackingMode;
        const wasAlreadyTracking = originalMode === Word.ChangeTrackingMode.trackAll
            || originalMode === Word.ChangeTrackingMode.trackMineOnly
            || originalMode === 'TrackAll'
            || originalMode === 'TrackMineOnly';

        console.log('[handleInsertOperation] Original tracking mode:', originalMode, '| Already tracking:', wasAlreadyTracking);
        console.log('[handleInsertOperation] Reference isListItem:', refParagraph.isListItem);

        // ═══════════════════════════════════════════════════════════
        // STEP 2: Insert EMPTY paragraph (Structural, not tracked yet)
        // ═══════════════════════════════════════════════════════════
        console.log('[INSERT DEBUG] Inserting EMPTY paragraph structure...');
        const newPara = refParagraph.insertParagraph('', 'After');
        context.trackedObjects.add(newPara);
        await context.sync();

        // ═══════════════════════════════════════════════════════════
        // STEP 3: Apply formatting (While empty, before tracking)
        // ═══════════════════════════════════════════════════════════

        // If reference is a list item and we want list formatting, try to inherit
        if (refParagraph.isListItem && op.list_level !== undefined) {
            try {
                // Try to attach to the same list as reference
                newPara.load('listItem');
                await context.sync();

                // Attempt to set list level
                if (newPara.listItem) {
                    newPara.listItem.level = op.list_level;
                    console.log('[handleInsertOperation] Set list level:', op.list_level);
                }
            } catch (listError) {
                console.warn('[handleInsertOperation] Could not set list formatting:', listError);
            }
        }

        // Apply style if specified
        if (op.style) {
            const normalizedStyle = normalizeStyleName(op.style, documentStyles);
            if (normalizedStyle) {
                newPara.style = normalizedStyle;
                console.log(`[handleInsertOperation] Applied style: ${normalizedStyle}`);
            } else {
                console.warn(`[handleInsertOperation] Style "${op.style}" not found`);
            }
        }

        // Apply explicit font properties
        // SMART FONT SIZE: Override AI's fontSize if it doesn't match document body style
        const bodyStyle = documentStyles.find(s => s.usedForBody);
        const headingStyle = documentStyles.find(s => s.usedForHeadings);
        const isHeadingContent = /^\d+\.\s+[A-Z]{2,}/.test(cleanContent); // e.g., "7. CONFIDENTIALITY"

        // DEBUG: Trace style detection
        console.log('[INSERT STYLE DEBUG] documentStyles count:', documentStyles.length);
        console.log('[INSERT STYLE DEBUG] bodyStyle found:', !!bodyStyle, bodyStyle?.name, bodyStyle?.font?.size);
        console.log('[INSERT STYLE DEBUG] headingStyle found:', !!headingStyle, headingStyle?.name);
        console.log('[INSERT STYLE DEBUG] isHeadingContent:', isHeadingContent);
        console.log('[INSERT STYLE DEBUG] op.fontSize:', op.fontSize);

        let effectiveFontSize = op.fontSize;
        if (bodyStyle?.font?.size && !isHeadingContent) {
            // For body/sub-clause text, use the detected body font size
            if (op.fontSize && op.fontSize !== bodyStyle.font.size) {
                console.log(`[handleInsertOperation] Overriding AI fontSize ${op.fontSize} → ${bodyStyle.font.size} (body style)`);
                effectiveFontSize = bodyStyle.font.size;
            }
        } else if (headingStyle?.font?.size && isHeadingContent) {
            // For heading text, use detected heading font size
            if (op.fontSize && op.fontSize !== headingStyle.font.size) {
                console.log(`[handleInsertOperation] Overriding AI fontSize ${op.fontSize} → ${headingStyle.font.size} (heading style)`);
                effectiveFontSize = headingStyle.font.size;
            }
        }

        if (op.font) {
            newPara.font.name = op.font;
        }
        if (effectiveFontSize) {
            newPara.font.size = effectiveFontSize;
        }
        if (op.bold !== undefined) {
            newPara.font.bold = op.bold;
        }
        if (op.italic !== undefined) {
            newPara.font.italic = op.italic;
        }
        if (op.underline !== undefined) {
            newPara.font.underline = op.underline ? 'Single' : 'None';
        }
        if (op.color && op.color !== 'auto') {
            newPara.font.color = op.color;
        }

        await context.sync(); // Commit formatting

        // ═══════════════════════════════════════════════════════════
        // STEP 4: Enable Track Changes & Insert Content
        // ═══════════════════════════════════════════════════════════

        // Force Enable TC
        console.log('[INSERT DEBUG] Setting changeTrackingMode to trackAll...');
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        console.log('[handleInsertOperation] Track changes forcibly ENABLED');
        // CRITICAL SYNC
        await context.sync();

        // Check if mode stuck
        context.document.load('changeTrackingMode');
        await context.sync();
        console.log('[INSERT DEBUG] Mode verified:', context.document.changeTrackingMode);

        // Insert Text into the styled empty paragraph
        // Parse markdown **bold** markers and apply formatting
        console.log('[INSERT DEBUG] Inserting content into empty para...');

        // Parse markdown bold markers: **text** -> bold
        const boldPattern = /\*\*(.+?)\*\*/g;
        const boldRanges: { start: number; end: number; text: string }[] = [];
        let plainContent = cleanContent;
        let offset = 0;

        // Find all bold markers and calculate positions in plain text
        let match;
        while ((match = boldPattern.exec(cleanContent)) !== null) {
            const markerStart = match.index - offset;
            const boldText = match[1];
            boldRanges.push({
                start: markerStart,
                end: markerStart + boldText.length,
                text: boldText
            });
            offset += 4; // Account for removed ** markers (2 at start, 2 at end)
        }

        // Remove markdown markers from content
        plainContent = cleanContent.replace(/\*\*(.+?)\*\*/g, '$1');

        console.log('[INSERT DEBUG] Bold ranges found:', boldRanges.length);
        if (boldRanges.length > 0) {
            console.log('[INSERT DEBUG] Bold sections:', boldRanges.map(r => `"${r.text}"`).join(', '));
        }

        // Insert plain text first
        const range = newPara.getRange('Content');
        range.insertText(plainContent, 'Replace');
        await context.sync();

        // Apply bold formatting to marked ranges
        if (boldRanges.length > 0) {
            for (const boldRange of boldRanges) {
                try {
                    // Search for the bold text within the paragraph
                    const searchResults = newPara.search(boldRange.text, { matchCase: true });
                    searchResults.load('items');
                    await context.sync();

                    if (searchResults.items.length > 0) {
                        searchResults.items[0].font.bold = true;
                        await context.sync();
                        console.log('[INSERT DEBUG] Applied bold to:', boldRange.text);
                    }
                } catch (boldError: any) {
                    console.warn('[INSERT DEBUG] Could not apply bold to:', boldRange.text, boldError?.message);
                }
            }
        }

        // ═══════════════════════════════════════════════════════════
        // STEP 5: Restore original state
        // ═══════════════════════════════════════════════════════════
        if (!wasAlreadyTracking) {
            console.log('[INSERT DEBUG] Restoring original mode...');
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
            console.log('[handleInsertOperation] Track changes DISABLED (restored)');
        } else {
            console.log('[handleInsertOperation] Track changes left ON (user preference)');
        }

        console.log('[handleInsertOperation] SUCCESS - Empty-First Strategy applied');
        console.log('[INSERT DEBUG] ========== END ==========');
        return true;

    } catch (e: any) {
        console.error('[handleInsertOperation] ERROR:', e?.message || e);
        console.error('[handleInsertOperation] Error stack:', e?.stack);
        return false;
    }
}



/**
 * Handle DELETE operation - Uses Native Track Changes with State Preservation
 * 
 * Instead of OOXML manipulation, we:
 * 1. Detect user's current track changes state
 * 2. Enable track changes if not already on
 * 3. Delete the range (Word shows as strikethrough)
 * 4. Restore original state
 */
async function handleDeleteOperation(
    context: any,
    op: any,
    lookup: { [id: number]: any }
): Promise<boolean> {
    const deleteId = op.target_id;
    let deletePara = lookup[deleteId];

    // FALLBACK: If not found by ID, try content-based matching
    if (!deletePara && op.original_text) {
        console.log('[handleDeleteOperation] ID lookup failed, trying content match...');
        deletePara = await findParagraphByContent(context, op.original_text);
    }

    if (!deletePara) {
        console.error('[handleDeleteOperation] Delete target not found:', deleteId);
        return false;
    }

    console.log('[handleDeleteOperation] Deleting P' + deleteId + ' with native track changes');

    try {
        // ═══════════════════════════════════════════════════════════
        // STEP 1: Detect user's current track changes state
        // ═══════════════════════════════════════════════════════════
        context.document.load('changeTrackingMode');
        await context.sync();

        const originalMode = context.document.changeTrackingMode;
        const wasAlreadyTracking = originalMode === Word.ChangeTrackingMode.trackAll
            || originalMode === Word.ChangeTrackingMode.trackMineOnly
            || originalMode === 'TrackAll'
            || originalMode === 'TrackMineOnly';

        console.log('[handleDeleteOperation] Original tracking mode:', originalMode, '| Already tracking:', wasAlreadyTracking);

        // ═══════════════════════════════════════════════════════════
        // STEP 2: Enable track changes if not already on
        // ═══════════════════════════════════════════════════════════
        // Always force enable track changes and SYNC to ensure it's active
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        console.log('[handleDeleteOperation] Track changes forcibly ENABLED');
        // CRITICAL: Always sync before delete so Word knows TC state
        await context.sync();

        // ═══════════════════════════════════════════════════════════
        // STEP 3: Delete the paragraph (Word shows as strikethrough)
        // ═══════════════════════════════════════════════════════════
        deletePara.delete();
        await context.sync();  // Commit the delete immediately
        console.log('[handleDeleteOperation] Paragraph deleted with track changes');

        // ═══════════════════════════════════════════════════════════
        // STEP 4: Restore original state (only if we changed it)
        // ═══════════════════════════════════════════════════════════
        if (!wasAlreadyTracking) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
            console.log('[handleDeleteOperation] Track changes DISABLED (restored)');
        } else {
            console.log('[handleDeleteOperation] Track changes left ON (user preference)');
        }

        console.log('[handleDeleteOperation] SUCCESS - deleted with native track changes');
        return true;

    } catch (e: any) {
        console.error('[handleDeleteOperation] ERROR:', e?.message || e);
        console.error('[handleDeleteOperation] Error stack:', e?.stack);
        return false;
    }
}
