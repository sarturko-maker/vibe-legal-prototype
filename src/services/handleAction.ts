/**
 * Handle Action - Rewritten with Legacy Architecture
 * 
 * KEY DIFFERENCES FROM PREVIOUS VERSION:
 * 1. Uses lookup{} object with trackedObjects.add() for live paragraph references
 * 2. All operations execute in SINGLE Word.run context
 * 3. Uses sectionRange.expandTo() for multi-paragraph AMEND (legacy pattern)
 * 4. Uses applyRedlineToOxml for diff application (ported from legacy)
 * 5. Uses insertBlocks for insertions (ported from legacy)
 * 
 * Source: Ported from taskpane.legacy.tsx handleAction pattern
 */

import { log, logError, logWarn } from '../utils/logger';
import { buildRouterSystemPrompt, buildSideInstruction, buildDealContextSection, SideState, DetectedParties, DealContextState } from '../prompts/systemPrompt';
import { callGeminiRouter } from './gemini/client';
import { ContractMap, ParagraphInfo, Message } from '../types';
import { buildSimpleContractMap } from './document/contractMap';
import { applyRedlineToOxml } from './ooxml/applyRedline';
import { insertBlocks, StyleToken } from './ooxml/insertBlocks';
import { stripFormattingMarkers } from './formatting/markdownParser';
import { validateAIResponse, ValidationResult, formatValidationError } from './validation';
import { buildRiskToleranceInstruction } from '../prompts/riskTolerancePrompt';
import { RiskTolerance } from '../types/state';

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
        const styleMap = new Map<string, {
            count: number;
            exampleIndex: number;
            font: any;
            paragraphFormat: any;
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
                        size: p.font.size,
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
                    }
                });
            } else {
                styleMap.get(styleName)!.count++;
            }
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
    styleMenu: StyleInfo[] = [] // Default to empty array
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

        // 2.2 AMENDs - reverse order
        const sortedAmends = [...amends].sort((a, b) =>
            (b.target_id || 0) - (a.target_id || 0)
        );

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
                const result = await handleAmendOperation(context, op, lookup, author);
                if (result) successCount++; else errorCount++;
            } catch (e: any) {
                console.error('AMEND failed:', e);
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
 * Handle AMEND operation - uses legacy sectionRange.expandTo pattern
 */
async function handleAmendOperation(
    context: any,
    op: any,
    lookup: { [id: number]: any },
    author: string
): Promise<boolean> {
    const targetId = op.target_id;
    console.log('[handleAmendOperation] Looking for target_id:', targetId, '| Type:', typeof targetId);
    console.log('[handleAmendOperation] Lookup keys sample:', Object.keys(lookup).slice(0, 5).join(', '));

    let startPara = lookup[targetId];
    console.log('[handleAmendOperation] Direct lookup result:', startPara ? 'FOUND' : 'NOT FOUND');

    // FALLBACK: If not found by ID, try content-based matching
    if (!startPara && op.original_text) {
        console.log('[handleAmendOperation] ID lookup failed, trying content match...');
        console.log('[handleAmendOperation] original_text snippet:', op.original_text?.substring(0, 50));
        startPara = await findParagraphByContent(context, op.original_text);
    }

    if (!startPara) {
        console.error('[handleAmendOperation] Paragraph not found by ID or content:', targetId);
        console.error('[handleAmendOperation] Operation details:', JSON.stringify(op, null, 2));
        return false;
    }

    // Get range (for AMEND with range, use expandTo like legacy)
    const endId = op.end_id || targetId;
    const endPara = lookup[endId] || startPara;

    console.log('[handleAmendOperation] Range: P' + targetId + ' to P' + endId);

    // Create range spanning start to end (LEGACY PATTERN)
    const range = startPara.getRange('Start').expandTo(endPara.getRange('End'));
    range.load('text');
    const rangeOxml = range.getOoxml();
    await context.sync();

    console.log('[handleAmendOperation] Original text (' + range.text.length + ' chars):', range.text.substring(0, 100));
    console.log('[handleAmendOperation] Modified text (' + (op.amended_text?.length || 0) + ' chars):', op.amended_text?.substring(0, 100));

    // CRITICAL: Strip markdown markers before applying to document
    const cleanAmendedText = stripFormattingMarkers(op.amended_text || '');

    try {
        console.log('[handleAmendOperation] Calling applyRedlineToOxml...');
        console.log('[handleAmendOperation] OOXML length:', rangeOxml.value?.length || 'undefined');

        // Apply redline using legacy engine pattern
        const result = applyRedlineToOxml(
            rangeOxml.value,
            range.text,
            cleanAmendedText,
            author
        );

        console.log('[handleAmendOperation] applyRedlineToOxml result - hasChanges:', result.hasChanges);

        if (result.hasChanges) {
            console.log('[handleAmendOperation] Inserting modified OOXML (length:', result.oxml?.length, ')');
            range.insertOoxml(result.oxml, Word.InsertLocation.replace);
            await context.sync();
            console.log('[handleAmendOperation] SUCCESS');
            return true;
        } else {
            console.log('[handleAmendOperation] No changes detected');
            return true;  // Not an error, just no changes
        }
    } catch (e: any) {
        console.error('[handleAmendOperation] ERROR in redline/insert phase:', e?.message || e);
        console.error('[handleAmendOperation] Error stack:', e?.stack);
        return false;
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
 * Handle INSERT operation
 * Supports numbered lists, bullet points, and plain text
 * Enforces track changes
 */
async function handleInsertOperation(
    context: any,
    op: any,
    lookup: { [id: number]: any },
    documentStyles: StyleInfo[], // Changed from styleCache for rich style support
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

    console.log(`[handleInsertOperation] Inserting after P${insertAfterId}:`, content.substring(0, 50));

    // CHECK FOR LIST ITEM (Bullet/Numbering)
    if (op.list_level !== undefined) {
        // Use existing insertAsListItem helper
        const result = await insertAsListItem(context, refParagraph, content, op.list_level, author);
        return result.success;
    }

    // Check if content looks like a bullet
    if (/^[\u2022\u2023\u25E6\u2043\u2219•]\s/.test(content)) {
        const cleanContent = content.replace(/^[\u2022\u2023\u25E6\u2043\u2219•]\s/, '');
        return (await insertAsListItem(context, refParagraph, cleanContent, 0, author)).success;
    }

    // PLAIN TEXT / STYLE APPLICATION
    try {
        console.log('[handleInsertOperation] Phase 1: Inserting paragraph');
        const newPara = refParagraph.insertParagraph(content, 'After');

        // APPLY STYLE
        if (op.style) {
            const normalizedStyle = normalizeStyleName(op.style, documentStyles);
            if (normalizedStyle) {
                // Check if style exists/is valid by trying to leverage checking or just set it.
                // Word will trigger error if style doesn't exist? Or just ignore? usually valid style name required.
                // Our normalization ensures we use an existing name from `documentStyles`.
                newPara.style = normalizedStyle;
                // Important: Load style to verify/sync
                newPara.load('style');
                console.log(`[handleInsertOperation] Applied style: ${normalizedStyle}`);
            } else {
                console.warn(`[handleInsertOperation] Style "${op.style}" not found in document`);
            }
        }

        // APPLY EXPLICIT FONT (for consistent INSERT formatting - ADR-011)
        if (op.font) {
            newPara.font.name = op.font;
            console.log(`[handleInsertOperation] Applied font: ${op.font}`);
        }
        if (op.fontSize) {
            newPara.font.size = op.fontSize;
            console.log(`[handleInsertOperation] Applied fontSize: ${op.fontSize}`);
        }

        // APPLY EXTENDED FORMATTING
        if (op.bold !== undefined) {
            newPara.font.bold = op.bold;
            console.log(`[handleInsertOperation] Applied bold: ${op.bold}`);
        }
        if (op.italic !== undefined) {
            newPara.font.italic = op.italic;
            console.log(`[handleInsertOperation] Applied italic: ${op.italic}`);
        }
        if (op.underline !== undefined) {
            newPara.font.underline = op.underline ? 'Single' : 'None';
            console.log(`[handleInsertOperation] Applied underline: ${op.underline}`);
        }
        if (op.color && op.color !== 'auto') {
            newPara.font.color = op.color;
            console.log(`[handleInsertOperation] Applied color: ${op.color}`);
        }

        context.trackedObjects.add(newPara);
        await context.sync();

        console.log('[handleInsertOperation] Phase 2: Applying track changes');
        await applyTrackChangesToParagraph(context, newPara, author);

        console.log('[handleInsertOperation] SUCCESS');
        return true;

    } catch (e: any) {
        console.error('[handleInsertOperation] Error:', e?.message || e);
        return false;
    }
}

/**
 * Insert content as a list item with track changes
 * Two-phase approach: 1) Insert structure, 2) Apply track changes
 * CRITICAL: Every code path MUST apply track changes
 */
async function insertAsListItem(
    context: any,
    refParagraph: any,
    content: string,
    level: number,
    author: string
): Promise<{ success: boolean; inserted: number }> {

    try {
        refParagraph.load('isListItem');
        await context.sync();

        console.log('[insertAsListItem] Reference isListItem:', refParagraph.isListItem);

        if (!refParagraph.isListItem) {
            // Not a list - use plain insert WITH TRACK CHANGES
            console.warn('[insertAsListItem] Reference is not a list item, using plain insert with track changes');
            return await insertPlainWithTrackChanges(context, refParagraph, content, author);
        }

        // Get OOXML and extract numId
        const ooxmlResult = refParagraph.getOoxml();
        await context.sync();

        const refOoxml = ooxmlResult.value;
        console.log('[insertAsListItem] OOXML length:', refOoxml.length);

        // DEBUG: Log actual context around numId
        const numIdIndex = refOoxml.indexOf('numId');
        console.log('[insertAsListItem] numId found at index:', numIdIndex);
        if (numIdIndex !== -1) {
            console.log('[insertAsListItem] numId context:', refOoxml.substring(numIdIndex - 10, numIdIndex + 50));
        }

        // More flexible regex - handles various attribute formats
        const numIdMatch = refOoxml.match(/w:numId[^>]*w:val\s*=\s*"(\d+)"/);

        if (!numIdMatch) {
            // numId not found - use plain insert WITH TRACK CHANGES
            console.warn('[insertAsListItem] numId not found, using plain insert with track changes');
            return await insertPlainWithTrackChanges(context, refParagraph, content, author);
        }

        const numId = parseInt(numIdMatch[1], 10);
        console.log('[insertAsListItem] Extracted numId:', numId);

        // ═══════════════════════════════════════════════════════════
        // PHASE 1: Insert bullet structure
        // ═══════════════════════════════════════════════════════════
        const bulletOoxml = buildMinimalBulletOoxml(content, numId, 0);
        refParagraph.getRange('End').insertOoxml(bulletOoxml, 'After');
        await context.sync();
        console.log('[insertAsListItem] Phase 1 complete: Bullet structure inserted');

        // ═══════════════════════════════════════════════════════════
        // PHASE 2: Apply track changes
        // ═══════════════════════════════════════════════════════════
        const nextPara = refParagraph.getNext();
        nextPara.load('text');
        context.trackedObjects.add(nextPara);
        await context.sync();

        console.log('[insertAsListItem] Phase 2: New paragraph text:', nextPara.text?.substring(0, 50) || '(empty)');
        await applyTrackChangesToParagraph(context, nextPara, author);

        console.log('[insertAsListItem] SUCCESS - inserted with minimal OOXML + track changes');
        return { success: true, inserted: 1 };

    } catch (e: any) {
        console.error('[insertAsListItem] Error:', e?.message || e);
        // CRITICAL: Fallback MUST also apply track changes
        return await insertPlainWithTrackChanges(context, refParagraph, content, author);
    }
}

/**
 * Insert plain paragraph with track changes
 * Two-phase approach ensures track changes are ALWAYS applied
 * Used as fallback when list insertion fails
 */
async function insertPlainWithTrackChanges(
    context: any,
    refParagraph: any,
    content: string,
    author: string
): Promise<{ success: boolean; inserted: number }> {

    try {
        console.log('[insertPlainWithTrackChanges] Phase 1: Inserting paragraph');
        const newPara = refParagraph.insertParagraph(content, 'After');
        context.trackedObjects.add(newPara);
        await context.sync();

        console.log('[insertPlainWithTrackChanges] Phase 2: Applying track changes');
        await applyTrackChangesToParagraph(context, newPara, author);

        console.log('[insertPlainWithTrackChanges] SUCCESS');
        return { success: true, inserted: 1 };

    } catch (e: any) {
        console.error('[insertPlainWithTrackChanges] Error:', e?.message || e);
        // Last resort - insert without track changes (better than nothing)
        try {
            refParagraph.insertParagraph(content, 'After');
            await context.sync();
            console.warn('[insertPlainWithTrackChanges] Inserted WITHOUT track changes (last resort)');
            return { success: true, inserted: 1 };
        } catch (lastError) {
            console.error('[insertPlainWithTrackChanges] Complete failure:', lastError);
            return { success: false, inserted: 0 };
        }
    }
}

/**
 * Build minimal OOXML for a bullet point paragraph
 * Much smaller (~800 bytes) than cloning full document package (175KB)
 */
function buildMinimalBulletOoxml(content: string, numId: number, level: number): string {
    const escapedContent = content
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');

    // Note: We don't include w:ins here - track changes will be applied separately
    // after the paragraph is inserted, because Word Online ignores w:ins in fresh inserts
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr>
              <w:numPr>
                <w:ilvl w:val="${level}"/>
                <w:numId w:val="${numId}"/>
              </w:numPr>
            </w:pPr>
            <w:r>
              <w:t>${escapedContent}</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text: string): string {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

/**
 * Apply track changes to a paragraph by modifying its OOXML
 * Wraps all runs in w:ins markup
 * Uses getRange('Whole').insertOoxml pattern for proper replacement
 */
async function applyTrackChangesToParagraph(context: any, paragraph: any, author: string): Promise<void> {
    try {
        const ooxmlResult = paragraph.getOoxml();
        await context.sync();

        const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const xmlDoc = parser.parseFromString(ooxmlResult.value, 'text/xml');

        // Check for parse errors
        const parseError = xmlDoc.querySelector('parsererror');
        if (parseError) {
            console.warn('[applyTrackChangesToParagraph] XML parse error');
            return;
        }

        // Find all <w:r> runs in the document
        const runs = xmlDoc.getElementsByTagNameNS(WORD_NS, 'r');
        const runsArray = Array.from(runs);

        if (runsArray.length === 0) {
            console.log('[applyTrackChangesToParagraph] No runs to wrap');
            return;
        }

        console.log(`[applyTrackChangesToParagraph] Found ${runsArray.length} runs to wrap`);

        const timestamp = new Date().toISOString();
        let insId = Math.floor(Math.random() * 10000000);

        for (const run of runsArray) {
            // Check if already inside w:ins by walking up parent hierarchy
            let parent = run.parentElement;
            let alreadyWrapped = false;
            while (parent) {
                if (parent.localName === 'ins') {
                    alreadyWrapped = true;
                    break;
                }
                parent = parent.parentElement;
            }

            if (alreadyWrapped) {
                console.log('[applyTrackChangesToParagraph] Run already wrapped, skipping');
                continue;
            }

            // Create w:ins wrapper
            const insNode = xmlDoc.createElementNS(WORD_NS, 'w:ins');
            insNode.setAttribute('w:id', String(insId++));
            insNode.setAttribute('w:author', author);
            insNode.setAttribute('w:date', timestamp);

            // Wrap: insert ins before run, then move run into ins
            run.parentNode?.insertBefore(insNode, run);
            insNode.appendChild(run);
        }

        console.log(`[applyTrackChangesToParagraph] Wrapped ${insId - Math.floor(Math.random() * 10000000)} runs`);

        // Serialize and apply using getRange('Whole').insertOoxml pattern
        const modifiedOoxml = serializer.serializeToString(xmlDoc);
        const paraRange = paragraph.getRange('Whole');
        paraRange.insertOoxml(modifiedOoxml, 'Replace');
        await context.sync();

        console.log('[applyTrackChangesToParagraph] Track changes applied successfully');

    } catch (e) {
        console.warn('[applyTrackChangesToParagraph] Error:', e);
        // Non-fatal - paragraph was still inserted
    }
}

/**
 * Handle DELETE operation
 * CRITICAL: Uses track changes (w:del markup) instead of actual deletion
 * This preserves paragraph IDs so subsequent operations target correct content
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

    console.log('[handleDeleteOperation] Marking as deleted (track change) P' + deleteId);

    // Get the paragraph's OOXML and text
    deletePara.load('text');
    const rangeOxml = deletePara.getOoxml();
    await context.sync();

    const originalText = deletePara.text || '';

    // Use applyRedlineToOxml with empty amended_text to create <w:del> track changes
    // This marks content as deleted WITHOUT actually removing it
    const result = applyRedlineToOxml(
        rangeOxml.value,
        originalText,
        '',  // Empty string = delete all content
        'Vibe AI'
    );

    if (result.hasChanges) {
        console.log('[handleDeleteOperation] Inserting track-change deleted OOXML');
        deletePara.insertOoxml(result.oxml, Word.InsertLocation.replace);
        await context.sync();
        console.log('[handleDeleteOperation] SUCCESS - marked as deleted with track changes');
        return true;
    } else {
        console.log('[handleDeleteOperation] No changes detected');
        return true;
    }
}

/**
 * Build style cache for insertBlocks
 */
async function buildStyleCache(context: any): Promise<Map<string, StyleToken>> {
    const styleCache = new Map<string, StyleToken>();

    // Add default Normal style
    styleCache.set('Normal', {
        token: 'Normal',
        styleId: 'Normal',
        styleName: 'Normal'
    });

    // We could expand this to load document styles if needed
    // For now, basic Normal support covers most insertion cases

    return styleCache;
}
