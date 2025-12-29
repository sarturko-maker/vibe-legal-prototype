/**
 * AMEND Tree Handler for Vibe Legal
 * Processes paragraph operations (MODIFY, INSERT_AFTER, DELETE)
 * Source: vibe-legal-beta-0.2.yaml L10455-10650
 */

import { AmendTreeOperation, ParaOperation, ExecutionResult } from '../../types';
import { computeSurgicalDiff } from '../diff';
import { applyDiffsToOoxml, extractTextFromOoxml } from '../ooxml';
import { getParagraphOoxml, insertOoxmlAtParagraph } from '../document';
import { log, logWarn } from '../../utils/logger';

/**
 * Handle AMEND (tree-based) operation.
 * Processes multiple paragraph operations in order.
 */
export async function handleAmendTree(
    context: any,
    operation: AmendTreeOperation,
    author: string
): Promise<ExecutionResult> {
    const paraOps = operation.para_operations;

    if (!paraOps || paraOps.length === 0) {
        return {
            success: false,
            operationType: 'AMEND',
            error: 'No paragraph operations provided'
        };
    }

    try {
        // Sort operations: MODIFY first, then INSERT_AFTER, then DELETE last
        // This ensures indices remain valid during processing
        const modifyOps = paraOps.filter(op => op.action === 'MODIFY');
        const insertOps = paraOps.filter(op => op.action === 'INSERT_AFTER');
        const deleteOps = paraOps.filter(op => op.action === 'DELETE');

        const orderedOps = [...modifyOps, ...insertOps, ...deleteOps];

        let successCount = 0;
        let errorCount = 0;

        for (const paraOp of orderedOps) {
            const result = await processParaOperation(context, paraOp, author);
            if (result) {
                successCount++;
            } else {
                errorCount++;
            }
        }

        return {
            success: errorCount === 0,
            operationType: 'AMEND',
            description: `Processed ${successCount} operations, ${errorCount} errors`
        };
    } catch (error: any) {
        return {
            success: false,
            operationType: 'AMEND',
            error: error.message
        };
    }
}

/**
 * Process a single paragraph operation.
 */
async function processParaOperation(
    context: any,
    op: ParaOperation,
    author: string
): Promise<boolean> {
    const paraId = op.para_id;

    if (paraId === undefined || paraId < 0) {
        logWarn('Invalid para_id:', paraId);
        return false;
    }

    try {
        switch (op.action) {
            case 'MODIFY':
                return await processModify(context, paraId, op.new_text || '', author);

            case 'INSERT_AFTER':
                return await processInsertAfter(context, paraId, op.new_text || '', author);

            case 'DELETE':
                return await processDelete(context, paraId, author);

            default:
                logWarn('Unknown action:', op.action);
                return false;
        }
    } catch (error: any) {
        logWarn(`Para operation failed for P${paraId}:`, error.message);
        return false;
    }
}

/**
 * Process MODIFY operation - surgical diff on paragraph.
 */
async function processModify(
    context: any,
    paraId: number,
    newText: string,
    author: string
): Promise<boolean> {
    if (!newText) {
        logWarn(`MODIFY P${paraId}: missing new_text`);
        return false;
    }

    log(`MODIFY P${paraId}`);

    // Get current OOXML
    const currentOoxml = await getParagraphOoxml(context, paraId);
    const currentText = extractTextFromOoxml(currentOoxml);

    if (currentText === newText) {
        log(`MODIFY P${paraId}: no changes detected`);
        return true;
    }

    // Compute diff
    const diffs = computeSurgicalDiff(currentText, newText);

    if (diffs.length === 0) {
        return true;
    }

    // Apply diffs with track changes
    const result = applyDiffsToOoxml(currentOoxml, diffs, author);

    if (result.hasChanges) {
        await insertOoxmlAtParagraph(context, paraId, result.oxml, 'Replace');
        log(`MODIFY P${paraId}: applied ${result.changeCount} changes`);
    }

    return true;
}

/**
 * Process INSERT_AFTER operation - add new paragraph after target.
 */
async function processInsertAfter(
    context: any,
    paraId: number,
    newText: string,
    author: string
): Promise<boolean> {
    if (!newText) {
        logWarn(`INSERT_AFTER P${paraId}: missing new_text`);
        return false;
    }

    log(`INSERT_AFTER P${paraId}`);

    // Get reference paragraph for styling
    const refOoxml = await getParagraphOoxml(context, paraId);

    // Build simple insertion paragraph
    // TODO: Add track change wrapper for insertion
    const newOoxml = buildSimpleParagraph(newText);

    await insertOoxmlAtParagraph(context, paraId, newOoxml, 'After');
    log(`INSERT_AFTER P${paraId}: inserted new paragraph`);

    return true;
}

/**
 * Process DELETE operation - mark paragraph as deleted.
 */
async function processDelete(
    context: any,
    paraId: number,
    author: string
): Promise<boolean> {
    log(`DELETE P${paraId}`);

    // Import delete handler logic
    const { handleDelete } = await import('./delete');

    const result = await handleDelete(context, {
        type: 'DELETE',
        target_id: paraId
    }, author);

    return result.success;
}

/**
 * Build a simple paragraph OOXML.
 */
function buildSimpleParagraph(text: string): string {
    const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return `<w:p xmlns:w="${WORD_NS}"><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

/**
 * Escape XML special characters.
 */
function escapeXml(text: string): string {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}
