/**
 * AMEND_SIMPLE Handler for Vibe Legal
 * Full paragraph replacement with surgical diff
 */

import { AmendSimpleOperation, ExecutionResult } from '../../types';
import { computeSurgicalDiff } from '../diff';
import { applyDiffsToOoxml, extractTextFromOoxml } from '../ooxml';
import { getParagraphOoxml, insertOoxmlAtParagraph } from '../document';

/**
 * Handle AMEND_SIMPLE operation.
 * Replaces paragraph content using surgical diff with track changes.
 */
export async function handleAmendSimple(
    context: any,
    operation: AmendSimpleOperation,
    author: string
): Promise<ExecutionResult> {
    const targetId = operation.target_id;

    if (targetId === undefined || targetId < 0) {
        return {
            success: false,
            operationType: 'AMEND_SIMPLE',
            error: 'Invalid target_id'
        };
    }

    try {
        console.log(`[handleAmendSimple] target_id=${targetId}, amended_text="${operation.amended_text?.substring(0, 80)}..."`);

        // Get current OOXML
        const currentOoxml = await getParagraphOoxml(context, targetId);
        const currentText = extractTextFromOoxml(currentOoxml);
        console.log(`[handleAmendSimple] Current text: "${currentText.substring(0, 100)}..."`);

        // Compute diff
        const diffs = computeSurgicalDiff(currentText, operation.amended_text);

        if (diffs.length === 0) {
            return {
                success: true,
                operationType: 'AMEND_SIMPLE',
                targetId,
                description: 'No changes needed'
            };
        }

        // Apply diffs to OOXML with track changes
        const result = applyDiffsToOoxml(currentOoxml, diffs, author);

        if (!result.hasChanges) {
            return {
                success: true,
                operationType: 'AMEND_SIMPLE',
                targetId,
                description: 'No changes applied'
            };
        }

        // Insert modified OOXML
        console.log(`[handleAmendSimple] Inserting modified OOXML for P${targetId}`);
        await insertOoxmlAtParagraph(context, targetId, result.oxml, 'Replace');
        console.log(`[handleAmendSimple] SUCCESS for P${targetId}`);

        return {
            success: true,
            operationType: 'AMEND_SIMPLE',
            targetId,
            description: operation.description || `Modified paragraph ${targetId}`
        };
    } catch (error: any) {
        return {
            success: false,
            operationType: 'AMEND_SIMPLE',
            targetId,
            error: error.message
        };
    }
}
