/**
 * DELETE Handler for Vibe Legal
 * Handles paragraph/clause deletion with track changes
 */

import { DeleteOperation, ExecutionResult } from '../../types';
import { getParagraphOoxml, insertOoxmlAtParagraph } from '../document';
import { createDeletionWrapper, cloneRunWithDelText, WORD_NS, generateRsid } from '../ooxml';

// Declare DOMParser for XML operations
declare var DOMParser: any;
declare var XMLSerializer: any;

/**
 * Handle DELETE operation.
 * Wraps paragraph content in deletion markers.
 */
export async function handleDelete(
    context: any,
    operation: DeleteOperation,
    author: string
): Promise<ExecutionResult> {
    const targetId = operation.target_id;

    if (targetId === undefined || targetId < 0) {
        return {
            success: false,
            operationType: 'DELETE',
            error: 'Invalid target_id'
        };
    }

    try {
        // Get current OOXML
        const currentOoxml = await getParagraphOoxml(context, targetId);

        // Wrap all content in deletion
        const deletedOoxml = wrapInDeletion(currentOoxml, author);

        // Replace paragraph with deleted version
        await insertOoxmlAtParagraph(context, targetId, deletedOoxml, 'Replace');

        return {
            success: true,
            operationType: 'DELETE',
            targetId,
            description: operation.description || `Deleted paragraph ${targetId}`
        };
    } catch (error: any) {
        return {
            success: false,
            operationType: 'DELETE',
            targetId,
            error: error.message
        };
    }
}

/**
 * Wrap entire paragraph content in deletion markers.
 */
function wrapInDeletion(oxml: string, author: string): string {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(oxml, 'application/xml');
        const sessionRsid = generateRsid();
        const date = new Date().toISOString();

        // Find paragraph
        const paragraphs = xmlDoc.getElementsByTagNameNS(WORD_NS, 'p');
        if (paragraphs.length === 0) return oxml;

        const para = paragraphs[0] as Element;

        // Get all runs
        const runs = para.getElementsByTagNameNS(WORD_NS, 'r');

        // Wrap each run in deletion
        for (let i = runs.length - 1; i >= 0; i--) {
            const run = runs[i] as Element;
            const parent = run.parentNode;
            if (!parent) continue;

            // Get text content
            const textEls = run.getElementsByTagNameNS(WORD_NS, 't');
            let text = '';
            for (let j = 0; j < textEls.length; j++) {
                text += textEls[j].textContent || '';
            }

            if (text.length > 0) {
                // Create deletion wrapper
                const delWrapper = createDeletionWrapper(xmlDoc, author, date);
                const delRun = cloneRunWithDelText(xmlDoc, run, text, sessionRsid);
                delWrapper.appendChild(delRun);

                // Replace run with deletion
                parent.replaceChild(delWrapper, run);
            }
        }

        const serializer = new XMLSerializer();
        return serializer.serializeToString(xmlDoc);
    } catch {
        return oxml;
    }
}
