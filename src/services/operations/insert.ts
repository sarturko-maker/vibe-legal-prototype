/**
 * INSERT Handlers for Vibe Legal
 * Handles INSERT and INSERT_BLOCK operations
 */

import { InsertOperation, InsertBlockOperation, ExecutionResult } from '../../types';
import { getParagraphOoxml, insertOoxmlAtParagraph } from '../document';
import { createTextRun, createInsertionWrapper, WORD_NS } from '../ooxml';

// Declare DOMParser for XML operations
declare var DOMParser: any;
declare var XMLSerializer: any;

/**
 * Handle INSERT operation.
 * Inserts content after a target paragraph.
 */
export async function handleInsert(
    context: any,
    operation: InsertOperation,
    author: string
): Promise<ExecutionResult> {
    const insertAfter = operation.insert_after;

    if (insertAfter === undefined || insertAfter < 0) {
        return {
            success: false,
            operationType: 'INSERT',
            error: 'Invalid insert_after'
        };
    }

    try {
        console.log(`[handleInsert] insert_after=${insertAfter}, content="${operation.content?.substring(0, 80)}..."`);

        // Get reference paragraph OOXML for styling
        const refOoxml = await getParagraphOoxml(context, insertAfter);

        // Build new paragraph with insertion tracking
        const newOoxml = buildInsertionOoxml(refOoxml, operation.content, author);

        // Insert after target paragraph
        console.log(`[handleInsert] Inserting after P${insertAfter}`);
        await insertOoxmlAtParagraph(context, insertAfter, newOoxml, 'After');
        console.log(`[handleInsert] SUCCESS for P${insertAfter}`);

        return {
            success: true,
            operationType: 'INSERT',
            targetId: insertAfter,
            description: operation.description || `Inserted content after paragraph ${insertAfter}`
        };
    } catch (error: any) {
        return {
            success: false,
            operationType: 'INSERT',
            targetId: insertAfter,
            error: error.message
        };
    }
}

/**
 * Handle INSERT_BLOCK operation.
 * Inserts multi-paragraph content after a target.
 */
export async function handleInsertBlock(
    context: any,
    operation: InsertBlockOperation,
    author: string
): Promise<ExecutionResult> {
    const insertAfter = operation.insert_after;

    if (insertAfter === undefined || insertAfter < 0) {
        return {
            success: false,
            operationType: 'INSERT_BLOCK',
            error: 'Invalid insert_after'
        };
    }

    try {
        // Get reference paragraph OOXML for styling
        const refOoxml = await getParagraphOoxml(context, insertAfter);

        // Build block OOXML
        const blockOoxml = buildBlockOoxml(refOoxml, operation.block, author);

        // Insert after target paragraph
        await insertOoxmlAtParagraph(context, insertAfter, blockOoxml, 'After');

        return {
            success: true,
            operationType: 'INSERT_BLOCK',
            targetId: insertAfter,
            description: operation.description || `Inserted block after paragraph ${insertAfter}`
        };
    } catch (error: any) {
        return {
            success: false,
            operationType: 'INSERT_BLOCK',
            targetId: insertAfter,
            error: error.message
        };
    }
}

/**
 * Build OOXML for a single insertion with track changes.
 */
function buildInsertionOoxml(refOoxml: string, content: string, author: string): string {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(refOoxml, 'application/xml');

        // Create new paragraph
        const newPara = xmlDoc.createElementNS(WORD_NS, 'w:p');

        // Create insertion wrapper
        const insWrapper = createInsertionWrapper(xmlDoc, author);

        // Create text run
        const run = createTextRun(xmlDoc, content);
        insWrapper.appendChild(run);

        newPara.appendChild(insWrapper);

        // Serialize
        const serializer = new XMLSerializer();
        return serializer.serializeToString(newPara);
    } catch {
        // Fallback: simple text
        return `<w:p xmlns:w="${WORD_NS}"><w:r><w:t>${content}</w:t></w:r></w:p>`;
    }
}

/**
 * Build OOXML for multi-paragraph block insertion.
 */
function buildBlockOoxml(refOoxml: string, block: string, author: string): string {
    // Split by newlines to create multiple paragraphs
    const lines = block.split('\n').filter(l => l.trim().length > 0);

    const paragraphs = lines.map(line => buildInsertionOoxml(refOoxml, line, author));

    return paragraphs.join('');
}
