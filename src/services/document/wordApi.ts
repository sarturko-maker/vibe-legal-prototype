/**
 * Word API Wrapper for Vibe Legal
 * Abstracts Office.js Word API calls
 * 
 * IMPORTANT: This module uses 1-indexed paragraph IDs throughout.
 * - loadParagraphs returns paragraphs with id = 1, 2, 3...
 * - getParagraphOoxml expects 1-indexed ID
 * - insertOoxmlAtParagraph expects 1-indexed ID
 * The conversion to 0-indexed array access happens internally.
 */

import { ParagraphInfo } from '../../types';

// Declare Word namespace (provided by Office.js)
declare var Word: any;

/**
 * Load all paragraphs from the document body.
 * NOTE: Paragraph IDs are 1-indexed (matching AI expectations from document map)
 */
export async function loadParagraphs(context: any): Promise<ParagraphInfo[]> {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    const result: ParagraphInfo[] = [];

    for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load('text,style,isListItem');
        await context.sync();

        // IMPORTANT: Use 1-indexed IDs to match legacy behavior and AI expectations
        result.push({
            id: i + 1,  // 1-indexed!
            text: para.text || '',
            style: para.style || 'Normal',
            isListItem: para.isListItem || false
        });
    }

    return result;
}

/**
 * Get the current selection text.
 */
export async function getSelectionText(context: any): Promise<string> {
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    return selection.text || '';
}

/**
 * Get OOXML for a specific paragraph.
 * @param paragraphId 1-indexed paragraph ID
 */
export async function getParagraphOoxml(context: any, paragraphId: number): Promise<string> {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    // Convert 1-indexed ID to 0-indexed array index
    const arrayIndex = paragraphId - 1;

    console.log(`[getParagraphOoxml] ID=${paragraphId} → arrayIndex=${arrayIndex}, total=${paragraphs.items.length}`);

    if (arrayIndex < 0 || arrayIndex >= paragraphs.items.length) {
        throw new Error(`Paragraph ${paragraphId} not found (array index ${arrayIndex}, total ${paragraphs.items.length})`);
    }

    const para = paragraphs.items[arrayIndex];
    const ooxml = para.getOoxml();
    await context.sync();

    console.log(`[getParagraphOoxml] Got OOXML for P${paragraphId}, length=${ooxml.value?.length}`);

    return ooxml.value;
}

/**
 * Insert OOXML at a specific paragraph.
 * @param paragraphId 1-indexed paragraph ID
 */
export async function insertOoxmlAtParagraph(
    context: any,
    paragraphId: number,
    ooxml: string,
    insertLocation: 'Replace' | 'Before' | 'After' = 'Replace'
): Promise<void> {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    // Convert 1-indexed ID to 0-indexed array index
    const arrayIndex = paragraphId - 1;

    console.log(`[insertOoxmlAtParagraph] ID=${paragraphId} → arrayIndex=${arrayIndex}, mode=${insertLocation}, total=${paragraphs.items.length}`);

    if (arrayIndex < 0 || arrayIndex >= paragraphs.items.length) {
        throw new Error(`Paragraph ${paragraphId} not found (array index ${arrayIndex}, total ${paragraphs.items.length})`);
    }

    const para = paragraphs.items[arrayIndex];
    para.insertOoxml(ooxml, insertLocation);
    await context.sync();

    console.log(`[insertOoxmlAtParagraph] Complete for P${paragraphId}`);
}

/**
 * Get paragraph count.
 */
export async function getParagraphCount(context: any): Promise<number> {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();
    return paragraphs.items.length;
}

/**
 * Run a Word operation with context.
 */
export async function runWord<T>(callback: (context: any) => Promise<T>): Promise<T> {
    return Word.run(callback);
}
