/**
 * OOXML Processor for Vibe Legal
 * Main entry point for OOXML manipulation with track changes
 * Source: vibe-legal-beta-0.2.yaml (various sections)
 */

import { WORD_NS } from './namespaces';
import { OoxmlModificationResult, TrackChangeInfo } from './types';
import { StructuredDiff } from '../diff/surgicalDiff';
import { buildStructuralMap, findSegmentsForRange, findInsertionPoint } from './segmentMap';
import { cloneRunWithText, cloneRunWithDelText } from './runBuilder';
import { createInsertionWrapper, createDeletionWrapper, generateSessionRsid } from './trackChanges';

// DOMParser for Node/browser compatibility
declare var DOMParser: any;
declare var XMLSerializer: any;

/**
 * Apply surgical diffs to OOXML with track changes.
 * This is the main entry point for OOXML modification.
 */
export function applyDiffsToOoxml(
    oxml: string,
    diffs: StructuredDiff[],
    author: string,
    date?: string
): OoxmlModificationResult {
    if (diffs.length === 0) {
        return { oxml, hasChanges: false, changeCount: 0 };
    }

    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(oxml, 'application/xml');

        // Check for parse errors
        const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
        if (parseError) {
            return {
                oxml,
                hasChanges: false,
                changeCount: 0,
                errors: ['XML parse error: ' + parseError.textContent]
            };
        }

        // Find the paragraph element
        const paragraphs = xmlDoc.getElementsByTagNameNS(WORD_NS, 'p');
        if (paragraphs.length === 0) {
            return { oxml, hasChanges: false, changeCount: 0 };
        }

        const paragraph = paragraphs[0] as Element;
        const map = buildStructuralMap(paragraph);
        const sessionRsid = generateSessionRsid();
        const changeDate = date || new Date().toISOString();

        let changeCount = 0;

        // Apply diffs in reverse order to preserve positions
        const reversedDiffs = [...diffs].sort((a, b) => b.start - a.start);

        for (const diff of reversedDiffs) {
            const applied = applyDiff(xmlDoc, map, diff, author, changeDate, sessionRsid);
            if (applied) changeCount++;
        }

        // Serialize result
        const serializer = new XMLSerializer();
        const resultOxml = serializer.serializeToString(xmlDoc);

        return {
            oxml: resultOxml,
            hasChanges: changeCount > 0,
            changeCount
        };
    } catch (error: any) {
        return {
            oxml,
            hasChanges: false,
            changeCount: 0,
            errors: [error.message]
        };
    }
}

/**
 * Apply a single diff to the document.
 */
function applyDiff(
    xmlDoc: Document,
    map: ReturnType<typeof buildStructuralMap>,
    diff: StructuredDiff,
    author: string,
    date: string,
    sessionRsid: string
): boolean {
    switch (diff.type) {
        case 'delete':
            return applyDeletion(xmlDoc, map, diff.start, diff.end!, author, date, sessionRsid);
        case 'insert':
            return applyInsertion(xmlDoc, map, diff.start, diff.text!, author, date, sessionRsid);
        case 'replace':
            // Replace = delete + insert
            const deleted = applyDeletion(xmlDoc, map, diff.start, diff.end!, author, date, sessionRsid);
            const inserted = applyInsertion(xmlDoc, map, diff.start, diff.text!, author, date, sessionRsid);
            return deleted || inserted;
        default:
            return false;
    }
}

/**
 * Apply a deletion at the specified range.
 */
function applyDeletion(
    xmlDoc: Document,
    map: ReturnType<typeof buildStructuralMap>,
    start: number,
    end: number,
    author: string,
    date: string,
    sessionRsid: string
): boolean {
    const affected = findSegmentsForRange(map.segments, start, end);
    if (affected.length === 0) return false;

    for (const affectedSeg of affected) {
        const segment = affectedSeg.segment;
        const originalText = segment.text;
        const runElement = segment.runElement;

        const beforeText = originalText.substring(0, affectedSeg.deleteStart);
        const deletedText = originalText.substring(affectedSeg.deleteStart, affectedSeg.deleteEnd);
        const afterText = originalText.substring(affectedSeg.deleteEnd);

        const parent = runElement.parentNode;
        if (!parent) continue;

        // Build replacement elements
        const newElements: Element[] = [];

        if (beforeText.length > 0) {
            newElements.push(cloneRunWithText(xmlDoc, runElement, beforeText));
        }

        // Deleted content wrapped in <w:del>
        const delWrapper = createDeletionWrapper(xmlDoc, author, date);
        const delRun = cloneRunWithDelText(xmlDoc, runElement, deletedText, sessionRsid);
        delWrapper.appendChild(delRun);
        newElements.push(delWrapper);

        if (afterText.length > 0) {
            newElements.push(cloneRunWithText(xmlDoc, runElement, afterText));
        }

        // Replace original run with new elements
        for (const el of newElements) {
            parent.insertBefore(el, runElement);
        }
        parent.removeChild(runElement);
    }

    return true;
}

/**
 * Apply an insertion at the specified position.
 */
function applyInsertion(
    xmlDoc: Document,
    map: ReturnType<typeof buildStructuralMap>,
    position: number,
    text: string,
    author: string,
    date: string,
    sessionRsid: string
): boolean {
    const insertPoint = findInsertionPoint(map.segments, position);
    if (insertPoint.segmentIndex < 0) return false;

    const segment = map.segments[insertPoint.segmentIndex];
    const runElement = segment.runElement;
    const parent = runElement.parentNode;
    if (!parent) return false;

    const originalText = segment.text;
    const beforeText = originalText.substring(0, insertPoint.insertOffset);
    const afterText = originalText.substring(insertPoint.insertOffset);

    const newElements: Element[] = [];

    if (beforeText.length > 0) {
        newElements.push(cloneRunWithText(xmlDoc, runElement, beforeText));
    }

    // Inserted content wrapped in <w:ins>
    const insWrapper = createInsertionWrapper(xmlDoc, author, date);
    const insRun = cloneRunWithText(xmlDoc, runElement, text, sessionRsid);
    insWrapper.appendChild(insRun);
    newElements.push(insWrapper);

    if (afterText.length > 0) {
        newElements.push(cloneRunWithText(xmlDoc, runElement, afterText));
    }

    // Replace original run with new elements
    for (const el of newElements) {
        parent.insertBefore(el, runElement);
    }
    parent.removeChild(runElement);

    return true;
}

/**
 * Get text content from OOXML paragraph.
 */
export function extractTextFromOoxml(oxml: string): string {
    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(oxml, 'application/xml');

        const paragraphs = xmlDoc.getElementsByTagNameNS(WORD_NS, 'p');
        if (paragraphs.length === 0) return '';

        const map = buildStructuralMap(paragraphs[0] as Element);
        return map.fullText;
    } catch {
        return '';
    }
}
