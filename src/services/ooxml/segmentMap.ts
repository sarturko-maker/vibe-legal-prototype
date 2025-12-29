/**
 * Segment Map Builder for Vibe Legal
 * Builds structural maps of paragraphs for surgical editing
 * Source: vibe-legal-beta-0.2.yaml L2045-2112
 */

import { WORD_NS } from './namespaces';
import { TextSegment, StructuralMap, AffectedSegment, InsertionPoint } from './types';

/**
 * Build a structural map from a paragraph element.
 * Extracts all text runs with their positions.
 */
export function buildStructuralMap(paragraphElement: Element): StructuralMap {
    const segments: TextSegment[] = [];
    let currentOffset = 0;

    // Get all w:r (run) elements
    const runs = paragraphElement.getElementsByTagNameNS(WORD_NS, 'r');

    for (let i = 0; i < runs.length; i++) {
        const run = runs[i] as Element;

        // Get text content from w:t elements
        const textElements = run.getElementsByTagNameNS(WORD_NS, 't');
        for (let j = 0; j < textElements.length; j++) {
            const text = textElements[j].textContent || '';
            if (text.length > 0) {
                const rPr = run.getElementsByTagNameNS(WORD_NS, 'rPr')[0] as Element | undefined;

                segments.push({
                    text,
                    startOffset: currentOffset,
                    endOffset: currentOffset + text.length,
                    runElement: run,
                    rPr
                });

                currentOffset += text.length;
            }
        }
    }

    const fullText = segments.map(s => s.text).join('');

    return {
        paragraphElement,
        segments,
        fullText
    };
}

/**
 * Find all segments that overlap with a character range.
 * Used for deletions that may span multiple runs.
 */
export function findSegmentsForRange(
    segments: TextSegment[],
    start: number,
    end: number
): AffectedSegment[] {
    const affected: AffectedSegment[] = [];

    for (const segment of segments) {
        // Check if segment overlaps with range [start, end)
        if (segment.endOffset > start && segment.startOffset < end) {
            affected.push({
                segment,
                deleteStart: Math.max(segment.startOffset, start) - segment.startOffset,
                deleteEnd: Math.min(segment.endOffset, end) - segment.startOffset
            });
        }
    }

    return affected;
}

/**
 * Find the insertion point for a given position.
 */
export function findInsertionPoint(
    segments: TextSegment[],
    position: number
): InsertionPoint {
    for (let i = 0; i < segments.length; i++) {
        const segment = segments[i];

        // Position is within this segment
        if (position >= segment.startOffset && position <= segment.endOffset) {
            return {
                segmentIndex: i,
                insertOffset: position - segment.startOffset
            };
        }
    }

    // Position is at end
    return {
        segmentIndex: segments.length - 1,
        insertOffset: segments[segments.length - 1]?.text.length || 0
    };
}

/**
 * Get the full text from a structural map.
 */
export function getFullText(map: StructuralMap): string {
    return map.fullText;
}
