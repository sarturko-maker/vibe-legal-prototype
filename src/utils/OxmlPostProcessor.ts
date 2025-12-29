
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

/**
 * Corrects double-encoded XML entities that LLMs sometimes produce.
 * e.g., &amp;amp; -> &amp;
 */
export function correctDoubleEncodedEntities(oxml: string): string {
    return oxml
        .replace(/&amp;amp;/g, '&amp;')
        .replace(/&amp;lt;/g, '&lt;')
        .replace(/&amp;gt;/g, '&gt;')
        .replace(/&amp;quot;/g, '&quot;')
        .replace(/&amp;apos;/g, '&apos;');
}

/**
 * Detects new paragraphs in a list that lack numbering and clones it from siblings.
 * This fixes the "Bullet Cloning" issue where LLMs forget <w:numPr>.
 */
export function cloneBulletNumbering(oxml: string): string {
    const doc = new DOMParser().parseFromString(oxml, 'text/xml');
    const paragraphs = Array.from(doc.getElementsByTagName('w:p'));

    for (let i = 1; i < paragraphs.length; i++) {
        const currentP = paragraphs[i];
        const prevP = paragraphs[i - 1];

        // Check if current paragraph lacks numPr
        const currentNumPr = currentP.getElementsByTagName('w:numPr')[0];

        if (!currentNumPr) {
            // Check if previous paragraph HAS numPr
            const prevNumPr = prevP.getElementsByTagName('w:numPr')[0];

            if (prevNumPr) {
                // Heuristic: Check indentation. If similar or absent in both, assume continuation.
                // For simplicity in this sprint, we assume if it follows a list item and has no numPr, it's a split.
                // We clone the numPr.

                const pPr = currentP.getElementsByTagName('w:pPr')[0] || doc.createElement('w:pPr');
                if (!currentP.getElementsByTagName('w:pPr')[0]) {
                    currentP.insertBefore(pPr, currentP.firstChild);
                }

                const newNumPr = prevNumPr.cloneNode(true);
                pPr.appendChild(newNumPr);
            }
        }
    }
    return new XMLSerializer().serializeToString(doc);
}

/**
 * Remaps duplicate w:id attributes to ensure uniqueness.
 * Strategy: Collect all IDs. If duplicates found, reassign them to high-entropy integers.
 */
/**
 * Remaps all w:id attributes to a safe, high-entropy range to prevent collisions.
 * Preserves relationships between elements (e.g., bookmarkStart/End) by mapping identical old IDs to the same new ID.
 */
export function remapIds(oxml: string): string {
    const doc = new DOMParser().parseFromString(oxml, 'text/xml');
    const elementsWithId = Array.from(doc.documentElement.getElementsByTagName('*')).filter(el => el.hasAttribute('w:id'));

    const idMap = new Map<string, string>();
    const SAFE_START = 90000; // Start high to avoid standard Word IDs
    let counter = 0;

    for (const el of elementsWithId) {
        const oldId = el.getAttribute('w:id')!;

        if (!idMap.has(oldId)) {
            // Generate a new safe ID
            // Using a deterministic offset + counter for stability, or random for entropy.
            // Let's use a high base + counter to ensure uniqueness within the session.
            const newId = (SAFE_START + Math.floor(Math.random() * 10000) + counter++).toString();
            idMap.set(oldId, newId);
        }

        el.setAttribute('w:id', idMap.get(oldId)!);
    }

    return new XMLSerializer().serializeToString(doc);
}

/**
 * Removes "ghost" comment anchors (empty ranges) left behind by deletions.
 */
export function removeGhostComments(oxml: string): string {
    const doc = new DOMParser().parseFromString(oxml, 'text/xml');

    // Find commentRangeStart elements
    const starts = Array.from(doc.getElementsByTagName('w:commentRangeStart'));

    for (const start of starts) {
        const id = start.getAttribute('w:id');
        // Check if immediately followed by End with same ID
        let next = start.nextSibling;
        while (next && next.nodeType !== 1) { // Skip text/whitespace
            next = next.nextSibling;
        }

        if (next && next.nodeName === 'w:commentRangeEnd' && (next as Element).getAttribute('w:id') === id) {
            // It's an empty range. Remove both.
            start.parentNode?.removeChild(start);
            next.parentNode?.removeChild(next);
        }
    }

    return new XMLSerializer().serializeToString(doc);
}
