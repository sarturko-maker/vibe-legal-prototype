/**
 * Track Changes Generator for Vibe Legal
 * Creates <w:ins> and <w:del> wrapper elements
 * Source: vibe-legal-beta-0.2.yaml L2304-2600
 */

import { WORD_NS, DEFAULT_AUTHOR } from './namespaces';

/**
 * Generate a unique ID for track changes.
 */
let trackChangeIdCounter = 0;
export function generateTrackChangeId(): number {
    return trackChangeIdCounter++;
}

/**
 * Reset track change ID counter (for testing).
 */
export function resetTrackChangeIdCounter(): void {
    trackChangeIdCounter = 0;
}

/**
 * Create an insertion track change wrapper (<w:ins>).
 */
export function createInsertionWrapper(
    xmlDoc: Document,
    author: string = DEFAULT_AUTHOR,
    date?: string
): Element {
    const ins = xmlDoc.createElementNS(WORD_NS, 'w:ins');
    ins.setAttribute('w:id', String(generateTrackChangeId()));
    ins.setAttribute('w:author', author);
    ins.setAttribute('w:date', date || new Date().toISOString());
    return ins;
}

/**
 * Create a deletion track change wrapper (<w:del>).
 */
export function createDeletionWrapper(
    xmlDoc: Document,
    author: string = DEFAULT_AUTHOR,
    date?: string
): Element {
    const del = xmlDoc.createElementNS(WORD_NS, 'w:del');
    del.setAttribute('w:id', String(generateTrackChangeId()));
    del.setAttribute('w:author', author);
    del.setAttribute('w:date', date || new Date().toISOString());
    return del;
}

/**
 * Wrap a run element in an insertion track change.
 */
export function wrapInInsertion(
    xmlDoc: Document,
    run: Element,
    author: string = DEFAULT_AUTHOR,
    date?: string
): Element {
    const ins = createInsertionWrapper(xmlDoc, author, date);
    ins.appendChild(run);
    return ins;
}

/**
 * Wrap a run element in a deletion track change.
 */
export function wrapInDeletion(
    xmlDoc: Document,
    run: Element,
    author: string = DEFAULT_AUTHOR,
    date?: string
): Element {
    const del = createDeletionWrapper(xmlDoc, author, date);
    del.appendChild(run);
    return del;
}

/**
 * Generate session RSID (random hex string).
 */
export function generateSessionRsid(): string {
    return Math.random().toString(16).substring(2, 10).toUpperCase();
}
