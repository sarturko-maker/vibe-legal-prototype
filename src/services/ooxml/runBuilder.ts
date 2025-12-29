/**
 * Run Builder for Vibe Legal
 * Creates and clones <w:r> (run) elements for OOXML manipulation
 * Source: vibe-legal-beta-0.2.yaml L2114-2200
 */

import { WORD_NS } from './namespaces';

/**
 * Extract rsidR attribute from a run element.
 */
export function extractRsidR(run: Element): string | null {
    return run.getAttribute('w:rsidR') || null;
}

/**
 * Clone a run element with new text content.
 * Preserves formatting (rPr) from the original run.
 */
export function cloneRunWithText(
    xmlDoc: Document,
    originalRun: Element,
    newText: string,
    rsid?: string
): Element {
    const newRun = xmlDoc.createElementNS(WORD_NS, 'w:r');

    // Copy RSID attributes
    const originalRsidR = extractRsidR(originalRun);
    if (originalRsidR) {
        newRun.setAttribute('w:rsidR', originalRsidR);
    } else if (rsid) {
        newRun.setAttribute('w:rsidR', rsid);
    }

    // Copy run properties (rPr)
    const rPr = originalRun.getElementsByTagNameNS(WORD_NS, 'rPr')[0] as Element;
    if (rPr) {
        newRun.appendChild(rPr.cloneNode(true));
    }

    // Create text element
    const textEl = xmlDoc.createElementNS(WORD_NS, 'w:t');
    textEl.textContent = newText;

    // Preserve whitespace if needed
    if (newText.startsWith(' ') || newText.endsWith(' ') || newText.includes('  ')) {
        textEl.setAttribute('xml:space', 'preserve');
    }

    newRun.appendChild(textEl);

    return newRun;
}

/**
 * Clone a run element with delText content (for deletions).
 * Used inside <w:del> track change markers.
 */
export function cloneRunWithDelText(
    xmlDoc: Document,
    originalRun: Element,
    deletedText: string,
    sessionRsid: string
): Element {
    const newRun = xmlDoc.createElementNS(WORD_NS, 'w:r');

    // Keep original rsidR, add rsidDel
    const originalRsidR = extractRsidR(originalRun);
    if (originalRsidR) {
        newRun.setAttribute('w:rsidR', originalRsidR);
    }
    newRun.setAttribute('w:rsidDel', sessionRsid);

    // Copy run properties (rPr)
    const rPr = originalRun.getElementsByTagNameNS(WORD_NS, 'rPr')[0] as Element;
    if (rPr) {
        newRun.appendChild(rPr.cloneNode(true));
    }

    // Create delText element (not w:t)
    const delText = xmlDoc.createElementNS(WORD_NS, 'w:delText');
    delText.textContent = deletedText;

    if (deletedText.startsWith(' ') || deletedText.endsWith(' ') || deletedText.includes('  ')) {
        delText.setAttribute('xml:space', 'preserve');
    }

    newRun.appendChild(delText);

    return newRun;
}

/**
 * Create a simple text run with optional formatting.
 */
export function createTextRun(
    xmlDoc: Document,
    text: string,
    formatting?: { bold?: boolean; italic?: boolean; underline?: boolean },
    rsid?: string
): Element {
    const run = xmlDoc.createElementNS(WORD_NS, 'w:r');

    if (rsid) {
        run.setAttribute('w:rsidR', rsid);
    }

    // Add run properties if formatting specified
    if (formatting && (formatting.bold || formatting.italic || formatting.underline)) {
        const rPr = xmlDoc.createElementNS(WORD_NS, 'w:rPr');

        if (formatting.bold) {
            rPr.appendChild(xmlDoc.createElementNS(WORD_NS, 'w:b'));
        }
        if (formatting.italic) {
            rPr.appendChild(xmlDoc.createElementNS(WORD_NS, 'w:i'));
        }
        if (formatting.underline) {
            const u = xmlDoc.createElementNS(WORD_NS, 'w:u');
            u.setAttribute('w:val', 'single');
            rPr.appendChild(u);
        }

        run.appendChild(rPr);
    }

    // Create text element
    const textEl = xmlDoc.createElementNS(WORD_NS, 'w:t');
    textEl.textContent = text;

    if (text.startsWith(' ') || text.endsWith(' ') || text.includes('  ')) {
        textEl.setAttribute('xml:space', 'preserve');
    }

    run.appendChild(textEl);

    return run;
}
