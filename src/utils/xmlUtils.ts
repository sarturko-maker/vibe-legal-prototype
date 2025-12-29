/**
 * XML utilities for Vibe Legal
 * OOXML namespace constants and XML helpers
 */

// Word Processing ML namespace
export const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Other OOXML namespaces (for future use)
export const RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/**
 * Get elements by tag name in Word namespace
 */
export function getWordElements(parent: Element, tagName: string): Element[] {
    const nodeList = parent.getElementsByTagNameNS(WORD_NS, tagName);
    return Array.from(nodeList) as Element[];
}

/**
 * Get first element by tag name in Word namespace
 */
export function getWordElement(parent: Element, tagName: string): Element | null {
    const elements = parent.getElementsByTagNameNS(WORD_NS, tagName);
    return elements.length > 0 ? elements[0] as Element : null;
}

/**
 * Create element in Word namespace
 */
export function createWordElement(doc: Document, tagName: string): Element {
    return doc.createElementNS(WORD_NS, `w:${tagName}`);
}

/**
 * Set attribute with w: prefix
 */
export function setWordAttr(element: Element, name: string, value: string): void {
    element.setAttribute(`w:${name}`, value);
}

/**
 * Get attribute with w: prefix
 */
export function getWordAttr(element: Element, name: string): string | null {
    return element.getAttribute(`w:${name}`);
}

/**
 * Build element path for error reporting (e.g., "w:p[0]/w:ins[1]/w:r[0]")
 */
export function buildElementPath(el: Element): string {
    const parts: string[] = [];
    let current: Element | null = el;

    while (current && current.parentElement) {
        const parent = current.parentElement;
        const siblings = Array.from(parent.children).filter(c => c.nodeName === current!.nodeName);
        const index = siblings.indexOf(current);
        parts.unshift(`${current.nodeName}[${index}]`);
        current = parent;
    }

    return parts.join('/');
}
