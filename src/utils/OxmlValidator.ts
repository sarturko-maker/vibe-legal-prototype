
import { DOMParser } from '@xmldom/xmldom';
import * as crypto from 'crypto';

export interface ValidationResult {
    isValid: boolean;
    error?: string;
}

/**
 * Validates the OOXML string against a "Schema-Light" set of rules.
 */
export function validateOxml(oxml: string): ValidationResult {
    // 1. Basic XML Parsing (Checks for unclosed tags, etc.)
    const parser = new DOMParser({
        errorHandler: {
            warning: () => { }, // Ignore warnings
            error: (msg) => { throw new Error(msg); },
            fatalError: (msg) => { throw new Error(msg); }
        }
    });

    let doc: Document;
    try {
        doc = parser.parseFromString(oxml, 'text/xml');
    } catch (e: any) {
        return { isValid: false, error: `XML Parsing Error: ${e.message}` };
    }

    // xmldom parser returns a document with <parsererror> if parsing fails (sometimes)
    const parserError = doc.getElementsByTagName('parsererror')[0];
    if (parserError) {
        return { isValid: false, error: `XML Parsing Error: ${parserError.textContent}` };
    }

    // 2. No text outside <w:t> or <w:delText>
    // We traverse the tree and check text nodes.
    const invalidTextNodes = findInvalidTextNodes(doc);
    if (invalidTextNodes.length > 0) {
        return { isValid: false, error: `Found text outside <w:t> or <w:delText>: "${invalidTextNodes[0].substring(0, 20)}..."` };
    }

    // 3. <w:ins> and <w:del> must have attributes
    const revisions = [...Array.from(doc.getElementsByTagName('w:ins')), ...Array.from(doc.getElementsByTagName('w:del'))];
    for (const rev of revisions) {
        if (!rev.getAttribute('w:id') || !rev.getAttribute('w:author') || !rev.getAttribute('w:date')) {
            return { isValid: false, error: `Revision tag <${rev.nodeName}> missing required attributes (id, author, date).` };
        }
    }

    // 4. No nested paragraphs (<w:p> inside <w:p>)
    const paragraphs = Array.from(doc.getElementsByTagName('w:p'));
    for (const p of paragraphs) {
        if (p.getElementsByTagName('w:p').length > 0) {
            return { isValid: false, error: `Found nested <w:p> inside another <w:p>.` };
        }
    }

    return { isValid: true };
}

function findInvalidTextNodes(node: Node): string[] {
    const invalidTexts: string[] = [];

    if (node.nodeType === 3) { // Text Node
        const text = node.textContent?.trim();
        if (text && text.length > 0) {
            const parentName = node.parentNode?.nodeName;
            // console.log(`[Validator] Found text "${text}" in parent "${parentName}"`);
            // Allowed parents for text
            const allowedParents = ['w:t', 'w:delText', 'w:instrText', 'm:t'];
            if (!allowedParents.includes(parentName || '')) {
                console.log(`[Validator] Invalid text found: "${text}" in "${parentName}"`);
                invalidTexts.push(text);
            }
        }
    } else {
        for (let i = 0; i < node.childNodes.length; i++) {
            invalidTexts.push(...findInvalidTextNodes(node.childNodes[i]));
        }
    }

    return invalidTexts;
}

/**
 * Calculates a structural checksum by hashing the tag structure (ignoring text content).
 */
export function calculateStructuralChecksum(oxml: string): string {
    // Remove all text content between tags
    // Regex strategy: Replace >[^<]+< with >< (simplified)
    // Or better: Use DOM parser and walk structure.
    // Regex is faster for "ignoring text content" if we just want structure.
    // Remove content of w:t, w:delText, etc.

    // Simple regex approach: Remove everything between > and <
    const structureOnly = oxml.replace(/>([^<]+)</g, '><');

    return crypto.createHash('md5').update(structureOnly).digest('hex');
}
