/**
 * Contract Map Builder for Vibe Legal
 * Builds structural map of document for AI context
 * Source: vibe-legal-beta-0.2.yaml L8215-8870
 */

import { ContractMap, ContractMapState, ClauseInfo, ParagraphInfo } from '../../types';

/**
 * Build contract map context section for router prompts.
 * This creates a string representation of the document structure for AI.
 * CRITICAL: Must include ALL paragraphs, not just detected headings.
 */
export function buildContractMapContext(contractMap: ContractMap | null): string {
    if (!contractMap) return '';

    // Build CLAUSE INDEX first - maps clause numbers to paragraph IDs
    const clauseIndexLines: string[] = [];
    for (const c of contractMap.clauses) {
        if (c.number) {
            clauseIndexLines.push(`[P${c.paragraphId}] ${c.number} ${c.title}`);
        }
    }

    const clauseIndexSection = clauseIndexLines.length > 0
        ? '=== CLAUSE INDEX (use [P##] as target_id) ===\n' + clauseIndexLines.join('\n')
        : '';

    // Build COMPLETE paragraph map (legacy format)
    // This is the addressing system AI uses to target operations
    const paragraphMapLines: string[] = [];
    for (const p of contractMap.paragraphs) {
        const text = p.text.substring(0, 120).replace(/\n/g, ' ').trim();
        if (text.length === 0) continue;

        // Build metadata string
        let meta = `{Style: ${p.style || 'Normal'}}`;
        if (p.isListItem && p.listLevel !== undefined) {
            meta += ` {List: Lvl ${p.listLevel}}`;
        } else if (p.isListItem) {
            meta += ` {List: true}`;
        }
        // Use [P##] format prominently
        paragraphMapLines.push(`[P${p.id}] ${meta} ${text}`);
    }

    const paragraphMapSection = paragraphMapLines.length > 0
        ? 'PARAGRAPH MAP:\n' + paragraphMapLines.join('\n')
        : '';

    let result = '\n=== CONTRACT INTELLIGENCE ===\n';
    result += `Total Paragraphs: ${contractMap.totalParagraphs}\n\n`;
    result += clauseIndexSection;
    result += '\n\n' + paragraphMapSection;
    result += '\n\n=== TARGETING REMINDER ===\n';
    result += '- The number after [P] is the PARAGRAPH ID (target_id)\n';
    result += '- Example: [P33] 5.4 Warranty Remedy → use target_id: 33\n';
    result += '- DO NOT use the clause number (5.4) as target_id!\n';

    return result;
}

/**
 * Build a simple contract map from paragraphs using heuristics.
 * Detects clause titles by pattern matching.
 */
export function buildSimpleContractMap(paragraphs: ParagraphInfo[]): ContractMap {
    const clauses: ClauseInfo[] = [];

    // Clause number pattern: 1., 1.1, 1.1.1, etc.
    const numberedClausePattern = /^(\d+(?:\.\d+)*)[.\s]+(.+?)(?:\.|$)/;

    // Common legal clause titles (for detection without numbers)
    const legalHeadings = [
        'GOVERNING LAW', 'JURISDICTION', 'DEFINITIONS', 'CONFIDENTIALITY',
        'TERMINATION', 'INDEMNITY', 'LIABILITY', 'FORCE MAJEURE', 'NOTICES',
        'ASSIGNMENT', 'ENTIRE AGREEMENT', 'AMENDMENTS', 'WAIVER', 'SEVERABILITY',
        'GENERAL PROVISIONS', 'MISCELLANEOUS', 'DISPUTE RESOLUTION'
    ];

    for (const para of paragraphs) {
        const text = para.text.trim();
        if (!text) continue;

        // Check for numbered clause pattern first
        const numberedMatch = text.match(numberedClausePattern);
        if (numberedMatch) {
            clauses.push({
                number: numberedMatch[1],
                title: numberedMatch[2].trim(),
                paragraphId: para.id,
                startOffset: 0,
                endOffset: para.text.length,
                font: para.font,
                fontSize: para.fontSize
            });
            continue;
        }

        // Check for heading-style clauses (ALL CAPS or known legal headings)
        const upperText = text.toUpperCase();
        const isAllCaps = text === text.toUpperCase() && text.length > 3 && text.length < 80;
        const isLegalHeading = legalHeadings.some(h => upperText.includes(h));

        if ((isAllCaps || isLegalHeading) && text.length < 100) {
            clauses.push({
                number: '',
                title: text,
                paragraphId: para.id,
                startOffset: 0,
                endOffset: para.text.length,
                font: para.font,
                fontSize: para.fontSize
            });
        }
    }

    return {
        paragraphs,
        clauses,
        totalParagraphs: paragraphs.length
    };
}

/**
 * Find clause by number in contract map.
 */
export function findClauseByNumber(contractMap: ContractMap | null, number: string): ClauseInfo | null {
    if (!contractMap) return null;
    return contractMap.clauses.find(c => c.number === number) || null;
}

/**
 * Find clause by title (fuzzy match) in contract map.
 */
export function findClauseByTitle(contractMap: ContractMap | null, title: string): ClauseInfo | null {
    if (!contractMap) return null;

    const normalizedTitle = title.toLowerCase().trim();

    // Exact match first
    const exact = contractMap.clauses.find(c =>
        c.title.toLowerCase() === normalizedTitle
    );
    if (exact) return exact;

    // Partial match
    return contractMap.clauses.find(c =>
        c.title.toLowerCase().includes(normalizedTitle) ||
        normalizedTitle.includes(c.title.toLowerCase())
    ) || null;
}

/**
 * Get paragraph by ID.
 */
export function getParagraphById(contractMap: ContractMap | null, id: number): ParagraphInfo | null {
    if (!contractMap) return null;
    return contractMap.paragraphs.find(p => p.id === id) || null;
}

/**
 * Create initial contract map state.
 */
export function createInitialContractMapState(): ContractMapState {
    return {
        map: null,
        documentHash: null,
        totalParagraphCount: null,
        isAnalyzing: false,
        generatedAt: null
    };
}
