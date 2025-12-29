/**
 * Document Types for Vibe Legal
 * Extracted from vibe-legal-beta-0.2.yaml
 */

// Contract Map - document structure representation
export interface ContractMap {
    paragraphs: ParagraphInfo[];
    clauses: ClauseInfo[];
    totalParagraphs: number;
}

export interface ParagraphInfo {
    id: number;
    text: string;
    style?: string;
    isListItem: boolean;
    listString?: string;
    level?: number;
    // Rich data for AI context
    listLevel?: number;       // Word's ilvl (0, 1, 2)
    alignment?: string;       // "Left", "Center", "Justified"
    isAllCaps?: boolean;      // Detect ALL CAPS patterns
    isHeadingPattern?: boolean; // Short text or numbered heading pattern
    // Font info for contract map enrichment
    font?: string;            // Font name: "Times New Roman", "Arial"
    fontSize?: number;        // Font size in points: 12, 14, etc.
}

export interface ClauseInfo {
    number: string;
    title: string;
    paragraphId: number;
    startOffset: number;
    endOffset: number;
    // Font info for contract map enrichment
    font?: string;            // Font name: "Times New Roman", "Arial"
    fontSize?: number;        // Font size in points: 12, 14, etc.
}

// Contract Map State (for React)
export interface ContractMapState {
    map: ContractMap | null;
    documentHash: string | null;
    totalParagraphCount: number | null;
    isAnalyzing: boolean;
    generatedAt: number | null;
}

// Focused clause for Pro Mode
export interface FocusedClause {
    number: string | null;
    title: string | null;
    text: string;
    paragraphId: number | null;
}

// Style information
export interface StyleInfo {
    name: string;
    font: {
        name: string;
        size: number;
        bold: boolean;
        italic: boolean;
        underline: boolean;
        color: string | null;
    };
    paragraphFormat: {
        alignment: string;
        lineSpacing: number;
        spaceBefore: number;
        spaceAfter: number;
    };
    usageCount: number;
    exampleParagraphId: number;
    usedForHeadings: boolean;
    usedForBody: boolean;
    hasDirectFormattingOverrides?: boolean;
}

// OOXML segment for surgical editing
export interface TextSegment {
    text: string;
    startOffset: number;
    endOffset: number;
    runElement: Element;
    rPr?: Element;
}

export interface StructuralMap {
    paragraphElement: Element;
    segments: TextSegment[];
    fullText: string;
}

// Affected segment during diff
export interface AffectedSegment {
    segment: TextSegment;
    deleteStart: number;
    deleteEnd: number;
}

// Insertion point
export interface InsertionPoint {
    segmentIndex: number;
    insertOffset: number;
}
