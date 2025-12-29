/**
 * OOXML Types for Vibe Legal
 * Internal types for OOXML manipulation
 */

// Text segment with run reference
export interface TextSegment {
    text: string;
    startOffset: number;
    endOffset: number;
    runElement: Element;
    rPr?: Element;
}

// Structural map of a paragraph
export interface StructuralMap {
    paragraphElement: Element;
    segments: TextSegment[];
    fullText: string;
}

// Affected segment info during diff application
export interface AffectedSegment {
    segment: TextSegment;
    deleteStart: number;
    deleteEnd: number;
}

// Insertion point location
export interface InsertionPoint {
    segmentIndex: number;
    insertOffset: number;
}

// Track change info
export interface TrackChangeInfo {
    author: string;
    date: string;
    sessionRsid: string;
}

// Result of OOXML modification
export interface OoxmlModificationResult {
    oxml: string;
    hasChanges: boolean;
    changeCount: number;
    errors?: string[];
}
