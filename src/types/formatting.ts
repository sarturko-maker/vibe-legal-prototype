/**
 * Formatting Types for Vibe Legal
 * Extracted from vibe-legal-beta-0.2.yaml
 */

// Formatting range with positions
export interface FormattingRange {
    start: number;
    end: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
}

// Text segment with formatting info
export interface Segment {
    text: string;
    formatting: {
        bold: boolean;
        italic: boolean;
        underline: boolean;
    };
}

// Markdown pattern types
export type FormattingType = 'bold' | 'italic' | 'underline';

// Formatting marker match
export interface FormattingMatch {
    type: FormattingType;
    start: number;
    end: number;
    content: string;
}
