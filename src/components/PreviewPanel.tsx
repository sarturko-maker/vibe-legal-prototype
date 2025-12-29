import React from 'react';
import { stripFormattingMarkers } from '../services/formatting/markdownParser';
import './PreviewPanel.css';

declare var diff_match_patch: any;

interface Operation {
    type: string;
    target_id?: number;
    amended_text?: string;
    original_text?: string;
    content?: string;
    description?: string;
}

/**
 * Generate a human-readable summary of a change operation
 */
function generateChangeSummary(op: Operation, original: string, modified: string): string {
    // Use AI-provided description if available
    if (op.description) {
        return op.description;
    }

    // Generate based on operation type
    const opType = op.type?.toUpperCase() || 'MODIFY';
    const preview = (modified || original || '').substring(0, 60).trim();
    const ellipsis = preview.length < (modified || original || '').length ? '...' : '';

    switch (opType) {
        case 'DELETE':
            return `Deleted: "${preview}${ellipsis}"`;
        case 'INSERT':
        case 'INSERT_AFTER':
            return `Inserted: "${preview}${ellipsis}"`;
        case 'AMEND':
        case 'MODIFY':
        default:
            return `Modified: "${preview}${ellipsis}"`;
    }
}

export { generateChangeSummary };

interface ReviewedSummary {
    index: number;
    summary: string;
    action: 'accepted' | 'rejected';
}

interface PreviewPanelProps {
    preview: {
        operations: Operation[];
        originalTexts?: { [key: number]: string };
    } | null;
    onAcceptOne: (index: number) => void;
    onRejectOne: (index: number) => void;
    onAcceptAll: () => void;
    onRejectAll: () => void;
    acceptedIndices: Set<number>;
    rejectedIndices: Set<number>;
    reviewedSummaries?: ReviewedSummary[];
    onClose?: () => void;
}

/**
 * Word-level diff for PREVIEW DISPLAY ONLY
 * This does NOT affect actual document changes (that's handled by applyRedline.ts)
 */
function computeWordDiff(original: string, modified: string): Array<[number, string]> {
    if (typeof diff_match_patch === 'undefined') {
        // Fallback if diff_match_patch not available
        if (original === modified) return [[0, original]];
        return [[-1, original], [1, modified]];
    }

    const dmp = new diff_match_patch();

    // Split into words (preserving whitespace)
    const splitIntoWords = (text: string): string[] => {
        return text.split(/(\s+)/).filter(s => s.length > 0);
    };

    const origWords = splitIntoWords(original);
    const modWords = splitIntoWords(modified);

    // Encode words as single characters for diff algorithm
    const wordArray: string[] = [''];
    const wordHash: { [key: string]: number } = {};

    const encodeWords = (words: string[]): string => {
        let encoded = '';
        for (const word of words) {
            if (word in wordHash) {
                encoded += String.fromCharCode(wordHash[word]);
            } else {
                const code = wordArray.length;
                wordArray.push(word);
                wordHash[word] = code;
                encoded += String.fromCharCode(code);
            }
        }
        return encoded;
    };

    const chars1 = encodeWords(origWords);
    const chars2 = encodeWords(modWords);

    const diffs = dmp.diff_main(chars1, chars2);
    dmp.diff_cleanupSemantic(diffs);

    // Decode back to words
    const result: Array<[number, string]> = [];
    for (const [op, chars] of diffs) {
        const words = chars.split('').map((c: string) => wordArray[c.charCodeAt(0)]).join('');
        if (words.length > 0) {
            result.push([op, words]);
        }
    }

    return result;
}

/**
 * Visual display component - PREVIEW ONLY
 * Does NOT affect actual document changes
 */
const WordDiff: React.FC<{ original: string; modified: string }> = ({ original, modified }) => {
    const cleanOrig = stripFormattingMarkers(original || '');
    const cleanMod = stripFormattingMarkers(modified || '');

    if (!cleanOrig && cleanMod) {
        return <span className="diff-added">{cleanMod}</span>;
    }
    if (cleanOrig && !cleanMod) {
        return <span className="diff-deleted">{cleanOrig}</span>;
    }
    if (!cleanOrig && !cleanMod) return null;

    try {
        const diffs = computeWordDiff(cleanOrig, cleanMod);

        return (
            <span className="word-diff">
                {diffs.map((diff, idx) => {
                    const [op, text] = diff;
                    if (op === 0) return <span key={idx}>{text}</span>;
                    if (op === -1) return <span key={idx} className="diff-deleted">{text}</span>;
                    if (op === 1) return <span key={idx} className="diff-added">{text}</span>;
                    return null;
                })}
            </span>
        );
    } catch (e) {
        console.error("Preview diff failed:", e);
        return <span>{cleanMod}</span>;
    }
};

export const PreviewPanel: React.FC<PreviewPanelProps> = ({
    preview,
    onAcceptOne,
    onRejectOne,
    onAcceptAll,
    onRejectAll,
    acceptedIndices,
    rejectedIndices,
    reviewedSummaries = [],
    onClose
}) => {
    if (!preview?.operations?.length) return null;

    // Count only valid (non-empty) pending operations
    const pendingCount = preview.operations.filter((op, i) => {
        const original = op.original_text || preview.originalTexts?.[op.target_id || 0] || '';
        const modified = op.amended_text || op.content || '';
        const hasContent = original.trim() || modified.trim();
        return hasContent && !acceptedIndices.has(i) && !rejectedIndices.has(i);
    }).length;

    // Separate reviewed summaries by action
    const appliedSummaries = reviewedSummaries.filter(s => s.action === 'accepted');
    const discardedSummaries = reviewedSummaries.filter(s => s.action === 'rejected');
    const hasReviewedItems = reviewedSummaries.length > 0;
    const isReviewComplete = pendingCount === 0 && hasReviewedItems;

    return (
        <div className="preview-panel">
            <div className="preview-header">
                <span>{isReviewComplete ? 'Review Complete' : `Proposed Changes (${pendingCount} pending)`}</span>
                <div className="bulk-actions">
                    {isReviewComplete ? (
                        <button className="done-btn" onClick={onClose}>Done</button>
                    ) : (
                        <>
                            <button className="accept-all-btn" onClick={onAcceptAll} disabled={pendingCount === 0}>Accept All</button>
                            <button className="reject-all-btn" onClick={onRejectAll} disabled={pendingCount === 0}>Reject All</button>
                        </>
                    )}
                </div>
            </div>

            {/* Running Summary Section - shows as user reviews */}
            {hasReviewedItems && (
                <div className="review-summary">
                    {appliedSummaries.length > 0 && (
                        <div className="summary-group summary-group--applied">
                            <div className="summary-group__header">Applied ({appliedSummaries.length})</div>
                            {appliedSummaries.map((s, i) => (
                                <div key={`applied-${i}`} className="summary-item summary-item--applied">
                                    {s.summary}
                                </div>
                            ))}
                        </div>
                    )}
                    {discardedSummaries.length > 0 && (
                        <div className="summary-group summary-group--discarded">
                            <div className="summary-group__header">Discarded ({discardedSummaries.length})</div>
                            {discardedSummaries.map((s, i) => (
                                <div key={`discarded-${i}`} className="summary-item summary-item--discarded">
                                    {s.summary}
                                </div>
                            ))}
                        </div>
                    )}
                </div>
            )}

            {/* Pending items list */}
            {pendingCount > 0 && (
                <div className="preview-list">
                    {preview.operations.map((op, idx) => {
                        const original = op.original_text || preview.originalTexts?.[op.target_id || 0] || '';
                        const modified = op.amended_text || op.content || '';

                        // Skip rendering empty or already-reviewed operations
                        if (!original.trim() && !modified.trim()) {
                            return null;
                        }

                        const isAccepted = acceptedIndices.has(idx);
                        const isRejected = rejectedIndices.has(idx);
                        const isPending = !isAccepted && !isRejected;

                        // Only show pending items in the list
                        if (!isPending) return null;

                        return (
                            <div key={idx} className="preview-item">
                                <div className="preview-diff">
                                    <WordDiff
                                        original={original}
                                        modified={modified}
                                    />
                                </div>
                                <div className="preview-item-actions">
                                    <button className="accept-btn-small" onClick={() => onAcceptOne(idx)} title="Accept this change">Yes</button>
                                    <button className="reject-btn-small" onClick={() => onRejectOne(idx)} title="Reject this change">No</button>
                                </div>
                            </div>
                        );
                    })}
                </div>
            )}
        </div>
    );
};

export default PreviewPanel;
