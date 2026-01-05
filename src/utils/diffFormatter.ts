/**
 * Diff Formatter
 * Converts technical diff output to human-readable preview for AI validation.
 */

export interface TextChange {
    type: 'equal' | 'delete' | 'insert';
    text: string;
}

/**
 * Format diff changes into human-readable preview for AI validation.
 * Groups delete+insert pairs as "Changed" operations.
 */
export function formatDiffForAI(changes: TextChange[]): string {
    const lines: string[] = [];

    for (let i = 0; i < changes.length; i++) {
        const change = changes[i];

        // Skip EQUAL changes (unchanged text)
        if (change.type === 'equal') continue;

        if (change.type === 'delete') {
            // Check if this is part of a replacement (delete followed by insert)
            if (i + 1 < changes.length && changes[i + 1].type === 'insert') {
                const nextChange = changes[i + 1];
                lines.push(`- Changed: "${change.text}" → "${nextChange.text}"`);
                i++; // Skip the insert since we handled it
            } else {
                lines.push(`- Deleted: "${change.text}"`);
            }
        } else if (change.type === 'insert') {
            lines.push(`- Inserted: "${change.text}"`);
        }
    }

    return lines.length > 0
        ? lines.join('\n')
        : '(No changes detected)';
}

/**
 * Convert diff-match-patch output format to our TextChange format.
 * diff-match-patch uses: [0=equal, -1=delete, 1=insert]
 */
export function convertDmpToTextChanges(dmpDiffs: Array<[number, string]>): TextChange[] {
    return dmpDiffs.map(([op, text]) => ({
        type: op === 0 ? 'equal' : op === -1 ? 'delete' : 'insert',
        text
    }));
}
