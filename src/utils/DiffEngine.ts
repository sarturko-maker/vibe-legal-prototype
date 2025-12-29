import { diff_match_patch } from 'diff-match-patch';

export function calculateRedline(original: string, modified: string): Array<{ op: number, text: string }> {
    const dmp = new diff_match_patch();
    const diffs = dmp.diff_main(original, modified);
    dmp.diff_cleanupSemantic(diffs);

    return diffs.map(diff => ({
        op: diff[0],
        text: diff[1]
    }));
}
