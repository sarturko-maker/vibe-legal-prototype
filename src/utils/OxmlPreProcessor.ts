
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

/**
 * Merges adjacent <w:r> elements that have identical <w:rPr> properties.
 * This reduces token count and helps the LLM see continuous text.
 */
export function mergeIdenticalRuns(oxml: string): string {
    const doc = new DOMParser().parseFromString(oxml, 'text/xml');
    const paragraphs = doc.getElementsByTagName('w:p');

    for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const runs = Array.from(p.childNodes).filter(n => n.nodeName === 'w:r');

        for (let j = 0; j < runs.length - 1; j++) {
            const currentRun = runs[j] as Element;
            const nextRun = runs[j + 1] as Element;

            // Get rPr strings for comparison
            const currentRPr = getRPrString(currentRun);
            const nextRPr = getRPrString(nextRun);

            if (currentRPr === nextRPr) {
                // Merge nextRun into currentRun
                const currentT = currentRun.getElementsByTagName('w:t')[0];
                const nextT = nextRun.getElementsByTagName('w:t')[0];

                if (currentT && nextT) {
                    // Append text
                    currentT.textContent = (currentT.textContent || '') + (nextT.textContent || '');

                    // Remove nextRun
                    p.removeChild(nextRun);

                    // Adjust index to check this run again against the *new* next run
                    runs.splice(j + 1, 1);
                    j--;
                }
            }
        }
    }
    return new XMLSerializer().serializeToString(doc);
}

function getRPrString(run: Element): string {
    const rPr = run.getElementsByTagName('w:rPr')[0];
    return rPr ? new XMLSerializer().serializeToString(rPr) : '';
}

/**
 * Strips w:rsid* attributes to prevent "drift" and reduce noise.
 * Word will regenerate these upon save.
 */
export function stripRsidAttributes(oxml: string): string {
    // Regex to remove w:rsidR="...", w:rsidRPr="...", w:rsidP="...", w:rsidRDefault="..."
    return oxml.replace(/ w:rsid[A-Za-z]+="[0-9A-F]+"/g, '');
}
