/**
 * Table Formatting Test
 */

import { applyRedlineToOxml } from './OxmlEngine';
import { DOMParser, XMLSerializer } from 'xmldom';

(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("🔬 Table Formatting Test\n");

const tableWithBold = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:tbl>
        <w:tr>
            <w:tc>
                <w:p>
                    <w:r>
                        <w:rPr>
                            <w:b/>
                            <w:sz w:val="28"/>
                        </w:rPr>
                        <w:t>Bold Text</w:t>
                    </w:r>
                </w:p>
            </w:tc>
        </w:tr>
    </w:tbl>
</w:body>
</w:document>`;

const result = applyRedlineToOxml(tableWithBold, "Bold Text", "Bold Modified");

console.log("OUTPUT:\n");
console.log(result.oxml);

const doc = new DOMParser().parseFromString(result.oxml, "text/xml");

// Check for w:ins
const inserts = doc.getElementsByTagName("w:ins");
console.log(`\nInserts found: ${inserts.length}`);

if (inserts.length > 0) {
    const firstIns = inserts[0];
    console.log("\nFirst INSERT content:");
    console.log(firstIns.outerHTML || firstIns.toString());

    // Check if it has w:rPr
    const rPrs = firstIns.getElementsByTagName("w:rPr");
    console.log(`\nw:rPr in INSERT: ${rPrs.length > 0 ? '✅ EXISTS' : '❌ MISSING'}`);

    if (rPrs.length > 0) {
        console.log("rPr content:", rPrs[0].outerHTML || rPrs[0].toString());
    }
}
