/**
 * Formatting Preservation Diagnostic
 * 
 * Tests whether bold, italic, font, and other formatting is preserved
 */

import { applyRedlineToOxml } from './OxmlEngine';
import { DOMParser, XMLSerializer } from 'xmldom';

(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("🔬 Formatting Preservation Diagnostic\n");
console.log("=".repeat(60));

// Test 1: Bold text modification
console.log("\n📝 TEST 1: Bold Text Modification\n");

const boldOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:p>
        <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>Bold Text</w:t>
        </w:r>
    </w:p>
</w:body>
</w:document>`;

const result1 = applyRedlineToOxml(boldOxml, "Bold Text", "Bold Modified");

console.log("INPUT: Bold Text with <w:b/>");
console.log("OUTPUT:");
console.log(result1.oxml);

const doc1 = new DOMParser().parseFromString(result1.oxml, "text/xml");
const boldTags = doc1.getElementsByTagName("w:b");
console.log(`\nBold tags found: ${boldTags.length} (expected: at least 1)`);

if (boldTags.length === 0) {
    console.log("❌ FAIL: Bold formatting lost!");
} else {
    console.log("✅ PASS: Bold formatting preserved");
}

// Test 2: Multiple formatting properties
console.log("\n" + "=".repeat(60));
console.log("\n📝 TEST 2: Multiple Formatting Properties\n");

const multiFormatOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:p>
        <w:r>
            <w:rPr>
                <w:b/>
                <w:i/>
                <w:u w:val="single"/>
                <w:sz w:val="28"/>
                <w:color w:val="FF0000"/>
            </w:rPr>
            <w:t>Formatted Text</w:t>
        </w:r>
    </w:p>
</w:body>
</w:document>`;

const result2 = applyRedlineToOxml(multiFormatOxml, "Formatted Text", "Formatted Modified");

console.log("INPUT: Text with bold, italic, underline, size 14pt, red color");
console.log("OUTPUT (abbreviated):");
console.log(result2.oxml.substring(0, 800));

const doc2 = new DOMParser().parseFromString(result2.oxml, "text/xml");
const bold2 = doc2.getElementsByTagName("w:b").length;
const italic2 = doc2.getElementsByTagName("w:i").length;
const underline2 = doc2.getElementsByTagName("w:u").length;
const size2 = doc2.getElementsByTagName("w:sz").length;
const color2 = doc2.getElementsByTagName("w:color").length;

console.log(`\nFormatting preserved:`);
console.log(`  Bold: ${bold2 > 0 ? '✅' : '❌'}`);
console.log(`  Italic: ${italic2 > 0 ? '✅' : '❌'}`);
console.log(`  Underline: ${underline2 > 0 ? '✅' : '❌'}`);
console.log(`  Size: ${size2 > 0 ? '✅' : '❌'}`);
console.log(`  Color: ${color2 > 0 ? '✅' : '❌'}`);

const allPreserved = bold2 > 0 && italic2 > 0 && underline2 > 0 && size2 > 0 && color2 > 0;

if (!allPreserved) {
    console.log("\n❌ CRITICAL: Formatting is being lost!");
    process.exit(1);
} else {
    console.log("\n✅ All formatting preserved");
}
