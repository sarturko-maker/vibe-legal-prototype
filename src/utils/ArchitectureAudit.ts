/**
 * Test S-Real: The Verbose Payload
 * 
 * This test simulates a real-world scenario where the LLM returns
 * a large OOXML string inside a JSON object.
 */

// Simulate a 50-line OOXML string with double quotes in attributes
function generateVerboseOoxml(): string {
    let oxml = '<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n';

    for (let i = 1; i <= 50; i++) {
        oxml += `  <w:p w:rsidR="00${i.toString().padStart(4, '0')}" w:rsidRDefault="00${i.toString().padStart(4, '0')}">\n`;
        oxml += `    <w:pPr>\n`;
        oxml += `      <w:pStyle w:val="Heading${(i % 3) + 1}"/>\n`;
        oxml += `      <w:jc w:val="center"/>\n`;
        oxml += `    </w:pPr>\n`;
        oxml += `    <w:r w:rsidRPr="00${i.toString().padStart(4, '0')}">\n`;
        oxml += `      <w:rPr>\n`;
        oxml += `        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>\n`;
        oxml += `        <w:sz w:val="24"/>\n`;
        oxml += `      </w:rPr>\n`;
        oxml += `      <w:t xml:space="preserve">Paragraph ${i} with "quoted text" and special chars: § ¶ — & < ></w:t>\n`;
        oxml += `    </w:r>\n`;
        oxml += `  </w:p>\n`;
    }

    oxml += '</w:body>';
    return oxml;
}

// Test 1: Can we parse OXML with quotes inside JSON?
function testJsonParsing() {
    console.log("=== Test S-Real: The Verbose Payload ===\n");

    const verboseOxml = generateVerboseOoxml();
    console.log(`Generated OXML: ${verboseOxml.length} characters, ${verboseOxml.split('\n').length} lines\n`);

    // Simulate what the LLM would return
    const mockLlmResponse = {
        analysis: "Made 50 paragraph changes",
        changes: ["Changed all paragraphs"],
        oxml: verboseOxml
    };

    // Scenario 1: JSON.stringify + JSON.parse (Best Case)
    console.log("--- Scenario 1: JSON.stringify → JSON.parse ---");
    try {
        const jsonString = JSON.stringify(mockLlmResponse);
        console.log(`JSON String Length: ${jsonString.length} characters`);

        const parsed = JSON.parse(jsonString);
        console.log(`✅ PASS: JSON.parse succeeded`);
        console.log(`   Parsed oxml length: ${parsed.oxml.length} characters`);
    } catch (e: any) {
        console.log(`❌ FAIL: ${e.message}`);
    }

    // Scenario 2: Manual JSON construction (What LLM might do - BROKEN)
    console.log("\n--- Scenario 2: Manual JSON Construction (Simulating LLM) ---");

    // This is what we FEAR the LLM might do - not escaping quotes
    const brokenJson = `{
        "analysis": "Made changes",
        "changes": ["test"],
        "oxml": "${verboseOxml}"
    }`;

    try {
        JSON.parse(brokenJson);
        console.log(`✅ PASS: JSON.parse succeeded`);
    } catch (e: any) {
        console.log(`❌ FAIL: JSON.parse failed - ${e.message}`);
        console.log(`   This is the EXPECTED failure mode!`);
    }

    // Scenario 3: Properly escaped (What LLM SHOULD do)
    console.log("\n--- Scenario 3: Properly Escaped JSON ---");

    const escapedOxml = verboseOxml
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\n/g, '\\n');

    const goodJson = `{
        "analysis": "Made changes",
        "changes": ["test"],
        "oxml": "${escapedOxml}"
    }`;

    try {
        const parsed = JSON.parse(goodJson);
        console.log(`✅ PASS: JSON.parse succeeded with escaped content`);
        console.log(`   Parsed oxml length: ${parsed.oxml.length} characters`);
    } catch (e: any) {
        console.log(`❌ FAIL: ${e.message}`);
    }

    // Scenario 4: Delimiter-based alternative
    console.log("\n--- Scenario 4: Delimiter-Based Alternative ---");

    const delimitedResponse = `
ANALYSIS: Made changes to 50 paragraphs

---START_OXML---
${verboseOxml}
---END_OXML---
`;

    const oxmlMatch = delimitedResponse.match(/---START_OXML---\n([\s\S]*?)\n---END_OXML---/);
    if (oxmlMatch) {
        console.log(`✅ PASS: Delimiter extraction succeeded`);
        console.log(`   Extracted oxml length: ${oxmlMatch[1].length} characters`);
        console.log(`   Quotes preserved: ${oxmlMatch[1].includes('"quoted text"')}`);
    } else {
        console.log(`❌ FAIL: Delimiter extraction failed`);
    }

    // VERDICT
    console.log("\n========================================");
    console.log("VERDICT: OOXML-inside-JSON Architecture");
    console.log("========================================");
    console.log("\n❌ HIGH RISK: The architecture is fragile.");
    console.log("\nReasons:");
    console.log("1. LLMs do NOT reliably escape quotes in large strings");
    console.log("2. Newlines inside OOXML break JSON string literals");
    console.log("3. Special chars (& < >) may cause double-escaping issues");
    console.log("\n✅ RECOMMENDED: Option B (Hybrid Engine)");
    console.log("   - LLM returns TEXT redlines (which it does well)");
    console.log("   - diff-match-patch + applyRedlineToOxml handles XML");
    console.log("   - This is ALREADY IMPLEMENTED and working!");
}

testJsonParsing();
