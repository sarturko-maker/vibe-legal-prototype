
import { constructOxmlPrompt, runOxmlAgent } from './Agent_Oxml_Redline';
import * as geminiApi from './geminiApi';

// Mock callGemini
const originalCallGemini = geminiApi.callGemini;
let mockResponse: string = "";

// Monkey patch callGemini
(geminiApi as any).callGemini = async (apiKey: string, model: string, prompt: string, mode: string) => {
    return mockResponse;
};

function assert(condition: boolean, message: string) {
    if (condition) {
        console.log(`✅ PASS: ${message}`);
    } else {
        console.error(`❌ FAIL: ${message}`);
        process.exit(1);
    }
}

async function runTests() {
    console.log("🚀 Running V3.1 Stress Tests (Visual Context & JSON)...\n");

    // Test S1: The Bold Header (Prompt Verification)
    console.log("--- Test S1: The Bold Header (Prompt & Parsing) ---");
    const s1Input = '<w:p><w:r><w:b/><w:t>Note:</w:t></w:r><w:r><w:t> content...</w:t></w:r></w:p>';
    const s1Context = "Para 1: Style='Normal', Font='Bold', StartsWith='Note: content...'";

    // 1. Verify Prompt Construction
    const promptS1 = constructOxmlPrompt(s1Input, "Soften the tone", s1Context);
    assert(promptS1.includes("<VISUAL_CONTEXT>\nPara 1: Style='Normal', Font='Bold'"), "Prompt includes Visual Context");
    assert(promptS1.includes("OUTPUT FORMAT: JSON"), "Prompt requests JSON");

    // 2. Verify JSON Parsing & Style Preservation (Simulated)
    // We simulate a "Good" LLM response that respects the instruction
    mockResponse = JSON.stringify({
        analysis: "User wants to soften tone. Visual context shows Bold 'Note:'. Preserving formatting.",
        changes: ["Softened content", "Preserved Bold on 'Note:'"],
        oxml: '<w:p><w:r><w:b/><w:t>Note:</w:t></w:r><w:r><w:t> gentle content...</w:t></w:r></w:p>'
    });

    const resultS1 = await runOxmlAgent("key", "model", s1Input, "Soften the tone", s1Context);
    assert(!resultS1.error, "Agent completed without error");
    assert(resultS1.oxml.includes("<w:b/>"), "Output preserved Bold tag");
    assert(resultS1.oxml.includes("gentle content"), "Output applied text change");


    // Test S2: The Style Preservation
    console.log("\n--- Test S2: The Style Preservation ---");
    const s2Input = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Header</w:t></w:r></w:p>';
    const s2Context = "Para 1: Style='Heading1', Font='Normal', StartsWith='Header'";

    // 1. Verify Prompt
    const promptS2 = constructOxmlPrompt(s2Input, "Fix typos", s2Context);
    assert(promptS2.includes("Style='Heading1'"), "Prompt includes Heading1 style in context");

    // 2. Verify JSON Parsing
    mockResponse = JSON.stringify({
        analysis: "Fixing typos. Preserving Heading1 style.",
        changes: ["Fixed typo"],
        oxml: '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Header Fixed</w:t></w:r></w:p>'
    });

    const resultS2 = await runOxmlAgent("key", "model", s2Input, "Fix typos", s2Context);
    assert(resultS2.oxml.includes('w:val="Heading1"'), "Output preserved Heading1 style");

    console.log("\n✅ PASSED: V3.1 Architecture Verified.");
}

runTests().catch(e => console.error(e));
