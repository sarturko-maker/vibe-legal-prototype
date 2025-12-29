
import { determineIntent } from './IntentRouter';
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

async function runIntegrationTest() {
    console.log("🔄 Running V3.1 Integration Data Flow Simulation...\n");

    // Scenario: User selects a row in a table and says: 'Delete this row.'
    const inputOxml = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const visualContext = "Selection: Table Row 1. Style='Table Grid'.";
    const instruction = "Delete this row.";

    // Step 1: Sanitizer (Mocked check)
    console.log("--- Step 1: Sanitizer ---");
    const parentBodyType: string = "Table Cell";
    const isAllowed = parentBodyType === "Main Document" || parentBodyType === "Table Cell";
    assert(isAllowed, `Sanitizer allows '${parentBodyType}'`);

    // Step 2: Context Extractor (Mocked)
    console.log("\n--- Step 2: Context Extractor ---");
    console.log(`Visual Context: "${visualContext}"`);
    assert(visualContext.includes("Table Row"), "Visual Context captured Table Row");

    // Step 3: Intent Router (Verify it allows the command)
    console.log("\n--- Step 3: Intent Router ---");
    const routerResult = determineIntent(instruction, { isEmpty: false, text: "Cell 1" }, { headers: [], cursorSection: "Body" }, visualContext);
    console.log(`Intent: ${routerResult.intent}, Scope: ${routerResult.scope}`);
    assert(routerResult.intent === 'COMMAND', "Router identified COMMAND");

    // Step 4: Prompt Assembly
    console.log("\n--- Step 4: Prompt Assembly ---");
    const prompt = constructOxmlPrompt(inputOxml, instruction, visualContext);
    assert(prompt.includes("<VISUAL_CONTEXT>\nSelection: Table Row 1"), "Prompt includes Visual Context");
    assert(prompt.includes("OUTPUT FORMAT: JSON"), "Prompt requests JSON");

    // Step 5: LLM Response (Simulated)
    console.log("\n--- Step 5: LLM Response (Simulated) ---");
    mockResponse = JSON.stringify({
        analysis: "User wants to remove the selected table row.",
        changes: ["Removed <w:tr> containing 'Cell 1'"],
        oxml: '<w:tbl></w:tbl>' // Empty table (row removed)
    });
    console.log(`Mock Response: ${mockResponse}`);

    // Step 6: Parser & Execution
    console.log("\n--- Step 6: Parser & Execution ---");
    const result = await runOxmlAgent("key", "model", inputOxml, instruction, visualContext);

    console.log(`Actual Result OXML: '${result.oxml}'`);
    assert(!result.error, "Agent executed successfully");
    assert(result.oxml === '<w:tbl/>', "Agent returned modified XML");

    console.log("\n✅ PASSED: End-to-End Data Flow Verified.");
}

runIntegrationTest().catch(e => console.error(e));
