
import { constructOxmlPrompt, runOxmlAgent } from './Agent_Oxml_Redline';
import * as geminiApi from './geminiApi';

// Mock callGemini manually
const originalCallGemini = geminiApi.callGemini;
let mockCallGemini: any = null;

// Override the export (this is hacky in CommonJS/TS-Node but might work if we modify the property)
// Since we can't easily spy on ES modules without Jest, we'll test constructOxmlPrompt directly
// and trust runOxmlAgent's simple logic, or try to mock if possible.
// Actually, runOxmlAgent just calls callGemini. The most important part is the prompt construction.

function assert(condition: boolean, message: string) {
    if (condition) {
        console.log(`✅ PASS: ${message}`);
    } else {
        console.error(`❌ FAIL: ${message}`);
        process.exit(1);
    }
}

console.log("Running Agent_Oxml_Redline Test Suite...");

// Test 1: Prompt Construction
const oxml = '<w:p><w:t>Hello</w:t></w:p>';
const instruction = 'Change Hello to Hi';
const prompt = constructOxmlPrompt(oxml, instruction);

assert(prompt.includes('<INSTRUCTION>\nChange Hello to Hi\n</INSTRUCTION>'), 'Prompt contains instruction');
assert(prompt.includes('<INPUT_OOXML>\n<w:p><w:t>Hello</w:t></w:p>\n</INPUT_OOXML>'), 'Prompt contains encapsulated OOXML');
assert(prompt.includes('NATIVE TRACK CHANGES:'), 'Prompt contains Track Changes instruction');

// Test 2: Verify runOxmlAgent calls callGemini (Mocking via monkey patching if possible, else skip)
// In this environment, we can't easily mock the import. 
// We will rely on the prompt construction test.

console.log("SUMMARY: All tests passed.");
