
import { validateOxml, calculateStructuralChecksum } from './OxmlValidator';
import { runOxmlAgent } from './Agent_Oxml_Redline';
import * as geminiApi from './geminiApi';

// Mock callGemini manually
const originalCallGemini = geminiApi.callGemini;
let mockCallGemini: any = null;

// Override the export (hacky but works for this environment)
(geminiApi as any).callGemini = async (apiKey: string, model: string, prompt: string, mode: any) => {
    if (mockCallGemini) return mockCallGemini(apiKey, model, prompt, mode);
    return originalCallGemini(apiKey, model, prompt, mode);
};

function assert(condition: boolean, message: string) {
    if (condition) {
        console.log(`✅ PASS: ${message}`);
    } else {
        console.error(`❌ FAIL: ${message}`);
        process.exit(1);
    }
}

console.log("Running Oxml Safeguards Test Suite...");

// --- Validator Tests ---
console.log("\n[TEST] Validator");

const invalidXml1 = '<w:p><w:t>Unclosed';
const res1 = validateOxml(invalidXml1);
assert(res1.isValid === false && !!res1.error, 'Detects unclosed tags');

const invalidXml2 = '<w:p>Illegal Text<w:r><w:t>Legal</w:t></w:r></w:p>';
const res2 = validateOxml(invalidXml2);
assert(res2.isValid === false && res2.error?.includes('Found text outside') || false, 'Detects text outside w:t');

const invalidXml3 = '<w:p><w:ins><w:r><w:t>New</w:t></w:r></w:ins></w:p>';
const res3 = validateOxml(invalidXml3);
assert(res3.isValid === false && res3.error?.includes('missing required attributes') || false, 'Detects missing attributes');

const validXml = '<w:p><w:ins w:id="1" w:author="A" w:date="D"><w:r><w:t>New</w:t></w:r></w:ins></w:p>';
const res4 = validateOxml(validXml);
assert(res4.isValid === true, 'Passes valid OOXML');


// --- Checksum Tests ---
console.log("\n[TEST] Checksum");

const xml1 = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>';
const xml2 = '<w:p><w:r><w:t>World</w:t></w:r></w:p>';
assert(calculateStructuralChecksum(xml1) === calculateStructuralChecksum(xml2), 'Ignores text content');

const xml3 = '<w:p><w:r><w:b/><w:t>Hello</w:t></w:r></w:p>';
assert(calculateStructuralChecksum(xml1) !== calculateStructuralChecksum(xml3), 'Detects structural changes');


// --- Retry Protocol Tests ---
console.log("\n[TEST] Retry Protocol");

async function testRetry() {
    // Mock Attempt 1: Invalid XML
    let attempts = 0;
    mockCallGemini = async () => {
        attempts++;
        if (attempts === 1) return '<w:p>Invalid';
        return '<w:p><w:r><w:t>Valid</w:t></w:r></w:p>';
    };

    const result = await runOxmlAgent('key', 'model', '<w:p/>', 'Edit');
    console.log(`Attempts: ${attempts}`);
    console.log(`Result OXML: ${result.oxml}`);
    assert(attempts === 2, 'Retried once');
    assert(result.oxml.includes('Valid'), 'Returned valid result');
    assert(!result.error, 'No error in final result');
}

async function testAbort() {
    // Mock 3 Invalid Attempts
    let attempts = 0;
    mockCallGemini = async () => {
        attempts++;
        return '<w:p>Invalid';
    };

    const result = await runOxmlAgent('key', 'model', '<w:p/>', 'Edit');
    assert(attempts === 3, 'Retried max times');
    assert(result.oxml === '<w:p/>', 'Returned original on abort');
    assert(!!result.error, 'Returned error message');
}

// Run async tests
(async () => {
    try {
        await testRetry();
        await testAbort();
        console.log("\nSUMMARY: All Safeguards Tests Passed.");
    } catch (e) {
        console.error(e);
        process.exit(1);
    }
})();
