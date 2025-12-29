
import { buildRouterSystemPrompt } from '../prompts/systemPrompt';

console.log('--- Verifying System Prompt Constraints ---');
const prompt = buildRouterSystemPrompt(null, '', null);

if (prompt.includes('NO OTHER OPERATION TYPES ARE PERMITTED')) {
    console.log('✅ Strict operation warning present');
} else {
    console.error('❌ Strict operation warning MISSING');
}

if (prompt.includes('AMEND_MULTI')) {
    console.error('❌ AMEND_MULTI or fake types invented');
}

if (!prompt.includes('Use multiple INSERT operations')) {
    console.error('❌ Complex request instructions MISSING');
} else {
    console.log('✅ Complex request instructions present');
}

if (prompt.includes('INSERT_BLOCK')) {
    console.log('✅ INSERT_BLOCK mentioned in prohibition example (expected)');
}
