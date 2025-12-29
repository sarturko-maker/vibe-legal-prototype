import { MindMapTopic, MindMapSubTopic } from './documentAnalysis';

export async function generateTopicAdvice(
    apiKey: string,
    topic: MindMapTopic,
    subTopic: MindMapSubTopic | null,
    documentText: string,
    dealContext: string | null,
    selectedSide: string | null,
    userQuestion: string | null
): Promise<string> {

    console.log('[generateTopicAdvice] Topic:', topic.title, 'SubTopic:', subTopic?.name || 'none');

    const prompt = `You are a commercial lawyer advising on a contract.

CONTRACT EXCERPT:
${documentText.substring(0, 12000)}

${dealContext ? `DEAL CONTEXT:\n${dealContext}\n` : ''}
${selectedSide ? `You are advising: ${selectedSide}\n` : ''}

TOPIC: ${topic.title}
${topic.summary}
${topic.keyFigures.length > 0 ? `Key figures: ${topic.keyFigures.join(', ')}` : ''}

${subTopic ? `SUB-TOPIC: ${subTopic.name}\n${subTopic.summary}` : ''}

${userQuestion
            ? `USER QUESTION: ${userQuestion}

Answer the user's specific question about this topic.`
            : `Provide practical advice about this ${subTopic ? 'sub-topic' : 'topic'}.
Consider:
- What should the reader understand about this?
- Are there any concerns or risks?
- What might they want to negotiate or clarify?`}

Be concise and practical. Use plain language.
3-4 paragraphs maximum. No bullet points unless essential.`;

    const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${apiKey}`,
        {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.4 }
            })
        }
    );

    const data = await response.json();
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';

    console.log('[generateTopicAdvice] Response length:', text.length);

    return text;
}
