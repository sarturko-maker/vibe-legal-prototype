export async function callGemini(
    apiKey: string,
    model: string,
    prompt: string,
    mode: "ASK" | "REDLINE" | "DRAFT" | "OOXML",
): Promise<string> {
    if (!apiKey) throw new Error("API Key missing.");
    const cleanModel = (model || "gemini-1.5-flash").replace(/^models\//, "");
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${cleanModel}:generateContent?key=${apiKey}`;

    let systemInstruction = "";
    if (mode === "REDLINE") {
        systemInstruction =
            "SYSTEM: You are a strict legal editor. Return ONLY the modified legal text. No markdown. No quotes. Do not use LaTeX. Preserve placeholders like [Name].";
    } else if (mode === "DRAFT") {
        systemInstruction =
            "SYSTEM: You are an expert legal drafter. Write a clean, professional clause based on the instruction. Return ONLY the clause text.";
    } else if (mode === "OOXML") {
        systemInstruction = "";
    } else {
        systemInstruction =
            "SYSTEM: You are a Senior Legal Counsel. Format response with Markdown (**bold**, lists). Be concise.";
    }

    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contents: [{ parts: [{ text: `${systemInstruction}\n\nUSER: ${prompt}` }] }] }),
    });

    if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error?.message || "Gemini API Error");
    }
    const data = await response.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "";
}

