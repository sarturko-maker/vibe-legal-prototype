// ==========================================
// VIBE LEGAL v2.0 (RE-FLOW)
// ==========================================
declare var diff_match_patch: any;
declare var React: any;
declare var ReactDOM: any;

// ==========================================
// 1. OXML ENGINE (Multi-Paragraph)
// ==========================================

function calculateRedline(original: string, modified: string): Array<{ op: number; text: string }> {
  const dmp = new diff_match_patch();
  const diffs = dmp.diff_main(original, modified);
  dmp.diff_cleanupSemantic(diffs);
  return diffs.map((diff) => ({ op: diff[0], text: diff[1] }));
}

function applyRedlineToOxml(oxml: string, originalText: string, modifiedText: string): string {
  const parser = new DOMParser();
  const xmlSerializer = new XMLSerializer();
  let xmlDoc;

  try {
    xmlDoc = parser.parseFromString(oxml, "text/xml");
  } catch (e) {
    console.error("OXML Parse Error");
    return oxml;
  }

  if (xmlDoc.getElementsByTagName("parsererror").length > 0) return oxml;

  // 1. EXTRACT PARAGRAPHS
  // We get ALL paragraphs in the selection, not just the first one.
  const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));
  if (paragraphs.length === 0) return oxml;

  // The container to append new content to (usually w:body or w:tc)
  const container = paragraphs[0].parentNode;

  // 2. FLATTEN TEXT & MAP PROPERTIES
  // We build a single text string but remember where the formatting came from.
  let originalFullText = "";
  const propertyMap: { start: number; end: number; rPr: Element | null }[] = [];

  paragraphs.forEach((p, pIndex) => {
    const runs = Array.from(p.getElementsByTagName("w:r"));
    runs.forEach((r) => {
      const rPr = r.getElementsByTagName("w:rPr")[0] || null;
      const texts = Array.from(r.getElementsByTagName("w:t"));
      texts.forEach((t) => {
        const textContent = t.textContent || "";
        if (textContent.length > 0) {
          propertyMap.push({
            start: originalFullText.length,
            end: originalFullText.length + textContent.length,
            rPr: rPr,
          });
          originalFullText += textContent;
        }
      });
    });

    // Treat paragraph breaks as newlines for the diff engine
    if (pIndex < paragraphs.length - 1) {
      originalFullText += "\n";
    }
  });

  // 3. CLEAN & DIFF
  const cleanModified = modifiedText
    .replace(/^(Here is the redline:|Here is the text:|Sure, I can help:|Here's the updated text:)\s*/i, "")
    .replace(/\$\\text\{/g, "")
    .replace(/\}\$/g, "")
    .replace(/\$([^0-9\n]+?)\$/g, "$1"); // Remove LaTeX artifacts

  const diffs = calculateRedline(originalFullText, cleanModified);

  // 4. RECONSTRUCTION
  // We create a DocumentFragment to hold the new structure
  const newContentFragment = xmlDoc.createDocumentFragment();

  // Helper to create a new paragraph (cloning style from the FIRST original paragraph)
  // In a V3, this would clone the style of the *closest* paragraph.
  const firstPara = paragraphs[0];
  const createNewParagraph = () => {
    const newP = xmlDoc.createElement("w:p");
    const pPr = firstPara.getElementsByTagName("w:pPr")[0];
    if (pPr) newP.appendChild(pPr.cloneNode(true));
    return newP;
  };

  let currentParagraph = createNewParagraph();
  newContentFragment.appendChild(currentParagraph);

  let currentOriginalIndex = 0;

  // Helper: Find style at a specific character index
  const getRunProperties = (index: number): Element | null => {
    const match = propertyMap.find((m) => index >= m.start && index < m.end);
    return match ? match.rPr : null;
  };

  const appendTextToCurrent = (text: string, type: "equal" | "insert" | "delete", rPr: Element | null) => {
    const parts = text.split("\n"); // Handle newlines -> New Paragraphs

    parts.forEach((part, index) => {
      if (index > 0) {
        // Newline found: Split paragraph
        if (type !== "delete") {
          currentParagraph = createNewParagraph();
          newContentFragment.appendChild(currentParagraph);
        }
      }

      if (part.length > 0) {
        if (type === "delete") {
          const del = xmlDoc.createElement("w:del");
          del.setAttribute("w:id", Math.floor(Math.random() * 100000).toString());
          del.setAttribute("w:author", "Vibe Legal");
          del.setAttribute("w:date", new Date().toISOString());

          const run = xmlDoc.createElement("w:r");
          if (rPr) run.appendChild(rPr.cloneNode(true));

          const delText = xmlDoc.createElement("w:delText");
          delText.setAttribute("xml:space", "preserve");
          delText.textContent = part;

          run.appendChild(delText);
          del.appendChild(run);
          currentParagraph.appendChild(del);
        } else if (type === "insert") {
          const ins = xmlDoc.createElement("w:ins");
          ins.setAttribute("w:id", Math.floor(Math.random() * 100000).toString());
          ins.setAttribute("w:author", "Vibe Legal");
          ins.setAttribute("w:date", new Date().toISOString());

          const run = xmlDoc.createElement("w:r");
          if (rPr) run.appendChild(rPr.cloneNode(true));

          const t = xmlDoc.createElement("w:t");
          t.setAttribute("xml:space", "preserve");
          t.textContent = part;

          run.appendChild(t);
          ins.appendChild(run);
          currentParagraph.appendChild(ins);
        } else {
          // EQUAL
          const run = xmlDoc.createElement("w:r");
          if (rPr) run.appendChild(rPr.cloneNode(true));

          const t = xmlDoc.createElement("w:t");
          t.setAttribute("xml:space", "preserve");
          t.textContent = part;

          run.appendChild(t);
          currentParagraph.appendChild(run);
        }
      }
    });
  };

  // 5. EXECUTE DIFF LOOP
  for (const diff of diffs) {
    const [op, text] = [diff.op, diff.text];

    if (op === 0) {
      // EQUAL
      let offset = 0;
      while (offset < text.length) {
        // Process chunk-by-chunk to preserve changing properties
        const rPr = getRunProperties(currentOriginalIndex + offset);

        // Find how long this property matches
        const currentPropRange = propertyMap.find(
          (m) => currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end,
        );

        let length = 1;
        if (currentPropRange) {
          const remainingInRun = currentPropRange.end - (currentOriginalIndex + offset);
          const remainingInDiff = text.length - offset;
          length = Math.min(remainingInRun, remainingInDiff);
        }

        const chunk = text.substring(offset, offset + length);
        appendTextToCurrent(chunk, "equal", rPr || null);
        offset += length;
      }
      currentOriginalIndex += text.length;
    } else if (op === 1) {
      // INSERT
      // Use property of nearest neighbor (start of insert position)
      const rPr = getRunProperties(currentOriginalIndex);
      appendTextToCurrent(text, "insert", rPr);
    } else if (op === -1) {
      // DELETE
      let offset = 0;
      while (offset < text.length) {
        const rPr = getRunProperties(currentOriginalIndex + offset);

        const currentPropRange = propertyMap.find(
          (m) => currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end,
        );

        let length = 1;
        if (currentPropRange) {
          const remainingInRun = currentPropRange.end - (currentOriginalIndex + offset);
          const remainingInDiff = text.length - offset;
          length = Math.min(remainingInRun, remainingInDiff);
        }

        const chunk = text.substring(offset, offset + length);
        appendTextToCurrent(chunk, "delete", rPr || null);
        offset += length;
      }
      currentOriginalIndex += text.length;
    }
  }

  // 6. SWAP
  // Remove old paragraphs
  paragraphs.forEach((p) => {
    if (p.parentNode) p.parentNode.removeChild(p);
  });
  // Inject new structure
  container.appendChild(newContentFragment);

  return xmlSerializer.serializeToString(xmlDoc);
}

// ==========================================
// 2. NETWORK LAYER (TRUE SYNC)
// ==========================================

async function fetchModelsFromGoogle(apiKey: string): Promise<any[]> {
  if (!apiKey) throw new Error("No API Key");
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  const response = await fetch(url);
  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || "Failed to fetch models");
  }
  const data = await response.json();
  return (data.models || [])
    .filter((m) => m.supportedGenerationMethods?.includes("generateContent"))
    .sort((a, b) => b.name.localeCompare(a.name));
}

async function callGemini(
  apiKey: string,
  model: string,
  prompt: string,
  mode: "ASK" | "REDLINE" | "DRAFT",
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
  } else {
    systemInstruction =
      "SYSTEM: You are a Senior Legal Counsel. Format response with Markdown (**bold**, lists). Be concise.";
  }

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ contents: [{ parts: [{ text: `${systemInstruction}\n\nUSER: ${prompt}` }] }] }),
  });

  if (!response.ok) throw new Error("Gemini API Error");
  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "";
}

function renderMarkdown(text: string): string {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.*?)\*/g, "<em>$1</em>")
    .replace(/\n/g, "<br>");
}

// ==========================================
// 3. UI COMPONENTS
// ==========================================

const fontStack = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';

const Header = ({ onSettingsClick }) =>
  React.createElement(
    "div",
    {
      style: {
        padding: "14px 16px",
        borderBottom: "1px solid #e5e5e5",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        background: "#fff",
      },
    },
    React.createElement(
      "div",
      null,
      React.createElement(
        "h2",
        { style: { margin: 0, fontSize: "14px", fontWeight: "600", color: "#111", fontFamily: fontStack } },
        "Vibe Legal",
      ),
    ),
    React.createElement(
      "button",
      {
        onClick: onSettingsClick,
        style: {
          background: "none",
          border: "none",
          cursor: "pointer",
          fontSize: "16px",
          opacity: 0.7,
          padding: "4px",
        },
      },
      "⚙️",
    ),
  );

const Settings = ({ apiKey, setApiKey, selectedModel, setSelectedModel, onBack }) => {
  const [status, setStatus] = React.useState("");
  const [models, setModels] = React.useState([]);

  React.useEffect(() => {
    if (apiKey) handleSync();
  }, []);

  const handleSync = async () => {
    setStatus("Loading...");
    try {
      const list = await fetchModelsFromGoogle(apiKey);
      setModels(list);
      setStatus(list.length + " models found");
      const currentExists = list.find((m) => m.name.replace("models/", "") === selectedModel);
      if (!currentExists && list.length > 0) {
        setSelectedModel(list[0].name.replace("models/", ""));
      }
    } catch (e) {
      setStatus("Error: Check Key");
      setModels([]);
    }
  };

  return React.createElement(
    "div",
    { style: { padding: "20px", background: "#f9f9f9", height: "100%", fontFamily: fontStack, fontSize: "13px" } },
    React.createElement("h3", { style: { marginTop: 0, fontSize: "14px", marginBottom: "20px" } }, "Settings"),
    React.createElement(
      "div",
      { style: { marginBottom: "20px" } },
      React.createElement(
        "label",
        { style: { display: "block", fontWeight: "600", marginBottom: "6px", color: "#444" } },
        "API Key",
      ),
      React.createElement("input", {
        type: "password",
        value: apiKey,
        onChange: (e) => setApiKey(e.target.value),
        style: { width: "100%", padding: "10px", border: "1px solid #ddd", borderRadius: "6px", fontSize: "13px" },
        placeholder: "Enter Gemini API Key",
      }),
    ),
    React.createElement(
      "div",
      { style: { marginBottom: "20px" } },
      React.createElement(
        "div",
        { style: { display: "flex", justifyContent: "space-between", marginBottom: "6px" } },
        React.createElement("label", { style: { fontWeight: "600", color: "#444" } }, "Model Selection"),
        React.createElement(
          "button",
          {
            onClick: handleSync,
            style: {
              fontSize: "11px",
              background: "none",
              border: "none",
              color: "#000",
              textDecoration: "underline",
              cursor: "pointer",
            },
          },
          "Refresh",
        ),
      ),
      React.createElement(
        "select",
        {
          value: selectedModel,
          onChange: (e) => setSelectedModel(e.target.value),
          style: {
            width: "100%",
            padding: "10px",
            border: "1px solid #ddd",
            borderRadius: "6px",
            fontSize: "13px",
            background: "#fff",
          },
        },
        models.length > 0
          ? models.map((m) =>
              React.createElement(
                "option",
                { key: m.name, value: m.name.replace("models/", "") },
                m.displayName || m.name,
              ),
            )
          : React.createElement("option", { value: selectedModel }, selectedModel || "Default"),
      ),
      React.createElement("div", { style: { fontSize: "11px", color: "#666", marginTop: "6px" } }, status),
    ),
    React.createElement(
      "button",
      {
        onClick: () => {
          localStorage.setItem("vibe_api_key", apiKey);
          localStorage.setItem("vibe_model", selectedModel);
          onBack();
        },
        style: {
          width: "100%",
          background: "#000",
          color: "#fff",
          padding: "12px",
          border: "none",
          borderRadius: "6px",
          cursor: "pointer",
          fontWeight: "600",
          fontSize: "13px",
        },
      },
      "Done",
    ),
  );
};

const ModeSwitch = ({ mode, setMode }) => {
  const btn = (active) => ({
    flex: 1,
    padding: "8px",
    border: active ? "1px solid #000" : "1px solid #e5e5e5",
    background: active ? "#000" : "#fff",
    color: active ? "#fff" : "#666",
    cursor: "pointer",
    fontSize: "12px",
    fontWeight: "500",
    transition: "0.2s",
    fontFamily: fontStack,
  });
  return React.createElement(
    "div",
    { style: { display: "flex", padding: "15px 16px 0", gap: "8px" } },
    React.createElement(
      "button",
      { style: { ...btn(mode === "ASK"), borderRadius: "6px" }, onClick: () => setMode("ASK") },
      "Ask",
    ),
    React.createElement(
      "button",
      { style: { ...btn(mode === "REDLINE"), borderRadius: "6px" }, onClick: () => setMode("REDLINE") },
      "Redline",
    ),
  );
};

const Chat = ({ messages }) => {
  const bottomRef = React.useRef(null);
  React.useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);
  return React.createElement(
    "div",
    {
      style: {
        flexGrow: 1,
        overflowY: "auto",
        padding: "20px",
        display: "flex",
        flexDirection: "column",
        gap: "16px",
        fontFamily: fontStack,
      },
    },
    messages.map((msg, i) =>
      React.createElement(
        "div",
        { key: i, style: { alignSelf: msg.role === "user" ? "flex-end" : "flex-start", maxWidth: "90%" } },
        React.createElement("div", {
          style: {
            background: msg.role === "user" ? "#000" : "#f4f4f4",
            color: msg.role === "user" ? "#fff" : "#111",
            padding: "10px 14px",
            borderRadius: "8px",
            fontSize: "13px",
            lineHeight: "1.5",
          },
          dangerouslySetInnerHTML: { __html: renderMarkdown(msg.content) },
        }),
      ),
    ),
    React.createElement("div", { ref: bottomRef }),
  );
};

// ==========================================
// 4. MAIN APP LOGIC
// ==========================================

const App = () => {
  const [view, setView] = React.useState("main");
  const [mode, setMode] = React.useState("ASK");
  const [apiKey, setApiKey] = React.useState("");
  const [selectedModel, setSelectedModel] = React.useState("gemini-1.5-flash");
  const [messages, setMessages] = React.useState([{ role: "bot", content: "Ready." }]);
  const [inputValue, setInputValue] = React.useState("");
  const [isProcessing, setIsProcessing] = React.useState(false);

  React.useEffect(() => {
    const k = localStorage.getItem("vibe_api_key");
    const m = localStorage.getItem("vibe_model");
    if (k) setApiKey(k);
    if (m) setSelectedModel(m);
  }, []);

  const handleAction = async () => {
    if (!apiKey) return setView("settings");
    if (!inputValue.trim()) return;

    const currentInput = inputValue;
    setInputValue("");
    setMessages((p) => [...p, { role: "user", content: currentInput }]);
    setIsProcessing(true);

    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        if (mode === "ASK") {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          const answer = await callGemini(
            apiKey,
            selectedModel,
            `CONTEXT:\n"${body.text.substring(0, 30000)}"\n\nQUESTION:\n${currentInput}`,
            "ASK",
          );
          setMessages((p) => [...p, { role: "bot", content: answer }]);
        } else {
          selection.load("text");
          await context.sync();
          const originalText = selection.text;

          if (!originalText || originalText.trim().length === 0) {
            const draft = await callGemini(apiKey, selectedModel, currentInput, "DRAFT");
            context.document.changeTrackingMode = "TrackAll";
            selection.insertText(draft, "Replace");
            await context.sync();
            context.document.changeTrackingMode = "Off";
            setMessages((p) => [...p, { role: "bot", content: "Draft inserted." }]);
          } else {
            const oxmlResult = selection.getOoxml();
            await context.sync();
            const modifiedText = await callGemini(
              apiKey,
              selectedModel,
              `INSTRUCTION: ${currentInput}\n\nORIGINAL TEXT:\n${originalText}`,
              "REDLINE",
            );
            const newOxml = applyRedlineToOxml(oxmlResult.value, originalText, modifiedText);
            selection.insertOoxml(newOxml, "Replace");
            await context.sync();
            setMessages((p) => [...p, { role: "bot", content: "Redline applied." }]);
          }
        }
      });
    } catch (e) {
      console.error(e);
      setMessages((p) => [...p, { role: "bot", content: "Error: " + e.message }]);
    } finally {
      setIsProcessing(false);
    }
  };

  if (view === "settings")
    return React.createElement(Settings, {
      apiKey,
      setApiKey,
      selectedModel,
      setSelectedModel,
      onBack: () => setView("main"),
    });

  return React.createElement(
    "div",
    {
      className: "oscar-chat-layout",
      style: { fontFamily: fontStack, height: "100vh", display: "flex", flexDirection: "column", background: "#fff" },
    },
    React.createElement(Header, { onSettingsClick: () => setView("settings") }),
    React.createElement(ModeSwitch, { mode, setMode }),
    React.createElement(Chat, { messages }),
    React.createElement(
      "div",
      {
        style: {
          padding: "16px 20px",
          borderTop: "1px solid #f0f0f0",
          display: "flex",
          gap: "10px",
          background: "#fff",
          alignItems: "flex-end",
        },
      },
      React.createElement("textarea", {
        placeholder: mode === "ASK" ? "Ask about document..." : "Type instruction...",
        value: inputValue,
        onChange: (e) => setInputValue(e.target.value),
        onKeyDown: (e) => {
          if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            handleAction();
          }
        },
        style: {
          flexGrow: 1,
          padding: "12px",
          border: "1px solid #ddd",
          borderRadius: "8px",
          resize: "none",
          height: "48px",
          fontSize: "13px",
          outline: "none",
          fontFamily: "inherit",
        },
      }),
      React.createElement(
        "button",
        {
          onClick: handleAction,
          disabled: isProcessing,
          style: {
            width: "48px",
            height: "48px",
            background: isProcessing ? "#eee" : "#000",
            color: "#fff",
            border: "none",
            borderRadius: "8px",
            cursor: "pointer",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          },
        },
        isProcessing ? "..." : "➤",
      ),
    ),
  );
};

Office.onReady(() => ReactDOM.render(React.createElement(App), document.getElementById("root")));
