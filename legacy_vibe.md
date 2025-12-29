// ==========================================
// VIBE LEGAL 3.3 - RESTORED VISUALS + LOGIC PATCH
// Based on V3.2 Hybrid Architecture
// ==========================================

declare var diff_match_patch: any;
declare var React: any;
declare var ReactDOM: any;

// ==========================================
// 1. UTILS: DIFF ENGINE
// ==========================================

function calculateRedline(original: string, modified: string): Array<{ op: number; text: string }> {
  const dmp = new diff_match_patch();
  const diffs = dmp.diff_main(original, modified);
  dmp.diff_cleanupSemantic(diffs);
  return diffs.map((diff) => ({ op: diff[0], text: diff[1] }));
}

// ==========================================
// 2. CORE: ROBUST OXML ENGINE (V3.0 Logic)
// ==========================================

function applyRedlineToOxml(oxml: string, originalText: string, modifiedText: string): string {
  const parser = new DOMParser();
  const xmlSerializer = new XMLSerializer();
  let xmlDoc;

  try {
    xmlDoc = parser.parseFromString(oxml, "text/xml");
  } catch (e) {
    console.error("OXML Parse Error:", e);
    return oxml;
  }

  if (xmlDoc.getElementsByTagName("parsererror").length > 0) return oxml;

  // 1. EXTRACT PARAGRAPHS
  const body = xmlDoc.getElementsByTagName("w:body")[0] || xmlDoc.documentElement;
  const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));

  if (paragraphs.length === 0) return oxml;

  // 2. BUILD CONTEXT & MAPS
  let originalFullText = "";
  const propertyMap: { start: number; end: number; rPr: Element | null; wrapper?: Node }[] = [];
  const paragraphMap: { start: number; end: number; pPr: Element | null; container: Node }[] = [];
  const sentinelMap: { start: number; node: Node; isTextBox?: boolean; originalContainer?: Node }[] = [];
  const referenceMap = new Map<string, Node>();
  const tokenToCharMap = new Map<string, string>();
  let nextCharCode = 0xe000;

  const uniqueContainers = new Set<Node>();
  const replacementContainers = new Map<Node, Node>();

  paragraphs.forEach((p, pIndex) => {
    const pStart = originalFullText.length;
    const children = Array.from(p.childNodes);

    children.forEach((child) => {
      if (child.nodeName === "w:r") {
        const r = child as Element;
        const rPr = r.getElementsByTagName("w:rPr")[0] || null;
        const runChildren = Array.from(r.childNodes);

        runChildren.forEach((rc) => {
          if (rc.nodeName === "w:t") {
            const textContent = rc.textContent || "";
            if (textContent.length > 0) {
              propertyMap.push({
                start: originalFullText.length,
                end: originalFullText.length + textContent.length,
                rPr: rPr,
              });
              originalFullText += textContent;
            }
          } else if (["w:drawing", "w:pict", "w:object", "w:fldChar", "w:instrText"].includes(rc.nodeName)) {
            // SENTINEL LOGIC
            const rcElement = rc as Element;
            const txbxContent = rcElement.getElementsByTagName("w:txbxContent")[0];
            const hasTextBox = rc.nodeName === "w:pict" && !!txbxContent;

            if (hasTextBox) {
              sentinelMap.push({
                start: originalFullText.length,
                node: rc,
                isTextBox: true,
                originalContainer: txbxContent,
              });
              originalFullText += "\uFFFC";
              propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
            } else {
              sentinelMap.push({ start: originalFullText.length, node: rc });
              originalFullText += "\uFFFC";
              propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
            }
          } else if (rc.nodeName === "w:footnoteReference" || rc.nodeName === "w:endnoteReference") {
            // REFERENCE LOGIC
            const ref = rc as Element;
            const id = ref.getAttribute("w:id");
            if (id) {
              const type = rc.nodeName === "w:footnoteReference" ? "FN" : "EN";
              const tokenString = `{{__${type}_${id}__}}`;
              const char = String.fromCharCode(nextCharCode++);
              referenceMap.set(char, rc);
              tokenToCharMap.set(tokenString, char);
              originalFullText += char;
              propertyMap.push({ start: originalFullText.length - 1, end: originalFullText.length, rPr: rPr });
            }
          }
        });
      } else if (child.nodeName === "w:hyperlink") {
        const h = child as Element;
        const hChildren = Array.from(h.childNodes);
        hChildren.forEach((hc) => {
          if (hc.nodeName === "w:r") {
            const r = hc as Element;
            const rPr = r.getElementsByTagName("w:rPr")[0] || null;
            const texts = Array.from(r.getElementsByTagName("w:t"));
            texts.forEach((t) => {
              const textContent = t.textContent || "";
              if (textContent.length > 0) {
                propertyMap.push({
                  start: originalFullText.length,
                  end: originalFullText.length + textContent.length,
                  rPr: rPr,
                  wrapper: h,
                });
                originalFullText += textContent;
              }
            });
          }
        });
      } else if (["w:sdt", "w:oMath", "m:oMath", "w:bookmarkStart", "w:bookmarkEnd"].includes(child.nodeName)) {
        sentinelMap.push({ start: originalFullText.length, node: child });
        originalFullText += "\uFFFC";
      }
    });

    if (pIndex < paragraphs.length - 1) {
      originalFullText += "\n";
    }

    const pEnd = originalFullText.length;
    const pPr = p.getElementsByTagName("w:pPr")[0] || null;
    const container = p.parentNode;
    if (container) uniqueContainers.add(container);

    paragraphMap.push({
      start: pStart,
      end: pEnd,
      pPr: pPr,
      container: container || body,
    });
  });

  // 3. SANITIZE & PROCESS INPUT
  const cleanAiResponse = (text: string): string => {
    let cleaned = text;
    cleaned = cleaned.replace(
      /^(Here is the redline:|Here is the text:|Sure, I can help:|Here's the updated text:)\s*/i,
      "",
    );
    cleaned = cleaned.replace(/\$\\text\{/g, "").replace(/\}\$/g, "");
    cleaned = cleaned.replace(/\$([^0-9\n]+?)\$/g, "$1");
    return cleaned;
  };

  const cleanModifiedText = cleanAiResponse(modifiedText);
  let processedModifiedText = cleanModifiedText;

  tokenToCharMap.forEach((char, tokenString) => {
    const escapedToken = tokenString.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&");
    processedModifiedText = processedModifiedText.replace(new RegExp(escapedToken, "g"), char);
  });

  // 4. DIFF CALCULATION
  const dmp = new diff_match_patch();
  const diffs = dmp.diff_main(originalFullText, processedModifiedText);
  dmp.diff_cleanupSemantic(diffs);

  // 5. RECONSTRUCTION
  const containerFragments = new Map<Node, DocumentFragment>();
  uniqueContainers.forEach((c) => containerFragments.set(c, xmlDoc.createDocumentFragment()));
  if (!containerFragments.has(body)) containerFragments.set(body, xmlDoc.createDocumentFragment());

  const getParagraphInfo = (index: number): { pPr: Element | null; container: Node } => {
    const match = paragraphMap.find((m) => index >= m.start && index < m.end);
    if (!match && paragraphMap.length > 0) {
      const last = paragraphMap[paragraphMap.length - 1];
      return { pPr: last.pPr, container: last.container };
    }
    return match ? { pPr: match.pPr, container: match.container } : { pPr: null, container: body };
  };

  const createNewParagraph = (pPr: Element | null) => {
    const newP = xmlDoc.createElement("w:p");
    if (pPr) newP.appendChild(pPr.cloneNode(true));
    return newP;
  };

  let startInfo = getParagraphInfo(0);
  let currentParagraph = createNewParagraph(startInfo.pPr);
  let currentContainer = startInfo.container;
  let currentFragment = containerFragments.get(currentContainer);
  if (currentFragment) currentFragment.appendChild(currentParagraph);

  let currentOriginalIndex = 0;
  const usedReferenceTokens = new Set<string>();

  const getRunProperties = (index: number): { rPr: Element | null; wrapper?: Node } => {
    const match = propertyMap.find((m) => index >= m.start && index < m.end);
    return match ? { rPr: match.rPr, wrapper: match.wrapper } : { rPr: null };
  };

  const appendTextToCurrent = (
    text: string,
    type: "equal" | "insert" | "delete",
    rPr: Element | null,
    wrapper: Node | undefined,
    baseIndex: number,
  ) => {
    const parts = text.split(/([\n\uFFFC]|[\uE000-\uF8FF])/);
    let localOffset = 0;

    parts.forEach((part) => {
      if (part === "\n") {
        if (type !== "delete") {
          let pPr: Element | null = null;
          let targetContainer: Node = currentContainer;

          if (type === "equal") {
            const lookupIndex = baseIndex + localOffset + 1;
            const info = getParagraphInfo(lookupIndex);
            pPr = info.pPr;
            targetContainer = info.container;
          } else {
            const info = getParagraphInfo(baseIndex);
            pPr = info.pPr;
            targetContainer = info.container;
          }

          currentParagraph = createNewParagraph(pPr);
          if (targetContainer !== currentContainer) {
            currentContainer = targetContainer;
            currentFragment = containerFragments.get(currentContainer);
          }
          if (currentFragment) currentFragment.appendChild(currentParagraph);
        }
        localOffset += 1;
      } else if (part === "\uFFFC") {
        const sentinelIndex = baseIndex + localOffset;
        const sentinel = sentinelMap.find((s) => s.start === sentinelIndex);

        if (sentinel) {
          const clone = sentinel.node.cloneNode(true) as Element;
          if (sentinel.isTextBox && sentinel.originalContainer) {
            const newContainer = clone.getElementsByTagName("w:txbxContent")[0];
            if (newContainer) {
              while (newContainer.firstChild) newContainer.removeChild(newContainer.firstChild);
              replacementContainers.set(sentinel.originalContainer, newContainer);
            }
          }
          if (sentinel.node.nodeName === "w:r" || sentinel.node.parentNode?.nodeName === "w:r") {
            const run = xmlDoc.createElement("w:r");
            if (rPr) run.appendChild(rPr.cloneNode(true));
            run.appendChild(clone);
            currentParagraph.appendChild(run);
          } else {
            currentParagraph.appendChild(clone);
          }
        }
        localOffset += 1;
      } else if (referenceMap.has(part)) {
        if (type !== "delete") {
          const refNode = referenceMap.get(part);
          if (refNode) {
            const clone = refNode.cloneNode(true);
            const run = xmlDoc.createElement("w:r");
            if (rPr) run.appendChild(rPr.cloneNode(true));
            run.appendChild(clone);
            currentParagraph.appendChild(run);
            let tokenString = "";
            tokenToCharMap.forEach((c, t) => {
              if (c === part) tokenString = t;
            });
            usedReferenceTokens.add(tokenString);
          }
        }
        localOffset += part.length;
      } else if (part.length > 0) {
        const run = xmlDoc.createElement("w:r");
        if (rPr) run.appendChild(rPr.cloneNode(true));

        let parent: Node = currentParagraph;
        if (wrapper) {
          const wrapperClone = wrapper.cloneNode(false);
          parent = wrapperClone;
          currentParagraph.appendChild(wrapperClone);
        }

        const t = type === "delete" ? xmlDoc.createElement("w:delText") : xmlDoc.createElement("w:t");
        t.setAttribute("xml:space", "preserve");
        t.textContent = part;

        run.appendChild(t);

        if (type === "delete") {
          const del = xmlDoc.createElement("w:del");
          del.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
          del.setAttribute("w:author", "Vibe");
          del.setAttribute("w:date", new Date().toISOString());
          del.appendChild(run);
          parent.appendChild(del);
        } else if (type === "insert") {
          const ins = xmlDoc.createElement("w:ins");
          ins.setAttribute("w:id", Math.floor(Math.random() * 10000).toString());
          ins.setAttribute("w:author", "Vibe");
          ins.setAttribute("w:date", new Date().toISOString());
          ins.appendChild(run);
          parent.appendChild(ins);
        } else {
          parent.appendChild(run);
        }
        localOffset += part.length;
      }
    });
  };

  for (const diff of diffs) {
    const [op, text] = diff;
    if (op === 0) {
      let offset = 0;
      while (offset < text.length) {
        const props = getRunProperties(currentOriginalIndex + offset);
        const currentPropRange = propertyMap.find(
          (m) => currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end,
        );
        let length = currentPropRange
          ? Math.min(currentPropRange.end - (currentOriginalIndex + offset), text.length - offset)
          : 1;
        const chunk = text.substring(offset, offset + length);
        appendTextToCurrent(chunk, "equal", props.rPr || null, props.wrapper, currentOriginalIndex + offset);
        offset += length;
      }
      currentOriginalIndex += text.length;
    } else if (op === 1) {
      const isStartOfParagraph = paragraphMap.some((p) => p.start === currentOriginalIndex);
      const props =
        currentOriginalIndex > 0 && !isStartOfParagraph
          ? getRunProperties(currentOriginalIndex - 1)
          : getRunProperties(currentOriginalIndex);
      appendTextToCurrent(text, "insert", props.rPr || null, props.wrapper, currentOriginalIndex);
    } else if (op === -1) {
      let offset = 0;
      while (offset < text.length) {
        const props = getRunProperties(currentOriginalIndex + offset);
        const currentPropRange = propertyMap.find(
          (m) => currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end,
        );
        let length = currentPropRange
          ? Math.min(currentPropRange.end - (currentOriginalIndex + offset), text.length - offset)
          : 1;
        const chunk = text.substring(offset, offset + length);
        appendTextToCurrent(chunk, "delete", props.rPr || null, props.wrapper, currentOriginalIndex + offset);
        offset += length;
      }
      currentOriginalIndex += text.length;
    }
  }

  tokenToCharMap.forEach((char, token) => {
    if (!usedReferenceTokens.has(token)) console.warn(`[AUDIT] Reference Deleted: ${token}`);
  });

  paragraphs.forEach((p) => {
    if (p.parentNode) p.parentNode.removeChild(p);
  });

  containerFragments.forEach((fragment, container) => {
    const replacement = replacementContainers.get(container);
    const target = replacement || container;
    target.appendChild(fragment);
  });

  return xmlSerializer.serializeToString(xmlDoc);
}

// ==========================================
// 3. NETWORK LAYER
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
    .filter((m: any) => m.supportedGenerationMethods?.includes("generateContent"))
    .sort((a: any, b: any) => b.name.localeCompare(a.name));
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

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || "Gemini API Error");
  }
  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "";
}

const renderMarkdown = (text: string) => {
  if (!text) return "";
  return text
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.*?)\*/g, "<em>$1</em>")
    .replace(/\n/g, "<br>");
};

// ==========================================
// 4. UI COMPONENTS
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
      "h2",
      { style: { margin: 0, fontSize: "14px", fontWeight: "600", color: "#111", fontFamily: fontStack } },
      "Vibe Legal 3.3",
    ),
    React.createElement(
      "button",
      {
        onClick: onSettingsClick,
        style: { background: "none", border: "none", cursor: "pointer", fontSize: "16px", opacity: 0.7 },
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
      if (list.length > 0 && !selectedModel) setSelectedModel(list[0].name.replace("models/", ""));
    } catch (e) {
      setStatus("Error: Check Key");
    }
  };

  return React.createElement(
    "div",
    { style: { padding: "20px", background: "#f9f9f9", height: "100%", fontFamily: fontStack } },
    React.createElement("h3", { style: { marginTop: 0 } }, "Settings"),
    React.createElement("label", { style: { display: "block", fontWeight: "600", marginBottom: "5px" } }, "API Key"),
    React.createElement("input", {
      type: "password",
      value: apiKey,
      onChange: (e) => setApiKey(e.target.value),
      style: { width: "100%", padding: "8px", marginBottom: "15px" },
    }),
    React.createElement("label", { style: { display: "block", fontWeight: "600", marginBottom: "5px" } }, "Model"),
    React.createElement(
      "div",
      { style: { display: "flex", gap: "10px", marginBottom: "5px" } },
      React.createElement(
        "select",
        {
          value: selectedModel,
          onChange: (e) => setSelectedModel(e.target.value),
          style: { flexGrow: 1, padding: "8px" },
        },
        models.length > 0
          ? models.map((m) =>
              React.createElement(
                "option",
                { key: m.name, value: m.name.replace("models/", "") },
                m.displayName || m.name,
              ),
            )
          : React.createElement("option", { value: selectedModel }, selectedModel),
      ),
      React.createElement("button", { onClick: handleSync }, "↻"),
    ),
    React.createElement("div", { style: { fontSize: "11px", color: "#666", marginBottom: "20px" } }, status),
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
          padding: "10px",
          border: "none",
          borderRadius: "4px",
          cursor: "pointer",
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
    border: active ? "1px solid #000" : "1px solid #ddd",
    background: active ? "#000" : "#fff",
    color: active ? "#fff" : "#666",
    cursor: "pointer",
    fontWeight: "500",
  });
  return React.createElement(
    "div",
    { style: { display: "flex", padding: "10px 16px", gap: "10px" } },
    React.createElement("button", { style: btn(mode === "ASK"), onClick: () => setMode("ASK") }, "Ask"),
    React.createElement("button", { style: btn(mode === "REDLINE"), onClick: () => setMode("REDLINE") }, "Redline"),
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
        padding: "16px",
        display: "flex",
        flexDirection: "column",
        gap: "12px",
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
            padding: "10px",
            borderRadius: "8px",
            fontSize: "13px",
          },
          dangerouslySetInnerHTML: { __html: renderMarkdown(msg.content) },
        }),
      ),
    ),
    React.createElement("div", { ref: bottomRef }),
  );
};

// ==========================================
// 5. MAIN APP
// ==========================================

const App = () => {
  const [view, setView] = React.useState("main");
  const [mode, setMode] = React.useState("ASK");
  const [apiKey, setApiKey] = React.useState("");
  const [selectedModel, setSelectedModel] = React.useState("gemini-1.5-flash");
  const [messages, setMessages] = React.useState([{ role: "bot", content: "Ready. Select text to begin." }]);
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
        selection.load("parentBody");
        await context.sync();

        // FIX: Whitelist Tables
        if (
          selection.parentBody.type !== Word.BodyType.mainDoc &&
          selection.parentBody.type !== Word.BodyType.tableCell
        ) {
          setMessages((p) => [
            ...p,
            { role: "bot", content: "⚠️ Please select text inside the main document or a table." },
          ]);
          return;
        }

        // --- FIX 3: SEPARATION OF CONCERNS ---
        // Only flush changes if we are EDITING (Redline/Draft)
        if (mode !== "ASK") {
          const trackedChanges = selection.getTrackedChanges();
          trackedChanges.load("items");
          await context.sync();

          if (trackedChanges.items.length > 0) {
            trackedChanges.items.forEach((c) => c.accept());
            await context.sync();
            // Optional: Notify user or stay silent
            // setMessages((p) => [...p, { role: "bot", content: "<i>Accepted existing track changes.</i>" }]);
          }
        }

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
          // REDLINE or DRAFT
          selection.load("text");

          // HYBRID INTELLIGENCE: Load Visual Context (V3.2 Logic)
          const paragraphs = selection.paragraphs;
          paragraphs.load("items");
          await context.sync();

          let visualContext = "";
          if (paragraphs.items.length > 0) {
            paragraphs.load("items/style, items/text, items/font/bold");
            await context.sync();
            visualContext = "VISUAL CONTEXT (Preserve formatting):\n";
            paragraphs.items.forEach((p, i) => {
              const style = p.style;
              const textStart = (p.text || "").substring(0, 40).replace(/\n/g, " ");
              const isBold = p.font.bold ? "Bold" : "Normal";
              visualContext += `- Para ${i + 1}: Style='${style}', Font='${isBold}', Text='${textStart}...'\n`;
            });
          }

          const originalText = selection.text;

          if (!originalText || originalText.trim().length === 0) {
            // INSERT MODE
            const draft = await callGemini(apiKey, selectedModel, currentInput, "DRAFT");
            selection.insertText(draft, "Replace");
            await context.sync();
            setMessages((p) => [...p, { role: "bot", content: "Draft inserted." }]);
          } else {
            // REDLINE MODE (Hybrid: Context -> Text -> Diff)
            const oxmlResult = selection.getOoxml();
            await context.sync();

            const prompt = `
${visualContext}

INSTRUCTION: ${currentInput}

ORIGINAL TEXT:
${originalText}
            `;

            const modifiedText = await callGemini(apiKey, selectedModel, prompt, "REDLINE");

            // --- FIX 2: DETECT SILENT FAILURE ---
            if (!modifiedText || modifiedText.trim() === originalText.trim()) {
              setMessages((p) => [...p, { role: "bot", content: "⚠️ The AI suggested no changes." }]);
              return;
            }

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
