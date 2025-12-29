import React, { useState, useEffect } from 'react';
import Header from './Header';
import Settings from './Settings';
import Chat from './Chat';
import ModeSwitch from './ModeSwitch';
import { applyRedlineToOxml } from '../../utils/OxmlEngine';
import { runOxmlAgent } from '../../utils/Agent_Oxml_Redline';
import { callGemini } from '../../utils/geminiApi';

/* global console, localStorage, Word */

const App: React.FC = () => {
    const [view, setView] = useState<'chat' | 'settings'>('chat');
    const [mode, setMode] = useState<"ASK" | "REDLINE" | "DRAFT">("ASK");
    const [apiKey, setApiKey] = useState<string>('');
    const [selectedModel, setSelectedModel] = useState<string>("gemini-1.5-flash");
    const [messages, setMessages] = useState<Array<{ role: 'user' | 'bot', content: string }>>([{ role: "bot", content: "Ready." }]);
    const [inputValue, setInputValue] = useState('');
    const [isProcessing, setIsProcessing] = useState(false);

    useEffect(() => {
        const k = localStorage.getItem('vibe_api_key');
        const m = localStorage.getItem('vibe_model');
        if (k) setApiKey(k);
        if (m) setSelectedModel(m);
    }, []);

    const handleAction = async () => {
        if (!apiKey) {
            setView('settings');
            return;
        }
        if (!inputValue.trim()) return;

        const currentInput = inputValue;
        setInputValue('');
        setMessages(prev => [...prev, { role: 'user', content: currentInput }]);
        setIsProcessing(true);

        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("parentBody");
                await context.sync();

                // ==========================================
                // PHASE 0: PRE-FLIGHT SANITIZER
                // ==========================================

                // 1. Rejection Logic: Ensure we are in the Main Body or Table
                if (selection.parentBody.type !== Word.BodyType.mainDoc && selection.parentBody.type !== Word.BodyType.tableCell) {
                    setMessages(prev => [...prev, { role: 'bot', content: 'Please select text inside the main document body or a table.' }]);
                    return;
                }

                // 2. Track Changes Flush
                let targetRange: Word.Range | Word.Body = selection;
                if (mode === "ASK") {
                    targetRange = context.document.body;
                }

                // Check for tracked changes in the target range
                const trackedChanges = targetRange.getTrackedChanges();
                trackedChanges.load("items");
                await context.sync();

                if (trackedChanges.items.length > 0) {
                    // Accept all changes to ensure clean state
                    trackedChanges.items.forEach((change) => {
                        change.accept();
                    });
                    await context.sync();

                    setMessages(prev => [...prev, {
                        role: 'bot',
                        content: 'Note: Existing track changes were accepted to process this request.'
                    }]);
                }

                // ==========================================
                // CORE LOGIC
                // ==========================================

                if (mode === "ASK") {
                    const body = context.document.body;
                    body.load("text");
                    await context.sync();
                    const answer = await callGemini(
                        apiKey,
                        selectedModel,
                        `CONTEXT:\n"${body.text.substring(0, 30000)}"\n\nQUESTION:\n${currentInput}`,
                        "ASK"
                    );
                    setMessages(prev => [...prev, { role: 'bot', content: answer }]);
                } else {
                    // REDLINE or DRAFT
                    selection.load("text");
                    await context.sync();
                    const originalText = selection.text;

                    if (!originalText || originalText.trim().length === 0) {
                        // DRAFT MODE (Insert)
                        const draft = await callGemini(apiKey, selectedModel, currentInput, "DRAFT");
                        context.document.changeTrackingMode = "TrackAll";
                        selection.insertText(draft, "Replace");
                        await context.sync();
                        context.document.changeTrackingMode = "Off";
                        setMessages(prev => [...prev, { role: 'bot', content: "Draft inserted." }]);
                    } else {
                        // REDLINE MODE (Modify) - V3.1 Architecture

                        // 1. Extract Visual Context
                        const paragraphs = selection.paragraphs;
                        paragraphs.load("items, style, text, font/bold");
                        await context.sync();

                        let visualContext = "Visual Context Summary:\n";
                        paragraphs.items.forEach((p, i) => {
                            const style = p.style;
                            const textStart = p.text.substring(0, 50).replace(/\n/g, " ");
                            const isBold = p.font.bold ? "Bold" : "Normal";
                            visualContext += `Para ${i + 1}: Style='${style}', Font='${isBold}', StartsWith='${textStart}...'\n`;
                        });

                        // 2. Get OOXML (Checkpoint)
                        const oxmlResult = selection.getOoxml();
                        await context.sync();
                        const originalOxml = oxmlResult.value;

                        // 3. Run Agent with Checkpoint Protection
                        try {
                            const agentResult = await runOxmlAgent(
                                apiKey,
                                selectedModel,
                                originalOxml,
                                currentInput,
                                visualContext
                            );

                            if (agentResult.error) {
                                throw new Error(agentResult.error);
                            }

                            selection.insertOoxml(agentResult.oxml, "Replace");
                            await context.sync();
                            setMessages(prev => [...prev, { role: 'bot', content: "Redline applied (V3.1)." }]);

                        } catch (e) {
                            console.error("Agent failed, restoring checkpoint.", e);
                            // Restore Checkpoint
                            selection.insertOoxml(originalOxml, "Replace");
                            await context.sync();
                            setMessages(prev => [...prev, {
                                role: 'bot',
                                content: `Error: ${(e as Error).message}. Changes reverted.`
                            }]);
                        }
                    }
                }
            });
        } catch (error) {
            console.error(error);
            setMessages(prev => [...prev, { role: 'bot', content: 'Error: ' + (error as Error).message }]);
        } finally {
            setIsProcessing(false);
        }
    };

    return (
        <div className="oscar-chat-layout" style={{ height: "100vh", display: "flex", flexDirection: "column", background: "#fff" }}>
            <Header onSettingsClick={() => setView('settings')} />

            {view === 'settings' ? (
                <Settings
                    apiKey={apiKey}
                    setApiKey={setApiKey}
                    selectedModel={selectedModel}
                    setSelectedModel={setSelectedModel}
                    onBack={() => setView('chat')}
                />
            ) : (
                <>
                    <ModeSwitch mode={mode} setMode={setMode} />
                    <Chat messages={messages} />
                    <div style={{
                        padding: "16px 20px",
                        borderTop: "1px solid #f0f0f0",
                        display: "flex",
                        gap: "10px",
                        background: "#fff",
                        alignItems: "flex-end",
                    }}>
                        <textarea
                            placeholder={mode === "ASK" ? "Ask about document..." : "Type instruction..."}
                            value={inputValue}
                            onChange={(e) => setInputValue(e.target.value)}
                            onKeyDown={(e) => {
                                if (e.key === 'Enter' && !e.shiftKey) {
                                    e.preventDefault();
                                    handleAction();
                                }
                            }}
                            style={{
                                flexGrow: 1,
                                padding: "12px",
                                border: "1px solid #ddd",
                                borderRadius: "8px",
                                resize: "none",
                                height: "48px",
                                fontSize: "13px",
                                outline: "none",
                                fontFamily: "inherit",
                            }}
                        />
                        <button
                            onClick={handleAction}
                            disabled={isProcessing}
                            style={{
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
                            }}
                        >
                            {isProcessing ? '...' : '➤'}
                        </button>
                    </div>
                </>
            )}
        </div>
    );
};

export default App;
