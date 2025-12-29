/**
 * App Component for Vibe Legal
 * Root component with context providers
 */

import React, { useState, useEffect } from 'react';
import { SettingsProvider } from '../state/SettingsContext';
import { DocumentProvider, useDocument } from '../state/DocumentContext';
import { ChatProvider, useChat } from '../state/ChatContext';
import { useSettings } from '../state/SettingsContext';
import { Toolbar } from './Toolbar';
import { FocusIndicator } from './FocusIndicator';
import { ChatArea } from './ChatArea';
import { PreviewPanel, generateChangeSummary } from './PreviewPanel';
import { InputArea } from './InputArea';
import { Settings } from './Settings';
import { Definitions } from './Definitions';
import { executeOperations } from '../services/handleAction';
import { analyzeDocument, DocumentAnalysis, DetectedParties, DefinedTerm, MindMapTopic } from '../services/documentAnalysis';
import { SideState } from './SideSelector';
import { DealContextState } from '../prompts/systemPrompt';
import { Negotiate } from './Negotiate';
import { MindMap } from './MindMap';
import { RiskTolerancePanel } from './RiskTolerancePanel';
import { NegotiateState, initialNegotiateState } from '../types/negotiate';
import { RiskTolerance, initialRiskTolerance } from '../types/state';
import './App.css';

declare var Word: any;

function AppContent() {
    const {
        pendingOperations,
        previewOriginalTexts,
        setPendingPreview,
        clearPendingOperations,
        addMessage,
        setIsProcessing
    } = useChat();
    const { getAuthorName, getCurrentApiKey, getCurrentModel } = useSettings();
    const { contractMap } = useDocument();

    // Track accepted/rejected indices for preview panel
    const [acceptedIndices, setAcceptedIndices] = useState<Set<number>>(new Set());
    const [rejectedIndices, setRejectedIndices] = useState<Set<number>>(new Set());
    // Track change summaries for final message
    const [changeSummaries, setChangeSummaries] = useState<{ index: number; summary: string; action: 'accepted' | 'rejected' }[]>([]);

    // Party/Side selector state
    const [selectedSide, setSelectedSide] = useState<SideState>({ selected: 'neutral' });
    const [sidesSelectorOpen, setSidesSelectorOpen] = useState(false);
    const [documentAnalysis, setDocumentAnalysis] = useState<DocumentAnalysis | null>(null);
    const [partiesLoading, setPartiesLoading] = useState(false);

    // Derived state for backwards compatibility
    const detectedParties: DetectedParties | null = documentAnalysis?.parties || null;
    const definedTerms: DefinedTerm[] = documentAnalysis?.definitions || [];

    // Deal Context state
    const [dealContext, setDealContext] = useState<string>('');
    const [contextOpen, setContextOpen] = useState(false);

    // Definitions state
    const [definitionsOpen, setDefinitionsOpen] = useState(false);

    // Mind Map state
    const [mindMapOpen, setMindMapOpen] = useState(false);
    const mindMapTopics: MindMapTopic[] = documentAnalysis?.mindMap || [];

    // Settings state
    const [settingsOpen, setSettingsOpen] = useState(false);

    // Negotiate state
    const [negotiateState, setNegotiateState] = useState<NegotiateState>(initialNegotiateState);

    // Risk Tolerance state (ADR-012)
    const [riskTolerance, setRiskTolerance] = useState<RiskTolerance>(initialRiskTolerance);
    const [riskPanelOpen, setRiskPanelOpen] = useState(false);

    const handleRiskToleranceUpdate = (updates: Partial<RiskTolerance>) => {
        setRiskTolerance(prev => ({ ...prev, ...updates }));
    };

    // Auto-disable Risk Tolerance when Sides is disabled
    React.useEffect(() => {
        if (selectedSide.selected === 'neutral' && riskTolerance.enabled) {
            setRiskTolerance(prev => ({ ...prev, enabled: false }));
        }
    }, [selectedSide.selected]);

    const handleNegotiateStateChange = (updates: Partial<NegotiateState> | ((prev: NegotiateState) => Partial<NegotiateState>)) => {
        setNegotiateState(prev => {
            const newUpdates = typeof updates === 'function' ? updates(prev) : updates;
            return { ...prev, ...newUpdates };
        });
    };

    const openNegotiate = () => {
        setNegotiateState({ ...initialNegotiateState, isOpen: true });
    };

    const closeNegotiate = () => {
        setNegotiateState(prev => ({ ...prev, isOpen: false }));
    };

    // Build context state object
    const contextState: DealContextState = {
        description: dealContext,
        isActive: dealContext.trim().length > 0
    };

    // Build document text for passing to SideSelector
    const documentText = contractMap?.paragraphs?.map(p => p.text).join('\n') || '';
    const apiKey = getCurrentApiKey();
    const model = getCurrentModel();

    // Handle Settings saved - load document and analyze (combined parties + definitions)
    const handleSettingsSaved = async (newApiKey: string) => {
        console.log('[App] Settings saved - loading document and analyzing...');
        console.log('[App] API key received:', !!newApiKey);

        try {
            // Step 1: Load document from Word
            let fullText = '';

            await Word.run(async (context: any) => {
                const body = context.document.body;
                const paragraphs = body.paragraphs;
                paragraphs.load("items");
                await context.sync();

                for (let i = 0; i < paragraphs.items.length; i++) {
                    paragraphs.items[i].load("text");
                }
                await context.sync();

                fullText = paragraphs.items.map((p: any) => p.text).join('\n');
                console.log('[App] Document loaded, length:', fullText.length);
            });

            // Step 2: Analyze document (parties + definitions) if we have enough text
            if (fullText.length > 100 && newApiKey) {
                console.log('[App] Starting document analysis...');
                setPartiesLoading(true);

                const analysis = await analyzeDocument(fullText, newApiKey, model);

                if (analysis) {
                    console.log('[App] Analysis complete:', analysis.parties.partyA.shortName, 'vs', analysis.parties.partyB.shortName);
                    console.log('[App] Definitions found:', analysis.definitions.length);
                    setDocumentAnalysis(analysis);
                } else {
                    console.log('[App] Could not analyze document');
                    setDocumentAnalysis(null);
                }

                setPartiesLoading(false);
            } else {
                console.log('[App] Document too short or no API key');
            }
        } catch (e) {
            console.error('[App] handleSettingsSaved error:', e);
            setPartiesLoading(false);
        }
    };

    const showPreview = pendingOperations.length > 0;

    // Build preview data for PreviewPanel
    const previewData = pendingOperations.length > 0
        ? {
            intent: 'MODIFY',
            operations: pendingOperations,
            originalTexts: previewOriginalTexts
        }
        : null;

    // Accept ONE operation
    const handleAcceptOne = async (index: number) => {
        const op = pendingOperations[index];
        if (!op) {
            console.error('[Accept] No operation at index:', index);
            return;
        }

        // Generate summary for this operation
        const original = op.original_text || previewOriginalTexts?.[op.target_id || 0] || '';
        const modified = (op as any).amended_text || (op as any).content || '';
        const summary = generateChangeSummary(op as any, original, modified);

        console.log('[Accept] Accepting operation index:', index);

        setIsProcessing(true);
        try {
            const author = getAuthorName();
            const result = await executeOperations([op], author);

            if (result.successCount > 0) {
                setAcceptedIndices(prev => new Set(prev).add(index));
                setChangeSummaries(prev => [...prev, { index, summary, action: 'accepted' }]);
            } else {
                console.error('[Accept] No changes applied - errorCount:', result.errorCount);
            }
        } catch (error: any) {
            console.error('[Accept] FAILED:', error);
            addMessage({
                role: 'bot',
                content: `Failed to apply change: ${error.message}`
            });
        } finally {
            setIsProcessing(false);
        }
    };

    // Reject ONE operation
    const handleRejectOne = (index: number) => {
        const op = pendingOperations[index];
        if (!op) return;

        // Generate summary for this operation
        const original = op.original_text || previewOriginalTexts?.[op.target_id || 0] || '';
        const modified = (op as any).amended_text || (op as any).content || '';
        const summary = generateChangeSummary(op as any, original, modified);

        setRejectedIndices(prev => new Set(prev).add(index));
        setChangeSummaries(prev => [...prev, { index, summary, action: 'rejected' }]);
    };

    // Accept ALL remaining operations
    const handleAcceptAll = async () => {
        const remainingOps = pendingOperations.filter((_, i) =>
            !acceptedIndices.has(i) && !rejectedIndices.has(i)
        );

        if (remainingOps.length === 0) {
            finishReview();
            return;
        }

        // Generate summaries for all remaining
        const newSummaries = remainingOps.map((op, i) => {
            const actualIndex = pendingOperations.indexOf(op);
            const original = op.original_text || previewOriginalTexts?.[op.target_id || 0] || '';
            const modified = (op as any).amended_text || (op as any).content || '';
            return { index: actualIndex, summary: generateChangeSummary(op as any, original, modified), action: 'accepted' as const };
        });

        setIsProcessing(true);
        try {
            const author = getAuthorName();
            const result = await executeOperations(remainingOps, author);

            setChangeSummaries(prev => [...prev, ...newSummaries]);
            finishReview(newSummaries.length, 0);
        } catch (error: any) {
            addMessage({
                role: 'bot',
                content: `Failed to apply changes: ${error.message}`
            });
        } finally {
            setIsProcessing(false);
        }
    };

    // Reject ALL remaining operations
    const handleRejectAll = () => {
        const remainingOps = pendingOperations.filter((_, i) =>
            !acceptedIndices.has(i) && !rejectedIndices.has(i)
        );

        const newSummaries = remainingOps.map((op) => {
            const actualIndex = pendingOperations.indexOf(op);
            const original = op.original_text || previewOriginalTexts?.[op.target_id || 0] || '';
            const modified = (op as any).amended_text || (op as any).content || '';
            return { index: actualIndex, summary: generateChangeSummary(op as any, original, modified), action: 'rejected' as const };
        });

        setChangeSummaries(prev => [...prev, ...newSummaries]);
        finishReview(0, newSummaries.length);
    };

    // Finish review and clean up - generate summary message
    const finishReview = (bulkAccepted = 0, bulkRejected = 0) => {
        // Combine tracked summaries with bulk actions
        const accepted = changeSummaries.filter(s => s.action === 'accepted');
        const rejected = changeSummaries.filter(s => s.action === 'rejected');
        const totalAccepted = accepted.length + bulkAccepted;
        const totalRejected = rejected.length + bulkRejected;

        // Build clean summary message without markdown bullets
        let summaryLines: string[] = [];
        summaryLines.push('**Review Complete**');
        summaryLines.push('');

        if (totalAccepted > 0) {
            summaryLines.push(`**Applied (${totalAccepted}):**`);
            accepted.forEach(s => summaryLines.push(`   ${s.summary}`));
            if (bulkAccepted > 0 && accepted.length === 0) {
                summaryLines.push(`   ${bulkAccepted} change(s) applied`);
            }
        }

        if (totalRejected > 0) {
            if (totalAccepted > 0) summaryLines.push('');
            summaryLines.push(`**Discarded (${totalRejected}):**`);
            rejected.forEach(s => summaryLines.push(`   ${s.summary}`));
            if (bulkRejected > 0 && rejected.length === 0) {
                summaryLines.push(`   ${bulkRejected} change(s) discarded`);
            }
        }

        if (totalAccepted === 0 && totalRejected === 0) {
            summaryLines.push('No changes were made.');
        }

        addMessage({
            role: 'bot',
            content: summaryLines.join('\n')
        });

        // Clear state
        clearPendingOperations();
        setPendingPreview(null);
        setAcceptedIndices(new Set());
        setRejectedIndices(new Set());
        setChangeSummaries([]);
    };

    // Auto-finish removed - we now let the user manually close via the Done button
    // This allows them to see the final summary before the panel disappears

    const handleClosePreview = () => {
        finishReview();
    };

    return (
        <div className="app-container">
            <header className="app-header">
                <h1 className="app-header__title">Vibe Legal</h1>

                <Toolbar
                    selectedSide={selectedSide}
                    onSideChange={setSelectedSide}
                    detectedParties={detectedParties}
                    partiesLoading={partiesLoading}
                    sidesSelectorOpen={sidesSelectorOpen}
                    onSidesSelectorChange={setSidesSelectorOpen}
                    contextOpen={contextOpen}
                    setContextOpen={setContextOpen}
                    dealContext={dealContext}
                    setDealContext={setDealContext}
                    contextState={contextState}
                    definitionsOpen={definitionsOpen}
                    setDefinitionsOpen={setDefinitionsOpen}
                    definedTerms={definedTerms}
                    onNegotiateOpen={openNegotiate}
                    negotiateHasHistory={negotiateState.messages.length > 0}
                    mindMapOpen={mindMapOpen}
                    setMindMapOpen={setMindMapOpen}
                    mindMapTopicCount={mindMapTopics.length}
                    riskTolerance={riskTolerance}
                    riskPanelOpen={riskPanelOpen}
                    onRiskPanelOpen={() => setRiskPanelOpen(true)}
                />
            </header>

            <FocusIndicator />

            <main className="app-main">
                <ChatArea />
                {showPreview && (
                    <PreviewPanel
                        preview={previewData}
                        onAcceptOne={handleAcceptOne}
                        onRejectOne={handleRejectOne}
                        onAcceptAll={handleAcceptAll}
                        onRejectAll={handleRejectAll}
                        acceptedIndices={acceptedIndices}
                        rejectedIndices={rejectedIndices}
                        reviewedSummaries={changeSummaries}
                        onClose={handleClosePreview}
                    />
                )}
            </main>

            <InputArea
                selectedSide={selectedSide}
                detectedParties={detectedParties}
                dealContext={contextState}
                riskTolerance={riskTolerance}
                onOpenSettings={() => setSettingsOpen(true)}
            />

            {/* Legal Disclaimer */}
            <div className="disclaimer">
                <p className="disclaimer__text">
                    This is not legal advice. AI can make mistakes. Always verify important information.
                </p>
            </div>

            <Negotiate
                state={negotiateState}
                onClose={closeNegotiate}
                onStateChange={handleNegotiateStateChange}
                apiKey={apiKey}
                dealContext={contextState}
                detectedParties={detectedParties}
            />
            <MindMap
                isOpen={mindMapOpen}
                onClose={() => setMindMapOpen(false)}
                topics={mindMapTopics}
                apiKey={apiKey}
                documentText={documentText}
                dealContext={dealContext}
                selectedSide={selectedSide.selected !== 'neutral' ? (selectedSide.selected === 'partyA' ? detectedParties?.partyA?.shortName : detectedParties?.partyB?.shortName) : null}
            />
            <Settings
                onSettingsSaved={handleSettingsSaved}
                isOpen={settingsOpen}
                onClose={() => setSettingsOpen(false)}
            />
            <RiskTolerancePanel
                isOpen={riskPanelOpen}
                onClose={() => setRiskPanelOpen(false)}
                riskTolerance={riskTolerance}
                selectedSide={selectedSide.selected === 'partyA'
                    ? detectedParties?.partyA?.shortName || 'Party A'
                    : selectedSide.selected === 'partyB'
                        ? detectedParties?.partyB?.shortName || 'Party B'
                        : null}
                onUpdate={handleRiskToleranceUpdate}
                isDisabled={selectedSide.selected === 'neutral'}
                onOpenSides={() => setSidesSelectorOpen(true)}
            />
            <Definitions
                isOpen={definitionsOpen}
                onClose={() => setDefinitionsOpen(false)}
                terms={definedTerms}
            />
        </div>
    );
}

export function App() {
    return (
        <SettingsProvider>
            <DocumentProvider>
                <ChatProvider>
                    <AppContent />
                </ChatProvider>
            </DocumentProvider>
        </SettingsProvider>
    );
}

export default App;
