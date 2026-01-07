/**
 * Negotiate Component
 * Modal for AI-powered debate prep
 */

import React, { useRef } from 'react';
import { NegotiateState, DebateMessage } from '../types/negotiate';
import { NegotiateSetup } from './NegotiateSetup';
import { DebateView } from './DebateView';
import { InteractiveInput } from './InteractiveInput';
import { NegotiateActions } from './NegotiateActions';
import { runAutoDebate, generateDebateArgument } from '../services/negotiate';
import { DealContextState } from '../prompts/systemPrompt';
import { DetectedParties } from '../services/documentAnalysis';
import { RiskTolerance } from '../types/state';
import { SideState } from './SideSelector';
import { AIProvider } from '../types';
import './Negotiate.css';

interface NegotiateProps {
    state: NegotiateState;
    onClose: () => void;
    onStateChange: (updates: Partial<NegotiateState> | ((prev: NegotiateState) => Partial<NegotiateState>)) => void;
    apiKey: string;
    model?: string;
    provider?: AIProvider;
    dealContext?: DealContextState | null;
    detectedParties?: DetectedParties | null;
    riskTolerance?: RiskTolerance | null;
    selectedSide?: SideState;
    onContinueInMainChat?: (summary: string) => void;
}

export const Negotiate: React.FC<NegotiateProps> = ({
    state,
    onClose,
    onStateChange,
    apiKey,
    model = 'gemini-2.0-flash',
    provider = 'gemini',
    dealContext,
    detectedParties,
    riskTolerance,
    selectedSide,
    onContinueInMainChat
}) => {
    const stopRef = useRef(false);

    if (!state.isOpen) return null;

    const handleStartAutoDebate = async (position: string, userSide: 'for' | 'against') => {
        stopRef.current = false;

        onStateChange({
            mode: 'auto',
            position,
            userSide,
            messages: [],
            isGenerating: true
        });

        const ctx = {
            position,
            dealContext,
            detectedParties,
            riskTolerance,
            userSide,
            selectedSide
        };

        await runAutoDebate(
            apiKey,
            ctx,
            5, // max 5 rounds per side
            (msg: DebateMessage) => {
                // Add message to state
                onStateChange((prev) => ({
                    messages: [...(prev.messages || []), msg]
                }));
            },
            () => stopRef.current,
            model,
            provider
        );

        onStateChange({ isGenerating: false });
    };

    const handleStop = () => {
        stopRef.current = true;
        onStateChange({ isGenerating: false });
    };

    const handleStartInteractive = async (position: string, userSide: 'for' | 'against') => {
        onStateChange({
            mode: 'interactive',
            position,
            userSide,
            messages: [],
            isGenerating: true
        });

        // Auto-start: AI immediately pushes back on the proposal (Opponent)
        try {
            const ctx = {
                position,
                dealContext,
                detectedParties,
                riskTolerance,
                userSide,
                selectedSide
            };

            const result = await generateDebateArgument(
                apiKey,
                ctx,
                [],
                'against',
                model,
                provider
            );

            const aiMsg: DebateMessage = {
                id: 1,
                side: 'against',
                author: 'ai',
                headline: result.headline,
                explanation: result.explanation,
                timestamp: new Date()
            };

            onStateChange({
                messages: [aiMsg],
                isGenerating: false
            });
        } catch (error) {
            console.error('[handleStartInteractive] Error:', error);
            onStateChange({ isGenerating: false });
        }
    };

    const handleUserArgument = async (text: string) => {
        // Add user's message
        const userMsg: DebateMessage = {
            id: state.messages.length + 1,
            side: state.userSide!,
            author: 'user',
            headline: text.length > 50 ? text.substring(0, 50) + '...' : text,
            explanation: text,
            timestamp: new Date()
        };

        onStateChange((prev) => ({
            messages: [...prev.messages, userMsg],
            isGenerating: true
        }));

        // Generate AI response (opposite side)
        const aiSide = state.userSide === 'for' ? 'against' : 'for';

        try {
            const ctx = {
                position: state.position,
                dealContext,
                detectedParties,
                riskTolerance,
                userSide: state.userSide!,
                selectedSide
            };

            const result = await generateDebateArgument(
                apiKey,
                ctx,
                [...state.messages, userMsg],
                aiSide,
                model,
                provider
            );

            const aiMsg: DebateMessage = {
                id: state.messages.length + 2,
                side: aiSide,
                author: 'ai',
                headline: result.headline,
                explanation: result.explanation,
                timestamp: new Date()
            };

            onStateChange((prev) => ({
                messages: [...prev.messages, aiMsg],
                isGenerating: false
            }));

        } catch (error) {
            console.error('[handleUserArgument] Error:', error);
            onStateChange({ isGenerating: false });
        }
    };

    const handleLetAIHelp = async () => {
        // AI generates argument for user's side
        onStateChange({ isGenerating: true });

        try {
            const ctx = {
                position: state.position,
                dealContext,
                detectedParties,
                riskTolerance,
                userSide: state.userSide!,
                selectedSide
            };

            const result = await generateDebateArgument(
                apiKey,
                ctx,
                state.messages,
                state.userSide!,
                model,
                provider
            );

            const aiMsg: DebateMessage = {
                id: state.messages.length + 1,
                side: state.userSide!,
                author: 'ai',  // AI helping user
                headline: result.headline,
                explanation: result.explanation,
                timestamp: new Date()
            };

            onStateChange((prev) => ({
                messages: [...prev.messages, aiMsg],
                isGenerating: false
            }));

        } catch (error) {
            console.error('[handleLetAIHelp] Error:', error);
            onStateChange({ isGenerating: false });
        }
    };

    const handleReset = () => {
        stopRef.current = true;
        onStateChange({
            mode: null,
            position: '',
            userSide: null,
            messages: [],
            isGenerating: false
        });
    };

    const sideName = selectedSide?.selected === 'partyA'
        ? detectedParties?.partyA?.shortName
        : (selectedSide?.selected === 'partyB' ? detectedParties?.partyB?.shortName : undefined);

    return (
        <div className="modal">
            <header className="modal__header">
                <div className="modal__header-row">
                    <h1 className="modal__title">Negotiate</h1>
                    <div className="modal__header-actions">
                        {state.isGenerating && (
                            <button onClick={handleStop} className="modal__header-btn">Stop</button>
                        )}
                        {state.mode && !state.isGenerating && (
                            <button onClick={handleReset} className="modal__header-btn">New</button>
                        )}
                        <button onClick={onClose} className="modal__header-btn">Close</button>
                    </div>
                </div>
            </header>

            {/* Setup Screen (when mode is null) */}
            {state.mode === null && (
                <NegotiateSetup
                    onStartAutoDebate={handleStartAutoDebate}
                    onStartInteractive={handleStartInteractive}
                    hasSideSelected={selectedSide?.selected !== 'neutral'}
                    sideName={sideName}
                />
            )}

            {/* Auto-Debate View */}
            {state.mode === 'auto' && (
                <>
                    <DebateView
                        position={state.position}
                        mode={state.mode}
                        messages={state.messages}
                        isGenerating={state.isGenerating}
                        userSide={state.userSide}
                    />
                    {!state.isGenerating && state.messages.length > 0 && (
                        <NegotiateActions
                            messages={state.messages}
                            position={state.position}
                            onContinueInChat={onContinueInMainChat ? (summary) => {
                                onClose();
                                onContinueInMainChat(summary);
                            } : undefined}
                        />
                    )}
                </>
            )}

            {/* Interactive Brainstorm View */}
            {state.mode === 'interactive' && (
                <>
                    <DebateView
                        position={state.position}
                        mode={state.mode}
                        messages={state.messages}
                        isGenerating={state.isGenerating}
                        userSide={state.userSide}
                    />
                    {!state.isGenerating && state.messages.length > 0 && (
                        <NegotiateActions
                            messages={state.messages}
                            position={state.position}
                            onContinueInChat={onContinueInMainChat ? (summary) => {
                                onClose();
                                onContinueInMainChat(summary);
                            } : undefined}
                        />
                    )}
                    <InteractiveInput
                        userSide={state.userSide!}
                        onSendArgument={handleUserArgument}
                        onLetAIRespond={handleLetAIHelp}
                        disabled={state.isGenerating}
                    />
                </>
            )}
        </div>
    );
};

export default Negotiate;
