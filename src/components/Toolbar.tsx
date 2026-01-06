/**
 * Toolbar Component
 * Container for toolbar buttons: Side Selector, Context, Risk Tolerance, Definitions, and Mind Map
 */

import React from 'react';
import { SideSelector, SideState, DetectedParties } from './SideSelector';
import { Context } from './Context';
import { DealContextState } from '../prompts/systemPrompt';
import { DefinedTerm } from '../services/documentAnalysis';
import { RiskTolerance } from '../types/state';
import './Toolbar.css';

interface ToolbarProps {
    selectedSide: SideState;
    onSideChange: (side: SideState) => void;
    detectedParties: DetectedParties | null;
    partiesLoading: boolean;
    // Side Selector control
    sidesSelectorOpen?: boolean;
    onSidesSelectorChange?: (open: boolean) => void;
    // Context props
    contextOpen: boolean;
    setContextOpen: (open: boolean) => void;
    dealContext: string;
    setDealContext: (value: string) => void;
    contextState: DealContextState;
    // Definitions props
    definitionsOpen: boolean;
    setDefinitionsOpen: (open: boolean) => void;
    definedTerms: DefinedTerm[];
    // Negotiate props
    onNegotiateOpen: () => void;
    negotiateHasHistory: boolean;
    // Mind Map props
    mindMapOpen: boolean;
    setMindMapOpen: (open: boolean) => void;
    mindMapTopicCount: number;
    // Risk Tolerance props
    riskTolerance: RiskTolerance;
    riskPanelOpen: boolean;
    onRiskPanelOpen: () => void;
}

export const Toolbar: React.FC<ToolbarProps> = ({
    selectedSide,
    onSideChange,
    detectedParties,
    partiesLoading,
    sidesSelectorOpen,
    onSidesSelectorChange,
    contextOpen,
    setContextOpen,
    dealContext,
    setDealContext,
    contextState,
    definitionsOpen,
    setDefinitionsOpen,
    definedTerms,
    onNegotiateOpen,
    negotiateHasHistory,
    mindMapOpen,
    setMindMapOpen,
    mindMapTopicCount,
    riskTolerance,
    riskPanelOpen,
    onRiskPanelOpen
}) => {
    return (
        <div className="toolbar">
            {/* Side Selector - wrapped for popover positioning */}
            <div className="toolbar__item">
                <SideSelector
                    selectedSide={selectedSide}
                    onSideChange={onSideChange}
                    detectedParties={detectedParties}
                    partiesLoading={partiesLoading}
                    forceOpen={sidesSelectorOpen}
                    onOpenChange={onSidesSelectorChange}
                />
            </div>

            {/* Context - wrapped for popover positioning */}
            <div className="toolbar__item">
                <button
                    onClick={() => setContextOpen(!contextOpen)}
                    className={`toolbar__btn ${contextOpen ? 'toolbar__btn--active' : ''}`}
                >
                    Context
                    {contextState.isActive && !contextOpen && <span className="toolbar__btn__dot" />}
                </button>
                <Context
                    isOpen={contextOpen}
                    onClose={() => setContextOpen(false)}
                    value={dealContext}
                    onSave={setDealContext}
                />
            </div>

            {/* Risk Tolerance - opens full-pane panel */}
            <button
                onClick={onRiskPanelOpen}
                className={`toolbar__btn ${riskPanelOpen ? 'toolbar__btn--active' : ''}`}
                title="Set negotiation risk tolerance"
            >
                Risk
                {riskTolerance.enabled && !riskPanelOpen && <span className="toolbar__btn__dot" />}
            </button>

            <button
                onClick={() => setDefinitionsOpen(!definitionsOpen)}
                className={`toolbar__btn ${definitionsOpen ? 'toolbar__btn--active' : ''}`}
            >
                Definitions
                {definedTerms.length > 0 && !definitionsOpen && <span className="toolbar__btn__dot" />}
            </button>

            <button
                onClick={onNegotiateOpen}
                className="toolbar__btn"
            >
                Negotiate
                {negotiateHasHistory && <span className="toolbar__btn__dot" />}
            </button>

            <button
                onClick={() => setMindMapOpen(true)}
                className={`toolbar__btn ${mindMapOpen ? 'toolbar__btn--active' : ''}`}
                disabled={mindMapTopicCount === 0}
            >
                Mind Map
                {mindMapTopicCount > 0 && !mindMapOpen && <span className="toolbar__btn__dot" />}
            </button>
        </div>
    );
};

export default Toolbar;


