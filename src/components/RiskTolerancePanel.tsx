/**
 * Risk Tolerance Panel Component
 * Full-pane panel for configuring AI risk calibration.
 * ADR-012: Risk Tolerance Feature
 */

import React, { useState, useEffect } from 'react';
import { RiskTolerance, getRiskSublabel } from '../types/state';
import './RiskTolerancePanel.css';

interface RiskTolerancePanelProps {
    isOpen: boolean;
    onClose: () => void;
    riskTolerance: RiskTolerance;
    selectedSide: string | null;
    onUpdate: (updates: Partial<RiskTolerance>) => void;
    isDisabled?: boolean;  // true when no Side selected
    onOpenSides?: () => void;  // callback to open Sides selector
}

export const RiskTolerancePanel: React.FC<RiskTolerancePanelProps> = ({
    isOpen,
    onClose,
    riskTolerance,
    selectedSide,
    onUpdate,
    isDisabled = false,
    onOpenSides
}) => {
    const [localLevel, setLocalLevel] = useState(riskTolerance.level);
    const [localBest, setLocalBest] = useState(riskTolerance.bestPosition);
    const [localFallback, setLocalFallback] = useState(riskTolerance.fallbackPosition);

    // Sync with props when panel opens
    useEffect(() => {
        if (isOpen) {
            setLocalLevel(riskTolerance.level);
            setLocalBest(riskTolerance.bestPosition);
            setLocalFallback(riskTolerance.fallbackPosition);
        }
    }, [isOpen, riskTolerance]);

    if (!isOpen) return null;

    const handleDone = () => {
        onUpdate({
            enabled: true,
            level: localLevel,
            bestPosition: localBest,
            fallbackPosition: localFallback
        });
        onClose();
    };

    const handleClear = () => {
        onUpdate({
            enabled: false,
            level: 50,
            bestPosition: '',
            fallbackPosition: ''
        });
        onClose();
    };

    const handleGoToSides = () => {
        onClose();
        if (onOpenSides) {
            onOpenSides();
        }
    };

    const riskSublabel = getRiskSublabel(localLevel);

    return (
        <div className="modal">
            <header className="modal__header">
                <div className="modal__header-row">
                    <h1 className="modal__title">Risk Tolerance</h1>
                    <div className="modal__header-actions">
                        {!isDisabled && riskTolerance.enabled && (
                            <button onClick={handleClear} className="modal__header-btn">Clear</button>
                        )}
                        <button onClick={onClose} className="modal__header-btn">Close</button>
                    </div>
                </div>
            </header>

            <div className="modal__content modal__content--padded">
                {/* Disabled state - show button to redirect to Sides */}
                {isDisabled && (
                    <button
                        className="risk-panel__sides-btn"
                        onClick={handleGoToSides}
                    >
                        Select a Side first to configure risk tolerance
                    </button>
                )}

                {/* Side indicator */}
                {selectedSide && (
                    <div className="risk-panel__side">
                        Acting for: <strong>{selectedSide}</strong>
                    </div>
                )}

                {/* Slider section */}
                <div className={`modal__section ${isDisabled ? 'risk-panel__disabled' : ''}`}>
                    <label className="modal__label">Risk Level</label>
                    <div className="risk-panel__slider-container">
                        <div className="risk-panel__slider-labels">
                            <span>Conservative</span>
                            <span className="risk-panel__level">{localLevel}%</span>
                            <span>Aggressive</span>
                        </div>
                        <input
                            type="range"
                            min="0"
                            max="100"
                            value={localLevel}
                            onChange={(e) => !isDisabled && setLocalLevel(parseInt(e.target.value))}
                            className="risk-panel__slider"
                            disabled={isDisabled}
                        />
                        <div className="risk-panel__badge">{riskSublabel}</div>
                    </div>
                </div>

                {/* Positions section - side by side */}
                <div className={`risk-panel__positions ${isDisabled ? 'risk-panel__disabled' : ''}`}>
                    <div className="risk-panel__position">
                        <label className="modal__label">Best Position</label>
                        <p className="risk-panel__hint">Examples of what we ideally want</p>
                        <textarea
                            value={localBest}
                            onChange={(e) => !isDisabled && setLocalBest(e.target.value)}
                            placeholder="e.g., Liability capped at 10% of contract value, no consequential damages..."
                            className="modal__textarea"
                            rows={5}
                            disabled={isDisabled}
                        />
                    </div>

                    <div className="risk-panel__position">
                        <label className="modal__label">Fallback Position</label>
                        <p className="risk-panel__hint">Examples of what we'd accept to close the deal</p>
                        <textarea
                            value={localFallback}
                            onChange={(e) => !isDisabled && setLocalFallback(e.target.value)}
                            placeholder="e.g., Liability up to 150%, consequentials capped at 50%..."
                            className="modal__textarea"
                            rows={5}
                            disabled={isDisabled}
                        />
                    </div>
                </div>

                {/* Guidance */}
                <div className="risk-panel__guidance">
                    These are examples, not an exhaustive list.
                    AI will infer your risk appetite and apply it to ALL commercial terms.
                </div>
            </div>

            <div className="modal__footer">
                {isDisabled ? (
                    /* When disabled, only show Cancel */
                    <button
                        className="modal__footer-btn modal__footer-btn--secondary"
                        onClick={onClose}
                        style={{ flex: 1 }}
                    >
                        Cancel
                    </button>
                ) : (
                    /* When enabled, show Cancel and Apply */
                    <>
                        <button
                            className="modal__footer-btn modal__footer-btn--secondary"
                            onClick={onClose}
                        >
                            Cancel
                        </button>
                        <button
                            className="modal__footer-btn modal__footer-btn--primary"
                            onClick={handleDone}
                        >
                            Apply
                        </button>
                    </>
                )}
            </div>
        </div>
    );
};

export default RiskTolerancePanel;
