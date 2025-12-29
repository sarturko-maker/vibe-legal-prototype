/**
 * NegotiateSetup Component
 * Setup screen for position input and mode selection
 */

import React, { useState } from 'react';

interface NegotiateSetupProps {
    onStartAutoDebate: (position: string) => void;
    onStartInteractive: (position: string, userSide: 'for' | 'against') => void;
}

export const NegotiateSetup: React.FC<NegotiateSetupProps> = ({
    onStartAutoDebate,
    onStartInteractive
}) => {
    const [position, setPosition] = useState('');
    const [selectedMode, setSelectedMode] = useState<'auto' | 'interactive' | null>(null);
    const [userSide, setUserSide] = useState<'for' | 'against'>('for');

    const canStart = position.trim().length >= 10;

    const handleStart = () => {
        if (!canStart) return;

        if (selectedMode === 'auto') {
            onStartAutoDebate(position.trim());
        } else if (selectedMode === 'interactive') {
            onStartInteractive(position.trim(), userSide);
        }
    };

    return (
        <div className="modal__content modal__content--padded">
            {/* Position Input */}
            <div className="modal__section">
                <label className="modal__label">Your Position</label>
                <textarea
                    value={position}
                    onChange={(e) => setPosition(e.target.value)}
                    placeholder="e.g., The indemnity cap should be 100% of the purchase price..."
                    className="modal__textarea"
                    style={{ height: '112px' }}
                    maxLength={500}
                />
            </div>

            {/* Mode Selection */}
            <div className="modal__section">
                <label className="modal__label">Mode</label>
                <div className="modal__toggle-group">
                    <button
                        className={`modal__toggle-btn ${selectedMode === 'auto' ? 'modal__toggle-btn--active' : ''}`}
                        onClick={() => setSelectedMode('auto')}
                    >
                        Auto-Debate
                    </button>
                    <button
                        className={`modal__toggle-btn ${selectedMode === 'interactive' ? 'modal__toggle-btn--active' : ''}`}
                        onClick={() => setSelectedMode('interactive')}
                    >
                        Brainstorm
                    </button>
                </div>
            </div>

            {/* Side Selection (only for interactive mode) */}
            {selectedMode === 'interactive' && (
                <div className="modal__section">
                    <label className="modal__label">Your Side</label>
                    <div className="modal__toggle-group">
                        <button
                            className={`modal__toggle-btn ${userSide === 'for' ? 'modal__toggle-btn--active' : ''}`}
                            onClick={() => setUserSide('for')}
                        >
                            For
                        </button>
                        <button
                            className={`modal__toggle-btn ${userSide === 'against' ? 'modal__toggle-btn--active' : ''}`}
                            onClick={() => setUserSide('against')}
                        >
                            Against
                        </button>
                    </div>
                </div>
            )}

            {/* Start Button */}
            <button
                className="modal__action-btn"
                onClick={handleStart}
                disabled={!canStart || !selectedMode}
            >
                Start
            </button>
        </div>
    );
};

export default NegotiateSetup;
