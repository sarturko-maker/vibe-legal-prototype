/**
 * NegotiateSetup Component
 * Setup screen for position input and mode selection
 * Requires side selection first (like Risk panel)
 */

import React, { useState } from 'react';

interface NegotiateSetupProps {
    onStartAutoDebate: (position: string, userSide: 'for' | 'against') => void;
    onStartInteractive: (position: string, userSide: 'for' | 'against') => void;
    hasSideSelected: boolean;
    sideName?: string;
}

export const NegotiateSetup: React.FC<NegotiateSetupProps> = ({
    onStartAutoDebate,
    onStartInteractive,
    hasSideSelected,
    sideName
}) => {
    const [position, setPosition] = useState('');
    const [selectedMode, setSelectedMode] = useState<'auto' | 'interactive' | null>(null);
    const [userSide, setUserSide] = useState<'for' | 'against'>('for');

    const canStart = position.trim().length >= 10;

    const handleStart = () => {
        if (!canStart) return;

        if (selectedMode === 'auto') {
            onStartAutoDebate(position.trim(), 'for');
        } else if (selectedMode === 'interactive') {
            onStartInteractive(position.trim(), 'for');
        }
    };

    // If no side selected in main toolbar, show prompt
    if (!hasSideSelected) {
        return (
            <div className="modal__content modal__content--padded">
                <div className="modal__section" style={{ textAlign: 'center', padding: '32px 16px' }}>
                    <p style={{ fontSize: '14px', color: 'var(--text-secondary)', marginBottom: '8px' }}>
                        ⚠️ Please select a side first
                    </p>
                    <p style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                        Use the <strong>Side</strong> button in the toolbar to choose which party you represent
                    </p>
                </div>
            </div>
        );
    }

    return (
        <div className="modal__content modal__content--padded">
            <div style={{
                background: 'var(--bg-secondary)',
                padding: '8px 12px',
                borderRadius: '4px',
                marginBottom: '16px',
                fontSize: '12px',
                color: 'var(--text-secondary)',
                borderLeft: '3px solid var(--accent-color)'
            }}>
                Negotiating as: <strong>{sideName || 'Selected Side'}</strong>
            </div>

            {/* Position Input */}
            <div className="modal__section">
                <label className="modal__label">Your Side's Proposal</label>
                <textarea
                    value={position}
                    onChange={(e) => setPosition(e.target.value)}
                    placeholder="e.g., The liability cap should be set at 200% of fees paid..."
                    className="modal__textarea"
                    style={{ height: '96px' }}
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

