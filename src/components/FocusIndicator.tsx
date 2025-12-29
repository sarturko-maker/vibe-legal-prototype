/**
 * FocusIndicator Component
 * Shows currently focused clause/selection
 */

import React from 'react';
import { useDocument } from '../state/DocumentContext';
import './FocusIndicator.css';

export function FocusIndicator() {
    const { focusedClause, setFocusedClause } = useDocument();

    if (!focusedClause) {
        return null;
    }

    const displayText = focusedClause.number
        ? `${focusedClause.number} ${focusedClause.title || ''}`.trim()
        : (focusedClause.title || 'Selected text');

    return (
        <div className="focus-banner">
            <div className="focus-banner__content">
                <p className="focus-banner__label">Focusing</p>
                <p className="focus-banner__text">{displayText}</p>
            </div>
            <button
                className="focus-banner__clear"
                onClick={() => setFocusedClause(null)}
            >
                ✕
            </button>
        </div>
    );
}
