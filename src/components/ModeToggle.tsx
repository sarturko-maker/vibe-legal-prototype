/**
 * ModeToggle Component
 * Toggles between Chat and Draft modes
 * - Chat: Answers questions, applies changes immediately with track changes
 * - Draft: Shows proposed changes in preview panel before applying
 */

import React from 'react';
import { useChat } from '../state/ChatContext';
import './ModeToggle.css';

export function ModeToggle() {
    const { appMode, setAppMode } = useChat();

    return (
        <div className="mode-toggle">
            <div className="mode-buttons">
                <button
                    className={`mode-btn ${appMode === 'chat' ? 'active' : ''}`}
                    onClick={() => setAppMode('chat')}
                    title="Chat mode: Answers questions and applies changes immediately"
                >
                    💬 Chat
                </button>
                <button
                    className={`mode-btn ${appMode === 'draft' ? 'active' : ''}`}
                    onClick={() => setAppMode('draft')}
                    title="Draft mode: Preview changes before applying"
                >
                    📝 Draft
                </button>
            </div>
        </div>
    );
}
