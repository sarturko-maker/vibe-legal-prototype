/**
 * InteractiveInput Component
 * Input area for user arguments in interactive brainstorm mode
 */

import React, { useState } from 'react';

interface InteractiveInputProps {
    userSide: 'for' | 'against';
    onSendArgument: (text: string) => void;
    onLetAIRespond: () => void;
    disabled: boolean;
}

export const InteractiveInput: React.FC<InteractiveInputProps> = ({
    userSide,
    onSendArgument,
    onLetAIRespond,
    disabled
}) => {
    const [input, setInput] = useState('');

    const handleSend = () => {
        if (input.trim().length < 5) return;
        onSendArgument(input.trim());
        setInput('');
    };

    const handleKeyDown = (e: React.KeyboardEvent) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSend();
        }
    };

    return (
        <div className="interactive-input">
            <div className="input-label">
                Your turn ({userSide === 'for' ? 'FOR' : 'AGAINST'}):
            </div>
            <textarea
                value={input}
                onChange={(e) => setInput(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="Type your argument..."
                disabled={disabled}
                rows={2}
                className="input-textarea"
            />
            <div className="input-actions">
                <button
                    onClick={onLetAIRespond}
                    disabled={disabled}
                    className="ai-respond-btn"
                >
                    Let AI help me
                </button>
                <button
                    onClick={handleSend}
                    disabled={disabled || input.trim().length < 5}
                    className="send-btn"
                >
                    Send
                </button>
            </div>
        </div>
    );
};

export default InteractiveInput;
