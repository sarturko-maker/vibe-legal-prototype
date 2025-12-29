/**
 * NegotiateActions Component
 * Copy and export actions for debate results
 */

import React from 'react';
import { DebateMessage } from '../types/negotiate';

interface NegotiateActionsProps {
    messages: DebateMessage[];
    position: string;
    onContinueInChat?: (summary: string) => void;
}

export const NegotiateActions: React.FC<NegotiateActionsProps> = ({
    messages,
    position,
    onContinueInChat
}) => {
    if (messages.length === 0) return null;

    const handleCopyAll = () => {
        const text = formatDebateForCopy(position, messages);
        navigator.clipboard.writeText(text);
        // Could add toast notification
        console.log('[NegotiateActions] Copied to clipboard');
    };

    const handleContinueInChat = () => {
        const summary = formatDebateSummary(position, messages);
        onContinueInChat?.(summary);
    };

    return (
        <div className="modal__footer">
            <button
                onClick={handleCopyAll}
                className="modal__footer-btn modal__footer-btn--secondary"
            >
                Copy all
            </button>
            {onContinueInChat && (
                <button
                    onClick={handleContinueInChat}
                    className="modal__footer-btn modal__footer-btn--primary"
                >
                    Continue in chat
                </button>
            )}
        </div>
    );
};

function formatDebateForCopy(position: string, messages: DebateMessage[]): string {
    let text = `NEGOTIATION: "${position}"\n\n`;

    for (const msg of messages) {
        const side = msg.side === 'for' ? 'FOR' : 'AGAINST';
        const author = msg.author === 'user' ? 'You' : 'AI';
        text += `[${side}] (${author})\n`;
        text += `${msg.headline}\n`;
        text += `${msg.explanation}\n\n`;
    }

    return text;
}

function formatDebateSummary(position: string, messages: DebateMessage[]): string {
    const forArgs = messages.filter(m => m.side === 'for').map(m => m.headline);
    const againstArgs = messages.filter(m => m.side === 'against').map(m => m.headline);

    return `I just explored the negotiation point: "${position}"

Key arguments FOR:
${forArgs.map(a => `- ${a}`).join('\n')}

Key arguments AGAINST:
${againstArgs.map(a => `- ${a}`).join('\n')}

What's your recommendation on how to approach this?`;
}

export default NegotiateActions;
