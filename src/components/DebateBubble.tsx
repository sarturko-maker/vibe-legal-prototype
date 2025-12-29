/**
 * DebateBubble Component
 * Individual debate message bubble
 */

import React from 'react';
import { DebateMessage } from '../types/negotiate';

interface DebateBubbleProps {
    message: DebateMessage;
}

export const DebateBubble: React.FC<DebateBubbleProps> = ({ message }) => {
    const isFor = message.side === 'for';

    return (
        <div className={`negotiate__argument negotiate__argument--${message.side}`}>
            <p className="negotiate__argument-label">
                {isFor ? 'For' : 'Against'}
            </p>
            <div className={`negotiate__argument-bubble negotiate__argument-bubble--${message.side}`}>
                <p className="negotiate__argument-headline">{message.headline}</p>
                <p className={`negotiate__argument-text negotiate__argument-text--${message.side}`}>
                    {message.explanation}
                </p>
            </div>
        </div>
    );
};

export default DebateBubble;
