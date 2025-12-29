/**
 * DebateView Component
 * Container for debate messages with position header and auto-scroll
 */

import React, { useRef, useEffect } from 'react';
import { DebateMessage } from '../types/negotiate';
import { DebateBubble } from './DebateBubble';

interface DebateViewProps {
    position: string;
    mode: 'auto' | 'interactive';
    messages: DebateMessage[];
    isGenerating: boolean;
    userSide?: 'for' | 'against' | null;
}

export const DebateView: React.FC<DebateViewProps> = ({
    position,
    mode,
    messages,
    isGenerating,
    userSide
}) => {
    const scrollRef = useRef<HTMLDivElement>(null);

    // Auto-scroll to bottom when new messages arrive
    useEffect(() => {
        if (scrollRef.current) {
            scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
        }
    }, [messages]);

    return (
        <div className="negotiate__debate" ref={scrollRef}>
            {messages.map((msg) => (
                <DebateBubble key={msg.id} message={msg} />
            ))}

            {/* Generating indicator */}
            {isGenerating && (
                <div className="negotiate__generating">
                    <div className="loading-dots">
                        <span className="loading-dots__dot" />
                        <span className="loading-dots__dot" />
                        <span className="loading-dots__dot" />
                    </div>
                    <span>Generating arguments...</span>
                </div>
            )}

            {/* Empty state */}
            {messages.length === 0 && !isGenerating && (
                <div className="negotiate__generating">
                    Waiting to start...
                </div>
            )}
        </div>
    );
};

export default DebateView;
