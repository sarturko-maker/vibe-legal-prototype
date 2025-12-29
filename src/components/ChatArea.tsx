/**
 * ChatArea Component
 * Displays chat messages with markdown rendering for bot responses
 */

import React, { useEffect, useRef } from 'react';
import { useChat } from '../state/ChatContext';
import { renderMarkdown } from '../utils/markdown';
import './ChatArea.css';

export function ChatArea() {
    const { messages, isProcessing, processingStage } = useChat();
    const bottomRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
    }, [messages]);

    return (
        <div className="chat-container">
            {messages.length === 0 ? (
                <div className="chat-empty">
                    <h3 className="chat-empty__title">Ready to assist</h3>
                    <p className="chat-empty__subtitle">Ask questions about your document or describe changes you'd like to make</p>
                </div>
            ) : (
                messages.map((msg, idx) => (
                    <div key={idx} className={`chat-message chat-message--${msg.role}`}>
                        {msg.role === 'bot' ? (
                            <div
                                className="chat-bubble chat-bubble--bot"
                                dangerouslySetInnerHTML={{ __html: renderMarkdown(msg.content) }}
                            />
                        ) : (
                            <div className="chat-bubble chat-bubble--user">
                                <p>{msg.content}</p>
                            </div>
                        )}
                    </div>
                ))
            )}

            {isProcessing && (
                <div className="chat-loading">
                    <div className="chat-loading__bubble">
                        <div className="loading-dots">
                            <span className="loading-dots__dot" />
                            <span className="loading-dots__dot" />
                            <span className="loading-dots__dot" />
                        </div>
                    </div>
                </div>
            )}

            <div ref={bottomRef} />
        </div>
    );
}
