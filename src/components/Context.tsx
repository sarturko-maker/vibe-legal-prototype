/**
 * Context Component
 * Popover for entering persistent deal context
 */

import React, { useState, useEffect, useRef } from 'react';
import './Context.css';

interface ContextProps {
    isOpen: boolean;
    onClose: () => void;
    value: string;
    onSave: (value: string) => void;
}

export const Context: React.FC<ContextProps> = ({ isOpen, onClose, value, onSave }) => {
    const [draft, setDraft] = useState(value);
    const textareaRef = useRef<HTMLTextAreaElement>(null);
    const popoverRef = useRef<HTMLDivElement>(null);

    // Sync draft with value when opening
    useEffect(() => {
        if (isOpen) {
            setDraft(value);
            // Focus textarea when opening
            setTimeout(() => textareaRef.current?.focus(), 100);
        }
    }, [isOpen, value]);

    // Close on click outside
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (popoverRef.current && !popoverRef.current.contains(event.target as Node)) {
                onClose();
            }
        };

        if (isOpen) {
            document.addEventListener('mousedown', handleClickOutside);
        }
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [isOpen, onClose]);

    if (!isOpen) return null;

    const handleDone = () => {
        onSave(draft);
        onClose();
    };

    const handleClear = () => {
        setDraft('');
        onSave('');
        onClose();
    };

    const charCount = draft.length;
    const maxChars = 1000;

    return (
        <div className="popover" ref={popoverRef}>
            <div className="popover__card">
                <div className="popover__header">
                    <h3 className="popover__title">Deal Context</h3>
                </div>

                <div className="context-popover__body">
                    <textarea
                        ref={textareaRef}
                        value={draft}
                        onChange={(e) => setDraft(e.target.value.slice(0, maxChars))}
                        placeholder="Describe the deal context, key priorities, or special considerations..."
                        className="context-popover__textarea"
                    />
                    <p className="context-popover__char-count">{charCount} characters</p>
                </div>

                <div className="popover__footer">
                    <button
                        className="popover__footer-btn popover__footer-btn--secondary"
                        onClick={handleClear}
                    >
                        Clear
                    </button>
                    <button
                        className="popover__footer-btn popover__footer-btn--primary"
                        onClick={handleDone}
                    >
                        Done
                    </button>
                </div>
            </div>
        </div>
    );
};

export default Context;
