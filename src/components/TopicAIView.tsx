import React, { useState, useEffect } from 'react';
import { MindMapTopic, MindMapSubTopic } from '../services/documentAnalysis';
import { parseMarkdown } from '../utils/parseMarkdown';

interface TopicAIViewProps {
    topic: MindMapTopic;
    subTopic: MindMapSubTopic | null;
    response: string | null;
    isGenerating: boolean;
    onSendQuery: (query: string | null) => void;
}

export const TopicAIView: React.FC<TopicAIViewProps> = ({
    topic,
    subTopic,
    response,
    isGenerating,
    onSendQuery
}) => {
    const [followUp, setFollowUp] = useState('');

    // Auto-generate on mount if no response
    useEffect(() => {
        if (!response && !isGenerating) {
            onSendQuery(null);  // null means "give me general advice"
        }
    }, []);

    const handleSendFollowUp = () => {
        if (followUp.trim().length < 5) return;
        onSendQuery(followUp.trim());
        setFollowUp('');
    };

    const handleKeyDown = (e: React.KeyboardEvent) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSendFollowUp();
        }
    };

    return (
        <div className="mindmap__ai-view">
            {/* Context Header */}
            <p className="mindmap__ai-context">
                {topic.title}
                {subTopic && ` › ${subTopic.name}`}
            </p>

            {/* Response Area */}
            {isGenerating ? (
                <div className="modal__loading" style={{ minHeight: '200px' }}>
                    <div className="loading-dots">
                        <span className="loading-dots__dot" />
                        <span className="loading-dots__dot" />
                        <span className="loading-dots__dot" />
                    </div>
                    <p className="modal__loading-text">Analyzing...</p>
                </div>
            ) : response ? (
                <>
                    <div className="mindmap__ai-response">
                        <p className="mindmap__ai-response-label">AI Analysis</p>
                        <div
                            className="mindmap__ai-response-text"
                            dangerouslySetInnerHTML={{ __html: parseMarkdown(response) }}
                        />
                    </div>

                    <div className="mindmap__followup">
                        <input
                            type="text"
                            value={followUp}
                            onChange={(e) => setFollowUp(e.target.value)}
                            onKeyDown={handleKeyDown}
                            placeholder="Ask a follow-up..."
                            className="mindmap__followup-input"
                        />
                        <button
                            onClick={handleSendFollowUp}
                            disabled={followUp.trim().length < 5}
                            className="mindmap__followup-btn"
                        >
                            Send
                        </button>
                    </div>
                </>
            ) : null}
        </div>
    );
};
