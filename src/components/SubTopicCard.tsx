import React from 'react';
import { MindMapSubTopic } from '../services/documentAnalysis';

interface SubTopicCardProps {
    subTopic: MindMapSubTopic;
    onAskAI: () => void;
    index?: number;
}

export const SubTopicCard: React.FC<SubTopicCardProps> = ({ subTopic, onAskAI, index = 0 }) => {
    return (
        <div
            className="mindmap__subtopic"
            style={{ animationDelay: `${index * 100}ms` }}
        >
            <div className="mindmap__subtopic-header">
                <p className="mindmap__subtopic-name">{subTopic.name}</p>
            </div>
            <p className="mindmap__subtopic-summary">{subTopic.summary}</p>
            <div className="mindmap__subtopic-actions">
                <button className="mindmap__subtopic-action" onClick={onAskAI}>
                    Ask AI
                </button>
            </div>
        </div>
    );
};
