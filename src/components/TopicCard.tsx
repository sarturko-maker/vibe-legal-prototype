import React from 'react';
import { MindMapTopic } from '../services/documentAnalysis';

interface TopicCardProps {
    topic: MindMapTopic;
    onClick: () => void;
    index?: number;
}

export const TopicCard: React.FC<TopicCardProps> = ({ topic, onClick, index = 0 }) => {
    // Defensive: ensure arrays exist
    const subTopics = topic.subTopics ?? [];
    const keyFigures = topic.keyFigures ?? [];

    return (
        <button
            className="mindmap__topic"
            onClick={onClick}
            style={{ animationDelay: `${index * 50}ms` }}
        >
            <div className="mindmap__topic-header">
                <p className="mindmap__topic-title">{topic.title}</p>
                <span className="mindmap__topic-count">
                    {subTopics.length} sub-topics
                </span>
            </div>
            <p className="mindmap__topic-summary">{topic.summary}</p>
            {keyFigures.length > 0 && (
                <div className="mindmap__figures">
                    {keyFigures.map((fig, i) => (
                        <span key={i} className="mindmap__figure">{fig}</span>
                    ))}
                </div>
            )}
        </button>
    );
};
