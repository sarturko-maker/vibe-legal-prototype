import React from 'react';
import { MindMapTopic } from '../services/documentAnalysis';
import { TopicCard } from './TopicCard';

interface TopicOverviewProps {
    topics: MindMapTopic[];
    onTopicClick: (topicId: string) => void;
}

export const TopicOverview: React.FC<TopicOverviewProps> = ({ topics, onTopicClick }) => {
    if (topics.length === 0) {
        return (
            <div className="modal__loading">
                <p className="modal__loading-text">No topics available. Open a contract document to generate the mind map.</p>
            </div>
        );
    }

    return (
        <div className="mindmap__topics">
            {topics.map((topic, index) => (
                <TopicCard
                    key={topic.id}
                    topic={topic}
                    onClick={() => onTopicClick(topic.id)}
                    index={index}
                />
            ))}
        </div>
    );
};
