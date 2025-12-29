import React from 'react';
import { MindMapTopic } from '../services/documentAnalysis';
import { SubTopicCard } from './SubTopicCard';

interface TopicExpandedProps {
    topic: MindMapTopic;
    onAskAITopic: () => void;
    onAskAISubTopic: (subTopicId: string) => void;
}

export const TopicExpanded: React.FC<TopicExpandedProps> = ({
    topic,
    onAskAITopic,
    onAskAISubTopic
}) => {
    return (
        <div className="mindmap__expanded">
            <div className="mindmap__expanded-header">
                <h2 className="mindmap__expanded-title">{topic.title}</h2>
                <p className="mindmap__expanded-summary">{topic.summary}</p>
                {topic.keyFigures.length > 0 && (
                    <div className="mindmap__figures">
                        {topic.keyFigures.map((fig, i) => (
                            <span key={i} className="mindmap__figure">{fig}</span>
                        ))}
                    </div>
                )}
            </div>

            {/* Sub-topics */}
            <div>
                <p className="mindmap__subtopics-label">Sub-topics</p>
                <div className="mindmap__subtopics">
                    {topic.subTopics.map((sub, index) => (
                        <SubTopicCard
                            key={sub.id}
                            subTopic={sub}
                            onAskAI={() => onAskAISubTopic(sub.id)}
                            index={index}
                        />
                    ))}
                </div>
            </div>
        </div>
    );
};
