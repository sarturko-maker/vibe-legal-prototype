import React, { useState } from 'react';
import { MindMapTopic } from '../services/documentAnalysis';
import { MindMapViewState, initialMindMapViewState } from '../types/mindmap';
import { TopicOverview } from './TopicOverview';
import { TopicExpanded } from './TopicExpanded';
import { TopicAIView } from './TopicAIView';
import { generateTopicAdvice } from '../services/mindmapAI';
import './MindMap.css';

interface MindMapProps {
    isOpen: boolean;
    onClose: () => void;
    topics: MindMapTopic[];
    apiKey: string;
    documentText: string;
    dealContext?: string | null;
    selectedSide?: string | null;
    onRefresh?: () => void;
}

export const MindMap: React.FC<MindMapProps> = ({
    isOpen,
    onClose,
    topics,
    apiKey,
    documentText,
    dealContext,
    selectedSide,
    onRefresh
}) => {
    const [viewState, setViewState] = useState<MindMapViewState>(initialMindMapViewState);

    if (!isOpen) return null;

    const handleBack = () => {
        if (viewState.view === 'ai') {
            setViewState(prev => ({
                ...prev,
                view: 'expanded',
                aiResponse: null,
                aiQuery: '',
                selectedSubTopicId: null
            }));
        } else if (viewState.view === 'expanded') {
            setViewState(prev => ({
                ...prev,
                view: 'overview',
                expandedTopicId: null
            }));
        }
    };

    const handleClose = () => {
        setViewState(initialMindMapViewState);
        onClose();
    };

    const handleTopicClick = (topicId: string) => {
        setViewState(prev => ({
            ...prev,
            view: 'expanded',
            expandedTopicId: topicId
        }));
    };

    const handleAskAITopic = () => {
        setViewState(prev => ({
            ...prev,
            view: 'ai',
            selectedSubTopicId: null,
            aiResponse: null,
            aiQuery: ''
        }));
    };

    const handleAskAISubTopic = (subTopicId: string) => {
        setViewState(prev => ({
            ...prev,
            view: 'ai',
            selectedSubTopicId: subTopicId,
            aiResponse: null,
            aiQuery: ''
        }));
    };

    const handleAIQuery = async (query: string | null) => {
        if (!currentTopic) return;

        setViewState(prev => ({ ...prev, isGenerating: true, aiQuery: query || '' }));

        try {
            const response = await generateTopicAdvice(
                apiKey,
                currentTopic,
                currentSubTopic || null,
                documentText,
                dealContext || null,
                selectedSide || null,
                query
            );

            setViewState(prev => ({
                ...prev,
                aiResponse: response,
                isGenerating: false
            }));
        } catch (error) {
            console.error('[MindMap] AI error:', error);
            setViewState(prev => ({
                ...prev,
                aiResponse: 'Sorry, there was an error generating advice. Please try again.',
                isGenerating: false
            }));
        }
    };

    const currentTopic = topics.find(t => t.id === viewState.expandedTopicId);
    const currentSubTopic = currentTopic?.subTopics.find(s => s.id === viewState.selectedSubTopicId);

    return (
        <div className="modal">
            <header className="modal__header">
                <div className="modal__header-row">
                    <h1 className="modal__title">
                        Mind Map <span className="modal__title-count">({topics.length})</span>
                    </h1>
                    <div className="modal__header-actions">
                        {viewState.view !== 'overview' && (
                            <button onClick={handleBack} className="modal__header-btn">Back</button>
                        )}
                        {viewState.view === 'overview' && onRefresh && (
                            <button onClick={onRefresh} className="modal__header-btn">Refresh</button>
                        )}
                        <button onClick={handleClose} className="modal__header-btn">Close</button>
                    </div>
                </div>
            </header>

            <div className="modal__content">
                {viewState.view === 'overview' && (
                    <TopicOverview
                        topics={topics}
                        onTopicClick={handleTopicClick}
                    />
                )}

                {viewState.view === 'expanded' && currentTopic && (
                    <TopicExpanded
                        topic={currentTopic}
                        onAskAITopic={handleAskAITopic}
                        onAskAISubTopic={handleAskAISubTopic}
                    />
                )}

                {viewState.view === 'ai' && currentTopic && (
                    <TopicAIView
                        topic={currentTopic}
                        subTopic={currentSubTopic || null}
                        response={viewState.aiResponse}
                        isGenerating={viewState.isGenerating}
                        onSendQuery={handleAIQuery}
                    />
                )}
            </div>
        </div>
    );
};
