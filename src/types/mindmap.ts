export interface MindMapViewState {
    view: 'overview' | 'expanded' | 'ai';
    expandedTopicId: string | null;
    selectedSubTopicId: string | null;
    aiResponse: string | null;
    aiQuery: string;
    isGenerating: boolean;
}

export const initialMindMapViewState: MindMapViewState = {
    view: 'overview',
    expandedTopicId: null,
    selectedSubTopicId: null,
    aiResponse: null,
    aiQuery: '',
    isGenerating: false
};
