/**
 * Negotiate Feature Types
 * Data structures for AI debate functionality
 */

export interface DebateMessage {
    id: number;
    side: 'for' | 'against';
    author: 'ai' | 'user';
    headline: string;
    explanation: string;
    timestamp: Date;
}

export interface NegotiateState {
    isOpen: boolean;
    mode: 'auto' | 'interactive' | null;  // null = setup screen
    position: string;
    userSide: 'for' | 'against' | null;
    messages: DebateMessage[];
    isGenerating: boolean;
}

export const initialNegotiateState: NegotiateState = {
    isOpen: false,
    mode: null,
    position: '',
    userSide: null,
    messages: [],
    isGenerating: false
};
