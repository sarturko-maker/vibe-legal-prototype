/**
 * Document Context for Vibe Legal
 * Manages contract map, style cache, document state
 */

import React, { createContext, useContext, useState, ReactNode } from 'react';
import { DocumentState, ContractMap, StyleInfo, FocusedClause } from '../types';

const defaultDocumentState: DocumentState = {
    contractMap: null,
    styleCache: new Map(),
    documentHash: null,
    isAnalyzing: false,
    focusedClause: null
};

interface DocumentContextType extends DocumentState {
    setContractMap: (map: ContractMap | null) => void;
    setStyleCache: (cache: Map<string, StyleInfo>) => void;
    setDocumentHash: (hash: string | null) => void;
    setIsAnalyzing: (analyzing: boolean) => void;
    setFocusedClause: (clause: FocusedClause | null) => void;
    invalidateMap: () => void;
}

const DocumentContext = createContext<DocumentContextType | null>(null);

export function DocumentProvider({ children }: { children: ReactNode }) {
    const [state, setState] = useState<DocumentState>(defaultDocumentState);

    const value: DocumentContextType = {
        ...state,
        setContractMap: (contractMap) => setState(s => ({ ...s, contractMap })),
        setStyleCache: (styleCache) => setState(s => ({ ...s, styleCache })),
        setDocumentHash: (documentHash) => setState(s => ({ ...s, documentHash })),
        setIsAnalyzing: (isAnalyzing) => setState(s => ({ ...s, isAnalyzing })),
        setFocusedClause: (focusedClause) => setState(s => ({ ...s, focusedClause })),
        invalidateMap: () => setState(s => ({ ...s, contractMap: null, documentHash: null }))
    };

    return (
        <DocumentContext.Provider value={value}>
            {children}
        </DocumentContext.Provider>
    );
}

export function useDocument() {
    const context = useContext(DocumentContext);
    if (!context) {
        throw new Error('useDocument must be used within DocumentProvider');
    }
    return context;
}
