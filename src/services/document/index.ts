// src/services/document/index.ts
// Document service exports

export {
    buildContractMapContext,
    buildSimpleContractMap,
    findClauseByNumber,
    findClauseByTitle,
    getParagraphById,
    createInitialContractMapState
} from './contractMap';

export {
    loadParagraphs,
    getSelectionText,
    getParagraphOoxml,
    insertOoxmlAtParagraph,
    getParagraphCount,
    runWord
} from './wordApi';
