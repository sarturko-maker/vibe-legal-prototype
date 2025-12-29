/**
 * Operation Types for Vibe Legal
 * Extracted from vibe-legal-beta-0.2.yaml
 */

// Operation type literals
export type OperationType =
    | 'INSERT'
    | 'INSERT_BLOCK'
    | 'AMEND_SIMPLE'
    | 'AMEND'
    | 'DELETE';

// Base operation interface
export interface BaseOperation {
    type: OperationType;
    target_id?: number;
    target?: {
        number?: string;
        title?: string;
    };
    description?: string;
    original_text?: string;  // For content-based paragraph matching
}

// AMEND_SIMPLE: Full paragraph replacement with surgical diff
export interface AmendSimpleOperation extends BaseOperation {
    type: 'AMEND_SIMPLE';
    scope: 'NODE';
    amended_text: string;
}

// AMEND: Tree-based paragraph operations
export interface AmendTreeOperation extends BaseOperation {
    type: 'AMEND';
    para_operations: ParaOperation[];
}

export interface ParaOperation {
    para_id: number;
    action: 'MODIFY' | 'INSERT_AFTER' | 'DELETE';
    new_text?: string;
}

// INSERT: Add content after a target
export interface InsertOperation extends BaseOperation {
    type: 'INSERT';
    insert_after: number;
    content: string;
    clone_from?: number;
}

// INSERT_BLOCK: Insert multiple paragraphs
export interface InsertBlockOperation extends BaseOperation {
    type: 'INSERT_BLOCK';
    insert_after: number;
    block: string;  // Multi-paragraph content
}

// DELETE: Remove a clause/paragraph
export interface DeleteOperation extends BaseOperation {
    type: 'DELETE';
    target_id: number;
}

// Union type for all operations
export type Operation =
    | AmendSimpleOperation
    | AmendTreeOperation
    | InsertOperation
    | InsertBlockOperation
    | DeleteOperation;

// AI Response wrapper
export interface AIRouterResponse {
    intent?: 'ANSWER' | 'MODIFY' | 'HYBRID';
    operations: Operation[];
    answer?: string;
    explanation?: string;
    error?: string;
}

// Execution result for each operation
export interface ExecutionResult {
    success: boolean;
    operationType: OperationType;
    targetId?: number;
    error?: string;
    description?: string;
}
