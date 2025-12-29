/**
 * Operations Router for Vibe Legal
 * Routes operations to appropriate handlers
 */

import { Operation, ExecutionResult } from '../../types';
import { handleAmendSimple } from './amendSimple';
import { handleAmendTree } from './amendTree';
import { handleInsert, handleInsertBlock } from './insert';
import { handleDelete } from './delete';
import { logError } from '../../utils/logger';

/**
 * Execute a list of operations.
 */
export async function executeOperations(
    context: any,
    operations: Operation[],
    author: string,
    onProgress?: (index: number, total: number) => void
): Promise<ExecutionResult[]> {
    const results: ExecutionResult[] = [];

    for (let i = 0; i < operations.length; i++) {
        const op = operations[i];

        if (onProgress) {
            onProgress(i, operations.length);
        }

        try {
            const result = await executeOperation(context, op, author);
            results.push(result);
        } catch (error: any) {
            logError(`Operation ${op.type} failed:`, error.message);
            results.push({
                success: false,
                operationType: op.type,
                targetId: op.target_id,
                error: error.message
            });
        }
    }

    return results;
}

/**
 * Execute a single operation.
 */
async function executeOperation(
    context: any,
    operation: Operation,
    author: string
): Promise<ExecutionResult> {
    switch (operation.type) {
        case 'AMEND_SIMPLE':
            return handleAmendSimple(context, operation, author);

        case 'INSERT':
            return handleInsert(context, operation, author);

        case 'INSERT_BLOCK':
            return handleInsertBlock(context, operation, author);

        case 'DELETE':
            return handleDelete(context, operation, author);

        case 'AMEND':
            return handleAmendTree(context, operation, author);

        default:
            return {
                success: false,
                operationType: (operation as any).type,
                error: `Unknown operation type: ${(operation as any).type}`
            };
    }
}
