export * from './types';
export * from './namespaces';
export * from './trackChanges';
export * from './runBuilder';
export * from './processor';
export * from './segmentMap';
export * from './applyRedline';
export * from './insertBlocks';

// Utility function for RSID generation
export function generateRsid(): string {
    return Math.random().toString(16).substring(2, 10).toUpperCase();
}
