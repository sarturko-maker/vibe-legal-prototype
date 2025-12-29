/**
 * Logger utility for Vibe Legal
 * Preserves debugging capability with production-safe output
 */

const DEBUG = process.env.NODE_ENV === 'development';

export const log = (...args: any[]): void => {
    if (DEBUG) console.log('[Vibe]', ...args);
};

export const logWarn = (...args: any[]): void => {
    if (DEBUG) console.warn('[Vibe]', ...args);
};

export const logError = (...args: any[]): void => {
    console.error('[Vibe Error]', ...args);  // Always log errors
};

// Startup message
export const logInit = (): void => {
    console.log('Vibe Legal Engine [Beta] initialized');
};
