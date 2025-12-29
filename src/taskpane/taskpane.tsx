/**
 * Vibe Legal - Office Add-in Entry Point
 * Modular React Architecture v1.0
 */

import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from '../components';
import './taskpane.css';

// Render immediately (works in browser AND Word)
const rootElement = document.getElementById('app');
if (rootElement) {
    const root = createRoot(rootElement);
    root.render(<App />);
    console.log('Vibe Legal Engine [Modular] initialized');
} else {
    console.error('Could not find #app element');
}

// Log when Office is ready (optional, for debugging)
if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
        console.log('Office.js ready:', info.host, info.platform);
    });
}
