import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './components/App';
import './taskpane.css';

/* global document, Office, module */

const title = 'Vibe Legal React';

const render = (Component: typeof App) => {
    const container = document.getElementById('container');
    if (container) {
        const root = createRoot(container);
        root.render(
            <React.StrictMode>
                <Component />
            </React.StrictMode>
        );
    }
};

/* Render application after Office initializes */
Office.onReady(() => {
    render(App);
});

if ((module as any).hot) {
    (module as any).hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}
