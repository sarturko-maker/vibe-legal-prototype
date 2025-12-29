/**
 * Office Mock for Browser-Only Testing
 * 
 * This mock allows running the add-in in a browser without Word.
 * Import this in development mode when Office is not available.
 * 
 * Usage:
 *   if (!window.Office) {
 *       await import('./officeMock');
 *   }
 */

// Create minimal Office namespace mock
const mockContext = {
    sync: async () => Promise.resolve(),
    trackedObjects: {
        add: () => { },
        remove: () => { }
    },
    document: {
        body: {
            text: "This is mock document text for browser testing.\n\n1. Example Clause.\nThe Provider shall deliver services in accordance with this Agreement.",
            load: () => { },
            paragraphs: {
                load: () => { },
                items: [
                    { text: "1. Example Clause.", isListItem: false, style: "Heading 1" },
                    { text: "The Provider shall deliver services in accordance with this Agreement.", isListItem: false, style: "Normal" }
                ]
            },
            getOoxml: () => ({ value: "<w:document><w:body><w:p><w:r><w:t>Mock OOXML</w:t></w:r></w:p></w:body></w:document>" })
        },
        getSelection: () => ({
            text: "",
            load: () => { },
            getRange: () => ({
                load: () => { },
                text: ""
            })
        }),
        onSelectionChanged: {
            add: (handler: Function) => {
                console.log("[OfficeMock] Selection handler registered");
                return Promise.resolve();
            }
        }
    }
};

const OfficeMock = {
    onReady: (callback: () => void) => {
        console.log("[OfficeMock] Office.onReady called - mocking Word environment");
        setTimeout(callback, 100);
    },
    context: mockContext
};

// Word namespace mock
const WordMock = {
    run: async (callback: (context: typeof mockContext) => Promise<void>) => {
        try {
            await callback(mockContext);
            console.log("[OfficeMock] Word.run completed successfully");
        } catch (error) {
            console.error("[OfficeMock] Word.run error:", error);
            throw error;
        }
    },
    InsertLocation: {
        before: "Before",
        after: "After",
        start: "Start",
        end: "End",
        replace: "Replace"
    }
};

// Expose to window for browser testing
declare global {
    interface Window {
        Office: typeof OfficeMock;
        Word: typeof WordMock;
    }
}

// Only install if we're in a browser without real Office
if (typeof window !== 'undefined' && !window.Office) {
    console.log("[OfficeMock] Installing Office mock for browser-only testing");
    (window as any).Office = OfficeMock;
    (window as any).Word = WordMock;
}

export { OfficeMock, WordMock };
