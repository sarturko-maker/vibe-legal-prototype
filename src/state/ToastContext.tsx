/**
 * ToastContext - Global toast notification system
 * Used for showing "Coming soon" messages and other brief notifications
 */

import React, { createContext, useContext, useState, useCallback, ReactNode } from 'react';

interface ToastContextType {
    showToast: (message: string) => void;
}

const ToastContext = createContext<ToastContextType | undefined>(undefined);

interface ToastProviderProps {
    children: ReactNode;
}

export const ToastProvider: React.FC<ToastProviderProps> = ({ children }) => {
    const [toast, setToast] = useState<{ message: string; id: number } | null>(null);

    const showToast = useCallback((message: string) => {
        const id = Date.now();
        setToast({ message, id });

        // Auto-dismiss after 2 seconds
        setTimeout(() => {
            setToast(current => current?.id === id ? null : current);
        }, 2000);
    }, []);

    return (
        <ToastContext.Provider value={{ showToast }}>
            {children}
            {toast && (
                <div className="toast" key={toast.id}>
                    {toast.message}
                </div>
            )}
        </ToastContext.Provider>
    );
};

export const useToast = (): ToastContextType => {
    const context = useContext(ToastContext);
    if (!context) {
        throw new Error('useToast must be used within a ToastProvider');
    }
    return context;
};
