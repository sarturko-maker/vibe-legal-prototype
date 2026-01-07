/**
 * DeployLocallyModal Component
 * Modal explaining local deployment option (preview feature)
 */

import React, { useRef, useEffect } from 'react';
import './DeployLocallyModal.css';

interface DeployLocallyModalProps {
    isOpen: boolean;
    onClose: () => void;
}

export const DeployLocallyModal: React.FC<DeployLocallyModalProps> = ({ isOpen, onClose }) => {
    const panelRef = useRef<HTMLDivElement>(null);

    // Close on Escape key
    useEffect(() => {
        const handleKeyDown = (event: KeyboardEvent) => {
            if (event.key === 'Escape') {
                onClose();
            }
        };

        if (isOpen) {
            document.addEventListener('keydown', handleKeyDown);
        }
        return () => document.removeEventListener('keydown', handleKeyDown);
    }, [isOpen, onClose]);

    if (!isOpen) return null;

    return (
        <div className="panel-overlay">
            <div className="panel-overlay__backdrop" onClick={onClose} />
            <div className="panel deploy-locally-modal" ref={panelRef}>
                <div className="panel__header">
                    <div className="panel__header-row">
                        <h2 className="panel__title">Deploy Locally</h2>
                        <button className="panel__close" onClick={onClose}>
                            Close
                        </button>
                    </div>
                </div>

                <div className="panel__content">
                    <p className="deploy-locally__intro">
                        Run VibeLegal entirely on your machine with no external connections.
                    </p>

                    <ul className="deploy-locally__list">
                        <li className="deploy-locally__item">
                            Download add-in files for local hosting
                        </li>
                        <li className="deploy-locally__item">
                            Use Ollama for AI processing
                        </li>
                        <li className="deploy-locally__item">
                            Documents never leave your computer
                        </li>
                    </ul>

                    <p className="deploy-locally__coming-soon">
                        Coming soon.
                    </p>
                </div>
            </div>
        </div>
    );
};

export default DeployLocallyModal;
