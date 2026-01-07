/**
 * Community Panel Component
 * Side panel showing community tools and playbooks (preview feature)
 * Matches Definitions panel styling
 */

import React, { useRef, useEffect } from 'react';
import { useToast } from '../state/ToastContext';
import './Community.css';

interface CommunityProps {
    isOpen: boolean;
    onClose: () => void;
}

interface CommunityItem {
    title: string;
    description: string;
}

const tools: CommunityItem[] = [
    {
        title: 'Check Defined Terms',
        description: 'Find terms used but not defined'
    },
    {
        title: 'Plain English Check',
        description: 'Identify overly complex clauses'
    },
    {
        title: 'GDPR Red Flags',
        description: 'Check for data protection issues'
    }
];

const playbooks: CommunityItem[] = [
    {
        title: 'SaaS Supplier Playbook',
        description: 'Negotiation positions for SaaS vendors'
    },
    {
        title: 'NDA Mutual Playbook',
        description: 'Balanced NDA negotiation guide'
    }
];

export const Community: React.FC<CommunityProps> = ({ isOpen, onClose }) => {
    const { showToast } = useToast();
    const panelRef = useRef<HTMLDivElement>(null);

    // Close on click outside
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (panelRef.current && !panelRef.current.contains(event.target as Node)) {
                onClose();
            }
        };

        if (isOpen) {
            document.addEventListener('mousedown', handleClickOutside);
        }
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [isOpen, onClose]);

    if (!isOpen) return null;

    const handleItemClick = () => {
        showToast('Coming soon');
    };

    return (
        <div className="panel-overlay">
            <div className="panel-overlay__backdrop" onClick={onClose} />
            <div className="panel" ref={panelRef}>
                <div className="panel__header">
                    <div className="panel__header-row">
                        <h2 className="panel__title">Community</h2>
                        <button className="panel__close" onClick={onClose}>
                            Close
                        </button>
                    </div>
                </div>

                <div className="panel__content">
                    {/* Tools Section */}
                    <div className="community__section">
                        <p className="community__section-title">Tools</p>
                        <div className="community__list">
                            {tools.map((tool, index) => (
                                <button
                                    key={index}
                                    className="community__item"
                                    onClick={handleItemClick}
                                >
                                    <p className="community__item-title">{tool.title}</p>
                                    <p className="community__item-desc">{tool.description}</p>
                                </button>
                            ))}
                        </div>
                    </div>

                    <div className="community__divider" />

                    {/* Playbooks Section */}
                    <div className="community__section">
                        <p className="community__section-title">Playbooks</p>
                        <div className="community__list">
                            {playbooks.map((playbook, index) => (
                                <button
                                    key={index}
                                    className="community__item"
                                    onClick={handleItemClick}
                                >
                                    <p className="community__item-title">{playbook.title}</p>
                                    <p className="community__item-desc">{playbook.description}</p>
                                </button>
                            ))}
                        </div>
                    </div>

                    <div className="community__divider" />

                    {/* Footer text */}
                    <p className="community__footer">
                        Coming soon - community-contributed tools and playbooks you can add to your workflow.
                    </p>
                </div>
            </div>
        </div>
    );
};

export default Community;
