/**
 * Definitions Component
 * Popover showing searchable list of defined terms
 */

import React, { useState, useEffect, useRef } from 'react';
import './Definitions.css';

interface DefinedTerm {
    term: string;
    definition: string;
}

interface DefinitionsProps {
    isOpen: boolean;
    onClose: () => void;
    terms: DefinedTerm[];
}

export const Definitions: React.FC<DefinitionsProps> = ({ isOpen, onClose, terms }) => {
    const [search, setSearch] = useState('');
    const [expandedTerm, setExpandedTerm] = useState<string | null>(null);
    const popoverRef = useRef<HTMLDivElement>(null);
    const searchRef = useRef<HTMLInputElement>(null);

    // Focus search on open
    useEffect(() => {
        if (isOpen) {
            setSearch('');
            setExpandedTerm(null);
            setTimeout(() => searchRef.current?.focus(), 100);
        }
    }, [isOpen]);

    // Close on click outside
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (popoverRef.current && !popoverRef.current.contains(event.target as Node)) {
                onClose();
            }
        };

        if (isOpen) {
            document.addEventListener('mousedown', handleClickOutside);
        }
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [isOpen, onClose]);

    if (!isOpen) return null;

    // Filter and sort alphabetically
    const filtered = terms
        .filter(t => t.term.toLowerCase().includes(search.toLowerCase()))
        .sort((a, b) => a.term.localeCompare(b.term));

    // Group by first letter
    const grouped: { [letter: string]: DefinedTerm[] } = {};
    filtered.forEach(t => {
        const letter = t.term[0]?.toUpperCase() || '#';
        if (!grouped[letter]) grouped[letter] = [];
        grouped[letter].push(t);
    });

    const handleToggleExpand = (term: string) => {
        setExpandedTerm(prev => prev === term ? null : term);
    };

    // Check if a definition is long enough to need expansion (more than ~100 chars or has ...)
    const needsExpansion = (definition: string) => {
        return definition.length > 80 || definition.endsWith('...');
    };

    return (
        <div className="panel-overlay">
            <div className="panel-overlay__backdrop" onClick={onClose} />
            <div className="panel" ref={popoverRef}>
                <div className="panel__header">
                    <div className="panel__header-row">
                        <h2 className="panel__title">
                            Defined Terms <span className="panel__title-count">({filtered.length})</span>
                        </h2>
                        <button className="panel__close" onClick={onClose}>
                            Close
                        </button>
                    </div>
                    <input
                        ref={searchRef}
                        type="text"
                        placeholder="Search definitions..."
                        value={search}
                        onChange={(e) => setSearch(e.target.value)}
                        className="panel__search"
                    />
                </div>

                <div className="panel__content">
                    {Object.keys(grouped).sort().map(letter => (
                        <div key={letter} className="definitions__group">
                            <p className="definitions__letter">{letter}</p>
                            <div className="definitions__list">
                                {grouped[letter].map((term, i) => {
                                    const isExpanded = expandedTerm === term.term;
                                    const canExpand = needsExpansion(term.definition);

                                    return (
                                        <div
                                            key={i}
                                            className={`definitions__item ${canExpand ? 'definitions__item--expandable' : ''}`}
                                            onClick={() => canExpand && handleToggleExpand(term.term)}
                                        >
                                            <p className="definitions__term">"{term.term}"</p>
                                            <p className={`definitions__text ${isExpanded ? 'definitions__text--expanded' : ''}`}>
                                                {term.definition}
                                            </p>
                                            {canExpand && (
                                                <button
                                                    className="definitions__expand-btn"
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        handleToggleExpand(term.term);
                                                    }}
                                                >
                                                    {isExpanded ? 'Show less' : 'Show more'}
                                                </button>
                                            )}
                                        </div>
                                    );
                                })}
                            </div>
                        </div>
                    ))}

                    {filtered.length === 0 && (
                        <div className="definitions__item">
                            <p className="definitions__text">
                                {search ? 'No matching terms' : 'No defined terms found'}
                            </p>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
};

export default Definitions;

