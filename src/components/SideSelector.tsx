/**
 * Side Selector Component
 * Allows user to select which party they represent
 * Receives detectedParties from parent (detected on Settings Done)
 */

import React, { useState, useRef, useEffect } from 'react';
import './SideSelector.css';

export interface DetectedParties {
    partyA: { shortName: string; fullName: string; role: string };
    partyB: { shortName: string; fullName: string; role: string };
}

export interface SideState {
    selected: 'neutral' | 'partyA' | 'partyB';
}

interface SideSelectorProps {
    selectedSide: SideState;
    onSideChange: (side: SideState) => void;
    detectedParties: DetectedParties | null;
    partiesLoading: boolean;
    forceOpen?: boolean;  // External control to open the selector
    onOpenChange?: (open: boolean) => void;  // Notify parent of open state changes
}

export const SideSelector: React.FC<SideSelectorProps> = ({
    selectedSide,
    onSideChange,
    detectedParties,
    partiesLoading,
    forceOpen,
    onOpenChange
}) => {
    const [isOpen, setIsOpen] = useState(false);

    // Sync with forceOpen prop
    useEffect(() => {
        if (forceOpen !== undefined && forceOpen !== isOpen) {
            setIsOpen(forceOpen);
        }
    }, [forceOpen]);

    // Notify parent of open state changes
    useEffect(() => {
        if (onOpenChange) {
            onOpenChange(isOpen);
        }
    }, [isOpen, onOpenChange]);
    const dropdownRef = useRef<HTMLDivElement>(null);

    // Close dropdown when clicking outside
    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
                setIsOpen(false);
            }
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, []);

    const getButtonLabel = (): string => {
        if (selectedSide.selected === 'neutral') return 'Side';
        if (selectedSide.selected === 'partyA' && detectedParties) {
            return detectedParties.partyA.shortName;
        }
        if (selectedSide.selected === 'partyB' && detectedParties) {
            return detectedParties.partyB.shortName;
        }
        return 'Side';
    };

    const handleSelect = (selected: 'neutral' | 'partyA' | 'partyB') => {
        onSideChange({ selected });
        setIsOpen(false);
    };

    const isActive = selectedSide.selected !== 'neutral';

    return (
        <div className="side-selector" ref={dropdownRef}>
            <button
                className={`toolbar__btn ${isOpen ? 'toolbar__btn--active' : ''}`}
                onClick={() => setIsOpen(!isOpen)}
            >
                {getButtonLabel()}
                {isActive && !isOpen && <span className="toolbar__btn__dot" />}
            </button>

            {isOpen && (
                <div className="popover" style={{ top: '100%', left: 0, right: 'auto', minWidth: '240px', marginTop: '8px' }}>
                    <div className="popover__card">
                        <div className="popover__header">
                            <h3 className="popover__title">Acting for</h3>
                        </div>

                        <div className="side-selector__list">
                            <button
                                className="side-selector__option"
                                onClick={() => handleSelect('neutral')}
                            >
                                <div>
                                    <p className="side-selector__option-label">Neutral</p>
                                    <p className="side-selector__option-desc">Balanced advice</p>
                                </div>
                                {selectedSide.selected === 'neutral' && <span className="side-selector__option-check" />}
                            </button>

                            {partiesLoading ? (
                                <div className="side-selector__option">
                                    <p className="side-selector__option-desc">Detecting parties...</p>
                                </div>
                            ) : detectedParties ? (
                                <>
                                    <button
                                        className="side-selector__option"
                                        onClick={() => handleSelect('partyA')}
                                    >
                                        <div>
                                            <p className="side-selector__option-label">{detectedParties.partyA.shortName}</p>
                                            <p className="side-selector__option-desc">{detectedParties.partyA.role}</p>
                                        </div>
                                        {selectedSide.selected === 'partyA' && <span className="side-selector__option-check" />}
                                    </button>

                                    <button
                                        className="side-selector__option"
                                        onClick={() => handleSelect('partyB')}
                                    >
                                        <div>
                                            <p className="side-selector__option-label">{detectedParties.partyB.shortName}</p>
                                            <p className="side-selector__option-desc">{detectedParties.partyB.role}</p>
                                        </div>
                                        {selectedSide.selected === 'partyB' && <span className="side-selector__option-check" />}
                                    </button>
                                </>
                            ) : (
                                <div className="side-selector__option">
                                    <p className="side-selector__option-desc">Save Settings to detect parties</p>
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

export default SideSelector;
