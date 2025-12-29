import React from 'react';

interface HeaderProps {
    onSettingsClick: () => void;
}

const Header: React.FC<HeaderProps> = ({ onSettingsClick }) => {
    return (
        <div className="header-container">
            <div className="brand-center">
                <h2 className="oscar-title">Vibe Legal v2</h2>
                <div id="doc-status" className="doc-status">Ready</div>
            </div>
            <button className="gear-icon-btn" title="Settings" onClick={onSettingsClick}>
                ⚙️
            </button>
        </div>
    );
};

export default Header;
