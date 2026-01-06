/**
 * Chat History Panel Component
 * Slide-in panel for managing chat conversations
 */

import React from 'react';
import { Chat, formatTimeAgo } from '../types/chatHistory';
import './ChatHistoryPanel.css';

interface ChatHistoryPanelProps {
    isOpen: boolean;
    onClose: () => void;
    currentChatId: string | null;
    chats: Chat[];
    onSelectChat: (chatId: string) => void;
    onNewChat: () => void;
    onDeleteChat: (chatId: string) => void;
    onClearAll?: () => void;
}

export const ChatHistoryPanel: React.FC<ChatHistoryPanelProps> = ({
    isOpen,
    onClose,
    currentChatId,
    chats,
    onSelectChat,
    onNewChat,
    onDeleteChat,
    onClearAll
}) => {
    if (!isOpen) return null;

    // Sort chats by updatedAt descending (most recent first)
    const sortedChats = [...chats].sort((a, b) =>
        new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime()
    );

    const handleNewChat = () => {
        onNewChat();
        onClose();
    };

    const handleSelectChat = (chatId: string) => {
        onSelectChat(chatId);
        onClose();
    };

    const handleDelete = (e: React.MouseEvent, chatId: string) => {
        e.stopPropagation();
        if (window.confirm('Delete this chat? This cannot be undone.')) {
            onDeleteChat(chatId);
        }
    };

    const handleClearAll = () => {
        // Note: Confirmation handled by the parent or skipped for Office Add-in compatibility
        onClearAll?.();
    };

    return (
        <div className="panel-overlay">
            <div className="panel-overlay__backdrop" onClick={onClose} />
            <div className="panel chat-history-panel">
                <div className="panel__header">
                    <div className="panel__header-row">
                        <h2 className="panel__title">Chat History</h2>
                        <button className="panel__close" onClick={onClose}>
                            Close
                        </button>
                    </div>
                </div>

                <div className="panel__content">
                    {/* New Chat Button */}
                    <button
                        className="chat-history__new-btn"
                        onClick={handleNewChat}
                    >
                        <span className="chat-history__new-icon">+</span>
                        New Chat
                    </button>

                    {/* Chat List */}
                    {sortedChats.length > 0 ? (
                        <div className="chat-history__list">
                            {sortedChats.map(chat => (
                                <div
                                    key={chat.id}
                                    className={`chat-history__item ${chat.id === currentChatId ? 'chat-history__item--active' : ''}`}
                                    onClick={() => handleSelectChat(chat.id)}
                                >
                                    <div className="chat-history__item-content">
                                        <p className="chat-history__title">
                                            {chat.id === currentChatId && (
                                                <span className="chat-history__active-dot">●</span>
                                            )}
                                            {chat.title}
                                        </p>
                                        <span className="chat-history__time">
                                            {formatTimeAgo(chat.updatedAt)}
                                        </span>
                                    </div>
                                    <button
                                        className="chat-history__delete-btn"
                                        onClick={(e) => handleDelete(e, chat.id)}
                                        title="Delete chat"
                                    >
                                        ×
                                    </button>
                                </div>
                            ))}
                        </div>
                    ) : (
                        <div className="chat-history__empty">
                            <p>No chats yet</p>
                            <p className="chat-history__empty-hint">
                                Start chatting to create your first conversation
                            </p>
                        </div>
                    )}

                    {/* Clear All Button */}
                    {sortedChats.length > 0 && onClearAll && (
                        <button
                            className="chat-history__clear-btn"
                            onClick={handleClearAll}
                        >
                            Clear All History
                        </button>
                    )}
                </div>
            </div>
        </div>
    );
};

export default ChatHistoryPanel;

