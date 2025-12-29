import React from 'react';

interface Message {
    role: 'user' | 'bot';
    content: string;
}

interface ChatProps {
    messages: Message[];
}

const Chat: React.FC<ChatProps> = ({ messages }) => {
    return (
        <div className="chat-window">
            {messages.map((msg, index) => (
                <div key={index} className={`message ${msg.role === 'user' ? 'user-message' : 'bot-message'}`}>
                    <div className="bubble">{msg.content}</div>
                </div>
            ))}
        </div>
    );
};

export default Chat;
