import React from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import type { ChatMessage, SlideAction } from '../types';

interface MessageBubbleProps {
  message: ChatMessage;
  onActionsApplied: (messageId: string) => void;
}

export const MessageBubble: React.FC<MessageBubbleProps> = ({
  message,
  onActionsApplied,
}) => {
  const isUser = message.role === 'user';
  const hasActions = message.actions && message.actions.length > 0;

  return (
    <div className={`msg-row ${isUser ? 'user' : 'ai'}`}>
      {!isUser && (
        <div className="avatar-container">
          AI
        </div>
      )}
      <div className={`bubble ${isUser ? 'user' : 'ai'}`}>
        <ReactMarkdown remarkPlugins={[remarkGfm]}>
          {message.content}
        </ReactMarkdown>

        {hasActions && (
          <ActionCard
            actions={message.actions!}
            applied={message.applied ?? false}
            onApply={() => onActionsApplied(message.id)}
          />
        )}
      </div>
    </div>
  );
};

/* 액션 카드 컴포넌트 */
const ActionCard: React.FC<{
  actions: SlideAction[];
  applied: boolean;
  onApply: () => void;
}> = ({ actions, applied, onApply }) => {
  return (
    <div className="action-card" style={{ marginTop: '10px', border: '1px solid #e2e8f0', borderRadius: '8px', overflow: 'hidden', background: 'white' }}>
      <div style={{ padding: '8px 12px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <span style={{ fontWeight: 600, fontSize: '12px', color: '#64748b' }}>AI 제안 ({actions.length}개)</span>
        {applied && <span style={{ color: '#10b981', fontSize: '11px', fontWeight: 600 }}>적용됨</span>}
      </div>
      <div style={{ padding: '12px' }}>
        <button 
          className="primary-btn" 
          onClick={onApply} 
          disabled={applied}
          style={{ padding: '8px', fontSize: '12px' }}
        >
          {applied ? '슬라이드에 반영됨' : '슬라이드에 적용하기'}
        </button>
      </div>
    </div>
  );
};
