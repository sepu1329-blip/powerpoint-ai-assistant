import React, { useState } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import type { ChatMessage, SlideAction, AgentStep } from '../types';

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
  const hasSteps = message.steps && message.steps.length > 0;

  return (
    <div className={`msg-row ${isUser ? 'user' : 'ai'}`}>
      {!isUser && (
        <div className="avatar-container">
          AI
        </div>
      )}
      <div className={`bubble ${isUser ? 'user' : 'ai'}`}>
        {/* AI 작업 단계 표시 (사용자 메시지엔 숨김) */}
        {!isUser && hasSteps && (
          <AgentStepsLog steps={message.steps!} />
        )}

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

/* ─────────────────────────────────────
   AI 에이전트 작업 단계 아코디언 컴포넌트
───────────────────────────────────── */
const AgentStepsLog: React.FC<{ steps: AgentStep[] }> = ({ steps }) => {
  const [expanded, setExpanded] = useState(false);

  const hasRunning = steps.some(s => s.status === 'running');
  const hasError   = steps.some(s => s.status === 'error');
  // 완료된 경우엔 기본적으로 접힘, 실행 중엔 펼침
  const [userToggled, setUserToggled] = useState(false);
  const isOpen = userToggled ? expanded : hasRunning;

  const handleToggle = () => {
    setUserToggled(true);
    setExpanded(!isOpen);
  };

  const summaryLabel = hasRunning
    ? '작업 진행 중...'
    : hasError
    ? '일부 작업 실패'
    : `작업 완료 (${steps.length}단계)`;

  const summaryColor = hasRunning ? '#854d0e' : hasError ? '#991b1b' : '#166534';
  const summaryBg   = hasRunning ? '#fef9c3' : hasError ? '#fee2e2' : '#dcfce7';
  const summaryBorder = hasRunning ? '#fde047' : hasError ? '#fca5a5' : '#86efac';

  return (
    <div style={{
      marginBottom: '10px',
      borderRadius: '8px',
      border: `1px solid ${summaryBorder}`,
      background: summaryBg,
      overflow: 'hidden',
      fontSize: '12px',
    }}>
      {/* 헤더 (클릭해서 펼치기/접기) */}
      <div
        onClick={handleToggle}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: '6px',
          padding: '6px 10px',
          cursor: 'pointer',
          color: summaryColor,
          fontWeight: 600,
          userSelect: 'none',
        }}
      >
        {/* 상태 아이콘 */}
        {hasRunning ? (
          <SpinnerIcon color={summaryColor} />
        ) : hasError ? (
          <ErrorIcon color={summaryColor} />
        ) : (
          <CheckAllIcon color={summaryColor} />
        )}

        <span style={{ flex: 1 }}>{summaryLabel}</span>

        {/* 펼치기 화살표 */}
        <svg
          width="12" height="12" viewBox="0 0 24 24"
          fill="none" stroke={summaryColor} strokeWidth="2.5"
          style={{
            transition: 'transform 0.2s',
            transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)',
            flexShrink: 0,
          }}
        >
          <polyline points="6 9 12 15 18 9" />
        </svg>
      </div>

      {/* 단계 목록 */}
      {isOpen && (
        <div style={{ borderTop: `1px solid ${summaryBorder}`, padding: '4px 0' }}>
          {steps.map((step, idx) => (
            <StepRow key={step.id} step={step} index={idx + 1} />
          ))}
        </div>
      )}
    </div>
  );
};

/* 개별 단계 행 */
const StepRow: React.FC<{ step: AgentStep; index: number }> = ({ step, index }) => {
  const isRunning = step.status === 'running';
  const isError   = step.status === 'error';
  const isDone    = step.status === 'done';

  const iconColor = isRunning ? '#854d0e' : isError ? '#991b1b' : '#166534';
  const textColor = isError ? '#991b1b' : '#374151';

  return (
    <div style={{
      display: 'flex',
      alignItems: 'flex-start',
      gap: '8px',
      padding: '5px 10px',
      borderBottom: '1px solid rgba(0,0,0,0.04)',
    }}>
      {/* 단계 번호 & 상태 아이콘 */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: '18px',
        height: '18px',
        flexShrink: 0,
        marginTop: '1px',
      }}>
        {isRunning ? (
          <SpinnerIcon color={iconColor} size={14} />
        ) : isError ? (
          <ErrorIcon color={iconColor} size={14} />
        ) : (
          <CheckIcon color={iconColor} size={14} />
        )}
      </div>

      {/* 단계 내용 */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{
          color: textColor,
          fontWeight: isRunning ? 600 : 500,
          lineHeight: 1.4,
        }}>
          <span style={{ color: '#9ca3af', marginRight: '4px' }}>#{index}</span>
          {step.label}
        </div>
        {step.detail && (
          <div style={{
            color: '#6b7280',
            fontSize: '11px',
            marginTop: '2px',
            lineHeight: 1.3,
            wordBreak: 'break-word',
          }}>
            {step.detail}
          </div>
        )}
      </div>
    </div>
  );
};

/* ─── 아이콘 헬퍼 ─── */
const SpinnerIcon: React.FC<{ color: string; size?: number }> = ({ color, size = 12 }) => (
  <div
    style={{
      width: size, height: size,
      border: `2px solid ${color}40`,
      borderTopColor: color,
      borderRadius: '50%',
      animation: 'spin 0.8s linear infinite',
      flexShrink: 0,
    }}
  />
);

const CheckIcon: React.FC<{ color: string; size?: number }> = ({ color, size = 14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2.5">
    <polyline points="20 6 9 17 4 12" />
  </svg>
);

const CheckAllIcon: React.FC<{ color: string; size?: number }> = ({ color, size = 14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2.5">
    <polyline points="20 6 9 17 4 12" />
    <polyline points="16 6 9 13" />
  </svg>
);

const ErrorIcon: React.FC<{ color: string; size?: number }> = ({ color, size = 14 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2.5">
    <circle cx="12" cy="12" r="10" />
    <line x1="12" y1="8" x2="12" y2="12" />
    <line x1="12" y1="16" x2="12.01" y2="16" />
  </svg>
);

/* ─── 액션 카드 컴포넌트 (기존 유지) ─── */
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
