import React, { useEffect, useRef, useState, useCallback } from 'react';
import { MessageBubble } from './MessageBubble';
import { ChatMessage, PresentationContext, AppSettings } from '../types';

interface ChatPanelProps {
  settings: AppSettings;
  messages: ChatMessage[];
  isLoading: boolean;
  statusText: string | null;
  error: string | null;
  context: PresentationContext | null;
  sendUserMessage: (text: string, useContext?: boolean) => Promise<void>;
  markActionsApplied: (messageId: string) => void;
  refreshContext: () => Promise<PresentationContext | null>;
  stopGeneration: () => void;
  clearChat: () => void;
}

const DEFAULT_QUICK_PROMPTS = [
  '현재 슬라이드를 요약해줘',
  '제목을 더 매력적으로 바꿔줘',
  '새 슬라이드 추가해줘',
  '내용을 3가지 핵심 포인트로 정리해줘',
];

export const ChatPanel: React.FC<ChatPanelProps> = ({ 
  settings, 
  messages, 
  isLoading, 
  statusText,
  error, 
  context, 
  sendUserMessage, 
  markActionsApplied, 
  refreshContext, 
  stopGeneration,
  clearChat
}) => {
  const [inputText, setInputText] = useState('');
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const isRefreshing = useRef(false);

  // 입력창 줄 수 계산 (최소 2줄, 최대 3줄)
  const inputRows = Math.max(2, Math.min(3, inputText.split('\n').length || 1));

  // 상단 프롬프트 리스트 (기본 + 사용자 정의)
  const allPrompts = [
    ...DEFAULT_QUICK_PROMPTS.map(p => ({ name: p, content: p })),
    ...(settings.customPrompts || [])
  ];

  // 메시지 스크롤
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // 초기 컨텍스트 로드 및 실시간 동기화
  useEffect(() => {
    const sync = async () => {
      if (isRefreshing.current) return;
      isRefreshing.current = true;
      await refreshContext();
      isRefreshing.current = false;
    };

    sync();
    const interval = setInterval(sync, 2000);
    return () => clearInterval(interval);
  }, [refreshContext]);

  const handleSendOrStop = useCallback(async () => {
    if (isLoading) {
      stopGeneration();
      return;
    }
    if (!inputText.trim()) return;
    const text = inputText;
    setInputText('');
    await sendUserMessage(text, settings.autoContext);
  }, [inputText, isLoading, sendUserMessage, stopGeneration, settings.autoContext]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    // 한글 입력 중 Enter 키 중복 방지 (IME)
    if (e.nativeEvent.isComposing) return;

    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendOrStop();
    }
  };

  const handlePromptSelect = (id: string) => {
    const selected = allPrompts.find(p => p.name === id || (p as any).id === id);
    if (selected) {
      setInputText(selected.content);
      inputRef.current?.focus();
    }
  };

  const currentSlide = context?.currentSlide;
  const selectedShapesCount = context?.selectedShapes?.length ?? 0;

  return (
    <div className="chat-panel" style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      {/* 상단 도구 모음 */}
      <div className="chat-top-toolbar" style={{ display: 'flex', flexDirection: 'column', gap: '8px', alignItems: 'flex-start', flexShrink: 0 }}>
        {/* 컨텍스트 스테이터스 (상단 이동 및 좌측 정렬) */}
        <div className="status-bar" style={{ marginTop: 0, justifyContent: 'flex-start', width: '100%' }}>
          <div className="context-pill" onClick={refreshContext} title="새로고침" style={{ padding: '4px 8px', borderRadius: '6px', background: '#f1f5f9' }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" style={{ marginRight: '6px' }}>
              <path d="M20 11a8.1 8.1 0 0 0-15.5-2m-.5 5v-5h5"/>
              <path d="M4 13a8.1 8.1 0 0 0 15.5 2m.5-5v5h-5"/>
            </svg>
            <span style={{ fontWeight: 600 }}>
              {context ? (
                (context.selectedSlides?.length ?? 0) >= 2
                  ? `${context.selectedSlides!.length}개 슬라이드 참조 중`
                  : context.selectedShapes && context.selectedShapes.length > 0
                    ? `객체 ${context.selectedShapes.length}개 참조 중`
                    : `슬라이드 1개, 객체 ${context.currentSlide?.shapes?.length ?? 0}개 참조 중`
              ) : '프레젠테이션 분석 중...'}
            </span>
          </div>
        </div>
        
        <select 
          className="styled-select" 
          onChange={(e) => handlePromptSelect(e.target.value)}
          value=""
        >
          <option value="" disabled>-- 자주 쓰는 프롬프트 선택 --</option>
          {allPrompts.map(p => (
            <option key={(p as any).id || p.name} value={(p as any).id || p.name}>
              {p.name}
            </option>
          ))}
        </select>
      </div>

      {/* 에러 메시지 표시 */}
      {error && (
        <div style={{ padding: '10px 15px', margin: '10px', backgroundColor: '#fee2e2', border: '1px solid #ef4444', borderRadius: '8px', color: '#b91c1c', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ flexShrink: 0 }}>
            <circle cx="12" cy="12" r="10" />
            <line x1="12" y1="8" x2="12" y2="12" />
            <line x1="12" y1="16" x2="12.01" y2="16" />
          </svg>
          <span style={{ flex: 1 }}>{error}</span>
        </div>
      )}




      {/* 메시지 영역 */}
      <div className="messages-container" style={{ flex: 1, overflowY: 'auto' }}>
        {messages.length <= 1 && (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', textAlign: 'center', color: '#94a3b8', padding: '40px 20px' }}>

            <p style={{ fontSize: '14px', fontWeight: 500, lineHeight: 1.6 }}>
              환영합니다! 무엇이든 물어보거나<br/>AI 모드를 통해 슬라이드를 조작해보세요.
            </p>
          </div>
        )}
        {messages.map((message) => (
          <MessageBubble
            key={message.id}
            message={message}
            onActionsApplied={markActionsApplied}
          />
        ))}
        {isLoading && messages[messages.length-1]?.role === 'user' && (
          <div className="message-bubble assistant loading">
            <div className="avatar-container">AI</div>
            <div className="message-content">
              <div className="typing-indicator" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <div style={{ display: 'flex', gap: '4px' }}>
                  <span></span><span></span><span></span>
                </div>
                {statusText && <span style={{ fontSize: '12px', color: '#64748b', marginLeft: '5px' }}>{statusText}</span>}
              </div>
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* 하단 입력 영역 */}
      <div className="bottom-area" style={{ flexShrink: 0, borderTop: '1px solid #e2e8f0', background: 'white' }}>
        <div className="input-row">
          <div className="input-field-container">
            <textarea
              ref={inputRef}
              className="chat-input"
              value={inputText}
              onChange={(e) => setInputText(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder="여기에 요청사항을 입력하세요..."
              rows={inputRows}
              disabled={isLoading && false}
            />
          </div>
          <button
            className="send-btn-circle"
            onClick={handleSendOrStop}
            disabled={!isLoading && !inputText.trim()}
            style={{ 
              width: '44px', 
              height: '44px',
              transition: 'transform 0.2s cubic-bezier(0.4, 0, 0.2, 1)',
              background: isLoading ? '#ef4444' : 'var(--primary)',
              boxShadow: isLoading ? '0 4px 12px rgba(239, 68, 68, 0.3)' : '0 4px 12px rgba(79, 70, 229, 0.3)',
              flexShrink: 0
            }}
          >
            {isLoading ? (
              <svg width="18" height="18" viewBox="0 0 24 24" fill="white">
                <rect x="5" y="5" width="14" height="14" rx="2" />
              </svg>
            ) : (
              <svg 
                width="22" 
                height="22" 
                viewBox="0 0 24 24" 
                fill="none" 
                stroke="currentColor" 
                strokeWidth="2.5" 
                style={{ transform: 'rotate(15deg) translate(-1px, 1px)' }}
              >
                <line x1="22" y1="2" x2="11" y2="13"/>
                <polygon points="22 2 15 22 11 13 2 9 22 2"/>
              </svg>
            )}
          </button>
        </div>
      </div>
      
      {/* Hidden button for App.tsx to clear */}
      <button className="clear-btn" onClick={clearChat} style={{ display: 'none' }}>초기화</button>
    </div>
  );
};
