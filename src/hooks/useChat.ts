import { useState, useCallback, useRef } from 'react';
import type { ChatMessage, SlideAction, PresentationContext } from '../types';
import { sendMessage, isGeminiReady } from '../services/gemini';
import { getPresentationContext, executeAction } from '../services/officeContext';

export function useChat() {
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'welcome',
      role: 'assistant',
      content: '안녕하세요! PowerPoint AI Assistant입니다 👋\n\n슬라이드 편집, 내용 개선, 디자인 제안 등 무엇이든 도와드리겠습니다.\n\n**사용 예시:**\n- "현재 슬라이드 제목을 더 임팩트 있게 바꿔줘"\n- "이 슬라이드에 3열 비교표를 추가해줘"\n- "새 슬라이드를 만들고 주요 성과 3가지를 정리해줘"',
      timestamp: new Date(),
    },
  ]);
  const [isLoading, setIsLoading] = useState(false);
  const [statusText, setStatusText] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [context, setContext] = useState<PresentationContext | null>(null);

  const abortRef = useRef<boolean>(false);

  // 컨텍스트 새로고침
  const refreshContext = useCallback(async () => {
    try {
      const ctx = await getPresentationContext();
      setContext(ctx);
      return ctx;
    } catch (err) {
      console.error('컨텍스트 로드 실패:', err);
      return null;
    }
  }, []);

  // 메시지 전송
  const sendUserMessage = useCallback(
    async (text: string, useContext: boolean = true) => {
      if (!text.trim() || isLoading) return;
      if (!isGeminiReady()) {
        setError('API 키를 먼저 설정해주세요. 우측 상단 설정 버튼을 눌러주세요.');
        return;
      }

      setError(null);
      abortRef.current = false;
      setStatusText('PowerPoint 슬라이드 분석 중...');

      // 사용자 메시지 추가
      const userMsg: ChatMessage = {
        id: `user-${Date.now()}`,
        role: 'user',
        content: text,
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, userMsg]);

      // AI 응답 메시지 준비 (스트리밍)
      const assistantMsgId = `assistant-${Date.now()}`;
      const assistantMsg: ChatMessage = {
        id: assistantMsgId,
        role: 'assistant',
        content: '',
        timestamp: new Date(),
        actions: [],
      };
      setMessages((prev) => [...prev, assistantMsg]);

      setIsLoading(true);

      try {
        // 컨텍스트 가져오기
        const ctx = useContext ? await refreshContext() : null;
        setStatusText('AI 응답 생성 중...');

        // 대화 히스토리 구성 (Gemini 형식)
        const history = messages
          .filter((m) => m.role !== 'system' && m.id !== 'welcome')
          .slice(-10) 
          .map((m) => ({
            role: m.role === 'user' ? 'user' : 'model' as 'user' | 'model',
            parts: [{ text: m.content }],
          }));

        await sendMessage(
          text,
          ctx,
          history,
          (chunk) => {
            if (abortRef.current) return;
            setStatusText('AI가 슬라이드 변경 사항을 구성 중...');
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantMsgId
                  ? { ...m, content: m.content + chunk }
                  : m
              )
            );
          }
        ).then(({ text: fullText, actions }) => {
          if (abortRef.current) return;
          
          let finalContent = fullText;
          if (!finalContent && actions && actions.length > 0) {
            finalContent = '요청하신 슬라이드 작업을 준비했습니다. 아래 버튼을 눌러 적용해 보세요.';
          } else if (!finalContent) {
            finalContent = '죄송합니다. 응답을 생성하는 중 오류가 발생했거나 내용이 비어 있습니다.';
          }

          setMessages((prev) =>
            prev.map((m) =>
              m.id === assistantMsgId
                ? { ...m, content: finalContent, actions, applied: false }
                : m
            )
          );
        });
      } catch (err) {
        const errMsg = err instanceof Error ? err.message : '알 수 없는 오류가 발생했습니다.';
        setError(errMsg);
        setMessages((prev) => prev.filter((m) => m.id !== assistantMsgId));
      } finally {
        setIsLoading(false);
        setStatusText(null);
      }
    },
    [isLoading, messages, refreshContext]
  );

  // 액션 적용 실행
  const markActionsApplied = useCallback(async (messageId: string) => {
    const msg = messages.find(m => m.id === messageId);
    if (!msg || !msg.actions || msg.applied) return;

    setIsLoading(true);
    setStatusText('슬라이드에 변경 사항 적용 중...');

    try {
      for (const action of msg.actions) {
        await executeAction(action);
      }
      
      setMessages((prev) =>
        prev.map((m) =>
          m.id === messageId ? { ...m, applied: true } : m
        )
      );
    } catch (err) {
      console.error('액션 실행 실패:', err);
      setError('슬라이드 수정 중 오류가 발생했습니다.');
    } finally {
      setIsLoading(false);
      setStatusText(null);
      await refreshContext();
    }
  }, [messages, refreshContext]);

  // 대화 초기화
  const clearChat = useCallback(() => {
    setMessages([
      {
        id: 'welcome',
        role: 'assistant',
        content: '대화가 초기화되었습니다. 새로운 요청을 입력해주세요!',
        timestamp: new Date(),
      },
    ]);
    setError(null);
    setStatusText(null);
  }, []);

  const stopGeneration = useCallback(() => {
    abortRef.current = true;
    setIsLoading(false);
    setStatusText(null);
  }, []);

  return {
    messages,
    isLoading,
    statusText,
    error,
    context,
    sendUserMessage,
    markActionsApplied,
    clearChat,
    stopGeneration,
    refreshContext,
  };
}
