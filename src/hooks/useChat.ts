import { useState, useCallback, useRef } from 'react';
import type { ChatMessage, SlideAction, PresentationContext, AgentStep } from '../types';
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
  // 현재 assistant 메시지 ID 추적 (중지 시 steps 정리용)
  const currentMsgIdRef = useRef<string | null>(null);

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

      // AI 응답 메시지 준비 (스트리밍 + steps 초기화)
      const assistantMsgId = `assistant-${Date.now()}`;
      currentMsgIdRef.current = assistantMsgId;
      const initialSteps: AgentStep[] = [
        { id: 'step-ctx', label: 'PowerPoint 슬라이드 콘텐츠 분석', status: 'running' },
      ];
      const assistantMsg: ChatMessage = {
        id: assistantMsgId,
        role: 'assistant',
        content: '',
        timestamp: new Date(),
        actions: [],
        steps: initialSteps,
      };
      setMessages((prev) => [...prev, assistantMsg]);

      // steps 업데이트 헬퍼
      const updateStep = (stepId: string, patch: Partial<AgentStep>) => {
        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistantMsgId
              ? {
                  ...m,
                  steps: (m.steps ?? []).map((s) =>
                    s.id === stepId ? { ...s, ...patch } : s
                  ),
                }
              : m
          )
        );
      };

      const addStep = (step: AgentStep) => {
        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistantMsgId
              ? { ...m, steps: [...(m.steps ?? []), step] }
              : m
          )
        );
      };

      setIsLoading(true);

      try {
        // 콘텍스트 가져오기
        const ctx = useContext ? await refreshContext() : null;
        updateStep('step-ctx', {
          status: 'done',
          detail: ctx
            ? `슬라이드 ${(ctx.currentSlideIndex ?? 0) + 1}/${ctx.slideCount}종, 객체 ${ctx.currentSlide?.shapes?.length ?? 0}개 인식`
            : '콘텍스트 없음 (오프라인 모드)',
        });

        addStep({ id: 'step-ai', label: 'AI 응답 생성 중', status: 'running' });
        setStatusText('AI 응답 생성 중...');

        // 대화 히스토리 구성 (Gemini 형식)
        const history = messages
          .filter((m) => m.role !== 'system' && m.id !== 'welcome')
          .slice(-10) 
          .map((m) => ({
            role: m.role === 'user' ? 'user' : 'model' as 'user' | 'model',
            parts: [{ text: m.content }],
          }));

        let streamChunkCount = 0;
        await sendMessage(
          text,
          ctx,
          history,
          (chunk) => {
            if (abortRef.current) return;
            streamChunkCount++;
            if (streamChunkCount === 1) {
              updateStep('step-ai', { label: 'AI 응답 수신 중' });
            }
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

          // AI 응답 단계 완료 표시
          updateStep('step-ai', {
            status: 'done',
            label: 'AI 응답 완료',
            detail: fullText.length > 0 ? `${fullText.length}자 스트리밍` : undefined,
          });

          // 액션이 있으면 단계 추가
          if (actions && actions.length > 0) {
            addStep({
              id: 'step-actions',
              label: `슬라이드 수정 액션 ${actions.length}개 준비됨`,
              status: 'done',
              detail: actions.map((a) => a.description || a.type).join(', '),
            });
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
        updateStep('step-ai', { status: 'error', detail: errMsg });
        setError(errMsg);
        setMessages((prev) => prev.filter((m) => m.id !== assistantMsgId));
      } finally {
        currentMsgIdRef.current = null;
        setIsLoading(false);
        setStatusText(null);
      }
    },
    [isLoading, messages, refreshContext]
  );

  // 액션 적용 실행 (isLoading과 독립적으로 동작)
  const markActionsApplied = useCallback(async (messageId: string) => {
    // 이미 applied된 메시지이거나 없으면 무시
    let targetMsg: ChatMessage | undefined;
    setMessages(prev => {
      targetMsg = prev.find(m => m.id === messageId);
      return prev;
    });
    // 최신 messages에서 직접 찾기
    targetMsg = messages.find(m => m.id === messageId);
    if (!targetMsg || !targetMsg.actions || targetMsg.applied) return;

    const actionsToRun = targetMsg.actions;
    const applyStepId = `step-apply-${Date.now()}`;

    // 낙관적 업데이트: applied = true로 먼저 마킹하여 중복 실행 방지
    setMessages((prev) =>
      prev.map((m) =>
        m.id === messageId
          ? {
              ...m,
              applied: true,
              steps: [
                ...(m.steps ?? []),
                { id: applyStepId, label: '슬라이드에 변경 사항 적용 중...', status: 'running' as const },
              ],
            }
          : m
      )
    );

    try {
      for (const action of actionsToRun) {
        await executeAction(action);
      }

      setMessages((prev) =>
        prev.map((m) =>
          m.id === messageId
            ? {
                ...m,
                steps: (m.steps ?? []).map((s) =>
                  s.id === applyStepId
                    ? { ...s, status: 'done' as const, label: `슬라이드 적용 완료 (${actionsToRun.length}개)` }
                    : s
                ),
              }
            : m
        )
      );
    } catch (err) {
      console.error('액션 실행 실패:', err);
      const errMsg = err instanceof Error ? err.message : '알 수 없는 오류';
      // 실패 시 applied 원복 + 오류 표시
      setMessages((prev) =>
        prev.map((m) =>
          m.id === messageId
            ? {
                ...m,
                applied: false,
                steps: (m.steps ?? []).map((s) =>
                  s.id === applyStepId
                    ? { ...s, status: 'error' as const, label: '슬라이드 수정 실패', detail: errMsg }
                    : s
                ),
              }
            : m
        )
      );
      setError(`슬라이드 수정 중 오류: ${errMsg}`);
    } finally {
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
    // 현재 진행 중인 메시지의 running steps를 취소 상태로 변경
    const msgId = currentMsgIdRef.current;
    if (msgId) {
      setMessages((prev) =>
        prev.map((m) =>
          m.id === msgId
            ? {
                ...m,
                content: m.content || '생성이 중단되었습니다.',
                steps: (m.steps ?? []).map((s) =>
                  s.status === 'running'
                    ? { ...s, status: 'error' as const, label: s.label.replace('중...', '중단됨').replace('중', '중단됨'), detail: '사용자가 중단함' }
                    : s
                ),
              }
            : m
        )
      );
      currentMsgIdRef.current = null;
    }
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
