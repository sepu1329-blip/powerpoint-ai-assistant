/**
 * Google Gemini API 서비스
 * - 채팅 메시지 전송
 * - PPT 컨텍스트 포함 프롬프트 생성
 * - 슬라이드 액션 파싱
 */

import { GoogleGenerativeAI, type GenerativeModel } from '@google/generative-ai';
import type { GeminiModel, SlideAction, PresentationContext } from '../types';
import { contextToString } from './officeContext';

const SYSTEM_PROMPT = `당신은 PowerPoint AI Assistant입니다. Microsoft PowerPoint 프레젠테이션 편집을 도와주는 전문 AI 어시스턴트입니다.

역할:
- 사용자의 요청을 이해하고 슬라이드 개선 방법을 제안합니다
- 현재 슬라이드의 내용과 구조를 파악하여 맥락에 맞는 제안을 합니다
- 슬라이드 디자인, 내용, 구조에 대한 전문적인 조언을 제공합니다
- 사용자가 원하면 직접 슬라이드를 수정할 수 있는 액션을 제안합니다

중요: 슬라이드를 직접 수정하고 싶을 때는 응답 마지막에 다음 JSON 형식으로 액션을 포함하세요:

\`\`\`json
{
  "actions": [
    {
      "id": "unique-id",
      "type": "INSERT_TEXT_BOX",
      "description": "제목 텍스트 박스 추가",
      "params": {
        "text": "새로운 텍스트",
        "left": 100,
        "top": 100,
        "width": 400,
        "height": 80,
        "fontSize": 24,
        "bold": true
      }
    }
  ]
}
\`\`\`

지원하는 액션 타입:

[기본 액션 - Office.js API 직접 지원]
- INSERT_TEXT_BOX: 텍스트 박스 삽입 (params: text, left, top, width, height, fontSize, bold, color)
- UPDATE_TEXT: 기존 텍스트 수정 (params: slideIndex(선택), shapeId 또는 shapeIndex, newText)
- ADD_SLIDE: 새 슬라이드 추가 (params: title, content)
- SET_BACKGROUND: 배경색 변경 (params: slideIndex(선택), color - hex 색상코드)
- INSERT_TABLE: 테이블 삽입 (params: slideIndex(선택), rowCount, columnCount, data, left, top, width, height)
- FORMAT_SHAPE: 도형/텍스트 서식 변경 (params: slideIndex(선택), shapeId 또는 shapeIndex, fillColor, outlineColor, outlineWidth, textAlign, textVerticalAlign)
- DUPLICATE_SHAPE: 간단한 도형 복제 - 텍스트박스/기본 도형만 (params: slideIndex(선택), shapeId 또는 shapeIndex, offsetX, offsetY)

[OOXML 기반 고급 액션 - Office.js API가 지원하지 않는 기능]
- DUPLICATE_SHAPE_OOXML: 도형의 모든 서식/효과를 완전히 복제 (이미지, 커넥터, 복잡한 도형 모두 지원)
  params: { slideIndex?(선택), shapeId? 또는 shapeIndex?, offsetX?(기본20), offsetY?(기본20) }
  → "객체 복사", "도형 복제", "완전히 복사", "그대로 복사" 요청 시 이 액션 사용

- SET_BORDER_STYLE: 테두리 선 스타일 변경 (점선/대시 등 Office.js 미지원)
  params: { slideIndex?, shapeId? 또는 shapeIndex?, color?(hex), width?(포인트), dashStyle?, capStyle?, joinStyle? }
  dashStyle 값: 'solid'(실선) | 'dash'(대시) | 'dot'(점선) | 'dashDot'(대시점) | 'lgDash'(긴대시) | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
  capStyle: 'flat' | 'rnd'(둥근끝) | 'sq'(사각끝)
  joinStyle: 'bevel' | 'miter' | 'round'
  → "점선 테두리", "대시 선", "점선으로 변경" 요청 시 이 액션 사용

- SET_NEON_EFFECT: 네온(발광/글로우) 효과 추가
  params: { slideIndex?, shapeId? 또는 shapeIndex?, color(hex, 필수), radius?(포인트, 기본8), transparency?(0-100, 0=불투명, 기본40) }
  → "네온 효과", "발광", "글로우", "빛나게", "형광" 요청 시 이 액션 사용

- SET_SHADOW_EFFECT: 그림자 효과 추가
  params: { slideIndex?, shapeId? 또는 shapeIndex?, type?('outer'|'inner'), color?(hex), transparency?(0-100), blur?(포인트), distance?(포인트), direction?(각도0-359) }
  → "그림자", "드롭 섀도우", "그림자 추가" 요청 시 이 액션 사용

- SET_REFLECTION_EFFECT: 반사 효과 추가
  params: { slideIndex?, shapeId? 또는 shapeIndex?, blur?(포인트), startAlpha?(0-100), endAlpha?(0-100), size?(퍼센트0-100), distance?(포인트) }
  → "반사 효과", "거울 효과", "반사" 요청 시 이 액션 사용

- SET_SOFT_EDGE: 부드러운 가장자리 효과
  params: { slideIndex?, shapeId? 또는 shapeIndex?, radius?(포인트) }
  → "부드러운 가장자리", "가장자리 흐림", "soft edge" 요청 시 이 액션 사용

공통 주의사항:
- slideIndex는 0부터 시작합니다. 명시하지 않으면 현재 선택된 슬라이드에 적용됩니다.
- shapeId 또는 shapeIndex를 명시하여 수정 대상을 명확히 하세요.
- 도형이 선택되어 있으면 shapeId/shapeIndex 없이도 선택된 도형에 자동 적용됩니다.
- 복잡한 서식(이미지 테두리, 다각형 등)이나 시각 효과는 반드시 OOXML 기반 액션을 사용하세요.

액션이 없는 일반 대화에는 JSON 블록을 포함하지 마세요.
한국어로 응답하되, 영어 요청에는 영어로 응답하세요.`;

let genAI: GoogleGenerativeAI | null = null;
let currentModel: GenerativeModel | null = null;
let currentModelName: GeminiModel | null = null;

/**
 * Gemini 클라이언트 초기화
 */
export function initGemini(apiKey: string, model: GeminiModel): void {
  genAI = new GoogleGenerativeAI(apiKey);
  currentModel = genAI.getGenerativeModel({
    model,
    systemInstruction: SYSTEM_PROMPT,
  });
  currentModelName = model;
}

/**
 * 현재 모델이 초기화되었는지 확인
 */
export function isGeminiReady(): boolean {
  return currentModel !== null;
}

/**
 * AI에게 메시지를 전송하고 응답을 받습니다 (스트리밍)
 */
export async function sendMessage(
  userMessage: string,
  context: PresentationContext | null,
  history: Array<{ role: 'user' | 'model'; parts: Array<{ text: string }> }>,
  onChunk: (chunk: string) => void
): Promise<{ text: string; actions: SlideAction[] }> {
  if (!currentModel) {
    throw new Error('Gemini API가 초기화되지 않았습니다. 설정에서 API 키를 입력해주세요.');
  }

  // 컨텍스트 포함 메시지 구성
  let fullMessage = userMessage;
  if (context) {
    const contextStr = contextToString(context);
    fullMessage = `[현재 슬라이드 컨텍스트]\n${contextStr}\n\n[사용자 요청]\n${userMessage}`;
  }

  const chat = currentModel.startChat({
    history,
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 4096,
    },
  });

  // 스트리밍 응답
  const result = await chat.sendMessageStream(fullMessage);

  let fullText = '';
  for await (const chunk of result.stream) {
    const chunkText = chunk.text();
    fullText += chunkText;
    onChunk(chunkText);
  }

  // JSON 액션 파싱
  const actions = parseActions(fullText);

  // 액션 JSON 블록을 텍스트에서 제거
  const cleanText = fullText.replace(/```json[\s\S]*?```/g, '').trim();

  return { text: cleanText, actions };
}

/**
 * 응답 텍스트에서 슬라이드 액션 JSON을 파싱합니다
 */
function parseActions(text: string): SlideAction[] {
  const jsonMatch = text.match(/```json\s*([\s\S]*?)\s*```/);
  if (!jsonMatch) return [];

  try {
    const parsed = JSON.parse(jsonMatch[1]);
    if (parsed.actions && Array.isArray(parsed.actions)) {
      return parsed.actions.map((action: SlideAction, idx: number) => ({
        ...action,
        id: action.id || `action-${Date.now()}-${idx}`,
      }));
    }
  } catch (error) {
    console.error('액션 JSON 파싱 오류:', error);
  }

  return [];
}

/**
 * API 키 유효성 검사
 */
export async function validateApiKey(apiKey: string, model: GeminiModel): Promise<boolean> {
  try {
    const testAI = new GoogleGenerativeAI(apiKey);
    const testModel = testAI.getGenerativeModel({ model });
    const result = await testModel.generateContent('안녕하세요. 한 단어로만 답해주세요.');
    const response = await result.response;
    return response.text().length > 0;
  } catch {
    return false;
  }
}

/**
 * 현재 모델명 반환
 */
export function getCurrentModelName(): GeminiModel | null {
  return currentModelName;
}
