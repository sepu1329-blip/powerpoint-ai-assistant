// ===========================
// 메시지 타입
// ===========================
export type MessageRole = 'user' | 'assistant' | 'system';

export type AgentStepStatus = 'running' | 'done' | 'error';

export interface AgentStep {
  id: string;
  label: string;          // 작업 단계 설명 (예: "슬라이드 분석 중...")
  status: AgentStepStatus;
  detail?: string;        // 추가 상세 정보 (선택)
}

export interface ChatMessage {
  id: string;
  role: MessageRole;
  content: string;
  timestamp: Date;
  actions?: SlideAction[];   // AI가 제안하는 슬라이드 수정 액션
  applied?: boolean;         // 액션 적용 여부
  steps?: AgentStep[];       // AI 에이전트 작업 단계 로그
}

// ===========================
// AI 슬라이드 액션 타입
// ===========================
export type SlideActionType =
  | 'INSERT_TEXT_BOX'
  | 'UPDATE_TEXT'
  | 'ADD_SLIDE'
  | 'SET_BACKGROUND'
  | 'INSERT_TABLE'
  | 'INSERT_SHAPE'
  | 'FORMAT_SHAPE'
  | 'DUPLICATE_SHAPE'
  // OOXML 기반 고급 액션 (Office.js API 미지원 기능)
  | 'DUPLICATE_SHAPE_OOXML'
  | 'SET_BORDER_STYLE'
  | 'SET_NEON_EFFECT'
  | 'SET_SHADOW_EFFECT'
  | 'SET_REFLECTION_EFFECT'
  | 'SET_SOFT_EDGE';

export interface DuplicateShapeParams {
  slideIndex?: number;    // 0-based slide index
  shapeId?: string;
  shapeIndex?: number;
  offsetX?: number;
  offsetY?: number;
}

export interface InsertTextBoxParams {
  text: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fontSize?: number;
  bold?: boolean;
  color?: string;
}

export interface UpdateTextParams {
  slideIndex?: number;    // 0-based slide index
  shapeId?: string;
  shapeIndex?: number;
  newText: string;
}

export interface AddSlideParams {
  title?: string;
  content?: string;
  layoutName?: string;
}

export interface SetBackgroundParams {
  color: string; // hex color e.g. "#FF5733"
}

export interface InsertTableParams {
  rowCount: number;
  columnCount: number;
  data?: string[][];
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export interface FormatShapeParams {
  slideIndex?: number;    // 0-based slide index
  shapeId?: string;
  shapeIndex?: number;
  fillColor?: string;    // e.g. "#FF0000"
  outlineColor?: string; // e.g. "#000000"
  outlineWidth?: number; // width in points
  textAlign?: 'Left' | 'Center' | 'Right' | 'Justify';
  textVerticalAlign?: 'Top' | 'Middle' | 'Bottom';
}

// ===========================
// OOXML 기반 고급 액션 파라미터
// ===========================

export interface DuplicateShapeOoxmlParams {
  slideIndex?: number;
  shapeId?: string;
  shapeIndex?: number;
  offsetX?: number;
  offsetY?: number;
}

export interface BorderStyleParams {
  slideIndex?: number;
  shapeId?: string;
  shapeIndex?: number;
  /** 선 색상 (hex, e.g. "#FF0000") */
  color?: string;
  /** 선 두께 (포인트) */
  width?: number;
  /**
   * 선 스타일: 'solid' | 'dash' | 'dot' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot'
   * | 'sysDash' | 'sysDot' | 'sysDashDot' | 'sysDashDotDot'
   */
  dashStyle?: string;
  /** 끝단 캡 스타일: 'flat' | 'rnd' | 'sq' */
  capStyle?: string;
  /** 연결 스타일: 'bevel' | 'miter' | 'round' */
  joinStyle?: string;
}

export interface NeonEffectParams {
  slideIndex?: number;
  shapeId?: string;
  shapeIndex?: number;
  /** 네온 색상 (hex) */
  color: string;
  /** 네온 반경 (포인트, 기본값 8) */
  radius?: number;
  /** 네온 투명도 (0-100, 0=불투명, 기본값 40) */
  transparency?: number;
}

export interface ShadowEffectParams {
  slideIndex?: number;
  shapeId?: string;
  shapeIndex?: number;
  /** 그림자 종류: 'outer' | 'inner' */
  type?: string;
  /** 그림자 색상 (hex) */
  color?: string;
  /** 투명도 (0-100) */
  transparency?: number;
  /** 흐림 반경 (포인트) */
  blur?: number;
  /** 거리 (포인트) */
  distance?: number;
  /** 방향 (각도, 0-359) */
  direction?: number;
}

export interface ReflectionEffectParams {
  slideIndex?: number;
  shapeId?: string;
  shapeIndex?: number;
  /** 흐림 반경 (포인트) */
  blur?: number;
  /** 시작 투명도 (0-100) */
  startAlpha?: number;
  /** 끝 투명도 (0-100) */
  endAlpha?: number;
  /** 크기 비율 (퍼센트, 0-100) */
  size?: number;
  /** 거리 (포인트) */
  distance?: number;
}

export interface SoftEdgeParams {
  slideIndex?: number;
  shapeId?: string;
  shapeIndex?: number;
  /** 부드러운 가장자리 반경 (포인트) */
  radius?: number;
}

export interface SlideAction {
  id: string;
  type: SlideActionType;
  description: string;
  params: InsertTextBoxParams | UpdateTextParams | AddSlideParams | SetBackgroundParams | InsertTableParams | FormatShapeParams | DuplicateShapeOoxmlParams | BorderStyleParams | NeonEffectParams | ShadowEffectParams | ReflectionEffectParams | SoftEdgeParams | Record<string, unknown>;
}

// ===========================
// PPT 컨텍스트 타입
// ===========================
export interface ShapeInfo {
  id: string;
  name: string;
  type: string;
  text?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

export interface SlideInfo {
  index: number;
  title?: string;
  shapes: ShapeInfo[];
  layoutName?: string;
}

export interface PresentationContext {
  slideCount: number;
  currentSlideIndex: number;
  currentSlide: SlideInfo | null;
  selectedShapes: ShapeInfo[];
  selectedSlides?: SlideInfo[];
  selectedText?: string;
  themeName?: string;
}

// ===========================
// 설정 타입
// ===========================
export type GeminiModel =
  | 'gemini-2.0-flash'
  | 'gemini-2.5-pro-exp-03-25'
  | 'gemini-1.5-pro'
  | 'gemini-1.5-flash';

export interface CustomPrompt {
  id: string;
  name: string;
  content: string;
}

export interface AppSettings {
  apiKey: string;
  model: GeminiModel;
  autoContext: boolean;
  customPrompts: CustomPrompt[];
}


// ===========================
// UI 상태 타입
// ===========================
export type ViewMode = 'chat' | 'settings';

export interface AppState {
  view: ViewMode;
  isLoading: boolean;
  isOfficeReady: boolean;
  error: string | null;
}
