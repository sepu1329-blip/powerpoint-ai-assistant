/**
 * Office.js PowerPoint 컨텍스트 서비스
 * - 현재 슬라이드/선택 객체 정보 읽기
 * - 슬라이드 조작 (텍스트 삽입, 슬라이드 추가 등)
 */

import type {
  PresentationContext,
  SlideInfo,
  ShapeInfo,
  SlideAction,
  InsertTextBoxParams,
  AddSlideParams,
  SetBackgroundParams,
  InsertTableParams,
  UpdateTextParams,
  FormatShapeParams,
  DuplicateShapeParams,
  DuplicateShapeOoxmlParams,
  BorderStyleParams,
  NeonEffectParams,
  ShadowEffectParams,
  ReflectionEffectParams,
  SoftEdgeParams,
} from '../types';
import {
  duplicateShapeOoxml,
  setBorderStyle,
  setNeonEffect,
  setShadowEffect,
  setReflectionEffect,
  setSoftEdge,
} from './ooxmlService';

// Office.js가 로드되었는지 확인
const isOfficeAvailable = (): boolean => {
  return typeof Office !== 'undefined' && typeof PowerPoint !== 'undefined';
};

/**
 * 현재 프레젠테이션 컨텍스트를 읽어옵니다.
 */
export async function getPresentationContext(): Promise<PresentationContext> {
  if (!isOfficeAvailable()) {
    return getMockContext();
  }

  return new Promise((resolve, reject) => {
    PowerPoint.run(async (context) => {
      try {
        const presentation = context.presentation;
        const slides = presentation.slides;
        slides.load('items');
        
        const selectedShapes = presentation.getSelectedShapes();
        selectedShapes.load('items');
        
        const selectedSlides = presentation.getSelectedSlides();
        selectedSlides.load('items');
        
        await context.sync();

        // 1. 필요한 기본 속성 로드
        for (const slide of selectedSlides.items) {
          slide.load('id,shapes/items');
        }
        for (const shape of selectedShapes.items) {
          shape.load('id,name,type,left,top,width,height');
        }
        await context.sync();

        // 2. 모든 도형의 기본 속성 및 텍스트 존재 여부 로드
        for (const slide of selectedSlides.items) {
          for (const s of slide.shapes.items) {
            s.load('id,name,type');
          }
        }
        await context.sync();

        // 3. 텍스트 프레임이 있는 경우 텍스트 내용까지 예약 로드
        // (sync를 최소화하기 위해 모든 가능한 경우를 예약)
        for (const slide of selectedSlides.items) {
          for (const s of slide.shapes.items) {
            try {
              if ((s as any).textFrame) {
                s.textFrame.textRange.load('text');
              }
            } catch (e) {}
          }
        }
        for (const shape of selectedShapes.items) {
          try {
            if ((shape as any).textFrame) {
              shape.textFrame.textRange.load('text');
            }
          } catch (e) {}
        }
        // 딱 한 번의 sync로 로딩 완료!
        await context.sync();

        // 4. 데이터 조립
        const selectedSlideInfos: SlideInfo[] = [];
        const allSlidesRaw = slides.items;

        for (const activeSlide of selectedSlides.items) {
          let slideIdx = -1;
          for (let i = 0; i < allSlidesRaw.length; i++) {
            if (allSlidesRaw[i].id === activeSlide.id) { slideIdx = i; break; }
          }

          const slideShapeInfos: ShapeInfo[] = [];
          for (const s of activeSlide.shapes.items) {
            let text: string | undefined;
            try { text = s.textFrame.textRange.text; } catch (e) {}
            slideShapeInfos.push({ id: s.id, name: s.name, type: s.type?.toString() ?? 'Unknown', text });
          }
          selectedSlideInfos.push({ index: slideIdx, shapes: slideShapeInfos });
        }

        const selectedShapeInfos: ShapeInfo[] = [];
        for (const shape of selectedShapes.items) {
          let text: string | undefined;
          try { text = shape.textFrame.textRange.text; } catch (e) {}
          selectedShapeInfos.push({
            id: shape.id, name: shape.name, type: shape.type?.toString() ?? 'Unknown', text,
            left: shape.left, top: shape.top, width: shape.width, height: shape.height
          });
        }

        resolve({
          slideCount: slides.items.length,
          currentSlideIndex: selectedSlideInfos.length > 0 ? selectedSlideInfos[0].index : 0,
          currentSlide: selectedSlideInfos.length > 0 ? selectedSlideInfos[0] : null,
          selectedShapes: selectedShapeInfos,
          selectedSlides: selectedSlideInfos
        });
      } catch (error) {
        console.error('getPresentationContext error:', error);
        reject(error);
      }
    });
  });
}

/**
 * 개발/테스트용 Mock 컨텍스트
 */
function getMockContext(): PresentationContext {
  return {
    slideCount: 5,
    currentSlideIndex: 0,
    currentSlide: {
      index: 0,
      title: '제목 슬라이드',
      shapes: [
        { id: '1', name: 'Title 1', type: 'TextBox', text: 'PowerPoint AI Assistant' },
        { id: '2', name: 'Subtitle 2', type: 'TextBox', text: '여기에 부제목을 입력하세요' },
      ],
      layoutName: 'Title Slide',
    },
    selectedShapes: [],
    selectedText: undefined,
    themeName: 'Office Theme',
  };
}

/**
 * 컨텍스트를 사람이 읽기 쉬운 문자열로 변환
 */
export function contextToString(context: PresentationContext): string {
  const lines: string[] = [
    `[프레젠테이션 정보]`,
    `- 전체 슬라이드 수: ${context.slideCount}`,
  ];

  if (context.selectedSlides && context.selectedSlides.length > 1) {
    lines.push(`- 현재 여러 슬라이드가 선택됨: ${context.selectedSlides.length}개 슬라이드`);
    context.selectedSlides.forEach(s => {
      lines.push(``, `[슬라이드 ${s.index + 1}번]`);
      s.shapes.forEach(sh => {
        lines.push(`- ID:[${sh.id}] ${sh.name}(${sh.type})${sh.text ? `: "${sh.text}"` : ''}`);
      });
    });
  } else if (context.currentSlide) {
    const slide = context.currentSlide;
    lines.push(`- 현재 슬라이드: ${slide.index + 1}번`);
    lines.push(``, `[현재 슬라이드 도형/텍스트]`);
    slide.shapes.forEach(s => {
      const isSelected = context.selectedShapes.some(ss => ss.id === s.id);
      lines.push(`- ID:[${s.id}] ${s.name}(${s.type})${s.text ? `: "${s.text}"` : ''} ${isSelected ? '(현재 선택됨, 바로 수정 가능)' : ''}`);
    });
  }

  if (context.selectedShapes.length > 0) {
    lines.push(``, `[유저가 명시적으로 선택한 도형 (최우선 수정 대상)]`);
    context.selectedShapes.forEach((s) => {
      lines.push(`- ${s.name}: "${s.text ?? '(텍스트 없음)'}"`);
    });
  }

  if (context.selectedText) {
    lines.push(``, `[선택된 텍스트]`, `"${context.selectedText}"`);
  }

  return lines.filter((l) => l !== null).join('\n');
}

// ===========================
// 슬라이드 조작 함수들
// ===========================

/**
 * 텍스트 박스 삽입
 */
export async function insertTextBox(params: InsertTextBoxParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] insertTextBox:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0);
    const textBox = slide.shapes.addTextBox(params.text, {
      left: params.left ?? 100,
      top: params.top ?? 100,
      width: params.width ?? 400,
      height: params.height ?? 100,
    });

    if (params.fontSize || params.bold || params.color) {
      const tf = textBox.textFrame;
      tf.load('textRanges');
      await context.sync();
    }

    await context.sync();
  });
}

/**
 * 새 슬라이드 추가
 */
export async function addNewSlide(params: AddSlideParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] addNewSlide:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    context.presentation.slides.add();
    await context.sync();

    // 마지막 슬라이드에 제목 추가
    if (params.title) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();

      const lastSlide = slides.items[slides.items.length - 1];
      lastSlide.shapes.addTextBox(params.title, {
        left: 50,
        top: 50,
        width: 600,
        height: 80,
      });
      await context.sync();
    }
  });
}

/**
 * 배경색 변경
 */
export async function setSlideBackground(params: SetBackgroundParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] setSlideBackground:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    // 현재 슬라이드 배경 변경
    const slide = slides.items[0];
    slide.load('background');
    await context.sync();

    await context.sync();
  });
}

/**
 * 텍스트 업데이트
 */
export async function updateText(params: UpdateTextParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] updateText:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    // 1. 대상 슬라이드 가져오기
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] || slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      if (selectedSlides.items.length === 0) throw new Error('작업할 슬라이드를 선택해 주세요.');
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();

    // 2. 검색을 위한 ID, Name 로드
    for (const s of shapes.items) { s.load('id,name,textFrame'); }
    await context.sync();

    // 3. 대상 도형 검색
    let targetShape = params.shapeId
      ? shapes.items.find((s) => s.id === params.shapeId || s.name === params.shapeId)
      : shapes.items[params.shapeIndex ?? 0];

    // 4. 검색 실패 시 선택된 도형에서 시도 (폴백)
    if (!targetShape) {
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load('items');
      await context.sync();
      if (selectedShapes.items.length > 0) {
        targetShape = selectedShapes.items[0];
        targetShape.load('id,name,textFrame');
        await context.sync();
      }
    }

    if (!targetShape) {
      throw new Error(`수정할 대상을 찾을 수 없습니다: ${params.shapeId || 'Index ' + (params.shapeIndex ?? 0)}`);
    }

    // 5. 텍스트 업데이트 적용
    if (targetShape.textFrame) {
      // 텍스트 강제 적용 (더 확실한 방식)
      targetShape.textFrame.textRange.text = params.newText;
      await context.sync();
      console.log(`Updated text for ${targetShape.id} to ${params.newText}`);
    } else {
      throw new Error('선택한 객체에 텍스트를 입력할 수 없습니다. (텍스트 프레임 없음)');
    }
  });
}

/**
 * 테이블 삽입
 */
export async function insertTable(params: InsertTableParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] insertTable:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(0);
    slide.shapes.addTable(params.rowCount, params.columnCount, {
      left: params.left ?? 100,
      top: params.top ?? 200,
      width: params.width ?? 500,
      height: params.height ?? 300,
    });
    await context.sync();
  });
}

/**
 * 도형 및 텍스트 서식 변경 (채우기, 테두리, 정렬 등)
 */
export async function formatShape(params: FormatShapeParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] formatShape:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    let targetShapes: any[] = [];

    // 1. 대상 도형 가져오기
    if (params.shapeId || params.shapeIndex !== undefined) {
      let slide: PowerPoint.Slide;
      if (params.slideIndex !== undefined) {
        const slides = context.presentation.slides;
        slides.load('items');
        await context.sync();
        slide = slides.items[params.slideIndex] || slides.items[0];
      } else {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load('items');
        await context.sync();
        if (selectedSlides.items.length === 0) return;
        slide = selectedSlides.items[0];
      }

      const shapes = slide.shapes;
      shapes.load('items');
      await context.sync();

      for (const s of shapes.items) { s.load('id,name'); }
      await context.sync();

      const ts = params.shapeId 
        ? shapes.items.find((s: any) => s.id === params.shapeId || s.name === params.shapeId)
        : shapes.items[params.shapeIndex ?? 0];
        
      if (ts) {
        targetShapes.push(ts);
      } else {
        // Fallback to currently selected shape if AI hallucinated the ID/Name
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load('items');
        await context.sync();
        if (selectedShapes.items.length > 0) {
          targetShapes = [...selectedShapes.items];
        }
      }
    } else {
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load('items');
      await context.sync();
      targetShapes = selectedShapes.items;
    }

    if (targetShapes.length === 0) {
      throw new Error(`수정할 대상을 찾을 수 없습니다: ${params.shapeId || '현재 선택된 객체 없음'}`);
    }

    // 2. 필요한 속성 로드
    for (const shape of targetShapes) {
      shape.load('fill,lineFormat,textFrame');
    }
    await context.sync();

    // 3. 서식 적용
    for (const shape of targetShapes) {
      if (params.fillColor && shape.fill) {
        const color = params.fillColor.startsWith('#') ? params.fillColor : `#${params.fillColor}`;
        shape.fill.setSolidColor(color);
      }

      if (shape.lineFormat) {
        if (params.outlineColor) {
          const color = params.outlineColor.startsWith('#') ? params.outlineColor : `#${params.outlineColor}`;
          shape.lineFormat.color = color;
        }
        if (params.outlineWidth !== undefined) {
          shape.lineFormat.weight = params.outlineWidth;
        }
      }

      if ((params.textAlign || params.textVerticalAlign) && shape.textFrame) {
        const textFrame = shape.textFrame;
        if (params.textAlign) {
          const alignMap: Record<string, string> = {
            'left': 'Left', 'center': 'Center', 'right': 'Right', 'justify': 'Justify'
          };
          const val = alignMap[params.textAlign.toLowerCase()] || params.textAlign;
          textFrame.textRange.paragraphFormat.horizontalAlignment = val as any;
        }
        if (params.textVerticalAlign) {
          const vAlignMap: Record<string, string> = {
            'top': 'Top', 'middle': 'Middle', 'bottom': 'Bottom'
          };
          const vVal = vAlignMap[params.textVerticalAlign.toLowerCase()] || params.textVerticalAlign;
          textFrame.verticalAlignment = vVal as any;
        }
      }
    }
    await context.sync();
  });
}


/**
 * 도형 복제
 */
export async function duplicateShape(params: DuplicateShapeParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] duplicateShape:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] || slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      if (selectedSlides.items.length === 0) return;
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();

    for (const s of shapes.items) { s.load('id,name'); }
    await context.sync();

    let targetShape = params.shapeId 
      ? shapes.items.find((s: any) => s.id === params.shapeId || s.name === params.shapeId)
      : shapes.items[params.shapeIndex ?? 0];

    if (!targetShape) {
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load('items');
      await context.sync();
      if (selectedShapes.items.length > 0) {
        targetShape = selectedShapes.items[0];
      }
    }

    if (targetShape) {
      console.log('Duplicating shape manually:', targetShape.id);
      
      // Office.js PowerPoint API는 shape.duplicate()를 지원하지 않으므로 수동으로 복제 (크기, 텍스트)
      targetShape.load('left,top,width,height,type,fill/foregroundColor,lineFormat/color');
      await context.sync();
      
      let textToCopy = '';
      try {
        if ((targetShape as any).textFrame) {
          targetShape.textFrame.textRange.load('text');
          await context.sync();
          textToCopy = targetShape.textFrame.textRange.text;
        }
      } catch(e) {}

      // 가능한 한 원본과 동일한 도형 생성 (기본은 사각형/텍스트박스)
      if (targetShape.type === PowerPoint.ShapeType.geometricShape) {
        const dup = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: targetShape.left + (params.offsetX ?? 20),
          top: targetShape.top + (params.offsetY ?? 20),
          width: targetShape.width,
          height: targetShape.height
        });
        if (textToCopy) {
           dup.textFrame.textRange.text = textToCopy;
        }
        try {
          if (targetShape.fill && targetShape.fill.foregroundColor) {
            dup.fill.setSolidColor(targetShape.fill.foregroundColor);
          }
          if (targetShape.lineFormat && targetShape.lineFormat.color) {
            dup.lineFormat.color = targetShape.lineFormat.color;
          }
        } catch(e) {}
      } else {
        slide.shapes.addTextBox(textToCopy, {
          left: targetShape.left + (params.offsetX ?? 20),
          top: targetShape.top + (params.offsetY ?? 20),
          width: targetShape.width,
          height: targetShape.height
        });
      }
      await context.sync();
    }
  });
}

/**
 * AI 액션 실행 (단일)
 */
export async function executeAction(action: SlideAction): Promise<void> {
  console.log('Executing action:', action.type, action.params);
  switch (action.type) {
    case 'INSERT_TEXT_BOX':
      await insertTextBox(action.params as InsertTextBoxParams);
      break;
    case 'UPDATE_TEXT':
      await updateText(action.params as UpdateTextParams);
      break;
    case 'ADD_SLIDE':
      await addNewSlide(action.params as AddSlideParams);
      break;
    case 'SET_BACKGROUND':
      await setSlideBackground(action.params as SetBackgroundParams);
      break;
    case 'INSERT_TABLE':
      await insertTable(action.params as InsertTableParams);
      break;
    case 'FORMAT_SHAPE':
      await formatShape(action.params as FormatShapeParams);
      break;
    case 'DUPLICATE_SHAPE':
      await duplicateShape(action.params as DuplicateShapeParams);
      break;
    // OOXML 기반 고급 액션 (Office.js API 미지원)
    case 'DUPLICATE_SHAPE_OOXML':
      await duplicateShapeOoxml(action.params as DuplicateShapeOoxmlParams);
      break;
    case 'SET_BORDER_STYLE':
      await setBorderStyle(action.params as BorderStyleParams);
      break;
    case 'SET_NEON_EFFECT':
      await setNeonEffect(action.params as NeonEffectParams);
      break;
    case 'SET_SHADOW_EFFECT':
      await setShadowEffect(action.params as ShadowEffectParams);
      break;
    case 'SET_REFLECTION_EFFECT':
      await setReflectionEffect(action.params as ReflectionEffectParams);
      break;
    case 'SET_SOFT_EDGE':
      await setSoftEdge(action.params as SoftEdgeParams);
      break;
    default:
      console.warn('지원하지 않는 액션 타입:', action.type);
  }
}

/**
 * AI 액션 목록 일괄 실행
 */
export async function executeActions(actions: SlideAction[]): Promise<void> {
  for (const action of actions) {
    await executeAction(action);
  }
}
