/**
 * OOXML 기반 고급 기능 서비스
 * Office.js API가 지원하지 않는 기능들을 OOXML을 직접 읽고 수정하여 구현합니다.
 *
 * 지원 기능:
 * - 도형 완전 복제 (OOXML 복사 방식)
 * - 테두리 선 스타일 변경 (점선, 대시 등)
 * - 네온(글로우) 효과 추가
 * - 그림자 효과 추가
 * - 반사 효과 추가
 * - 3D 회전 효과 추가
 */

// ===========================
// 파라미터 타입 정의
// ===========================

export interface OoxmlDuplicateParams {
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
  /** 선 두께 (포인트, 1pt = 12700 EMU) */
  width?: number;
  /**
   * 선 스타일:
   * 'solid' | 'dash' | 'dot' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot' | 'sysDashDot' | 'sysDashDotDot'
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
  /** 네온 색상 (hex, e.g. "#00FFFF") */
  color: string;
  /** 네온 반경 (포인트, 기본값 8) */
  radius?: number;
  /** 네온 투명도 (0-100, 기본값 60) */
  transparency?: number;
}

export interface ShadowEffectParams {
  slideIndex?: number;
  shapeId?: string;
  shapeIndex?: number;
  /** 그림자 종류: 'outer' | 'inner' | 'perspective' */
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
  /** 반사 방향 (각도) */
  direction?: number;
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

// ===========================
// 내부 유틸리티 함수
// ===========================

/** Office.js 가용 여부 확인 */
const isOfficeAvailable = (): boolean =>
  typeof Office !== 'undefined' && typeof PowerPoint !== 'undefined';

/** hex 색상을 RGB 배열로 변환 */
function hexToRgb(hex: string): [number, number, number] {
  const cleaned = hex.replace('#', '');
  const r = parseInt(cleaned.substring(0, 2), 16);
  const g = parseInt(cleaned.substring(2, 4), 16);
  const b = parseInt(cleaned.substring(4, 6), 16);
  return [r, g, b];
}

/** 투명도(0-100%)를 alpha 값(0-100000)으로 변환 */
function transparencyToAlpha(transparency: number): number {
  // transparency 0% → alpha 100000 (불투명)
  // transparency 100% → alpha 0 (완전 투명)
  return Math.round((100 - transparency) * 1000);
}

/** 포인트를 EMU(English Metric Unit)로 변환. 1pt = 12700 EMU */
function ptToEmu(pt: number): number {
  return Math.round(pt * 12700);
}



/**
 * 슬라이드 OOXML에서 특정 도형을 찾아 XML을 수정합니다.
 */
function modifyShapeInSlideOoxml(
  xml: string,
  shapeId: string,
  shapeName: string,
  modifier: (shapeXml: string, slideXml: string, shapeTag: string) => string
): string {
  // <p:sp>, <p:pic>, <p:graphicFrame>, <p:cxnSp> 등 다양한 도형 태그 처리
  const shapeTags = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp'];
  
  for (const tag of shapeTags) {
    const result = findAndModifyShapeBlock(xml, tag, shapeId, shapeName, modifier);
    if (result !== xml) return result; // 수정 성공
  }
  
  // 직접 찾기 방식으로 fallback
  console.warn('도형을 OOXML에서 찾을 수 없습니다:', shapeId, shapeName);
  return xml;
}

/**
 * XML 문자열에서 특정 태그 블록을 찾아 수정
 */
function findAndModifyShapeBlock(
  xml: string,
  tag: string,
  shapeId: string,
  shapeName: string,
  modifier: (shapeXml: string, slideXml: string, shapeTag: string) => string
): string {
  const openTag = `<${tag}`;
  const closeTag = `</${tag}>`;
  
  let searchStart = 0;
  while (true) {
    const startIdx = xml.indexOf(openTag, searchStart);
    if (startIdx === -1) break;
    
    // 중첩 태그 처리하며 블록 끝 찾기
    let depth = 1;
    let pos = startIdx + openTag.length;
    
    // 셀프 클로징 태그 체크
    const firstTagEnd = xml.indexOf('>', pos - 1);
    if (xml[firstTagEnd - 1] === '/') {
      // <p:sp ... /> 셀프 클로징
      searchStart = firstTagEnd + 1;
      continue;
    }
    
    while (depth > 0 && pos < xml.length) {
      const nextOpen = xml.indexOf(openTag, pos);
      const nextClose = xml.indexOf(closeTag, pos);
      
      if (nextClose === -1) break;
      
      if (nextOpen !== -1 && nextOpen < nextClose) {
        depth++;
        pos = nextOpen + openTag.length;
      } else {
        depth--;
        pos = nextClose + closeTag.length;
      }
    }
    
    const endIdx = pos;
    const shapeBlock = xml.substring(startIdx, endIdx);
    
    // id나 name이 일치하는지 확인
    const idMatch = shapeBlock.includes(`id="${shapeId}"`) || 
                   shapeBlock.includes(`id='${shapeId}'`);
    const nameMatch = shapeBlock.includes(`name="${shapeName}"`) || 
                     shapeBlock.includes(`name='${shapeName}'`) ||
                     (shapeName && shapeBlock.includes(shapeName));
    
    if (idMatch || nameMatch) {
      const modifiedBlock = modifier(shapeBlock, xml, tag);
      return xml.substring(0, startIdx) + modifiedBlock + xml.substring(endIdx);
    }
    
    searchStart = Math.min(startIdx + openTag.length, endIdx);
    if (searchStart >= xml.length) break;
  }
  
  return xml; // 변경 없음
}

/** 정규식 특수문자 이스케이프 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ===========================
// 고급 OOXML 헬퍼: 효과 XML 조각 생성
// ===========================

/**
 * 네온(글로우) 효과 XML 조각 생성
 * <a:outerShdw> 대신 <a:glow> 사용
 */
function buildNeonXml(color: string, radiusPt: number, transparency: number): string {
  const radiusEmu = ptToEmu(radiusPt);
  const [r, g, b] = hexToRgb(color);
  const hexColor = color.replace('#', '').toUpperCase().padEnd(6, '0');
  const alpha = transparencyToAlpha(transparency); // 0-100000
  
  return `<a:glow rad="${radiusEmu}"><a:srgbClr val="${hexColor}"><a:alpha val="${alpha}"/></a:srgbClr></a:glow>`;
}

/**
 * 그림자 효과 XML 조각 생성
 */
function buildShadowXml(params: ShadowEffectParams): string {
  const color = (params.color ?? '#000000').replace('#', '').toUpperCase();
  const alpha = transparencyToAlpha(params.transparency ?? 50);
  const blurEmu = ptToEmu(params.blur ?? 4);
  const distEmu = ptToEmu(params.distance ?? 3);
  const dir = Math.round((params.direction ?? 45) * 60000); // 각도를 60000분의 1도로
  
  if (params.type === 'inner') {
    return `<a:innerShdw blurRad="${blurEmu}" dist="${distEmu}" dir="${dir}"><a:srgbClr val="${color}"><a:alpha val="${alpha}"/></a:srgbClr></a:innerShdw>`;
  }
  
  // 기본: 외부 그림자
  return `<a:outerShdw blurRad="${blurEmu}" dist="${distEmu}" dir="${dir}" algn="tl" rotWithShape="0"><a:srgbClr val="${color}"><a:alpha val="${alpha}"/></a:srgbClr></a:outerShdw>`;
}

/**
 * 반사 효과 XML 조각 생성
 */
function buildReflectionXml(params: ReflectionEffectParams): string {
  const blurEmu = ptToEmu(params.blur ?? 0.5);
  const startAlpha = transparencyToAlpha(params.startAlpha ?? 0);   // 시작 투명도 → 불투명한 부분
  const endAlpha = transparencyToAlpha(params.endAlpha ?? 100);      // 끝 투명도 → 완전 투명
  const size = Math.round((params.size ?? 50) * 1000);               // 퍼센트 → 1/100000
  const dir = Math.round((params.direction ?? 5400000));             // 기본 90도 반사
  const distEmu = ptToEmu(params.distance ?? 0);
  
  return `<a:reflection blurRad="${blurEmu}" stA="${startAlpha}" stPos="0" endA="${endAlpha}" endPos="${size}" dist="${distEmu}" dir="${dir}" fadeDir="${dir}" rotWithShape="0"/>`;
}

/**
 * 부드러운 가장자리 XML 조각 생성
 */
function buildSoftEdgeXml(radiusPt: number): string {
  const radiusEmu = ptToEmu(radiusPt);
  return `<a:softEdge rad="${radiusEmu}"/>`;
}

/**
 * 도형의 <a:effectLst> 안에 효과를 추가/교체
 */
function upsertEffectInSpPr(shapeXml: string, effectXml: string, effectTag: string): string {
  // spPr 태그 찾기
  const spPrStart = shapeXml.indexOf('<p:spPr>') !== -1 ? '<p:spPr>' : '<p:spPr ';
  const spPrIdx = shapeXml.indexOf('<p:spPr');
  const spPrEndIdx = shapeXml.indexOf('</p:spPr>');
  
  if (spPrIdx === -1) {
    console.warn('spPr 태그를 찾을 수 없습니다.');
    return shapeXml;
  }
  
  const spPrContent = shapeXml.substring(spPrIdx, spPrEndIdx + '</p:spPr>'.length);
  
  // effectLst가 있는지 확인
  if (spPrContent.includes('<a:effectLst>')) {
    // effectLst 내에 해당 효과가 있으면 교체, 없으면 추가
    const effLstStart = spPrContent.indexOf('<a:effectLst>');
    const effLstEnd = spPrContent.indexOf('</a:effectLst>') + '</a:effectLst>'.length;
    let effLstContent = spPrContent.substring(effLstStart, effLstEnd);
    
    const closeEffTag = `</${effectTag}>`;
    const openEffTag = `<${effectTag}`;
    
    if (effLstContent.includes(openEffTag)) {
      // 기존 효과 교체
      const tagStart = effLstContent.indexOf(openEffTag);
      const tagEnd = effLstContent.indexOf('>', effLstContent.indexOf(openEffTag));
      // 셀프 클로징 여부 확인
      if (effLstContent[tagEnd - 1] === '/') {
        const selfClosingEnd = tagEnd + 1;
        effLstContent = effLstContent.substring(0, tagStart) + effectXml + effLstContent.substring(selfClosingEnd);
      } else {
        const fullEnd = effLstContent.indexOf(closeEffTag) + closeEffTag.length;
        effLstContent = effLstContent.substring(0, tagStart) + effectXml + effLstContent.substring(fullEnd);
      }
    } else {
      // 새 효과 추가
      effLstContent = effLstContent.replace('</a:effectLst>', effectXml + '</a:effectLst>');
    }
    
    const newSpPrContent = spPrContent.substring(0, effLstStart) + effLstContent + spPrContent.substring(effLstEnd);
    return shapeXml.substring(0, spPrIdx) + newSpPrContent + shapeXml.substring(spPrEndIdx + '</p:spPr>'.length);
  } else {
    // effectLst가 없으면 spPr 닫기 전에 추가
    const newEffLst = `<a:effectLst>${effectXml}</a:effectLst>`;
    const newSpPrContent = spPrContent.replace('</p:spPr>', newEffLst + '</p:spPr>');
    return shapeXml.substring(0, spPrIdx) + newSpPrContent + shapeXml.substring(spPrEndIdx + '</p:spPr>'.length);
  }
}

/**
 * 도형의 <a:ln> (선/테두리)을 추가 또는 교체
 */
function upsertLineInSpPr(shapeXml: string, lineXml: string): string {
  const spPrIdx = shapeXml.indexOf('<p:spPr');
  const spPrEndIdx = shapeXml.indexOf('</p:spPr>');
  
  if (spPrIdx === -1) return shapeXml;
  
  const spPrContent = shapeXml.substring(spPrIdx, spPrEndIdx + '</p:spPr>'.length);
  
  let newSpPrContent: string;
  if (spPrContent.includes('<a:ln')) {
    // 기존 <a:ln> 블록 찾아 교체
    const lnStart = spPrContent.indexOf('<a:ln');
    const lnEnd = spPrContent.indexOf('</a:ln>') + '</a:ln>'.length;
    if (lnEnd > lnStart) {
      newSpPrContent = spPrContent.substring(0, lnStart) + lineXml + spPrContent.substring(lnEnd);
    } else {
      // 셀프 클로징
      const selfEnd = spPrContent.indexOf('>', lnStart) + 1;
      newSpPrContent = spPrContent.substring(0, lnStart) + lineXml + spPrContent.substring(selfEnd);
    }
  } else {
    // </p:spPr> 직전에 삽입 (xfrm, prstGeom, 채우기 다음에 위치)
    newSpPrContent = spPrContent.replace('</p:spPr>', lineXml + '</p:spPr>');
  }
  
  return shapeXml.substring(0, spPrIdx) + newSpPrContent + shapeXml.substring(spPrEndIdx + '</p:spPr>'.length);
}

// ===========================
// OOXML 기반 공개 함수들
// ===========================

/**
 * OOXML을 사용하여 도형을 완전히 복제합니다.
 * Office.js API의 duplicate() 미지원 문제를 해결합니다.
 */
export async function duplicateShapeOoxml(params: OoxmlDuplicateParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] duplicateShapeOoxml:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    // 슬라이드 가져오기
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] ?? slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      if (selectedSlides.items.length === 0) throw new Error('슬라이드를 선택해주세요.');
      slide = selectedSlides.items[0];
    }

    // 도형 찾기
    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();
    for (const s of shapes.items) s.load('id,name,left,top,width,height');
    await context.sync();

    let targetShape: PowerPoint.Shape | undefined;
    if (params.shapeId) {
      targetShape = shapes.items.find(s => s.id === params.shapeId || s.name === params.shapeId);
    } else if (params.shapeIndex !== undefined) {
      targetShape = shapes.items[params.shapeIndex];
    }

    if (!targetShape) {
      const sel = context.presentation.getSelectedShapes();
      sel.load('items');
      await context.sync();
      if (sel.items.length > 0) {
        targetShape = sel.items[0];
        targetShape.load('id,name,left,top,width,height');
        await context.sync();
      }
    }

    if (!targetShape) throw new Error('복제할 도형을 찾을 수 없습니다.');

    const targetId = targetShape.id;
    const targetName = targetShape.name;
    const offsetX = params.offsetX ?? 20;
    const offsetY = params.offsetY ?? 20;

    // 현재 슬라이드 OOXML 가져오기
    const ooxml: string = await new Promise((resolve, reject) => {
      (Office.context.document as any).getSelectedDataAsync(
        'ooxml',
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error('OOXML 읽기 실패'));
        }
      );
    });

    // 대상 도형 블록 찾아 복제
    const modifiedOoxml = cloneShapeBlockInOoxml(ooxml, targetId, targetName, offsetX, offsetY);

    // 수정된 OOXML 적용
    await new Promise<void>((resolve, reject) => {
      (Office.context.document as any).setSelectedDataAsync(
        modifiedOoxml,
        { coercionType: 'ooxml' },
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error('OOXML 쓰기 실패: ' + result.error?.message));
        }
      );
    });

    await context.sync();
  });
}

/**
 * OOXML에서 도형 블록을 복제하여 새 ID로 삽입
 */
function cloneShapeBlockInOoxml(
  xml: string,
  shapeId: string,
  shapeName: string,
  offsetXPt: number,
  offsetYPt: number
): string {
  const shapeTags = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp'];
  
  for (const tag of shapeTags) {
    const openTag = `<${tag}`;
    const closeTag = `</${tag}>`;
    
    let searchStart = 0;
    while (true) {
      const startIdx = xml.indexOf(openTag, searchStart);
      if (startIdx === -1) break;
      
      // 블록 끝 찾기
      let depth = 1;
      let pos = startIdx + openTag.length;
      const firstTagEnd = xml.indexOf('>', pos - 1);
      if (xml[firstTagEnd - 1] === '/') {
        searchStart = firstTagEnd + 1;
        continue;
      }
      
      while (depth > 0 && pos < xml.length) {
        const nextOpen = xml.indexOf(openTag, pos);
        const nextClose = xml.indexOf(closeTag, pos);
        if (nextClose === -1) break;
        if (nextOpen !== -1 && nextOpen < nextClose) {
          depth++;
          pos = nextOpen + openTag.length;
        } else {
          depth--;
          pos = nextClose + closeTag.length;
        }
      }
      
      const endIdx = pos;
      const shapeBlock = xml.substring(startIdx, endIdx);
      
      const idMatch = shapeBlock.includes(`id="${shapeId}"`) || shapeBlock.includes(`id='${shapeId}'`);
      const nameMatch = shapeName && (shapeBlock.includes(`name="${shapeName}"`) || shapeBlock.includes(`name='${shapeName}'`));
      
      if (idMatch || nameMatch) {
        // 복제본 생성: 새 ID 부여, 위치 오프셋 적용
        const newId = `clone_${Date.now()}`;
        const newName = `${shapeName} 복사본`;
        
        let cloned = shapeBlock
          .replace(new RegExp(`id="${escapeRegex(shapeId)}"`, 'g'), `id="${newId}"`)
          .replace(new RegExp(`id='${escapeRegex(shapeId)}'`, 'g'), `id='${newId}'`)
          .replace(new RegExp(`name="${escapeRegex(shapeName)}"`, 'g'), `name="${newName}"`)
          .replace(new RegExp(`name='${escapeRegex(shapeName)}'`, 'g'), `name='${newName}'`);
        
        // 위치 오프셋 적용 (EMU 단위)
        const offsetXEmu = ptToEmu(offsetXPt);
        const offsetYEmu = ptToEmu(offsetYPt);
        cloned = offsetShapePosition(cloned, offsetXEmu, offsetYEmu);
        
        // 원본 블록 뒤에 복제본 삽입
        return xml.substring(0, endIdx) + cloned + xml.substring(endIdx);
      }
      
      searchStart = Math.min(startIdx + openTag.length, endIdx);
      if (searchStart >= xml.length) break;
    }
  }
  
  console.warn('OOXML에서 대상 도형 블록을 찾지 못했습니다:', shapeId);
  return xml;
}

/**
 * 도형 XML에서 위치(off x, off y)에 오프셋 추가
 */
function offsetShapePosition(shapeXml: string, offsetXEmu: number, offsetYEmu: number): string {
  // <a:off x="..." y="..."/> 패턴 찾아 값 조정
  return shapeXml.replace(/<a:off\s+x="(\d+)"\s+y="(\d+)"/g, (_match, x, y) => {
    const newX = parseInt(x) + offsetXEmu;
    const newY = parseInt(y) + offsetYEmu;
    return `<a:off x="${newX}" y="${newY}"`;
  }).replace(/<a:off\s+y="(\d+)"\s+x="(\d+)"/g, (_match, y, x) => {
    const newX = parseInt(x) + offsetXEmu;
    const newY = parseInt(y) + offsetYEmu;
    return `<a:off y="${newY}" x="${newX}"`;
  });
}

/**
 * 테두리 선 스타일 변경 (점선, 대시 등)
 * Office.js는 lineFormat.dashStyle을 지원하지 않으므로 OOXML로 처리
 */
export async function setBorderStyle(params: BorderStyleParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] setBorderStyle:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] ?? slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      if (selectedSlides.items.length === 0) throw new Error('슬라이드를 선택해주세요.');
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();
    for (const s of shapes.items) s.load('id,name');
    await context.sync();

    let targetShape: PowerPoint.Shape | undefined;
    if (params.shapeId) {
      targetShape = shapes.items.find(s => s.id === params.shapeId || s.name === params.shapeId);
    } else if (params.shapeIndex !== undefined) {
      targetShape = shapes.items[params.shapeIndex];
    }
    if (!targetShape) {
      const sel = context.presentation.getSelectedShapes();
      sel.load('items');
      await context.sync();
      if (sel.items.length > 0) {
        targetShape = sel.items[0];
        targetShape.load('id,name');
        await context.sync();
      }
    }
    if (!targetShape) throw new Error('대상 도형을 찾을 수 없습니다.');

    const targetId = targetShape.id;
    const targetName = targetShape.name;

    const ooxml: string = await new Promise((resolve, reject) => {
      (Office.context.document as any).getSelectedDataAsync('ooxml', (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error('OOXML 읽기 실패'));
      });
    });

    // 선 XML 구성
    const widthEmu = ptToEmu(params.width ?? 1) * 100; // 선 두께는 1/100pt 단위 아닌 1/12700pt
    // 실제로 <a:ln w> 속성은 EMU가 아닌 1/100000 인치 = 12700 * pt
    const lineWidthVal = ptToEmu(params.width ?? 0.75); // 12700 per point

    const colorHex = (params.color ?? '#000000').replace('#', '').toUpperCase();
    
    let dashXml = '';
    if (params.dashStyle && params.dashStyle !== 'solid') {
      dashXml = `<a:prstDash val="${params.dashStyle}"/>`;
    }
    
    let capXml = '';
    if (params.capStyle) {
      capXml = `cap="${params.capStyle}"`;
    }
    
    let joinXml = '';
    if (params.joinStyle === 'round') joinXml = '<a:round/>';
    else if (params.joinStyle === 'bevel') joinXml = '<a:bevel/>';
    else if (params.joinStyle === 'miter') joinXml = '<a:miter lim="800000"/>';
    
    const lineXml = `<a:ln w="${lineWidthVal}" ${capXml}><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>${dashXml}${joinXml}</a:ln>`;

    const modifiedOoxml = findAndModifyShapeBlock(ooxml, 'p:sp', targetId, targetName, (shapeXml) => {
      return upsertLineInSpPr(shapeXml, lineXml);
    });
    // 다른 태그도 시도
    let finalOoxml = modifiedOoxml;
    if (finalOoxml === ooxml) {
      for (const tag of ['p:pic', 'p:cxnSp']) {
        const result = findAndModifyShapeBlock(ooxml, tag, targetId, targetName, (shapeXml) => {
          return upsertLineInSpPr(shapeXml, lineXml);
        });
        if (result !== ooxml) { finalOoxml = result; break; }
      }
    }

    await new Promise<void>((resolve, reject) => {
      (Office.context.document as any).setSelectedDataAsync(finalOoxml, { coercionType: 'ooxml' }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error('OOXML 쓰기 실패: ' + result.error?.message));
      });
    });
    await context.sync();
  });
}

/**
 * 네온(글로우) 효과 적용
 * Office.js는 네온 효과를 지원하지 않으므로 OOXML로 처리
 */
export async function setNeonEffect(params: NeonEffectParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] setNeonEffect:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] ?? slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();
    for (const s of shapes.items) s.load('id,name');
    await context.sync();

    let targetShape: PowerPoint.Shape | undefined;
    if (params.shapeId) targetShape = shapes.items.find(s => s.id === params.shapeId || s.name === params.shapeId);
    else if (params.shapeIndex !== undefined) targetShape = shapes.items[params.shapeIndex];
    if (!targetShape) {
      const sel = context.presentation.getSelectedShapes();
      sel.load('items');
      await context.sync();
      if (sel.items.length > 0) { targetShape = sel.items[0]; targetShape.load('id,name'); await context.sync(); }
    }
    if (!targetShape) throw new Error('대상 도형을 찾을 수 없습니다.');

    const targetId = targetShape.id;
    const targetName = targetShape.name;

    const ooxml: string = await new Promise((resolve, reject) => {
      (Office.context.document as any).getSelectedDataAsync('ooxml', (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error('OOXML 읽기 실패'));
      });
    });

    const neonXml = buildNeonXml(params.color, params.radius ?? 8, params.transparency ?? 40);

    let modifiedOoxml = ooxml;
    for (const tag of ['p:sp', 'p:pic', 'p:cxnSp', 'p:graphicFrame']) {
      const result = findAndModifyShapeBlock(modifiedOoxml, tag, targetId, targetName, (shapeXml) => {
        return upsertEffectInSpPr(shapeXml, neonXml, 'a:glow');
      });
      if (result !== modifiedOoxml) { modifiedOoxml = result; break; }
    }

    await new Promise<void>((resolve, reject) => {
      (Office.context.document as any).setSelectedDataAsync(modifiedOoxml, { coercionType: 'ooxml' }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error('OOXML 쓰기 실패: ' + result.error?.message));
      });
    });
    await context.sync();
  });
}

/**
 * 그림자 효과 적용
 */
export async function setShadowEffect(params: ShadowEffectParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] setShadowEffect:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] ?? slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();
    for (const s of shapes.items) s.load('id,name');
    await context.sync();

    let targetShape: PowerPoint.Shape | undefined;
    if (params.shapeId) targetShape = shapes.items.find(s => s.id === params.shapeId || s.name === params.shapeId);
    else if (params.shapeIndex !== undefined) targetShape = shapes.items[params.shapeIndex];
    if (!targetShape) {
      const sel = context.presentation.getSelectedShapes();
      sel.load('items');
      await context.sync();
      if (sel.items.length > 0) { targetShape = sel.items[0]; targetShape.load('id,name'); await context.sync(); }
    }
    if (!targetShape) throw new Error('대상 도형을 찾을 수 없습니다.');

    const targetId = targetShape.id;
    const targetName = targetShape.name;

    const ooxml: string = await new Promise((resolve, reject) => {
      (Office.context.document as any).getSelectedDataAsync('ooxml', (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error('OOXML 읽기 실패'));
      });
    });

    const shadowXml = buildShadowXml(params);
    const effectTag = params.type === 'inner' ? 'a:innerShdw' : 'a:outerShdw';

    let modifiedOoxml = ooxml;
    for (const tag of ['p:sp', 'p:pic', 'p:cxnSp']) {
      const result = findAndModifyShapeBlock(modifiedOoxml, tag, targetId, targetName, (shapeXml) => {
        return upsertEffectInSpPr(shapeXml, shadowXml, effectTag);
      });
      if (result !== modifiedOoxml) { modifiedOoxml = result; break; }
    }

    await new Promise<void>((resolve, reject) => {
      (Office.context.document as any).setSelectedDataAsync(modifiedOoxml, { coercionType: 'ooxml' }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error('OOXML 쓰기 실패: ' + result.error?.message));
      });
    });
    await context.sync();
  });
}

/**
 * 반사 효과 적용
 */
export async function setReflectionEffect(params: ReflectionEffectParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] setReflectionEffect:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] ?? slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();
    for (const s of shapes.items) s.load('id,name');
    await context.sync();

    let targetShape: PowerPoint.Shape | undefined;
    if (params.shapeId) targetShape = shapes.items.find(s => s.id === params.shapeId || s.name === params.shapeId);
    else if (params.shapeIndex !== undefined) targetShape = shapes.items[params.shapeIndex];
    if (!targetShape) {
      const sel = context.presentation.getSelectedShapes();
      sel.load('items');
      await context.sync();
      if (sel.items.length > 0) { targetShape = sel.items[0]; targetShape.load('id,name'); await context.sync(); }
    }
    if (!targetShape) throw new Error('대상 도형을 찾을 수 없습니다.');

    const targetId = targetShape.id;
    const targetName = targetShape.name;

    const ooxml: string = await new Promise((resolve, reject) => {
      (Office.context.document as any).getSelectedDataAsync('ooxml', (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error('OOXML 읽기 실패'));
      });
    });

    const reflXml = buildReflectionXml(params);

    let modifiedOoxml = ooxml;
    for (const tag of ['p:sp', 'p:pic', 'p:cxnSp']) {
      const result = findAndModifyShapeBlock(modifiedOoxml, tag, targetId, targetName, (shapeXml) => {
        return upsertEffectInSpPr(shapeXml, reflXml, 'a:reflection');
      });
      if (result !== modifiedOoxml) { modifiedOoxml = result; break; }
    }

    await new Promise<void>((resolve, reject) => {
      (Office.context.document as any).setSelectedDataAsync(modifiedOoxml, { coercionType: 'ooxml' }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error('OOXML 쓰기 실패: ' + result.error?.message));
      });
    });
    await context.sync();
  });
}

/**
 * 부드러운 가장자리(Soft Edge) 효과 적용
 */
export async function setSoftEdge(params: SoftEdgeParams): Promise<void> {
  if (!isOfficeAvailable()) {
    console.log('[Mock] setSoftEdge:', params);
    return;
  }

  await PowerPoint.run(async (context) => {
    let slide: PowerPoint.Slide;
    if (params.slideIndex !== undefined) {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      slide = slides.items[params.slideIndex] ?? slides.items[0];
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load('items');
    await context.sync();
    for (const s of shapes.items) s.load('id,name');
    await context.sync();

    let targetShape: PowerPoint.Shape | undefined;
    if (params.shapeId) targetShape = shapes.items.find(s => s.id === params.shapeId || s.name === params.shapeId);
    else if (params.shapeIndex !== undefined) targetShape = shapes.items[params.shapeIndex];
    if (!targetShape) {
      const sel = context.presentation.getSelectedShapes();
      sel.load('items');
      await context.sync();
      if (sel.items.length > 0) { targetShape = sel.items[0]; targetShape.load('id,name'); await context.sync(); }
    }
    if (!targetShape) throw new Error('대상 도형을 찾을 수 없습니다.');

    const targetId = targetShape.id;
    const targetName = targetShape.name;

    const ooxml: string = await new Promise((resolve, reject) => {
      (Office.context.document as any).getSelectedDataAsync('ooxml', (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error('OOXML 읽기 실패'));
      });
    });

    const softEdgeXml = buildSoftEdgeXml(params.radius ?? 5);

    let modifiedOoxml = ooxml;
    for (const tag of ['p:sp', 'p:pic', 'p:cxnSp']) {
      const result = findAndModifyShapeBlock(modifiedOoxml, tag, targetId, targetName, (shapeXml) => {
        return upsertEffectInSpPr(shapeXml, softEdgeXml, 'a:softEdge');
      });
      if (result !== modifiedOoxml) { modifiedOoxml = result; break; }
    }

    await new Promise<void>((resolve, reject) => {
      (Office.context.document as any).setSelectedDataAsync(modifiedOoxml, { coercionType: 'ooxml' }, (result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error('OOXML 쓰기 실패: ' + result.error?.message));
      });
    });
    await context.sync();
  });
}
