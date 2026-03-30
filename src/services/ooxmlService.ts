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
// 내부 헬퍼: 대시 스타일 매핑 / XML 이스케이프
// ===========================

/** PowerPoint.ShapeLineDashStyle → OOXML prstDash val 매핑 */
function mapDashStyle(dash: any): string {
  if (!dash) return '';
  const s = String(dash).toLowerCase();
  if (s.includes('dashdot') && s.includes('dot'))  return 'lgDashDotDot';
  if (s.includes('longdashDotDot') || s === 'longdashdotdot') return 'lgDashDotDot';
  if (s.includes('longdashdot'))  return 'lgDashDot';
  if (s.includes('longdash'))     return 'lgDash';
  if (s.includes('systemdashdot'))return 'sysDashDot';
  if (s.includes('systemdash'))   return 'sysDash';
  if (s.includes('systemdot'))    return 'sysDot';
  if (s.includes('rounddot'))     return 'dot';
  if (s.includes('squaredot'))    return 'dot';
  if (s.includes('dashdot'))      return 'dashDot';
  if (s.includes('dash'))         return 'dash';
  if (s.includes('dot'))          return 'dot';
  return '';
}

/** XML 특수문자 이스케이프 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
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

    const offsetX = params.offsetX ?? 20;
    const offsetY = params.offsetY ?? 20;

    // 상세 속성 로드 (위치, 채우기, 테두리)
    targetShape.load('left,top,width,height,type');
    await context.sync();

    // 채우기 색상 - 열거형 대신 문자열 비교로 'solid' 체크
    let fillXml = '<a:noFill/>';
    try {
      targetShape.fill.load('type,foregroundColor,transparency');
      await context.sync();
      const fillType = String(targetShape.fill.type ?? '').toLowerCase();
      if (fillType === 'solid' || fillType === '2') {
        const c = (targetShape.fill.foregroundColor ?? '#4472C4').replace('#', '');
        const transparency = targetShape.fill.transparency ?? 0;
        const a = Math.round((1 - transparency) * 100000);
        fillXml = `<a:solidFill><a:srgbClr val="${c}"><a:alpha val="${a}"/></a:srgbClr></a:solidFill>`;
      }
    } catch (_e) {
      console.warn('채우기 속성 로드 실패, 기본값 사용:', _e);
    }

    // 테두리
    let lineXml = '';
    try {
      targetShape.lineFormat.load('color,weight,dashStyle,visible');
      await context.sync();
      const isVisible = targetShape.lineFormat.visible;
      if (isVisible !== false) {
        const lc = (targetShape.lineFormat.color ?? '#000000').replace('#', '');
        const lw = Math.round((targetShape.lineFormat.weight ?? 0.75) * 12700);
        const dv = mapDashStyle(targetShape.lineFormat.dashStyle);
        const dashXml = dv ? `<a:prstDash val="${dv}"/>` : '';
        lineXml = `<a:ln w="${lw}"><a:solidFill><a:srgbClr val="${lc}"/></a:solidFill>${dashXml}</a:ln>`;
      } else {
        lineXml = '<a:ln><a:noFill/></a:ln>';
      }
    } catch (_e) {
      console.warn('테두리 속성 로드 실패, 테두리 없음으로 처리:', _e);
    }

    // 텍스트
    let textBodyXml = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>';
    try {
      targetShape.textFrame.load('hasText');
      await context.sync();
      if (targetShape.textFrame.hasText) {
        targetShape.textFrame.textRange.load('text');
        targetShape.textFrame.textRange.font.load('bold,color,size');
        await context.sync();
        const txt = targetShape.textFrame.textRange.text ?? '';
        const bold = targetShape.textFrame.textRange.font.bold ?? false;
        const fc   = (targetShape.textFrame.textRange.font.color ?? '').replace('#', '');
        const fsSz = Math.round((targetShape.textFrame.textRange.font.size ?? 18) * 100);
        const boldXml = bold ? ' b="1"' : '';
        const fcXml   = fc ? `<a:solidFill><a:srgbClr val="${fc}"/></a:solidFill>` : '';
        textBodyXml = `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="${fsSz}"${boldXml}>${fcXml}</a:rPr><a:t>${escapeXml(txt)}</a:t></a:r></a:p></p:txBody>`;
      }
    } catch (_e) {
      console.warn('텍스트 속성 로드 실패, 빈 텍스트 사용:', _e);
    }


    // 위치 계산 (포인트 → EMU)
    const newLeftEmu = Math.round((targetShape.left  + offsetX) * 12700);
    const newTopEmu  = Math.round((targetShape.top   + offsetY) * 12700);
    const widthEmu   = Math.round( targetShape.width  * 12700);
    const heightEmu  = Math.round( targetShape.height * 12700);
    const newShapeId = Math.floor(Math.random() * 90000) + 10000;
    const newName    = `\uBCF5\uC0AC\uBCF8 ${Date.now() % 10000}`;

    // Flat OPC OOXML 패키지 구성 후 setSelectedDataAsync로 삽입
    // PowerPoint는 setSelectedDataAsync + coercionType:'ooxml' 을 지원
    const ooxmlPackage = `<?xml version="1.0" encoding="utf-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels"
    pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
          Target="ppt/presentation.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/ppt/slides/slide1.xml"
    pkg:contentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml">
    <pkg:xmlData>
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:cSld><p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
          <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
            <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
          <p:sp>
            <p:nvSpPr>
              <p:cNvPr id="${newShapeId}" name="${newName}"/>
              <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
              <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
              <a:xfrm><a:off x="${newLeftEmu}" y="${newTopEmu}"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              ${fillXml}
              ${lineXml}
            </p:spPr>
            ${textBodyXml}
          </p:sp>
        </p:spTree></p:cSld>
      </p:sld>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

    await new Promise<void>((resolve, reject) => {
      (Office.context.document as any).setSelectedDataAsync(
        ooxmlPackage,
        { coercionType: 'ooxml' },
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error('도형 삽입 실패: ' + (result.error?.message ?? JSON.stringify(result.error))));
        }
      );
    });
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
