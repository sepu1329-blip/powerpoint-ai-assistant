import React from 'react';
import type { PresentationContext } from '../types';

interface SlideContextBadgeProps {
  context: PresentationContext | null;
  onRefresh: () => void;
}

export const SlideContextBadge: React.FC<SlideContextBadgeProps> = ({
  context,
  onRefresh,
}) => {
  if (!context) {
    return (
      <button className="context-badge loading" onClick={onRefresh}>
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
        </svg>
        <span>슬라이드 로딩 중...</span>
      </button>
    );
  }

  const shapeCount = context.currentSlide?.shapes.length ?? 0;
  const slideNum = context.currentSlideIndex + 1;
  const total = context.slideCount;

  return (
    <button className="context-badge" onClick={onRefresh} title="클릭하여 새로고침">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <rect x="2" y="3" width="20" height="14" rx="2" ry="2"/>
        <path d="M8 21h8M12 17v4"/>
      </svg>
      <span>슬라이드 {slideNum}/{total}</span>
      <span className="context-divider">·</span>
      <span>도형 {shapeCount}개</span>
      <svg className="refresh-icon" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
        <polyline points="1 4 1 10 7 10"/>
        <path d="M3.51 15a9 9 0 1 0 .49-3"/>
      </svg>
    </button>
  );
};
