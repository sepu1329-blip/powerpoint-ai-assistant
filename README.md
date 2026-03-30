# PowerPoint AI Assistant

> Gemini AI를 활용한 Microsoft PowerPoint 추가 기능 (Add-in)

**Claude for PowerPoint**와 동일한 기능을 Google Gemini API로 구현한 PowerPoint 전용 AI 어시스턴트입니다.

## 주요 기능

- 🤖 **AI 채팅**: 자연어로 슬라이드 편집 요청
- 📊 **슬라이드 컨텍스트**: 현재 슬라이드 정보를 AI에 자동 전달
- ✏️ **슬라이드 수정**: AI 제안 → 사용자 승인 → 자동 적용
- ⚙️ **모델 선택**: Gemini 2.0 Flash / 2.5 Pro 선택 가능
- 🔑 **API 키 관리**: 설정창에서 직접 입력/저장

## 설치 방법 (사이드로딩)

### 1단계: manifest.xml 다운로드

[manifest.xml](https://raw.githubusercontent.com/sepu1329-blip/powerpoint-ai-assistant/main/manifest.xml)을 다운로드하세요.

### 2단계: 네트워크 공유 폴더 설정 (Windows)

1. 빈 폴더를 만들고 네트워크 공유 폴더로 설정:
   - 폴더 우클릭 → 속성 → 공유 탭 → 공유 클릭
   - "나 포함 모든 사용자" 읽기 권한 부여

2. 공유 폴더 경로 확인 (예: `\\YourPC\AddinFolder`)

### 3단계: PowerPoint에서 신뢰 폴더 등록

1. PowerPoint 실행
2. **파일 → 옵션 → 보안 센터 → 보안 센터 설정**
3. **신뢰할 수 있는 앱 카탈로그** 메뉴 클릭
4. **카탈로그 URL**에 공유 폴더 경로 입력 (`\\YourPC\AddinFolder`)
5. **카탈로그 추가** 클릭 → "메뉴에 표시" 체크 → 확인

### 4단계: Add-in 설치

1. PowerPoint 재시작
2. **삽입 탭 → 내 추가 기능** 클릭
3. **공유 폴더** 탭 선택
4. **PowerPoint AI Assistant** 선택 → 추가

### Mac에서 설치

1. `/Users/<사용자명>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` 폴더 생성
2. `manifest.xml`을 해당 폴더에 복사
3. PowerPoint 재시작 후 삽입 → 내 추가 기능에서 설치

## 개발 환경 설정

```bash
# 의존성 설치
npm install

# 로컬 개발 서버 시작 (HTTPS)
npm run dev

# 빌드
npm run build
```

로컬 테스트 시 manifest.xml의 `SourceLocation`을 `https://localhost:3000/`으로 변경하세요.

## 기술 스택

- **프레임워크**: React 18 + TypeScript
- **빌드**: Vite 6
- **AI**: Google Gemini API (`@google/generative-ai`)
- **Office**: Office.js (`@types/office-js`)
- **배포**: GitHub Pages

## 라이선스

MIT
