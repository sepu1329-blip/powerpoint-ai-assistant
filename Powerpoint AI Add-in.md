🖥️ 아래와 같은 기능을 파워포인트 추가 기능으로 만들고 싶어
1. React와 TypeScript 기반의 PowerPoint 전용 Add-in 프로젝트를 생성해
2. 생성 완료 후, 불필요한 기본 샘플 코드를 모두 삭제하고, 빈 화면의 메인 컴포넌트를 렌더링하도록 `App.tsx`를 초기화해.
3. `manifest.xml`을 수정하여 우측 작업 창(Taskpane)이 열리도록 UI 확장을 설정
4. Claude by Anthropic for PowerPoint 추가 기능과 완전히 동일한 기능을 하게 만들어줘.
5. 웹을 만들어야 하면 깃허브 페이지에 업로드해서 사용하게 준비해줘.
6. 설정창에서 api키를 입력 저장 할 수 있게 구성해.
7. ai는 gemini 최신 api로 pro, flash 선택할 수 있게 해줘.
8. 공식 스토어를 통한 Add-in 설치가 차단된 환경이므로, 매니페스트를 통한 네트워크 공유 폴더 사이드로딩(Sideloading) 방식으로 데스크톱 앱에서 테스트할 수 있도록 설정해.
9. Office.js API를 제어하여 PowerPoint 문서의 데이터를 읽어오는 로직을 작성해.
10. 사용자가 선택한 객체, 내용 또는 선택한 다중 슬라이드 들을 컨텍스트로 활용해. 선택이 없으면 현재 보고 있는 슬라이드를 컨텍스트로 활용해.
11. UI 디자인은 https://sepu1329-blip.github.io/excel-ai-assistant/ 와 동일하게 만들어줘.
12. 아래는 Claude by Anthropic for PowerPoint 의 설명 내용이야.
    Professionals who build presentations can use Claude for PowerPoint to collaborate with Claude directly where you work.

    Claude for PowerPoint accelerates slide development through intelligent, template-aware assistance. It reads your existing deck — layouts, fonts, colors, and slide masters — and makes edits that respect your formatting.

    Build new slides from a client or corporate template, iterate on a draft with pinpoint edits, or generate a full deck structure from a natural language description.

    Whether you're restructuring a storyline, converting bullets into diagrams, or adding native charts, Claude works as a co-author inside your deck — no copying and pasting between tools.

    Available in beta as a research preview to customers on the Claude Pro, Max, Team and Enterprise plans.

    Visit our website at https://claude.com/claude-in-powerpoint, our Help Center at https://support.claude.com/en/articles/13521390-use-claude-in-powerpoint, and our plan page at https://claude.com/pricing for more information. Get support: https://support.claude.com/en/articles/9015913-how-to-get-support.