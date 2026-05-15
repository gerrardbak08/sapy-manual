# CLAUDE.md — Work Platform
## Claude Code 마스터 지시 파일

> 이 파일은 Claude Code가 모든 대화 시작 시 자동으로 읽는다.
> **모든 작업 전 이 파일과 DESIGN.md를 반드시 완독한다.**

---

## 프로젝트 정보

```
서비스명:   Work Platform (전사 안전보건통합관리 플랫폼)
회사:       ㈜아성다이소 안전보건팀
스택:       Next.js 15 (App Router) + React 19 + Tailwind CSS + Radix UI
폰트:       Pretendard (sans) + DM Mono (숫자)
아이콘:     lucide-react (2D만, 3D 절대 금지)
애니메이션: motion/react (framer-motion v11+)
테마:       class-based dark mode (.dark on <html>)
배포:       Vercel (현재) → AWS (예정)
백엔드:     AWS Lambda + API Gateway + DynamoDB + S3 + Cognito
AI:         Amazon Bedrock (Claude)
```

---

## 작업 전 필수 체크리스트

모든 UI 작업 시작 전 순서대로 확인한다.

```
□ 1. DESIGN.md를 완독했는가?
□ 2. 이 화면의 목적을 한 문장으로 정의했는가?
□ 3. UI 우선인가, UX 우선인가 결정했는가?
□ 4. 5가지 상태(정상/빈값/로딩/에러/권한없음)를 모두 설계했는가?
□ 5. DESIGN.md에 없는 값을 쓰려는 것은 아닌가?
```

---

## 절대 금지 — 즉시 reject

아래 항목이 코드에 있으면 무조건 수정한다. 예외 없음.

```
# 색상
text-white (버튼)           → text-primary-foreground
bg-background (카드 배경)   → bg-card
bg-[#13245A]                → bg-primary
text-amber-400              → text-warning
임의 hex 하드코딩            → CSS 변수 사용
purple / pink / cyan 추가   → 4색 시스템만

# 레이아웃
Fragment(<>) as 페이지 루트  → PageShell div
이중 overflow-y-auto        → flex-col + flex-1 구조로
p-3/p-4/p-5/p-6 혼용        → p-4(모바일) / p-5(md+)

# 타이포그래피
text-[9px]                  → 삭제 (WCAG 위반)
text-[10px]                 → text-micro
text-[11px], text-[12px]    → text-caption
text-[13px]                 → text-body 또는 text-sm
font-bold (카드 h2 제목)     → font-semibold

# 컴포넌트
rounded-xl (카드 외곽)       → rounded-2xl
3D 아이콘                    → lucide-react 2D만
아이콘 겹침                  → 하나 제거 or ··· 통합
SectionHeader 중복 정의      → features/dashboard/SectionHeader.tsx 단일 사용

# 모션
transition-all               → transition-colors / transition-transform
200ms 초과 애니메이션         → 업무용 앱 흐름 방해
bounce / spring              → 금지

# 기타
색만으로 정보 전달            → 색 + 아이콘 + 텍스트 3중 필수
그림자 4단계 이상             → 3단계만 (elev-1/2/3)
```

---

## 핵심 판단 기준

> **"이 선택이 관리자의 판단을 더 빠르고 정확하게 만드는가?"**
> YES면 채택, NO면 반려. "예뻐서"는 이유가 아니다.

### 페이지 유형별 UI/UX 비중

| 페이지 유형 | 우선순위 | 이유 |
|---|---|---|
| 대시보드, 알림, 임원 보고 | **UI 우선** | 첫눈에 신뢰감, 즉각 인식 |
| 사고 신고 폼, 서류 제출 | **UX 우선** | 빠른 완료, 단계 최소화 |
| 서류 관리, 검색·필터 | **UX 우선** | 찾기 쉬운 구조가 핵심 |
| 교육 이수 현황 | **UI 우선** | 한눈에 현황 파악 |

### 사용자별 UI 방향

| 사용자 | UI 방향 | 데이터 범위 |
|---|---|---|
| 본사 안전보건팀 | 고밀도 대시보드 | 전체 |
| 현장 점포 직원 | 단순하고 큰 터치 UI | 자기 점포만 |
| 임원·경영진 | 히어로 숫자 중심 | 전체 요약 |
| 외부 협력사 | 명확한 폼 | 자사 관련만 |

---

## PageShell 패턴 — 모든 platform 라우트 필수

```tsx
export default function SomePage() {
  return (
    <div className="flex flex-col h-full overflow-y-auto overflow-x-hidden bg-background">
      <Header title="페이지명" />
      <div className="flex-1 p-4 md:p-5 space-y-3 min-w-0">
        {/* 콘텐츠 */}
      </div>
    </div>
  );
}
// ✗ 금지: Fragment(<>)로 페이지 루트 대체
// ✗ 금지: 이중 overflow-y-auto (스크롤 트랩)
```

---

## 현재 수정 우선순위

### P0 — 즉시 수정 (구조 오류)

```
1. app/(platform)/annual-kpi/page.tsx
   → 루트 div에 PageShell 클래스 추가

2. features/meetings/MeetingsShell.tsx
   → Fragment(<>) → div wrapper 교체

3. features/tasks/MyTasksColumn.tsx
   → 이중 스크롤 컨테이너 제거

4. features/accidents/cards/SectionHeader.tsx
   → 삭제, features/dashboard/SectionHeader로 import 통일

5. features/dashboard/AnnouncementCard.tsx
   → bg-background → bg-card
```

### P1 — 이번 주

```
- rounded-xl (카드 외곽) 전체 → rounded-2xl
- text-white (버튼) 3곳 → text-primary-foreground
- bg-[#13245A] → bg-primary (인증 페이지)
- text-amber-400 → text-warning
- content padding 6종 혼용 → p-4(모바일) / p-5(md+) 통일
```

### P2 — 다음 주 (git commit 먼저, 배치 처리)

```
text-[Npx] 하드코딩 620건 → 토큰으로 일괄 교체 (codemod 권장)
font-bold (카드 h2) → font-semibold
```

---

## 화면 설계 전 필수 4가지 질문

코드 작성 전에 반드시 이 4가지에 답한다.

```
1. 이 화면의 목적은?
   → 사용자가 3초 안에 이해할 수 있는 한 문장으로.

2. 주요 행동은?
   → 읽기 / 탐색 / 입력 / 확인 중 하나.
     하나의 화면에 4가지가 모두 있으면 사용자가 혼란스럽다.

3. 모든 상태가 설계됐는가?
   → 정상 / 빈값 / 로딩 / 에러 / 권한없음 전부 포함.
     이 중 하나라도 빠지면 미완성이다.

4. 레퍼런스 의도가 우리와 맞는가?
   → 사용자 맥락, 정보 밀도, 신뢰감이 동일한가?
```

---

## 안전보건 도메인 고정 규칙

```
# 위험도 색상 (불변)
위험도 70 이상 → danger  (red)
위험도 40~69   → warning (orange)
위험도 0~39    → success (green)

# 사고 심각도
사망·중상  → Status Line danger + 목록 최상단
경상       → Status Line warning
아차사고   → Status Line info

# 법정 서류 상태
유효       → success 배지
만료 7일 전 → warning 배지 + pulse 애니메이션
만료       → danger 배지
검토중     → info 배지

# 교육 이수율
90% 이상 → success
70~89%   → warning
70% 미만 → danger
```

---

## 빈 상태 문구 (Work Platform 전용)

```
사고 없음: 🎉 "오늘은 사고 신고가 없습니다" / "안전한 하루를 보내고 있습니다"
서류 없음: 📋 "등록된 서류가 없습니다" / "서류를 업로드해 주세요"
교육 없음: 📚 "완료된 교육이 없습니다" / "교육 일정을 확인해 주세요"
점검 없음: ✅ "예정된 점검이 없습니다"
```

---

## Claude Code 실행 프롬프트

### 전체 UI 자동 감사

```
DESIGN.md와 CLAUDE.md를 완독한 뒤,
workplatform.vercel.app을 Playwright로 열어서
모든 화면을 순회하며:

1. 두 파일의 기준 위반 항목 전부 찾기
2. P0 → P1 → P2 순서로 심각도 분류
3. 파일명 · 줄번호 · 위반내용 · 수정방법 순서로 보고

찾기만 해. 수정은 내가 "시작"이라고 하면 해.
수정할 때마다 무엇을 바꿨는지 말하고 내 확인 받기.
```

### 새 화면 생성

```
DESIGN.md를 따라 [화면명] 화면을 만들어줘.

목적: [3초 안에 이해할 수 있는 한 문장]
주요 행동: [읽기/탐색/입력/확인 중 하나]
설계해야 할 상태:
  - 정상 상태
  - 빈 상태 (감정적 메시지 포함)
  - 로딩 상태 (Skeleton)
  - 에러 상태
  - 권한 없음 상태
```

### P0 즉시 수정

```
DESIGN.md와 CLAUDE.md의 P0 항목을 순서대로 수정해줘.
수정할 때마다 어떤 파일을 어떻게 바꿨는지 말하고,
내가 "다음"이라고 하면 다음 항목으로 넘어가줘.
완료 후 npm run build 돌려서 에러 없는지 확인해줘.
```

---

## AI 분석 결과 표시

Bedrock/Claude 분석 결과는 항상 info 색 전용 블록으로 구분한다.

```tsx
<div className="rounded-2xl border border-info/20 bg-info/5 p-4">
  <div className="flex items-center gap-2 mb-2">
    <span className="text-info text-base">✦</span>
    <span className="text-caption font-semibold text-info">AI 분석</span>
  </div>
  <p className="text-body text-foreground leading-relaxed">{result}</p>
</div>
```

---

## 파일 구조 규칙

```
src/
├── app/
│   ├── (auth)/      → 인증 페이지 (bg-primary 배경)
│   └── (platform)/  → 메인 앱 (PageShell 필수)
├── components/
│   ├── ui/          → 기본 컴포넌트
│   ├── features/
│   │   └── dashboard/
│   │       └── SectionHeader.tsx  ← 전체에서 유일한 SectionHeader
│   └── layout/      → Header, Sidebar
├── lib/
│   └── api/         → API 함수 전용
└── styles/
    └── globals.css  → CSS 변수 전용 (인라인 hex 금지)
```

---

## 배포 전 최종 체크리스트

```
# 구조
□ 모든 platform 라우트에 PageShell div 있음
□ 이중 스크롤 컨테이너 없음
□ Fragment(<>) 페이지 루트 없음

# 디자인 토큰
□ 하드코딩 hex 없음
□ text-[Npx] 하드코딩 없음
□ 모든 카드 외곽 rounded-2xl
□ 모든 버튼 text-primary-foreground

# 컴포넌트
□ SectionHeader 중복 파일 없음
□ 3D 아이콘 없음
□ 아이콘 겹침 없음
□ 빈 상태 감정적 문구 있음

# 상태 설계
□ 로딩 → Skeleton 있음
□ 빈값 → Empty State 있음
□ 에러 → 에러 메시지 있음
□ 위험도 색상 danger/warning/success 정확

# 접근성
□ 9px 이하 폰트 없음
□ 모든 상태 색+아이콘+텍스트 3중

# 빌드
□ npm run build → exit 0
□ npm run lint  → 0 errors
```

---

*Work Platform CLAUDE.md v1.0 | 2026-04-14*
