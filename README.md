# P25 Triton Punch List Manager

Phoenix P25 Triton 프로젝트 펀치리스트 / 스냅리스트 관리 시스템

---

## 배포 방법 (Vercel — 권장, 무료)

### 1단계 — GitHub 업로드
1. https://github.com 접속 → 로그인 (없으면 무료 가입)
2. New Repository → 이름: `p25-punchlist` → Create
3. "uploading an existing file" 클릭
4. 이 폴더의 모든 파일을 드래그앤드롭 업로드
5. Commit changes

### 2단계 — Vercel 배포
1. https://vercel.com 접속 → GitHub 계정으로 로그인
2. "Add New Project" → p25-punchlist 선택
3. Framework Preset: **Vite** 선택
4. Deploy 클릭
5. 배포 완료 → URL 발급 (예: `https://p25-punchlist.vercel.app`)

### 3단계 — 모바일 홈화면 설치 (PWA)
- **iPhone**: Safari에서 URL 접속 → 공유버튼 → "홈 화면에 추가"
- **Android**: Chrome에서 URL 접속 → 메뉴 → "앱 설치" 또는 "홈 화면에 추가"

---

## 팀 공유 방법

배포된 URL을 팀원에게 공유하면 됩니다.
- CM팀: 같은 URL 접속
- 인도 팀원: 같은 URL 접속
- 각자 기기에서 데이터가 로컬 저장됩니다.

> **Note**: 팀 간 실시간 데이터 공유가 필요하면 Supabase 연동이 필요합니다.
> 별도 문의 시 설정 가이드 제공 가능합니다.

---

## 주요 기능

- 📋 스냅 등록 / 상태 관리
- 🌐 EN / 한국어 전환
- 👥 역할별 워크플로 (CM / Contractor / Approver)
- 📥 **Excel 내보내기** (P25 템플릿 형식, 3개 시트)
- 📤 JSON 내보내기 / 불러오기 (기기 간 데이터 이전)
- ⚙ 업체 이메일 설정
- 📱 PWA — 홈화면 앱 설치 가능

---

## Excel 내보내기 시트 구성

| 시트 | 내용 |
|---|---|
| P25 Snag List | 전체 스냅 목록 (원본 템플릿 형식) |
| Dashboard | 통계 요약 (상태별, 우선순위별, 업체별) |
| Activity Log | 전체 상태 변경 이력 |

---

## 로컬 개발 환경 실행

```bash
npm install
npm run dev
```

---

## 기술 스택

- React 18 + Vite
- SheetJS (xlsx) — Excel 내보내기
- localStorage — 데이터 저장
- PWA (manifest.json) — 모바일 앱 설치

---

*Phoenix P25 Triton Project — CM Team, JLCM*
