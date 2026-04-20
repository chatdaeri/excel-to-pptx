---
name: excel-to-pptx
description: 엑셀 → PPT 자동 변환 (표 전용 lite 패키지). 차트·지도·컨셉·3블록 자동·브랜드 커스터마이징은 제외하고 표지·목차·간지·풀폭표·2단표 5종 슬라이드만 다루는 가벼운 패키지. AI 는 ~15줄짜리 주문서(deck.cjs)로 슬라이드 순서·표지·간지를 작성하고, 시트 [표] 블록은 sheet 기반 자동 매핑. 셀 병합 보존(Stage 4)·한국식 헤더(네이비 박스)·Pretendard Static 고정.
---

# excel-to-pptx — 엑셀 → PPT 변환 (표 전용 lite)

엑셀 스캔 → 슬라이드 구성 결정 → 주문서(`deck.cjs`) 작성 → 런너 → 검증.

**excel-to-pptx (lite)** 는 **표 전용 패키지**로 표지·목차·간지·풀폭표·2단표 5종 슬라이드만 다룬다.

---

## ⛔ lite 가 지원하지 않는 것

| 기능 |
|---|
| 차트 일체 (bar/line/doughnut/barLine) |
| 지도 슬라이드 (T2 map) |
| 컨셉 박스 (T4 concept) |
| 3블록 자동 시트 (T6 auto) |
| 풀폭 차트 (T7 bigchart) |
| 브랜드 컬러·폰트 커스터마이징 |
| inline 차트 series 재구성 |

엑셀에 `[차트]` 블록이 있으면 런너가 거부한다.

---

## ✅ lite 가 지원하는 것

| 슬라이드 | 빌더 | 설명 |
|---|---|---|
| `cover` | T5 | 표지 (overline / title / subtitle / date 3층 구조) |
| `toc` | T5 | 목차 |
| `divider` | T3 | 챕터 디바이더 (Part N) |
| `bigtable` | T7b | 풀폭 표 1장 (1블록 [표] 시트) |
| `data2col` | T1 (table-only) | 좌·우 2단 표 (2블록 [표] 시트) |

공통:
- ✅ 셀 병합 자동 재현 (Stage 4 patch_table_merges)
- ✅ 표 스타일: 헤더 `#E8E8E8` / 좌측 라벨 `#EAEFF8` / 강조 행 빨강 외곽 / 외곽 좌·우 선 없음
- ✅ Stage 1 자동 좌표 검증
- ✅ 한국식 분양 헤더(네이비 박스 소목차 + 검정 메인 메시지) · Pretendard 폰트 고정

---

## 🗂 스킬 파일 구조

```
excel-to-pptx/
├── SKILL.md                                  ← (본 문서)
├── README.md                                 📖 사용자 안내서 — 엑셀 준비·결과물 모양·커스텀 요청 방법
├── customization-process.md                  🤖 AI 전용 — 사용자가 톤 변경 요청 시 따라야 할 작업 매뉴얼
│
├── builders/                                 ← 런타임 코어 (수정 금지)
│   ├── _helpers.cjs                            컬러·폰트·좌표 상수, 헤더·표 빌더
│   │                                           (chartOpts·AUTO_LAYOUTS 제외)
│   ├── validate.cjs                            Stage 1 Validator
│   ├── excel-sheet-parser.cjs                  시트 → {sectionLabel, main/sub, blocks[]}
│   │                                           ([표] 자동 인식 / [차트] 인식하되 런너가 거부)
│   ├── t1-data-table.cjs                       T1: 좌·우 2단 표 전용 (chart slot 거부)
│   ├── t3-divider.cjs                          T3: 챕터 디바이더
│   ├── t5-cover-toc.cjs                        T5: 표지 / 목차
│   ├── t7b-big-table.cjs                       T7b: 풀폭 대형 표
│   └── patch_table_merges.py                   ★ Stage 4 — 엑셀 셀 병합 → PPTX 표 병합 재현
│
└── scripts/
    ├── scan-sheets.py                          엑셀 스캐너 ([차트] 발견 시 lite 미지원 경고)
    └── run.cjs                                 deck.cjs 주문서 런너 (cover/toc/divider/bigtable/data2col)
```

---

## ⚡ 시작 전 확인

사용자에게 받을 것:

| # | 항목 | 형식 |
|---|---|---|
| 1 | **엑셀 파일 절대경로** | `/Users/.../data.xlsx` |
| 2 | **출력 파일명** | 기본 `deck.pptx` |
| 3 | **(선택) 표지 배경 이미지** | `cover-bg.png` 등 |

> 차트 요청은 lite 에서 지원하지 않으므로 강제로 진행하지 않는다.

> 🎨 **스킬 자체 톤 변경 요청 시** ("우리 회사 PPT 톤으로 lite 스타일 바꿔줘" 류): 빌드 흐름과 별개의 작업이다. **반드시 [customization-process.md](customization-process.md) 를 먼저 Read** 하고 그 5단계 절차를 따른다. 사용자에게 변경 항목 리스트를 보여주고 컨펌을 받기 전에 코드 수정 시작 금지.

---

### 📍 스킬 경로 — Claude Code · Claude Desktop 공통 처리

**작업 시작 시 첫 셸 명령으로 1회 export**:

```bash
# 스킬 호출 시 시스템이 안내한 "Base directory for this skill: …" 값을 그대로 넣음
export SKILL_DIR="<base-directory>"
```

| 환경 | 일반적인 base 경로 |
|---|---|
| Claude Code (로컬) | `/Users/{username}/.claude/skills/excel-to-pptx` |
| Claude Desktop / Claude.ai | `/mnt/skills/excel-to-pptx` 또는 `/mnt/user-data/skills/excel-to-pptx` |

이후 모든 명령은 `"$SKILL_DIR"/...` 형태로 실행 → 환경 무관 동작. 자세한 안내는 [README.md](README.md) 참조.

---

## 🎨 핵심 디자인 시스템 (요약)

| 용도 | 값 |
|---|---|
| 베이스 | `#FFFFFF` 화이트 |
| 디바이더 / 표지 | `#0A3D3A` 딥그린 · `#000000` 블랙 |
| 소목차 네이비 박스 | `#00205F` |
| 표 헤더 배경 | `#E8E8E8` |
| 표 맨왼쪽 컬럼 | `#EAEFF8` |
| 강조 (수치 / 표 외곽) | `#C0392B` 레드 |
| 본문 / 출처 | `#000000` / `#666666` |

- 폰트: **Pretendard Static** (Light · Regular · Medium · Bold) — 고정. ⚠️ `Pretendard Variable` 금지
- 슬라이드: **10 × 5.625 inch (16:9)**
- 사이즈: 대목차 10pt Light · 소목차 10pt Bold · 부제 11pt Medium · 메인 18pt Medium · 컬럼 헤더 12pt · 표 본문 9pt · 출처 8pt

---

## 🧭 제작 워크플로우 (5단계)

### 1. 스캐너 — 시트별 main / sub / table header (+샘플 3행)

```bash
python3 "$SKILL_DIR"/scripts/scan-sheets.py <엑셀>
```

- 시트 ≤10개 → stdout 전체 상세
- 시트 ≥11개 → `./bunyang-output/scan.txt` 저장 + stdout 요약

⚠ **[차트] 블록 검출 시**: scan 결과에 `[차트:...] ...` 표시가 있으면 즉시 사용자에게 안내:
> "이 시트(들)에 차트 블록이 있는데 lite 버전은 차트를 지원하지 않습니다. 엑셀에서 [차트] 블록을 [표] 로 바꾸거나 제거해주세요."

---

### 2. 전체 슬라이드 구성 결정

스캐너 결과 보고 **매핑표 초안** 작성:

| # | 슬라이드 | 타입 | 시트 | 비고 |
|---|---|---|---|---|
| 1 | 표지 | `cover` | — | overline·title·subtitle 작성 |
| 2 | 목차 | `toc` | — | 항목 작성 |
| 3 | Part 1 간지 | `divider` | — | 섹션 슬라이드 ≥2장일 때만 |
| 4 | 사업 개요 | `bigtable` | 사업개요 | 풀폭 |
| 5 | 분양 현황 | `data2col` | 분양현황 | 좌·우 |
| … | … | … | … | |

**판단 가이드 (lite 전용)**:
- 1블록 [표] → `bigtable`
- 2블록 [표][표] → `data2col`
- 3블록+ → 시트를 둘로 분리해야 함 (lite 는 1·2블록만)
- [차트] 포함 → 거부

**divider 삽입 기준**: 섹션 슬라이드 ≥2장일 때만.

---

### 3. 표지 / 목차 / 간지 작성

#### 표지(cover) 3층 구조

| 필드 | 시각 위치 | 내용 | 도출 방법 |
|---|---|---|---|
| `overline` | 맨 위 (선 박스 11pt) | **사업 주체 / 시행사명** | 엑셀에 회사명 있으면 그대로. 없으면 사업명으로 대체 |
| `title` | 중간 (21pt) | **지역 + 건물명** | 사업개요 시트의 대지위치 + 프로젝트명 조합 |
| `subtitle` | 큰 볼드 (45pt) | **보고서/제안서 종류** | 엑셀 파일명에서 추출 (`..._제안서_...` → "사업 제안서") |
| `date` | 하단 (11pt) | 작성 월 | 현재 연·월 |

> ⚠ **금지**: `overline`/`title`/`subtitle` 자리에 부제·홍보 문구·호재 나열 금지. "누가 / 어디를 / 무슨 문서" 세 줄로만.

```js
{ type: 'cover',
  overline: '(주)○○개발',
  title:    '군포 당동 오피스텔',
  subtitle: '간략 보고서',
  date:     '2026.04' }

// ⭐ toc.items 는 **생략 권장** — 같은 deck 안의 divider 들 title 로 자동 동기화됨
{ type: 'toc' }

{ type: 'divider', chapter: 'Part 1', title: '사업 개요' }
{ type: 'divider', chapter: 'Part 2', title: '분양 현황' }
// → 위 두 divider 가 자동으로 toc 에 "01. 사업 개요 / 02. 분양 현황" 으로 들어감
```

> ⚠ **목차 항목 = 간지(divider) 개수 == 자동 일치**: `toc.items` 를 직접 적으면 divider 와 개수가 어긋나기 쉽다. 비워두면 런너가 같은 deck 안의 divider title 들을 순서대로 넘버링한다. 정 직접 적고 싶다면 divider 개수와 정확히 같게 맞출 것.

---

### 4. 시트 기반 데이터 슬라이드 yaml 작성 ⭐

#### 4-A. 각 슬라이드 sectionLabel·subhead·mainMessage 명시

| 필드 | 내용 |
|---|---|
| `sectionLabel` | **반드시 `/` 구분자**. `"Part 1 / 사업 개요"`. 앞=대목차(회색 Light), 뒤=네이비 박스 안 흰 Bold. `/` 없으면 박스 사라짐 |
| `subhead` | 부제 (검정 Medium 11pt). **엑셀 2행 그대로 복붙**. 축약 금지 |
| `mainMessage` | **엑셀 1행 그대로 복붙**. 축약·재구성 금지 |

#### 4-B. deck.cjs 예시

```js
module.exports = {
  excel: './data.xlsx',
  out:   './deck.pptx',
  slides: [
    { type: 'cover',  overline: '(주)○○개발', title: '군포 당동',
      subtitle: '오피스텔', date: '2026.04' },

    // ⭐ items 생략 — 아래 divider 들의 title 로 자동 동기화
    { type: 'toc' },

    { type: 'divider', chapter: 'Part 1', title: '사업 개요' },

    // [표] 1개 → bigtable sheet 기반 (rows·병합·main/sub 전부 자동)
    { type: 'bigtable',
      sheet: '사업개요',
      sectionLabel: 'Part 1 / 사업 개요' },

    { type: 'divider', chapter: 'Part 2', title: '분양 현황' },

    // [표][표] 2개 → data2col sheet 기반
    { type: 'data2col',
      sheet: '군포시_분양현황',
      sectionLabel: 'Part 2 / 분양 현황' }
  ]
};
```

> ⚠ **[표] inline 작성 금지**: `sheet: '...'` 만 지정하면 파서가 엑셀 전체 행·열·병합을 자동 반영. inline rows 로 옮겨 적으면 데이터 손실 위험.

> ⚠ **표가 슬라이드 밖으로 살짝 넘치는 것 < 데이터 손실**. 엑셀 원본을 보존하고, 크기 조정은 사용자가 PPT 열어 직접.

---

### 5. PPT 생성

```bash
node "$SKILL_DIR"/scripts/run.cjs deck.cjs
```

런너가 자동 수행 (4 stage):
1. **엑셀 전체 덤프** (openpyxl · rows + 셀 병합 메타)
2. `slides[]` 순회 → T5/T3/T7b/T1-table 디스패치 → **Stage 1 좌표 검증 자동**
3. **PPTX 저장**
4. **Stage 4**: patch_table_merges.py (엑셀 셀 병합 → PPTX 표 병합 XML 주입)

차트 폰트 패치 단계 없음 (차트 미지원).

---

## 📇 deck.cjs 슬라이드 타입 레퍼런스

```js
// 표지 — overline/title/subtitle 3층 (단계 3 표 참조)
{ type: 'cover', overline?, title, subtitle?, date?, logoText?, bgImagePath?, bgColor? }

// 목차 — items 는 생략 권장 (같은 deck 안 divider title 들로 자동 동기화)
{ type: 'toc', project?, title?, items?: [...] }

// 챕터 디바이더
{ type: 'divider', chapter: 'Part 1', title: '사업 개요\n검토', bgImagePath?, bgColor? }

// T7b 풀폭 표 (블록 1개 [표] 시트 기본, sheet 기반 권장)
{ type: 'bigtable',
  sheet?: '사업개요',                           // 시트명 — 자동 매핑
  sectionLabel, subhead?, mainMessage?,
  tableHeader?: '사업지 개요',                   // 미지정 시 블록 제목 자동
  // sheet 없을 때만 inline:
  table?: {
    rows: [['구분','내용', ...], ['대지위치', '경기 ...', ...], ...],
    leftColLabel?: true,                        // 첫 컬럼 #EAEFF8 (default true)
    emphRowIndices?: [3],                       // 강조 행 (0-based, 헤더 제외)
    fontSize?: 9,
    colW?: [1.4, 3.5, 1.4, 2.86]                // 컬럼 너비 (합 = 9.16)
  },
  source? }

// T1 좌·우 2단 표 (블록 2개 [표] 시트 기본, sheet 기반 권장)
{ type: 'data2col',
  sheet?: '분양현황',                            // 시트명 — 좌/우 슬롯 자동
  sectionLabel, subhead?, mainMessage?,
  // sheet 없을 때만 inline:
  leftHeader,
  leftSlot:  { type: 'table', rows: [...], leftColLabel?, emphRowIndices?, ... },
  rightHeader,
  rightSlot: { type: 'table', rows: [...] },
  source? }
```

---

## 📁 엑셀 시트 작성 약속 (사용자가 지킬 것)

```
A1: [키 메시지]                ← 메인 메시지 (18pt Medium)
A2: [서브 분석 문장]           ← 부제 (11pt Medium)
(빈 행)
A4: [표] 평당가 추이           ← 표로 렌더 — PPT 엔 prefix 제거되어 나감
A5: 지역  2023  2024  2025    ← 헤더
A6~: 데이터
(선택) [출처: 부동산114]       ← 마지막 라인
```

### 블록 제목 prefix 규칙

| prefix | lite 처리 |
|---|---|
| `[표] 제목` | ✅ 표 (native table) |
| prefix 없음 | ✅ 표로 간주 |
| `[차트] ...` | ❌ 거부 (lite 미지원) |

---

## ✅ 디자인 체크리스트 (요약)

Stage 1(런너) 자동 검증.

### 헤더
- [ ] 대목차 — 박스 밖 검정 Light 10pt
- [ ] 소목차 — `#00205F` 네이비 박스 안 흰색 Bold 10pt
- [ ] 부제(검정 Medium 11pt) → 메인(검정 Medium 18pt)

### 본문
- [ ] 표 native
- [ ] 데이터 = 엑셀 원본 1:1
- [ ] 표 헤더 `#E8E8E8` + 검정 Bold
- [ ] 표 맨왼쪽 컬럼 `#EAEFF8`
- [ ] 표 외곽 — 좌·우 세로선 없음, 상·하 2px 검정, 내부 1px 연회색
- [ ] 강조 행 — 외곽 2px 빨강

### 하단·여백
- [ ] 출처 8pt `#666`
- [ ] 좌우 마진 0.42 inch
- [ ] Pretendard Static 패밀리만

---

## 🚫 피해야 할 것

- ❌ 차트 요청을 lite 로 강제 진행 → lite 는 차트 미지원
- ❌ 표 데이터를 scan-sheets 샘플 3행만 보고 그대로 넣기 → **sheet 기반 자동** 사용
- ❌ 엑셀 데이터 임의 가공·축약·재구성 — 원본 1:1 보존
- ❌ `sectionLabel` 에 `/` 없이 적기 → 네이비 박스 사라짐
- ❌ 엑셀 1·2행 메시지를 내 기준으로 요약·단축 → 원문 그대로 복붙
- ❌ 1장짜리 섹션에 divider 삽입
- ❌ 표 4개+ 시트를 lite 로 처리 → 시트 분리 필요
- ❌ **표지 3층 자리에 부제성 설명문·홍보 문구·호재 나열 넣기**
- ❌ `Pretendard Variable`

---

## 의존성

```bash
# Node
npm install -g pptxgenjs
# 또는 작업 폴더에서 npm install pptxgenjs

# Python
pip install openpyxl Pillow
```

