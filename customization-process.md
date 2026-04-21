# 커스텀 작업 매뉴얼 (AI 전용)

**🤖 이 문서는 AI(Claude) 가 사용자로부터 "회사 톤으로 excel-to-pptx 스킬 스타일을 바꿔줘" 요청을 받았을 때 따라야 할 작업 절차다. 사용자(사람) 가 읽으라고 만든 문서가 아니다 — 사용자 안내는 README.md 의 "회사 스타일에 맞게 바꾸기" 섹션에 있다.**

---

## 0. 언제 이 매뉴얼을 쓰나

사용자가 다음 같은 요청을 할 때:

- "우리 회사 PPT 톤에 맞게 excel-to-pptx 스킬 스타일을 바꿔줘"
- "이 PPTX 보고 컬러·폰트 우리 회사 톤으로 적용"
- 회사 가이드 이미지/PPTX 를 던지면서 "이 스타일로 톤 갈아줘"
- "표 헤더 색을 회색 말고 네이비로 바꾸고 싶어"

**중요**: excel-to-pptx 스킬은 차트·지도·컨셉 박스·3블록 자동 슬라이드·inline 차트 series 재구성 같은 기능이 없다. 만약 사용자가 "차트 색도 바꿔줘" 같은 요청을 하면 excel-to-pptx 가 차트를 지원하지 않는다고 안내한다.

---

## 1. 5단계 작업 흐름

### 1단계 — 사용자한테 받을 것

**필수**:
- 참고 PPTX 1~2장. 일반 데이터 슬라이드 한 장(메인 메시지·표가 보이는 것) + 표지 한 장 권장.
- 파일 형식: **PPTX 강력 권장**. JPG/PNG 도 가능하지만 PPTX 안에는 정확한 HEX 컬러·폰트 패밀리·폰트 사이즈가 XML 로 들어 있어서 추측 없이 정확하게 분석 가능.

**받으면 좋은 것 (없으면 기본값 유지 가능)**:

| 항목 | 기본값 | 변경 시 사용자에게 확인할 질문 |
|---|---|---|
| 메인 컬러 (네이비 박스·컬럼 헤더 막대) | `#00205F` | "박스 컬러는 어떤 HEX 로?" |
| 강조 컬러 (강조 행 외곽·텍스트) | `#C0392B` | "강조 컬러는?" |
| 표 헤더 배경 / 글씨 | 회색 `#E8E8E8` + 검정 글씨 | "표 헤더 배경 — 회색 유지? 네이비? 다른 색?" |
| 표 좌측 라벨 컬럼 배경 | 연파랑 `#EAEFF8` | "유지? 제거? 다른 색?" |
| 폰트 패밀리 | Pretendard Static | "Noto Sans KR / 자체 폰트 / Pretendard 유지?" — Static 만 지원 (Variable 금지) |
| 표지 배경 | 딥그린 `#0A3D3A` | "단색 / 이미지 / 그라데이션?" |
| 디바이더 배경 | 딥그린 | "표지와 같은 톤? 다른 톤?" |
| 소목차 박스 모양 | 네이비 사각 박스 + 흰 글씨 | "박스 유지 / 밑줄 / 수직 막대 / 연한 배경 박스?" |

폰트 패밀리를 바꾸는 경우 **Static 폰트 파일이 시스템에 설치되어 있어야 PPT 가 제대로 표시된다**. 사용자에게 "이 폰트가 PowerPoint 에서 보이게 하려면 폰트 파일이 필요한데, 시스템에 이미 설치돼 있나요?" 확인. 없으면 설치 방법 안내 (또는 assets 폴더에 폰트 파일 받아 두기).

### 2단계 — PPTX 분석

PPTX 는 ZIP 압축 파일이라 unzip 으로 풀어볼 수 있다.

```bash
mkdir /tmp/refppt && cd /tmp/refppt && unzip <ref.pptx>
# ppt/slides/slide1.xml, slide2.xml 등 확인
# ppt/theme/theme1.xml 도 폰트 정보 들어 있음
```

각 슬라이드 XML 에서 추출할 것:

| 정보 | XML 패턴 |
|---|---|
| 폰트 패밀리 | `<a:rPr ... typeface="Pretendard Bold">` 의 `typeface` 속성 |
| 폰트 크기 | `<a:rPr sz="1800">` → 18pt (×100 단위) |
| 굵기 | `<a:rPr b="1">` → Bold |
| 텍스트 색 | `<a:solidFill><a:srgbClr val="00205F"/></a:solidFill>` |
| 도형 배경 | `<a:solidFill>` 안 `srgbClr` |
| 표 셀 배경 | `<a:tcPr><a:fill><a:solidFill>` |
| 보더 굵기·색 | `<a:lnL/lnR/lnT/lnB w="25400">` (EMU 단위, 12700=1pt) |

위 정보를 **현재 excel-to-pptx 의 기본값과 대조**해서 변경 항목 리스트를 만든다. excel-to-pptx 의 현재 기본값은 `builders/_helpers.cjs` 상단의 `C` 객체와 `FONT_*`, `BORDER_*` 상수에 있다.

### 3단계 — 변경 리스트 사용자에게 보여주고 컨펌

이런 형태로 정리해서 보여준다:

```
참고 PPTX 분석 결과 — 변경할 항목 리스트:

[컬러]
  · 메인 네이비 컬러: #00205F → #1A4D8C  (조금 더 밝은 네이비)
  · 강조 컬러: #C0392B → #E94560  (브랜드 레드)
  · 표 헤더 배경: #E8E8E8 회색 → #1A4D8C 네이비 (글씨 흰색으로 반전)
  · 표 좌측 라벨 배경: #EAEFF8 → 그대로 유지

[폰트]
  · 패밀리: Pretendard → Noto Sans KR
  · 메인 메시지 사이즈: 18pt → 20pt Bold
  · 그 외 사이즈: 그대로 유지

[표지/디바이더]
  · 표지 배경: 딥그린 단색 → 네이비 그라데이션
  · 디바이더 배경: 딥그린 → 화이트 + 검정 텍스트
  · 표지에 회사 로고 추가 (assets/logo.png 위치 필요)

[기타]
  · 소목차 박스: 네이비 사각 박스 → 왼쪽 수직 막대 + 검정 글씨
  · 강조 행 처리: 빨간 외곽선 유지

→ 이대로 진행해도 될까요? 부분 수정 필요한 항목 있으면 짚어 주세요.
```

사용자 컨펌 받기 전에 코드 수정 절대 시작하지 말 것. 사용자가 "메인 컬러는 그대로 두고 폰트만 바꿔주세요" 같은 부분 요청을 하면 그 부분만 진행.

### 4단계 — 코드 수정

#### 4-A. 변경 지점 맵 (excel-to-pptx 전용)

| 변경 항목 | 위치 |
|---|---|
| 컬러 팔레트 (`C.navy`, `C.red`, `C.blueLight`, `C.grayHeader` 등) | `builders/_helpers.cjs` 상단 `C` 객체 |
| 폰트 패밀리 | `builders/_helpers.cjs` 상단 `FONT*` 상수 5개 |
| 표 보더 (굵기·색·없음 여부) | `builders/_helpers.cjs::BORDER_*` 상수 5개 |
| 표 보더 배치 (외곽 4변 감쌀지, 내부선 숨길지) | `builders/_helpers.cjs::decideBorders` 함수 |
| 표 헤더/좌측 라벨/강조 배경·글씨 | `builders/_helpers.cjs::buildTableRows` 안 `fill`/`color`/`ff` 삼항 분기 |
| 헤더 4단(대목차/소목차/부제/메인) 사이즈·여백 | `builders/_helpers.cjs::addAnalysisHeader` + `LAYOUT` 상수 |
| 소목차 박스 모양 (사각/밑줄/수직막대/배경색) | `builders/_helpers.cjs::addAnalysisHeader` 안 `addShape(rect)` 부분 교체 |
| 컬럼 헤더 (네이비 막대 + 텍스트) | `builders/_helpers.cjs::addColumnHeader` |
| 출처 (위치·크기·정렬) | `builders/_helpers.cjs::addSource` + `LAYOUT.SOURCE` |
| 표지 (배경·로고·타이포 계층) | `builders/t5-cover-toc.cjs::buildCover` |
| 목차 (배경·타이포) | `builders/t5-cover-toc.cjs::buildToc` |
| 디바이더 (배경·레이아웃) | `builders/t3-divider.cjs` 전체 |

#### 4-B. 핵심 원칙

1. **상수 이름은 절대 바꾸지 말 것**. `C.navy` 를 `C.primary` 로 리네임하면 `t1-data-table.cjs`, `t7b-big-table.cjs` 등 다른 빌더의 참조가 다 깨진다. **값만 교체**.
2. **새 컬러가 필요하면 `C` 객체에 추가만**. 예: `C.brandOrange = 'FF6F00'`. 기존 키는 건드리지 않는다.
3. **새 Weight 가 필요하면 상수 추가** + `module.exports` 에 등록. 예: `FONT_EXTRABOLD = 'Noto Sans KR ExtraBold'`.
4. **사이즈를 키우면 LAYOUT 의 `h` 도 같이 조정**. 예를 들어 `mainMessage` 사이즈를 18pt → 24pt 로 키우면 `LAYOUT.MAIN_MSG.h` 를 0.50 → 0.70 정도로. 안 키우면 텍스트가 잘림.
5. **excel-to-pptx 의 `decideBorders` 함수는 빨간 외곽선(emph) 로직이 들어 있다**. 강조 행 처리 방식을 외곽선 → 배경 tint 로 바꾸려면 `decideBorders` 의 `BORDER_EMPH` 사용 부분을 다 제거하고 `buildTableRows` 의 `fill` 분기에 `isEmph` 케이스 추가.

#### 4-C. 자주 쓰는 변형 예시

**예시 1 — 표 헤더 배경 네이비 + 흰 글씨**:

```js
// builders/_helpers.cjs::buildTableRows 안 (현재 코드)
const fill = isHeader ? { color: C.grayHeader }
            : isLeftLabel ? { color: C.blueLight }
            : undefined;
const color = isEmph ? C.red : C.text;

// 변경 후
const fill = isHeader ? { color: C.navy }            // 헤더 배경 네이비
            : isLeftLabel ? { color: C.blueLight }
            : undefined;
const color = isHeader ? C.white                      // 헤더 글씨 흰색
            : isEmph   ? C.red
            :            C.text;
```

**예시 2 — 소목차 네이비 박스 → 왼쪽 수직 막대**:

```js
// builders/_helpers.cjs::addAnalysisHeader 안, 소목차 그리는 addShape(rect) 부분 교체
slide.addShape(pres.ShapeType.rect, {
  x: LAYOUT.SUB_LABEL_BOX.x, y: LAYOUT.SUB_LABEL_BOX.y,
  w: 0.06, h: LAYOUT.SUB_LABEL_BOX.h,                 // 가는 세로 막대
  fill: { color: C.navy }, line: { color: C.navy, width: 0 }
});
slide.addText(subLabel, {
  x: LAYOUT.SUB_LABEL_BOX.x + 0.14,                   // 막대 옆 글씨
  y: LAYOUT.SUB_LABEL_BOX.y,
  w: estW - 0.14, h: LAYOUT.SUB_LABEL_BOX.h,
  color: C.black, fontFace: FONT_BOLD, fontSize: 10
});
```

**예시 3 — 강조 행을 빨간 외곽선 → 연노랑 배경 + 빨간 글씨**:

```js
// 1) builders/_helpers.cjs::decideBorders 에서 BORDER_EMPH 사용 부분 다 제거
//    (top/bottom/left/right 분기에서 isEmph 조건을 빼고 일반 BORDER_THIN/THICK 만 사용)

// 2) builders/_helpers.cjs::buildTableRows 의 fill 분기 수정
const fill = isHeader    ? { color: C.grayHeader }
           : isEmph      ? { color: 'FFF4E5' }       // 연노랑 배경 추가
           : isLeftLabel ? { color: C.blueLight }
           : undefined;
// color, ff 는 그대로 (강조 글씨는 여전히 빨간 Bold)
```

**예시 4 — 표지 단색 → 회사 로고 + 그라데이션**:

```js
// builders/t5-cover-toc.cjs::buildCover 안
slide.background = { color: bg };  // 단색은 그대로

// 로고 추가 (assets/logo.png 사용자에게 받아서 skill 폴더 안에 배치)
const path = require('path');
const fs = require('fs');
const logoPath = path.join(__dirname, '..', 'assets', 'logo.png');
if (fs.existsSync(logoPath)) {
  slide.addImage({ path: logoPath, x: 0.5, y: 0.5, w: 1.5, h: 0.5 });
}
```

### 5단계 — 테스트 빌드 + 시각 확인

수정 끝나면 사용자에게 미리 만들어둔 deck.cjs (또는 간단한 테스트용) 로 빌드 한 번 돌려서 결과 확인하라고 안내.

```bash
node "$SKILL_DIR"/scripts/run.cjs ./deck.cjs
```

PPT 파일을 사용자가 직접 열어서:
- 표지·디바이더 배경이 의도대로 나오는지
- 표 헤더 배경·글씨 대비가 충분한지
- 폰트가 깨지지 않고 잘 적용됐는지 (특히 폰트 바꿨을 때)
- 강조 행이 의도한 대로 보이는지

확인 후 추가 조정 1~2번 더 들어가는 게 보통이다.

---

## 2. 흔한 실수

1. **`C` 객체 키 이름 바꾸지 말 것** — 값만 교체. 이름 바꾸면 다른 빌더 참조 깨짐.
2. **폰트 바꿀 때 사용자 시스템에 폰트 설치 여부 확인** — PowerPoint 가 폰트를 못 찾으면 fallback 폰트로 표시돼서 의도가 안 살아남.
3. **사용자 컨펌 없이 코드 수정 시작 금지** — 변경 리스트 보여주고 OK 받은 다음 시작.
4. **차트 관련 요청은 거절** — excel-to-pptx 는 차트 자체가 없다. "차트 컬러도 바꿔줘" 같은 요청은 지원 불가 안내.
5. **Stage 1 좌표 검증 무시 금지** — LAYOUT 좌표를 임의로 슬라이드 경계 밖으로 이동시키면 빌드가 멈춘다. 마진(좌우 0.42 inch) 안에서 작업.
6. **테스트 빌드 없이 끝내기 금지** — 수정 후 한 번이라도 빌드 돌려서 시각 확인 권장.

---

## 3. 빠른 요약 — 어떤 요청에 어디 고치나

| 사용자 요청 | 먼저 열 파일 | 추가로 건드릴 곳 |
|---|---|---|
| "메인 컬러 바꿔줘" | `_helpers.cjs::C.navy` | — |
| "표 헤더 네이비 + 흰 글씨" | `_helpers.cjs::buildTableRows` 의 fill·color 분기 | `C.navy` 값 사용 |
| "강조 행 색 바꿔줘" | `_helpers.cjs::C.red` + `BORDER_EMPH` | `buildTableRows` color 분기 |
| "강조 행을 외곽선 → 배경색으로" | `_helpers.cjs::decideBorders` (BORDER_EMPH 제거) + `buildTableRows::fill` 분기 추가 | — |
| "폰트 Noto Sans KR 로" | `_helpers.cjs::FONT*` 상수 5개 | 시스템에 폰트 설치 여부 확인 |
| "메인 메시지 22pt Bold 로" | `_helpers.cjs::addAnalysisHeader` 안 메인 메시지 fontSize | `LAYOUT.MAIN_MSG.h` 같이 키우기 |
| "소목차 박스 → 밑줄/수직막대" | `_helpers.cjs::addAnalysisHeader` 안 addShape | `LAYOUT.SUB_LABEL_BOX` 좌표도 가능 |
| "표 외곽 4변 다 굵게" | `_helpers.cjs::decideBorders` left/right 분기 | `BORDER_THICK` 사용 |
| "표 내부선 숨겨" | `BORDER_THIN` 값을 `{ type: 'none' }` 으로 | 또는 `decideBorders` 에서 BORDER_NONE 반환 |
| "표 좌측 라벨 배경 제거" | `_helpers.cjs::buildTableRows` 의 isLeftLabel 분기 fill 제거 | — |
| "표지 배경 그라데이션 / 로고 추가" | `builders/t5-cover-toc.cjs::buildCover` | assets 폴더에 로고 파일 |
| "디바이더 디자인 새로" | `builders/t3-divider.cjs` 전체 | 참고 이미지 받기 |
| "출처 위치 / 크기 바꿔" | `_helpers.cjs::addSource` + `LAYOUT.SOURCE` | — |
| "컬럼 헤더 막대 → 언더라인" | `_helpers.cjs::addColumnHeader` 안 addShape | rect → line 으로 |

---

## 4. 적용 범위 안내 (사용자에게 미리 알려줄 것)

수정 끝낸 뒤 사용자에게 다음을 명시:

- **이 변경은 excel-to-pptx 스킬 자체에 적용된다**. 그 결과 excel-to-pptx 로 만드는 모든 PPT 가 같은 톤으로 빌드된다.
- 만약 회사 A 톤과 회사 B 톤을 둘 다 써야 한다면 excel-to-pptx 폴더를 통째로 복사해서 `excel-to-pptx-companyA` / `excel-to-pptx-companyB` 두 개로 분리하는 방법이 있다 (사용자가 원하면 안내).
- 변경된 파일을 git 으로 백업해두면 나중에 원래 톤으로 돌리기 쉽다.
