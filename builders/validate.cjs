/**
 * Stage 1 자동 좌표·속성 검증
 *
 * 빌더가 아래 훅을 호출하면 Validator 가 자동 등록·검사한다.
 *
 *   const v = new Validator();
 *   buildT1(pres, opts, v);   // 각 빌더가 내부에서 v.beginSlide() → v.element() 호출
 *   v.print();                // 결과 출력 (실패 시 false 반환 → process.exit 결정)
 *
 * 등록 훅:
 *   v.beginSlide()                                 // pres.addSlide() 직후 1회
 *   v.element({ type, x, y, w, h, fontSize, label }) // 현재 슬라이드 소속
 *   v.requireSource()                              // T1·T2·T4 — source 누락 검출
 *
 * 검증 항목:
 *   1. overflow  (슬라이드 영역 10×5.625 밖)
 *   2. margin    (좌우 0.42, 상 0.31, 하 0.31 미달)
 *   3. min font  (7pt 미만 — 출처 8pt 허용)
 *   4. 출처 누락 (requireSource 호출됐는데 label='source' element 없음)
 */

const SLIDE_W = 10;
const SLIDE_H = 5.625;
const MARGIN_X = 0.42;
const MARGIN_TOP = 0.31;
const MARGIN_BOTTOM = 0.31;
const MIN_FONT = 7;

class Validator {
  constructor() {
    this.currentSlideIdx = -1;
    this.elements = [];
    this.issues = [];
    this.requireSourceSlides = new Set();
    this.haveSourceSlides = new Set();
  }

  beginSlide() {
    this.currentSlideIdx++;
    return this.currentSlideIdx;
  }

  element({ type, x, y, w, h, fontSize, label }) {
    if (this.currentSlideIdx < 0) return;   // beginSlide 전 호출 무시
    const slideIdx = this.currentSlideIdx;
    this.elements.push({ slideIdx, type, x, y, w, h, fontSize, label });
    const id = `슬라이드 ${slideIdx + 1} [${label || type}]`;

    if (x + w > SLIDE_W + 0.05)
      this.issues.push(`${id}: 우측 overflow (x+w=${(x+w).toFixed(2)} > ${SLIDE_W})`);
    if (y + h > SLIDE_H + 0.05)
      this.issues.push(`${id}: 하단 overflow (y+h=${(y+h).toFixed(2)} > ${SLIDE_H})`);
    if (x < -0.02)
      this.issues.push(`${id}: 좌측 음수 좌표 (x=${x})`);
    if (y < -0.02)
      this.issues.push(`${id}: 상단 음수 좌표 (y=${y})`);

    if (['text', 'chart', 'table'].includes(type)) {
      if (x < MARGIN_X - 0.02)
        this.issues.push(`${id}: 좌측 마진 미달 (x=${x} < ${MARGIN_X})`);
      if (x + w > SLIDE_W - MARGIN_X + 0.02)
        this.issues.push(`${id}: 우측 마진 미달 (x+w=${(x+w).toFixed(2)} > ${(SLIDE_W-MARGIN_X).toFixed(2)})`);
    }

    if (type === 'text' && fontSize && fontSize < MIN_FONT)
      this.issues.push(`${id}: 폰트 ${fontSize}pt < 최소 ${MIN_FONT}pt`);

    if (label === 'source') this.haveSourceSlides.add(slideIdx);
  }

  requireSource(slideIdx) {
    if (slideIdx === undefined) slideIdx = this.currentSlideIdx;
    if (slideIdx >= 0) this.requireSourceSlides.add(slideIdx);
  }

  report() {
    for (const idx of this.requireSourceSlides) {
      if (!this.haveSourceSlides.has(idx))
        this.issues.push(`슬라이드 ${idx + 1}: 분석 장표인데 출처 누락`);
    }
    return this.issues;
  }

  print() {
    const issues = this.report();
    if (issues.length === 0) {
      console.log(`✅ Stage 1 좌표 검증 통과 (슬라이드 ${this.currentSlideIdx + 1}장 / element ${this.elements.length}건)`);
      return true;
    }
    console.log(`❌ Stage 1 좌표 검증 ${issues.length}건:`);
    issues.forEach(i => console.log(`   - ${i}`));
    return false;
  }
}

module.exports = { Validator };
