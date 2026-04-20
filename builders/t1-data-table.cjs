/**
 * T1 (lite): 데이터 2단 — 좌·우 슬롯 모두 표 전용
 *
 *   ✗ chart slot 미지원 (lite 는 차트 미지원)
 *   ✓ 셀 병합·강조 행·좌측 라벨 컬럼 모두 동작
 *
 * opts = {
 *   sectionLabel: "Part 1 / 분양 현황",
 *   subhead:      "동구·서구 동반 상승",
 *   mainMessage:  "광주 주요 지역 시세 상승세",
 *   leftHeader:   "주요 단지 시세",
 *   leftSlot:  { type: 'table', rows: [[...]], headers?: [...], leftColLabel?, emphRowIndices?, fontSize?, colAlign?, merges?, opts? },
 *   rightHeader: "...",
 *   rightSlot: { type: 'table', ... },
 *   source: "[출처: ...]"
 * }
 *
 * v (Validator, optional): 전달 시 Stage 1 검증 자동 참여
 */

const { LAYOUT, addAnalysisHeader, addColumnHeader, addSource, tableOpts, resolveTableRows } = require('./_helpers.cjs');

function buildT1Table(pres, opts, v) {
  const slide = pres.addSlide();
  if (v) v.beginSlide();

  addAnalysisHeader(slide, pres, opts, v);
  addColumnHeader(slide, pres, 'left',  opts.leftHeader,  v);
  addColumnHeader(slide, pres, 'right', opts.rightHeader, v);

  fillTable(slide, LAYOUT.SLOT_LEFT,  opts.leftSlot,  v, 'leftSlot');
  fillTable(slide, LAYOUT.SLOT_RIGHT, opts.rightSlot, v, 'rightSlot');

  addSource(slide, opts.source, v);
  if (v) v.requireSource();
  return slide;
}

function fillTable(slide, coord, slot, v, label) {
  if (!slot) return;
  if (slot.type && slot.type !== 'table') {
    throw new Error(
      `excel-to-pptx: data2col slot.type 은 'table' 만 지원합니다 (받은 값: '${slot.type}'). ` +
      `lite 는 차트를 지원하지 않습니다.`
    );
  }
  const rows = resolveTableRows(slot);
  slide.addTable(rows, tableOpts(coord, slot.opts || {}));
  if (v) v.element({ type: 'table', ...coord, label });
}

module.exports = { buildT1Table };
