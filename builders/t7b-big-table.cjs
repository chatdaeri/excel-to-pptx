/**
 * T7b: 풀폭 대형 표 1장 (T7 의 표 버전)
 *
 * 구성:
 *   [공통 헤더]  — 대목차 / 소목차 / 부제 / 메인 메시지
 *   [컬럼 헤더]  — 풀폭 (y=1.78)
 *   [대형 표 슬롯] — 풀폭 (x=0.42, y=2.06, w=9.16, h=3.00)
 *   [출처]
 *
 * 용도: [표] 블록이 1개뿐인 시트(사업개요·입지평가 등) 에서 표를
 *      좌/우 분할 없이 풀폭으로 전체 데이터 그대로 렌더링할 때.
 *      T1 에 억지로 축약해 넣지 말 것.
 *
 * opts = {
 *   sectionLabel, subhead, mainMessage, source,
 *   tableHeader: '사업지 개요',
 *   table: {
 *     rows: [[헤더행], [데이터행1], ...],       // 1행=헤더, 이후 데이터
 *     headers?: [...],                         // 명시적 헤더 (rows 는 순수 데이터)
 *     leftColLabel?: true,                     // 첫 컬럼 #EAEFF8 배경 (default true)
 *     emphRowIndices?: [3],                    // 강조 행 (빨간 외곽)
 *     fontSize?: 9,
 *     colW?: [1.2, 1.2, 4.2, 1.26, 1.3],      // 컬럼 너비 (inch). 합 = 9.16 권장
 *     opts?: { ... pptxgenjs tableOpts 오버라이드 }
 *   }
 * }
 *
 * v (Validator, optional)
 */

const { LAYOUT, addAnalysisHeader, addColumnHeader, addSource, tableOpts, resolveTableRows } = require('./_helpers.cjs');

function buildT7b(pres, opts, v) {
  const slide = pres.addSlide();
  if (v) v.beginSlide();

  addAnalysisHeader(slide, pres, opts, v);
  addColumnHeader(slide, pres, 'full', opts.tableHeader || '', v);

  const table = opts.table || { rows: [] };
  const rows = resolveTableRows(table);

  const baseOpts = { ...(table.opts || {}) };
  if (Array.isArray(table.colW)) baseOpts.colW = table.colW;

  slide.addTable(rows, tableOpts(LAYOUT.SLOT_FULL, baseOpts));
  if (v) v.element({ type: 'table', ...LAYOUT.SLOT_FULL, label: 'bigTable' });

  addSource(slide, opts.source, v);
  if (v) v.requireSource();
  return slide;
}

module.exports = { buildT7b };
