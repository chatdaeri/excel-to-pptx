/**
 * Auto 모드 파서 — 엑셀 시트 한 장 → 슬라이드 스펙 한 개
 *
 * 컨벤션 (사용자 약속):
 *   1행       → 키 메시지 (메인 메시지)
 *   2행       → 서브 메시지 (부제) — 비어있으면 AI가 섹션라벨/시트명으로 보완
 *   이후      → 빈 행 기준으로 블록 N개 분리
 *   각 블록   → 1행: 블록 제목 (보통 A열 단독)
 *                  · "[표] 제목"          → kind = 'table' 로 렌더
 *                  · "[차트] 제목"        → kind = 'chart' (default: bar)
 *                  · "[차트:line] 제목"   → kind = 'chart', chartType = 'line'
 *                                          (bar | line | doughnut | barLine)
 *                  · prefix 없으면 table 로 간주
 *              2행: 표 헤더
 *              3행~: 데이터
 *              (선택) 마지막 행이 "출처" 로 시작하면 해당 블록의 출처로 추출
 *
 * 시트명은 섹션 라벨 자동 생성에 사용.
 *   "01_사업환경_년도별분양물량" → "Ⅰ. 사업환경 / 년도별분양물량"
 *
 * 사용:
 *   const { parseSheet, deriveLayout } = require('./excel-sheet-parser.cjs');
 *   const spec = parseSheet('01_시장환경_거래량', rows2D);
 *   // spec = { sheetName, sectionLabel, mainMessage, subMessage, blocks: [...] }
 */

function parseSheet(sheetName, rows2D, merges = []) {
  const rows = (rows2D || []).map(r => Array.isArray(r) ? r : []);

  let idx = 0;
  idx = skipEmpty(rows, idx);

  // 1행: 키 메시지 (메인)
  let mainMessage = '';
  if (idx < rows.length) {
    mainMessage = getFirstNonEmpty(rows[idx]);
    idx++;
  }

  // 2행: 서브 메시지 (부제). 비어있어도 idx 는 전진
  let subMessage = '';
  if (idx < rows.length && !isEmptyRow(rows[idx])) {
    subMessage = getFirstNonEmpty(rows[idx]);
    idx++;
  }

  // 이후: 블록 수집
  const blocks = [];
  while (idx < rows.length) {
    idx = skipEmpty(rows, idx);
    if (idx >= rows.length) break;

    const blockStartRow = idx;  // absolute 0-indexed row in sheet
    const block = [];
    while (idx < rows.length && !isEmptyRow(rows[idx])) {
      block.push(rows[idx]);
      idx++;
    }
    const parsed = parseBlock(block, merges, blockStartRow);
    if (parsed) blocks.push(parsed);
  }

  return {
    sheetName,
    sectionLabel: deriveSectionLabel(sheetName),
    mainMessage,
    subMessage,
    blocks
  };
}

/**
 * 블록 내부 파싱
 *   - 최소 2행 필요 (제목 + 헤더 or 헤더 + 데이터)
 *   - 첫 행이 A열만 값 → 제목, 둘째 행 → 헤더
 *   - 첫 행이 여러 열에 값 → 헤더로 간주, 제목 없음
 *   - 마지막 행 텍스트가 "출처" 포함 + 단일 값 → 해당 블록 출처
 */
// [표] / [차트] / [차트:타입] prefix 파서
// 반환: { kind: 'table'|'chart', chartType?: string|null, title: string }
function parsePrefix(raw) {
  const s = String(raw || '').trim();
  const m = s.match(/^\[(표|차트)(?::([A-Za-z]+))?\]\s*/);
  if (!m) return { kind: 'table', chartType: null, title: s };
  const kind = m[1] === '차트' ? 'chart' : 'table';
  const chartType = m[2] ? m[2].toLowerCase() : null;
  const title = s.slice(m[0].length).trim();
  return { kind, chartType, title };
}

function parseBlock(block, allMerges = [], blockStartRow = 0) {
  if (block.length === 0) return null;

  let i = 0;
  let title = '';
  let kind = 'table';
  let chartType = null;

  const firstRow = block[0];
  const firstRowValues = firstRow.filter(v => v != null && String(v).trim() !== '');
  if (firstRowValues.length === 1 && block.length >= 2) {
    const parsed = parsePrefix(firstRowValues[0]);
    title = parsed.title;
    kind = parsed.kind;
    chartType = parsed.chartType;
    i = 1;
  }

  // 출처 라인 분리 (블록 마지막이 "출처" 단일 값)
  let source = '';
  let endIdx = block.length - 1;
  const lastRow = block[endIdx] || [];
  const lastText = getFirstNonEmpty(lastRow);
  const lastValues = lastRow.filter(v => v != null && String(v).trim() !== '');
  if (lastValues.length === 1 && /출처/.test(lastText)) {
    source = lastText;
    endIdx--;
  }

  // 헤더 + 데이터 필요
  if (endIdx < i + 1) return null;

  // colCount 결정:
  //   블록 내 전체 행(헤더 포함)의 실질 컬럼 수 최대값을 사용.
  //   헤더 한 줄만 기준으로 잡으면 뒤쪽에 넓은 행(세대구성 표 등)이 있을 때
  //   그 컬럼이 통째로 버려진다 — 엑셀 원본 데이터 유실 방지.
  let colCount = 0;
  for (let r = i; r <= endIdx; r++) {
    const trimmed = trimTrailingEmpty(block[r] || []);
    if (trimmed.length > colCount) colCount = trimmed.length;
  }
  if (colCount === 0) return null;

  // 헤더: 원본 행 그대로, colCount 에 맞춰 빈 셀 패딩
  const headerRaw = block[i] || [];
  const headers = [];
  for (let c = 0; c < colCount; c++) {
    headers.push(headerRaw[c] == null ? '' : String(headerRaw[c]));
  }

  const dataRows = [];
  for (let r = i + 1; r <= endIdx; r++) {
    const raw = block[r] || [];
    const padded = [];
    for (let c = 0; c < colCount; c++) {
      const cell = raw[c];
      padded.push(cell == null ? '' : cell);
    }
    dataRows.push(padded);
  }

  // merge 정보: 시트 절대좌표 → 표 상대좌표로 변환
  //   표의 row 0 = headers (block[i]), 시트 row = blockStartRow + i
  //   표의 row k = block[i + k], 시트 row = blockStartRow + i + k
  //   표 영역 밖으로 삐져나가는 merge 는 제외 또는 clamp
  const tableStartRow = blockStartRow + i;
  const tableEndRow   = blockStartRow + endIdx;
  const tableMerges = [];
  for (const m of (allMerges || [])) {
    if (!m) continue;
    const mRow = m.row, mCol = m.col, mRs = m.rowspan || 1, mCs = m.colspan || 1;
    // 표 영역과 교차하지 않으면 skip
    if (mRow > tableEndRow || mRow + mRs - 1 < tableStartRow) continue;
    if (mCol >= colCount) continue;
    // 표 내 상대 좌표
    const relRow = Math.max(0, mRow - tableStartRow);
    const bottomAbs = Math.min(mRow + mRs - 1, tableEndRow);
    const relRowEnd = bottomAbs - tableStartRow;
    const relCol = mCol;
    const relColEnd = Math.min(mCol + mCs - 1, colCount - 1);
    const rs = relRowEnd - relRow + 1;
    const cs = relColEnd - relCol + 1;
    if (rs < 1 || cs < 1) continue;
    if (rs === 1 && cs === 1) continue;  // 병합 아님
    tableMerges.push({ row: relRow, col: relCol, rowspan: rs, colspan: cs });
  }

  return { title, kind, chartType, headers, rows: dataRows, source, merges: tableMerges };
}

/**
 * 섹션 라벨 자동 생성
 *   "01_사업환경_년도별분양물량" → "Ⅰ. 사업환경 / 년도별분양물량"
 *   "사업환경"                    → "사업환경"
 *   "2.시장환경 - 거래량"         → "Ⅱ. 시장환경 / 거래량"
 */
function deriveSectionLabel(sheetName) {
  if (!sheetName) return '';
  const parts = String(sheetName).split(/[_\-.]/).map(s => s.trim()).filter(Boolean);
  if (parts.length === 0) return '';

  const roman = ['Ⅰ','Ⅱ','Ⅲ','Ⅳ','Ⅴ','Ⅵ','Ⅶ','Ⅷ','Ⅸ','Ⅹ'];
  let first = parts[0];
  let numPrefix = '';
  const m = first.match(/^(\d+)(.*)$/);
  if (m) {
    const n = parseInt(m[1]);
    numPrefix = (n >= 1 && n <= 10 ? roman[n - 1] : String(n)) + '. ';
    first = (m[2] || '').trim();
  }

  // 숫자 prefix 만 있는 토큰(예: "01") 이었다면 다음 토큰을 대목차로 승격
  let bigPart, restParts;
  if (first) {
    bigPart = numPrefix + first;
    restParts = parts.slice(1);
  } else if (parts.length >= 2) {
    bigPart = numPrefix + parts[1];
    restParts = parts.slice(2);
  } else {
    bigPart = numPrefix.trim() || parts[0];
    restParts = [];
  }

  const big = bigPart.trim();
  const rest = restParts.filter(Boolean);
  if (rest.length === 0) return big;
  return `${big} / ${rest.join(' / ')}`;
}

/**
 * 블록 수 → 추천 슬라이드 타입 + layout (T6 는 3블록만)
 *   1 → { type: 'bigchart' | 'data2col' }  — AI 가 차트/표 판단
 *   2 → { type: 'data2col' }                — T1 좌/우
 *   3 → { type: 'auto', layout: '3-row' (기본). AI 가 layout-rules.md 보고 선택 }
 *   4+ → { type: null, split: true }        — 슬라이드 분할 필요
 */
function deriveLayout(blockCount) {
  if (blockCount === 1) return { type: 'bigchart-or-data2col', layout: null };
  if (blockCount === 2) return { type: 'data2col', layout: null };
  if (blockCount === 3) return { type: 'auto', layout: '3-row' };
  return { type: null, split: true };
}

// ────────── 내부 유틸 ──────────
function isEmptyRow(row) {
  if (!row || row.length === 0) return true;
  return row.every(c => c == null || String(c).trim() === '');
}
function skipEmpty(rows, idx) {
  while (idx < rows.length && isEmptyRow(rows[idx])) idx++;
  return idx;
}
function getFirstNonEmpty(row) {
  for (const c of row || []) {
    if (c != null && String(c).trim() !== '') return String(c).trim();
  }
  return '';
}
function trimTrailingEmpty(row) {
  const arr = [...(row || [])];
  while (arr.length > 0 && (arr[arr.length - 1] == null || String(arr[arr.length - 1]).trim() === '')) arr.pop();
  return arr;
}

module.exports = { parseSheet, deriveLayout, deriveSectionLabel };
