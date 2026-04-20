/**
 * excel-to-pptx 공통 헬퍼 (차트 미지원 — 표 전용)
 *
 *   - 컬러/폰트/좌표 상수
 *   - 헤더(대목차·소목차·부제·메인메시지) 렌더
 *   - 표 보더 매니저 + buildTableRows (셀 병합 인지)
 *
 *   ✗ chartOpts / AUTO_LAYOUTS 미포함 (차트·T6 미지원)
 *
 * 모든 렌더 함수는 마지막 인자로 Validator v 를 옵셔널 수신.
 * v 가 있으면 element 를 자동 등록해 Stage 1 검증에 참여.
 */

const C = {
  navy:       '00205F',
  navyDark:   '001544',
  red:        'C0392B',
  blue:       '7DA1C4',
  blueLight:  'EAEFF8',
  blueDark:   '2C5F8D',
  gray:       '9CA3AF',
  grayHeader: 'E8E8E8',
  white:      'FFFFFF',
  black:      '000000',
  text:       '000000',
  midGray:    '666666',
  border:     'D1D1D1',
  borderLight:'E5E5E5',
  lightBg:    'F4F4F4',
  deepGreen:  '0A3D3A',
  yellow:     'FFD700',
  cyan:       '1FA8C9'
};

// ★ Pretendard Static 패밀리
const FONT        = 'Pretendard';
const FONT_LIGHT  = 'Pretendard Light';
const FONT_MEDIUM = 'Pretendard Medium';
const FONT_SEMI   = 'Pretendard SemiBold';
const FONT_BOLD   = 'Pretendard Bold';

const LAYOUT = {
  W: 10, H: 5.625,
  MARGIN_X: 0.42, MARGIN_TOP: 0.31, MARGIN_BOTTOM: 0.31,
  BIG_LABEL:        { x: 0.42, y: 0.28, w: 9.16, h: 0.20 },
  SUB_LABEL_BOX:    { x: 0.42, y: 0.50, w: 3.6,  h: 0.28 },
  SUBHEAD:          { x: 0.42, y: 0.88, w: 9.16, h: 0.22 },
  MAIN_MSG:         { x: 0.42, y: 1.12, w: 9.16, h: 0.50 },
  COL_LEFT_HEADER:  { x: 0.42, y: 1.78, w: 4.58, h: 0.22 },
  COL_RIGHT_HEADER: { x: 5.20, y: 1.78, w: 4.38, h: 0.22 },
  SLOT_LEFT:        { x: 0.42, y: 2.06, w: 4.58, h: 3.00 },
  SLOT_RIGHT:       { x: 5.20, y: 2.06, w: 4.38, h: 3.00 },
  // T7b 풀폭 단일 슬롯
  FULL_COL_HEADER:  { x: 0.42, y: 1.78, w: 9.16, h: 0.22 },
  SLOT_FULL:        { x: 0.42, y: 2.06, w: 9.16, h: 3.00 },
  SOURCE:           { x: 0.42, y: 5.24, w: 9.16, h: 0.18 }
};

// ============================================================
// 헤더 / 텍스트 헬퍼
// ============================================================

function addAnalysisHeader(slide, pres, { sectionLabel, subhead, mainMessage }, v) {
  const parts = String(sectionLabel || '').split('/').map(s => s.trim()).filter(Boolean);
  const bigLabel = parts[0] || '';
  const subLabel = parts.slice(1).join(' / ') || '';

  // 1. 대목차 — Pretendard Light 10pt
  if (bigLabel) {
    slide.addText(bigLabel, {
      ...LAYOUT.BIG_LABEL,
      color: C.black, fontSize: 10, fontFace: FONT_LIGHT,
      align: 'left', valign: 'middle', charSpacing: -1
    });
    if (v) v.element({ type: 'text', ...LAYOUT.BIG_LABEL, fontSize: 10, label: 'bigLabel' });
  }

  // 2. 소목차 — 네이비 박스 안 흰 Bold 10pt
  if (subLabel) {
    const textLen = [...subLabel].length;
    const estW = Math.min(9.16, Math.max(2.0, textLen * 0.14 + 0.4));
    slide.addShape(pres.ShapeType.rect, {
      x: LAYOUT.SUB_LABEL_BOX.x, y: LAYOUT.SUB_LABEL_BOX.y,
      w: estW, h: LAYOUT.SUB_LABEL_BOX.h,
      fill: { color: C.navy }, line: { color: C.navy, width: 0 }
    });
    slide.addText(subLabel, {
      x: LAYOUT.SUB_LABEL_BOX.x, y: LAYOUT.SUB_LABEL_BOX.y,
      w: estW, h: LAYOUT.SUB_LABEL_BOX.h,
      color: C.white, fontSize: 10, fontFace: FONT_BOLD,
      align: 'left', valign: 'middle', inset: 0.14
    });
    if (v) v.element({
      type: 'text', x: LAYOUT.SUB_LABEL_BOX.x, y: LAYOUT.SUB_LABEL_BOX.y,
      w: estW, h: LAYOUT.SUB_LABEL_BOX.h, fontSize: 10, label: 'subLabel'
    });
  }

  // 3. 부제 — Pretendard Medium 11pt
  if (subhead) {
    slide.addText(subhead, {
      ...LAYOUT.SUBHEAD,
      color: C.black, fontSize: 11, fontFace: FONT_MEDIUM,
      align: 'left', valign: 'middle'
    });
    if (v) v.element({ type: 'text', ...LAYOUT.SUBHEAD, fontSize: 11, label: 'subhead' });
  }

  // 4. 메인 메시지 — Pretendard Medium 18pt
  const msg = Array.isArray(mainMessage) ? mainMessage : [{ text: mainMessage }];
  slide.addText(msg, {
    ...LAYOUT.MAIN_MSG,
    color: C.black, fontSize: 18, fontFace: FONT_MEDIUM,
    align: 'left', valign: 'top', charSpacing: -1
  });
  if (v) v.element({ type: 'text', ...LAYOUT.MAIN_MSG, fontSize: 18, label: 'mainMessage' });
}

/**
 * 컬럼 헤더 — 네이비 직사각형 막대 + 텍스트
 * side: 'left' | 'right' | 'full'
 */
function addColumnHeader(slide, pres, side, text, v) {
  const coord = side === 'left'  ? LAYOUT.COL_LEFT_HEADER
              : side === 'right' ? LAYOUT.COL_RIGHT_HEADER
              :                    LAYOUT.FULL_COL_HEADER;
  const barW = 0.06;
  const barInset = 0.02;
  slide.addShape(pres.ShapeType.rect, {
    x: coord.x, y: coord.y + barInset, w: barW, h: coord.h - barInset * 2,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText(text, {
    x: coord.x + barW + 0.08, y: coord.y,
    w: coord.w - barW - 0.08, h: coord.h,
    color: C.black, fontSize: 12, bold: false, fontFace: FONT,
    align: 'left', valign: 'middle'
  });
  if (v) v.element({ type: 'text', ...coord, fontSize: 12, label: `${side}Header` });
}

function addSource(slide, text, v) {
  slide.addText(text || '', {
    ...LAYOUT.SOURCE,
    color: C.midGray, fontSize: 8, bold: false, fontFace: FONT,
    align: 'left', valign: 'middle'
  });
  if (v && text) v.element({ type: 'text', ...LAYOUT.SOURCE, fontSize: 8, label: 'source' });
}

// ============================================================
// 표 보더 시스템 — 정밀 위치 기반
// ============================================================

const BORDER_THICK    = { type: 'solid', color: C.black, pt: 2 };
const BORDER_HEADER_B = { type: 'solid', color: C.black, pt: 1 };
const BORDER_THIN     = { type: 'solid', color: C.borderLight, pt: 0.5 };
const BORDER_EMPH     = { type: 'solid', color: C.red, pt: 2 };
const BORDER_NONE     = { type: 'none' };

function decideBorders(i, j, totalRows, totalCols, emphSet) {
  const isHeaderRow = i === 0;
  const isFirstRow  = i === 0;
  const isLastRow   = i === totalRows - 1;
  const isFirstCol  = j === 0;
  const isLastCol   = j === totalCols - 1;

  const isEmph   = emphSet.has(i);
  const prevEmph = emphSet.has(i - 1);
  const nextEmph = emphSet.has(i + 1);

  let top;
  if (isEmph || prevEmph)  top = BORDER_EMPH;
  else if (isFirstRow)     top = BORDER_THICK;
  else                     top = BORDER_THIN;

  let bottom;
  if (isEmph || nextEmph)  bottom = BORDER_EMPH;
  else if (isHeaderRow)    bottom = BORDER_HEADER_B;
  else if (isLastRow)      bottom = BORDER_THICK;
  else                     bottom = BORDER_THIN;

  let left;
  if (isFirstCol && isEmph) left = BORDER_EMPH;
  else                      left = BORDER_NONE;

  let right;
  if (isLastCol && isEmph) right = BORDER_EMPH;
  else if (isLastCol)      right = BORDER_NONE;
  else                     right = BORDER_THIN;

  return [top, right, bottom, left];
}

function buildTableRows(headerCells, dataRows, options = {}) {
  const totalRows = dataRows.length + 1;
  const totalCols = headerCells.length;
  const emphSet = new Set((options.emphRowIndices || []).map(k => k + 1));
  const fontSize = options.fontSize || 9;

  const startMap = new Map();
  const hiddenMap = new Map();
  for (const m of (options.merges || [])) {
    const r = m.row | 0, c = m.col | 0;
    const rs = Math.max(1, (m.rowspan | 0) || 1);
    const cs = Math.max(1, (m.colspan | 0) || 1);
    if (rs === 1 && cs === 1) continue;
    startMap.set(`${r},${c}`, { row: r, col: c, rowspan: rs, colspan: cs });
    for (let dr = 0; dr < rs; dr++) {
      for (let dc = 0; dc < cs; dc++) {
        if (dr === 0 && dc === 0) continue;
        hiddenMap.set(`${r + dr},${c + dc}`, true);
      }
    }
  }

  function bordersFor(i, j) {
    const hidden = hiddenMap.has(`${i},${j}`);
    if (hidden) return [BORDER_NONE, BORDER_NONE, BORDER_NONE, BORDER_NONE];
    const merge = startMap.get(`${i},${j}`);
    if (!merge) return decideBorders(i, j, totalRows, totalCols, emphSet);
    const topBorders    = decideBorders(i, j, totalRows, totalCols, emphSet);
    const leftBorders   = decideBorders(i, j, totalRows, totalCols, emphSet);
    const bottomBorders = decideBorders(i + merge.rowspan - 1, j, totalRows, totalCols, emphSet);
    const rightBorders  = decideBorders(i, j + merge.colspan - 1, totalRows, totalCols, emphSet);
    return [topBorders[0], rightBorders[1], bottomBorders[2], leftBorders[3]];
  }

  const allRows = [headerCells, ...dataRows];

  return allRows.map((row, i) => row.map((raw, j) => {
    const isHeader = (i === 0);
    const text = (raw && typeof raw === 'object' && 'text' in raw)
      ? String(raw.text)
      : (raw == null ? '' : String(raw));
    const userOpts = (raw && typeof raw === 'object' && raw.options) ? raw.options : {};

    const isLeftLabel = options.leftColAsLabel && (j === 0) && !isHeader;
    const isEmph = emphSet.has(i);

    const fill = isHeader ? { color: C.grayHeader }
                : isLeftLabel ? { color: C.blueLight }
                : undefined;
    const color = isEmph ? C.red : C.text;
    const align = userOpts.align
                || (options.colAlign && options.colAlign[j])
                || (isHeader ? 'center' : (options.align || 'left'));

    const ff = (isHeader || isEmph || isLeftLabel) ? FONT_BOLD : FONT;

    return {
      text,
      options: {
        fontFace: userOpts.fontFace || ff,
        fontSize: userOpts.fontSize || fontSize,
        color: userOpts.color || color,
        fill: userOpts.fill || fill,
        align,
        valign: 'middle',
        border: bordersFor(i, j)
      }
    };
  }));
}

function tableOpts(coord, overrides = {}) {
  return {
    ...coord,
    fontFace: FONT, fontSize: 9, color: C.text,
    border: BORDER_NONE,
    valign: 'middle',
    ...overrides
  };
}

function resolveTableRows(slot) {
  if (!slot || !Array.isArray(slot.rows) || slot.rows.length === 0) return [];

  const firstRow = slot.rows[0];
  if (Array.isArray(firstRow) && firstRow.some(c => c && typeof c === 'object' && 'options' in c)) {
    return slot.rows;
  }

  let headers, dataRows;
  if (Array.isArray(slot.headers)) {
    headers  = slot.headers;
    dataRows = slot.rows;
  } else {
    headers  = firstRow;
    dataRows = slot.rows.slice(1);
  }

  const buildOpts = {
    leftColAsLabel: slot.leftColLabel !== false,
    emphRowIndices: slot.emphRowIndices || slot.emph || [],
    fontSize:       slot.fontSize || 9,
    align:          slot.align,
    colAlign:       slot.colAlign,
    merges:         slot.merges || []
  };
  return buildTableRows(headers, dataRows, buildOpts);
}

module.exports = {
  C, FONT, FONT_LIGHT, FONT_MEDIUM, FONT_SEMI, FONT_BOLD, LAYOUT,
  addAnalysisHeader, addColumnHeader, addSource,
  tableOpts, buildTableRows, resolveTableRows,
  BORDER_THICK, BORDER_THIN, BORDER_EMPH, BORDER_NONE
};
