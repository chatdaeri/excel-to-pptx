#!/usr/bin/env node
/**
 * excel-to-pptx 런너 — deck.cjs 한 장으로 PPTX 자동 빌드 (표 전용 / 차트 미지원)
 *
 * 사용:
 *   node scripts/run.cjs <deck.cjs 경로>
 *
 * deck.cjs 형식:
 *   module.exports = {
 *     excel: './data.xlsx',         // bigtable/data2col 에 sheet 쓰면 필수
 *     out:   './deck.pptx',
 *     slides: [
 *       { type: 'cover',    title, subtitle, date, overline?, logoText?, bgImagePath?, bgColor? },
 *       { type: 'toc',      project?, title?, items: [...] },
 *       { type: 'divider',  chapter, title, bgImagePath?, bgColor? },
 *       { type: 'bigtable', sheet?: '...', sectionLabel, subhead, mainMessage,
 *                           tableHeader, table?, source?, emphRowIndices?, fontSize?, colW? },
 *       { type: 'data2col', sheet?: '...', sectionLabel, subhead, mainMessage,
 *                           leftHeader,  leftSlot:  { type: 'table', rows, ... },
 *                           rightHeader, rightSlot: { type: 'table', rows, ... }, source? }
 *     ]
 *   };
 *
 * 흐름:
 *   1) 엑셀 덤프 (있으면)
 *   2) slides[] 순회 → T5/T3/T7b/T1-table 디스패치 (Validator 자동 주입)
 *   3) PPTX 저장
 *   4) 표 셀 병합 패치
 *
 * 특징:
 *   ✗ check-sync 단계 없음 (브랜드 spec 동기화 없음)
 *   ✗ patch_chart_fonts 단계 없음 (차트 없음)
 *   ✗ bigchart / map / concept / auto 슬라이드 타입 거부
 *   ✗ data2col 의 chart slot 거부
 *   ✗ [차트] 블록 만나면 거부
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const SKILL_ROOT = path.resolve(__dirname, '..');

// pptxgenjs 경로 fallback (환경 무관)
let pptxgen;
const tryPaths = [
  'pptxgenjs',
  '/opt/homebrew/lib/node_modules/pptxgenjs',
  '/usr/local/lib/node_modules/pptxgenjs',
  '/usr/lib/node_modules/pptxgenjs'
];
for (const p of tryPaths) {
  try { pptxgen = require(p); break; } catch (_) { /* try next */ }
}
if (!pptxgen) {
  console.error('pptxgenjs 모듈을 찾을 수 없습니다. `npm install -g pptxgenjs` 또는 deck.cjs 폴더에서 `npm install pptxgenjs` 후 재시도하세요.');
  process.exit(1);
}

// ── 인자 ──
const deckArg = process.argv[2];
if (!deckArg) {
  console.error('Usage: node run.cjs <deck.cjs>');
  process.exit(1);
}
const deckAbs = path.resolve(deckArg);
if (!fs.existsSync(deckAbs)) {
  console.error(`deck 파일 없음: ${deckAbs}`);
  process.exit(1);
}

// ── deck.cjs 로드 ──
const deck = require(deckAbs);
const { excel, out, slides } = deck;
if (!out || !Array.isArray(slides)) {
  console.error('deck.cjs 에 out 과 slides 가 있어야 함');
  process.exit(1);
}

// ── 단계 1. 엑셀 덤프 (필요 시) ──
let excelData = {};
const needsExcel = slides.some(s =>
  (s.type === 'bigtable' && s.sheet) ||
  (s.type === 'data2col' && s.sheet)
);

if (needsExcel && !excel) {
  console.error('sheet 기반 슬라이드가 있는데 deck.cjs 에 excel 경로가 없음');
  process.exit(1);
}

if (excel) {
  const excelAbs = path.isAbsolute(excel) ? excel
                   : path.resolve(path.dirname(deckAbs), excel);
  if (!fs.existsSync(excelAbs)) {
    console.error(`엑셀 파일 없음: ${excelAbs}`);
    process.exit(1);
  }
  console.log(`[1/4] 엑셀 덤프: ${path.basename(excelAbs)}`);
  const py = [
    'import openpyxl, json, sys',
    'from openpyxl.utils import range_boundaries',
    `wb = openpyxl.load_workbook(r"${excelAbs}", data_only=True)`,
    'out = {}',
    'for name in wb.sheetnames:',
    '    ws = wb[name]',
    '    rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]',
    '    merges = []',
    '    for mr in ws.merged_cells.ranges:',
    '        mc, mR, xc, xR = range_boundaries(str(mr))',
    '        merges.append({"row": mR-1, "col": mc-1, "rowspan": xR-mR+1, "colspan": xc-mc+1})',
    '    out[name] = {"rows": rows, "merges": merges}',
    'sys.stdout.write(json.dumps(out, ensure_ascii=False, default=str))'
  ].join('\n');
  const jsonStr = execFileSync('python3', ['-c', py], {
    encoding: 'utf8', maxBuffer: 200 * 1024 * 1024
  });
  excelData = JSON.parse(jsonStr);
  console.log(`   → ${Object.keys(excelData).length}개 시트 로드`);
} else {
  console.log('[1/4] 엑셀 없음 (인라인 슬라이드만)');
}

// ── 빌더 + 헬퍼 로드 ──
const { parseSheet }    = require(path.join(SKILL_ROOT, 'builders/excel-sheet-parser.cjs'));
const { buildT1Table }  = require(path.join(SKILL_ROOT, 'builders/t1-data-table.cjs'));
const { buildT3 }       = require(path.join(SKILL_ROOT, 'builders/t3-divider.cjs'));
const { buildT5 }       = require(path.join(SKILL_ROOT, 'builders/t5-cover-toc.cjs'));
const { buildT7b }      = require(path.join(SKILL_ROOT, 'builders/t7b-big-table.cjs'));
const { Validator }     = require(path.join(SKILL_ROOT, 'builders/validate.cjs'));

const pres = new pptxgen();
pres.defineLayout({ name: 'BUNLITE16x9', width: 10, height: 5.625 });
pres.layout = 'BUNLITE16x9';
const v = new Validator();

function sheetRows(name) {
  const e = excelData[name];
  if (!e) return null;
  return Array.isArray(e) ? e : e.rows;
}
function sheetMerges(name) {
  const e = excelData[name];
  return (e && !Array.isArray(e) && Array.isArray(e.merges)) ? e.merges : [];
}

const pendingMerges = [];

// ── toc 자동 동기화용: deck 안의 divider 들에서 목차 항목 추출 ──
// (\n 은 공백으로 치환 — 디바이더는 멀티라인 허용, 목차는 한 줄이 자연스러움)
const dividerTitles = slides
  .filter(s => s.type === 'divider')
  .map(s => String(s.title || s.chapter || '').replace(/\s*\n\s*/g, ' ').trim())
  .filter(Boolean);

// sectionLabel fallback 용: 슬라이드 순회 중 마지막 divider 의 chapter 추적
let currentDividerChapter = '';

// ── 단계 2. 슬라이드 빌드 ──
console.log('\n[2/4] PPTX 빌드');

slides.forEach((s, i) => {
  const n = i + 1;
  try {
    if (s.type === 'divider') currentDividerChapter = String(s.chapter || '').trim();
    dispatch(s, n);
    console.log(`   · 슬라이드 ${n}/${slides.length}  ${s.type}  OK`);
  } catch (e) {
    console.error(`\n슬라이드 ${n} (${s.type}) 빌드 실패:`, e.message);
    throw e;
  }
});

/**
 * sectionLabel fallback ladder:
 *   1) 시트명 파싱 결과가 `대 / 소` 형태 → 그대로
 *   2) 대목차만 있음 (`/` 없음)          → `대 / firstBlockTitle`
 *   3) 완전히 비어있음                   → `divider chapter / (firstBlockTitle || sheetName)`
 * deck.cjs 에 명시된 sectionLabel 이 있으면 이 함수는 호출되지 않음 (항상 그게 우선).
 */
function fallbackSectionLabel(specSectionLabel, firstBlockTitle, sheetName) {
  const label = String(specSectionLabel || '').trim();
  const hasSlash = label.includes('/');
  if (label && hasSlash) return label;
  if (label && !hasSlash) {
    return firstBlockTitle ? `${label} / ${firstBlockTitle}` : label;
  }
  // 완전히 비어있음
  const big = currentDividerChapter || '';
  const sub = firstBlockTitle || sheetName || '';
  if (big && sub) return `${big} / ${sub}`;
  return big || sub || '';
}

function warnIfAutoMessageNeeded(spec, slideIdx, ctx) {
  if (!spec.needsAutoMessage) return;
  const blockDesc = spec.blocks
    .map((b, i) => `[${i + 1}] ${b.kind === 'chart' ? '차트' : '표'} "${b.title || '(제목 없음)'}" (${b.rows.length}행 × ${b.headers.length}열)`)
    .join(', ');
  console.warn(
    `   ⚠ 슬라이드 ${slideIdx} (${ctx}, sheet="${spec.sheetName}"): ` +
    `엑셀 1·2행이 모두 비어있음. mainMessage/subhead 가 빈 상태로 렌더됩니다.\n` +
    `     → Claude 는 이 시트의 블록을 읽고 deck.cjs 에 mainMessage/subhead 를 직접 기입해야 합니다.\n` +
    `     블록: ${blockDesc || '(없음)'}`
  );
}

function dispatch(s, slideIdx) {
  switch (s.type) {
    case 'cover':    return buildT5(pres, 'cover', s, v);
    case 'toc': {
      // 목차 항목 자동 동기화:
      //   - items 미지정/빈 배열 → divider title 들로 자동 채움
      //   - items 명시 + divider 와 개수 불일치 → 경고만 (사용자 의도 존중, 명시값 사용)
      const userItems = Array.isArray(s.items) ? s.items.filter(Boolean) : [];
      let items;
      if (userItems.length === 0) {
        items = dividerTitles;
        if (dividerTitles.length === 0) {
          console.warn('   ⚠ toc.items 비어있고 divider 도 없음 — 목차 항목 0개로 빌드');
        } else {
          console.log(`   · toc 항목 ${dividerTitles.length}개 자동 동기화 (divider 기반)`);
        }
      } else {
        items = userItems;
        if (dividerTitles.length > 0 && userItems.length !== dividerTitles.length) {
          console.warn(
            `   ⚠ toc.items(${userItems.length}개) 와 divider 개수(${dividerTitles.length}개) 불일치. ` +
            `자동 동기화하려면 deck.cjs 에서 toc 의 items 줄을 지우세요.`
          );
        }
      }
      return buildT5(pres, 'toc', { ...s, items }, v);
    }
    case 'divider':  return buildT3(pres, {
      chapterLabel: s.chapter, chapterTitle: s.title,
      bgImagePath: s.bgImagePath, bgColor: s.bgColor
    }, v);
    case 'bigtable': return buildBigtable(s, slideIdx);
    case 'data2col': return buildData2col(s, slideIdx);

    // excel-to-pptx 미지원 타입 — 친절한 안내
    case 'bigchart':
    case 'map':
    case 'concept':
    case 'auto':
      throw new Error(
        `excel-to-pptx 는 '${s.type}' 슬라이드를 지원하지 않습니다 ` +
        `(차트/지도/컨셉박스/3블록 자동 미지원). ` +
        `사용 가능 타입: cover, toc, divider, bigtable, data2col`
      );

    default: throw new Error(`unknown slide type: ${s.type}`);
  }
}

function rejectChartBlocks(spec, sheetName, ctx) {
  const chartBlocks = spec.blocks.filter(b => b.kind === 'chart');
  if (chartBlocks.length === 0) return;
  const titles = chartBlocks.map(b => `[차트] ${b.title}`).join(', ');
  throw new Error(
    `${ctx}: sheet "${sheetName}" 에 [차트] 블록 ${chartBlocks.length}개 발견 (${titles}). ` +
    `excel-to-pptx 는 차트 미지원입니다. 엑셀에서 [차트] 블록을 제거해주세요.`
  );
}

function buildBigtable(s, slideIdx) {
  let table        = s.table;
  let source       = s.source;
  let tableHeader  = s.tableHeader;
  let sectionLabel = s.sectionLabel;
  let subhead      = s.subhead;
  let mainMessage  = s.mainMessage;
  let blockMerges  = [];

  if (s.sheet) {
    const raw = sheetRows(s.sheet);
    if (!raw) throw new Error(`bigtable: sheet not found "${s.sheet}"`);
    const spec = parseSheet(s.sheet, raw, sheetMerges(s.sheet));
    rejectChartBlocks(spec, s.sheet, 'bigtable');

    if (spec.blocks.length === 0) {
      throw new Error(`bigtable: sheet "${s.sheet}" 에서 [표] 블록 추출 실패`);
    }
    const b = spec.blocks[0];

    table = table || {
      rows: [b.headers, ...b.rows],
      leftColLabel: s.leftColLabel !== false,
      emphRowIndices: s.emphRowIndices || [],
      fontSize: s.fontSize,
      colW: s.colW,
      merges: b.merges || []
    };
    tableHeader  = tableHeader  || b.title;
    source       = source       || b.source;
    sectionLabel = sectionLabel || fallbackSectionLabel(spec.sectionLabel, spec.firstBlockTitle, s.sheet);
    mainMessage  = mainMessage  || spec.mainMessage;
    subhead      = subhead      || spec.subMessage;
    blockMerges  = b.merges || [];

    if (!s.mainMessage && !s.subhead) warnIfAutoMessageNeeded(spec, slideIdx, 'bigtable');
  }

  buildT7b(pres, {
    sectionLabel, subhead, mainMessage,
    tableHeader: tableHeader || '',
    table: table || { rows: [] },
    source: source || ''
  }, v);

  if (blockMerges.length) {
    pendingMerges.push({ slideIndex: slideIdx, tableIndex: 0, merges: blockMerges });
  } else if (table && Array.isArray(table.merges) && table.merges.length) {
    pendingMerges.push({ slideIndex: slideIdx, tableIndex: 0, merges: table.merges });
  } else if (s.merges && s.merges.length) {
    pendingMerges.push({ slideIndex: slideIdx, tableIndex: 0, merges: s.merges });
  }
}

function buildData2col(s, slideIdx) {
  let leftSlot     = s.leftSlot;
  let rightSlot    = s.rightSlot;
  let leftHeader   = s.leftHeader;
  let rightHeader  = s.rightHeader;
  let sectionLabel = s.sectionLabel;
  let subhead      = s.subhead;
  let mainMessage  = s.mainMessage;
  let source       = s.source;
  let b0 = null, b1 = null;

  if (s.sheet) {
    const raw = sheetRows(s.sheet);
    if (!raw) throw new Error(`data2col: sheet not found "${s.sheet}"`);
    const spec = parseSheet(s.sheet, raw, sheetMerges(s.sheet));
    rejectChartBlocks(spec, s.sheet, 'data2col');

    if (spec.blocks.length !== 2) {
      const hint = spec.blocks.length === 1 ? "1블록 → bigtable 사용"
                 : spec.blocks.length === 3 ? "3블록은 excel-to-pptx 미지원 — 시트를 둘로 분리하세요"
                 :                            "블록 수 확인 후 분할 (excel-to-pptx 는 1·2블록만 지원)";
      throw new Error(
        `data2col 은 2블록 전용. sheet "${s.sheet}" 블록 ${spec.blocks.length}개 → ${hint}`
      );
    }
    [b0, b1] = spec.blocks;
    leftSlot     = leftSlot  || blockToTableSlot(b0);
    rightSlot    = rightSlot || blockToTableSlot(b1);
    leftHeader   = leftHeader   || b0.title;
    rightHeader  = rightHeader  || b1.title;
    sectionLabel = sectionLabel || fallbackSectionLabel(spec.sectionLabel, spec.firstBlockTitle, s.sheet);
    subhead      = subhead      || spec.subMessage;
    mainMessage  = mainMessage  || spec.mainMessage;
    source       = source       || b0.source || b1.source;

    if (!s.mainMessage && !s.subhead) warnIfAutoMessageNeeded(spec, slideIdx, 'data2col');
  }

  buildT1Table(pres, {
    sectionLabel, subhead, mainMessage,
    leftHeader,  leftSlot,
    rightHeader, rightSlot,
    source: source || ''
  }, v);

  // 셀 병합 메타 수집 — T1 은 left 먼저, right 나중 순으로 addTable.
  // sheet 기반: b0.merges / b1.merges. inline: leftSlot.merges / rightSlot.merges.
  let tidx = 0;
  for (const [block, slot] of [[b0, leftSlot], [b1, rightSlot]]) {
    if (slot) {
      const m = (block && block.merges && block.merges.length) ? block.merges
              : (Array.isArray(slot.merges) && slot.merges.length) ? slot.merges
              : null;
      if (m) pendingMerges.push({ slideIndex: slideIdx, tableIndex: tidx, merges: m });
      tidx++;
    }
  }
}

function blockToTableSlot(block) {
  return {
    type: 'table',
    rows: [block.headers, ...block.rows],
    leftColLabel: true,
    merges: block.merges || []
  };
}

// Stage 1 검증
if (!v.print()) {
  console.error('\n검증 실패 → 중단');
  process.exit(1);
}

// ── 단계 3. 저장 ──
const outAbs = path.isAbsolute(out) ? out
               : path.resolve(path.dirname(deckAbs), out);

(async () => {
  console.log(`\n[3/4] 파일 저장 → ${outAbs}`);
  await pres.writeFile({ fileName: outAbs });

  // ── 단계 4. 셀 병합 패치 ──
  if (pendingMerges.length > 0) {
    console.log(`\n[4/4] 표 셀 병합 패치 (${pendingMerges.length}개 표)`);
    const mergesJsonPath = outAbs + '.merges.json';
    fs.writeFileSync(mergesJsonPath, JSON.stringify(pendingMerges, null, 2), 'utf8');
    try {
      execFileSync('python3',
        [path.join(SKILL_ROOT, 'builders/patch_table_merges.py'), outAbs, mergesJsonPath],
        { stdio: 'inherit' });
    } finally {
      try { fs.unlinkSync(mergesJsonPath); } catch (_) {}
    }
  } else {
    console.log('\n[4/4] 셀 병합 없음 — skip');
  }

  console.log(`\n✅ 완료: ${outAbs}`);
})().catch(e => { console.error(e); process.exit(1); });
