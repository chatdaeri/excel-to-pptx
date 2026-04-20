/**
 * T3: 챕터 디바이더 (검은 박스 없음, 36pt Bold)
 *
 * opts = {
 *   bgImagePath?: './assets/divider-1.jpg',
 *   bgColor?: '0A3D3A',
 *   chapterLabel: 'Part 1',
 *   chapterTitle: '사업환경\n검토'
 * }
 *
 * v (Validator, optional) — 디바이더는 풀블리드라 마진 검증 대상 아님.
 *                           beginSlide 만 호출해 슬라이드 idx 만 맞춰줌.
 */

const { C, FONT, FONT_BOLD, LAYOUT } = require('./_helpers.cjs');
const fs = require('fs');

function buildT3(pres, opts, v) {
  const slide = pres.addSlide();
  if (v) v.beginSlide();

  const bg = opts.bgColor || C.deepGreen;

  if (opts.bgImagePath && fs.existsSync(opts.bgImagePath)) {
    slide.background = { path: opts.bgImagePath };
    slide.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: LAYOUT.W, h: LAYOUT.H,
      fill: { color: '000000', transparency: 50 },
      line: { color: '000000', width: 0 }
    });
  } else {
    slide.background = { color: bg };
  }

  const baseX = 0.6;
  const baseY = 1.6;

  slide.addText(opts.chapterLabel || 'Part', {
    x: baseX, y: baseY, w: 3.5, h: 0.32,
    color: C.white, fontSize: 11, fontFace: FONT,
    charSpacing: 2,
    align: 'left', valign: 'middle'
  });

  slide.addShape(pres.ShapeType.line, {
    x: baseX, y: baseY + 0.4, w: 1.0, h: 0,
    line: { color: C.white, width: 1 }
  });

  const titleLines = (opts.chapterTitle || '').split('\n');
  const titleArr = titleLines.map((l, i) => ({
    text: l,
    options: { breakLine: i < titleLines.length - 1 }
  }));
  slide.addText(titleArr, {
    x: baseX, y: baseY + 0.6, w: 5.5, h: 2.0,
    color: C.white, fontSize: 29, fontFace: FONT_BOLD,
    charSpacing: 1, lineSpacingMultiple: 1.15,
    align: 'left', valign: 'top'
  });

  return slide;
}

module.exports = { buildT3 };
