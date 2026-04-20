/**
 * T5: 표지 / 목차 (2 mode)
 *
 * mode: 'cover' | 'toc'
 *
 * cover opts: { overline, title, subtitle, date, logoText, bgImagePath?, bgColor? }
 * toc   opts: { project, title, items: ['환경 점검', '가격 점검', ...] }
 *
 * v (Validator, optional). 표지·목차는 풀블리드라 element 등록 생략, beginSlide 만.
 *
 * (사업개요는 T1 `data2col` 로 직접 작성한다. overview 모드는 제거됨.)
 */

const { C, FONT, LAYOUT } = require('./_helpers.cjs');
const fs = require('fs');

function buildT5(pres, mode, opts, v) {
  if (mode === 'cover') return buildCover(pres, opts, v);
  if (mode === 'toc')   return buildToc(pres, opts, v);
  throw new Error(`Unknown T5 mode: ${mode}`);
}

function buildCover(pres, opts, v) {
  const slide = pres.addSlide();
  if (v) v.beginSlide();

  const bg = opts.bgColor || C.deepGreen;
  if (opts.bgImagePath && fs.existsSync(opts.bgImagePath)) {
    slide.background = { path: opts.bgImagePath };
    slide.addShape(pres.ShapeType.rect, {
      x: 0, y: 0, w: LAYOUT.W, h: LAYOUT.H,
      fill: { color: bg, transparency: 15 },
      line: { color: bg, width: 0 }
    });
  } else {
    slide.background = { color: bg };
  }

  if (opts.overline) {
    // 상/하 선 간격을 텍스트에 50% 더 가깝게
    slide.addShape(pres.ShapeType.line, { x: 3.5, y: 1.73, w: 3, h: 0, line: { color: C.white, width: 1 } });
    slide.addText(opts.overline, {
      x: 3, y: 1.78, w: 4, h: 0.22,
      color: C.white, fontSize: 11, fontFace: FONT,
      align: 'center', valign: 'middle'
    });
    slide.addShape(pres.ShapeType.line, { x: 3.5, y: 2.02, w: 3, h: 0, line: { color: C.white, width: 1 } });
  }

  if (opts.title) {
    slide.addText(opts.title, {
      x: 0.5, y: 2.3, w: 9, h: 0.5,
      color: C.white, fontSize: 21, fontFace: FONT,
      align: 'center', valign: 'bottom'
    });
  }

  if (opts.subtitle) {
    // title 과의 갭을 50% 축소 (2.85 → 2.70)
    slide.addText(opts.subtitle, {
      x: 0.5, y: 2.70, w: 9, h: 1.1,
      color: C.white, bold: true, fontSize: 45, fontFace: FONT,
      charSpacing: -2, align: 'center', valign: 'top'
    });
  }

  if (opts.date) {
    slide.addText(opts.date, {
      x: 0.5, y: 4.0, w: 9, h: 0.3,
      color: C.white, fontSize: 11, fontFace: FONT,
      align: 'center'
    });
  }

  if (opts.logoText) {
    slide.addText(opts.logoText, {
      x: 0.5, y: 5.1, w: 9, h: 0.35,
      color: C.white, bold: true, fontSize: 13, fontFace: FONT,
      charSpacing: 4, align: 'center'
    });
  }

  return slide;
}

function buildToc(pres, opts, v) {
  const slide = pres.addSlide();
  if (v) v.beginSlide();

  slide.background = { color: '0E1822' };

  if (opts.project) {
    slide.addText(opts.project, {
      x: 0.5, y: 1.7, w: 4.3, h: 0.3,
      color: C.midGray, fontSize: 10, fontFace: FONT
    });
  }
  slide.addText(opts.title || 'Contents', {
    x: 0.5, y: 2.0, w: 4.3, h: 0.9,
    color: C.white, fontSize: 40, fontFace: FONT,
    italic: false
  });
  slide.addShape(pres.ShapeType.line, {
    x: 0.5, y: 2.95, w: 2.5, h: 0,
    line: { color: C.white, width: 1 }
  });

  slide.addShape(pres.ShapeType.line, {
    x: 5.0, y: 0.7, w: 0, h: 4.2,
    line: { color: C.midGray, width: 0.5 }
  });

  const startY = 1.3;
  const itemH = 0.6;
  (opts.items || []).forEach((item, i) => {
    const y = startY + i * itemH;
    slide.addText(`0${i + 1}.`, {
      x: 5.5, y, w: 0.6, h: itemH,
      color: C.white, bold: true, fontSize: 18, fontFace: FONT,
      valign: 'middle'
    });
    slide.addText(item, {
      x: 6.2, y, w: 3.5, h: itemH,
      color: C.white, fontSize: 14, fontFace: FONT,
      valign: 'middle'
    });
  });

  return slide;
}

module.exports = { buildT5 };
