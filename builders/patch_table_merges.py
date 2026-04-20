#!/usr/bin/env python3
"""
Stage 4 — PPTX 표 셀 병합 후처리.

pptxgenjs 는 addTable() 에 셀 병합 API 가 없음. 이 스크립트는 엑셀 원본의
merged_cells 정보를 받아 PPTX 내부 slide XML 의 <a:tbl> 에 gridSpan / rowSpan /
hMerge / vMerge 속성을 직접 주입한다.

사용:
  python3 patch_table_merges.py <out.pptx> <merges.json>

merges.json 형식:
  [
    { "slideIndex": 4, "tableIndex": 0,
      "merges": [
        {"row": 0, "col": 2, "rowspan": 1, "colspan": 2},
        ...
      ]
    },
    ...
  ]
  - slideIndex: 1-based (slide1.xml 의 숫자)
  - tableIndex: 해당 슬라이드 내 <a:tbl> 순서 0-based
  - row/col:    표 관점 (헤더 row 0) 0-based
"""

import json
import os
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
ET.register_namespace('a', NS_A)
ET.register_namespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

TAG_TBL = f'{{{NS_A}}}tbl'
TAG_TR  = f'{{{NS_A}}}tr'
TAG_TC  = f'{{{NS_A}}}tc'


def apply_merges_to_tbl(tbl, merges):
    rows = tbl.findall(TAG_TR)
    total_rows = len(rows)
    for m in merges:
        r, c = m['row'], m['col']
        rs = max(1, m.get('rowspan', 1))
        cs = max(1, m.get('colspan', 1))
        if rs == 1 and cs == 1:
            continue
        if r >= total_rows:
            continue
        start_cells = rows[r].findall(TAG_TC)
        if c >= len(start_cells):
            continue
        start_cell = start_cells[c]

        # 시작 셀 속성
        if cs > 1:
            start_cell.set('gridSpan', str(cs))
        if rs > 1:
            start_cell.set('rowSpan', str(rs))

        # 같은 행의 오른쪽 병합 셀들
        for dc in range(1, cs):
            if c + dc < len(start_cells):
                cell = start_cells[c + dc]
                cell.set('hMerge', '1')

        # 아래쪽 행의 병합 셀들
        for dr in range(1, rs):
            if r + dr >= total_rows:
                continue
            lower_cells = rows[r + dr].findall(TAG_TC)
            if c >= len(lower_cells):
                continue
            for dc in range(0, cs):
                if c + dc >= len(lower_cells):
                    continue
                cell = lower_cells[c + dc]
                cell.set('vMerge', '1')
                if dc > 0:
                    cell.set('hMerge', '1')


def patch_pptx(pptx_path, merges_meta):
    if not merges_meta:
        return 0

    by_slide = {}
    for entry in merges_meta:
        by_slide.setdefault(int(entry['slideIndex']), []).append(entry)

    tmp_dir = tempfile.mkdtemp(prefix='bunyang_merge_')
    try:
        with zipfile.ZipFile(pptx_path, 'r') as z:
            z.extractall(tmp_dir)

        touched = 0
        for slide_idx, entries in by_slide.items():
            slide_xml = os.path.join(tmp_dir, 'ppt', 'slides', f'slide{slide_idx}.xml')
            if not os.path.exists(slide_xml):
                continue
            tree = ET.parse(slide_xml)
            root = tree.getroot()
            tbls = root.iter(TAG_TBL)
            tbl_list = list(tbls)
            for entry in entries:
                ti = int(entry.get('tableIndex', 0))
                if ti >= len(tbl_list):
                    continue
                apply_merges_to_tbl(tbl_list[ti], entry.get('merges', []))
                touched += 1
            # xml.etree 에는 standalone 인자가 없어 수동으로 선언 기록
            body = ET.tostring(root, encoding='UTF-8').decode('utf-8')
            with open(slide_xml, 'w', encoding='utf-8') as f:
                f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
                f.write(body)

        # 재압축 (원본 덮어쓰기)
        out_tmp = pptx_path + '.tmp'
        with zipfile.ZipFile(out_tmp, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, _, files in os.walk(tmp_dir):
                for f in files:
                    full = os.path.join(root_dir, f)
                    rel = os.path.relpath(full, tmp_dir)
                    z.write(full, rel)
        shutil.move(out_tmp, pptx_path)
        return touched
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    if len(sys.argv) != 3:
        print('Usage: patch_table_merges.py <out.pptx> <merges.json>', file=sys.stderr)
        sys.exit(1)
    pptx_path = sys.argv[1]
    merges_json = sys.argv[2]
    if not os.path.exists(pptx_path):
        print(f'pptx not found: {pptx_path}', file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(merges_json):
        print(f'merges json not found: {merges_json}', file=sys.stderr)
        sys.exit(1)
    with open(merges_json, 'r', encoding='utf-8') as f:
        meta = json.load(f)
    n = patch_pptx(pptx_path, meta)
    print(f'✓ 표 셀 병합 패치 ({n}개 표)')


if __name__ == '__main__':
    main()
