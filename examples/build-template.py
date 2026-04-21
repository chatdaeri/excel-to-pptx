#!/usr/bin/env python3
"""
sample-template.xlsx 를 생성하는 스크립트.

사용:
  python3 examples/build-template.py

결과:
  examples/sample-template.xlsx  — 스킬 컨벤션을 셀 설명으로 담은 템플릿 1시트.
  사용자는 이 파일을 열어서 내용만 자기 데이터로 교체하면 됨.
"""

from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

OUT = Path(__file__).parent / "sample-template.xlsx"

# 시트 이름 = 번호_대목차_소목차 패턴. 첫 슬라이드 한 장에 해당.
SHEET_NAME = "01_시장환경_거래량"

# (row, col, value) 리스트로 정리. col 은 1-based.
CELLS = [
    # 1행 — 메인 메시지
    (1, 1, "여기에는 각 슬라이드에 들어갈 메인 메시지를 입력합니다 (한 줄 결론·시사점)"),
    # 2행 — 서브 메시지
    (2, 1, "여기엔 서브 메시지를 입력합니다 (범위·조건). 메인·서브 모두 비었을 경우에는 AI 가 자동 생성합니다"),

    # 3행 = 빈 행 (블록 구분)

    # ── 표 1 ──────────────────────────────────────────
    (4, 1, "[표] 이게 표의 제목이 됩니다. 반드시 [표] 표의 제목 이런 식으로 입력합니다. 각 표는 빈 행으로 구분됩니다"),
    # 헤더
    (5, 1, "구분"),
    (5, 2, "2022년"),
    (5, 3, "2023년"),
    (5, 4, "2024년"),
    # 데이터 15행 (6~20) — A열엔 설명 가이드, B~D열엔 더미 숫자로 실제 렌더 모양 확인 가능
    (6,  1, "첫 열은 좌측 라벨로 자동 인식되어 연파랑 배경이 적용됩니다"),
    (6,  2, 2150), (6,  3, 2480), (6,  4, 2890),
    (7,  1, "원본 엑셀 셀 값을 그대로 입력하면 됩니다 (숫자·텍스트 모두 가능)"),
    (7,  2, 1820), (7,  3, 1950), (7,  4, 2210),
    (8,  1, "빈 셀은 자동으로 공백 처리됩니다"),
    (8,  2, ""),    (8,  3, 1420), (8,  4, 1680),
    (9,  1, "엑셀 병합 셀도 지원 — 병합 정보가 그대로 반영됩니다"),
    (9,  2, 1320), (9,  3, 1580), (9,  4, 1780),
    (10, 1, "강조하고 싶은 행은 deck.cjs 의 emphRowIndices 로 지정 (빨간 외곽)"),
    (10, 2, 5410), (10, 3, 5820), (10, 4, 6200),
    (11, 1, "행 수는 표당 10~15행 권장 (너무 많으면 폰트가 줄어듦)"),
    (11, 2, 980),  (11, 3, 1120), (11, 4, 1340),
    (12, 1, "A열에는 카테고리·지역·연도 같은 라벨을 넣는 게 일반적"),
    (12, 2, 620),  (12, 3, 710),  (12, 4, 845),
    (13, 1, "16행 이상이면 폰트 size 를 줄이거나 슬라이드를 분할"),
    (13, 2, 420),  (13, 3, 510),  (13, 4, 670),
    (14, 1, "한 시트에 표가 1개면 풀폭 bigtable 로 렌더"),
    (14, 2, 1850), (14, 3, 2040), (14, 4, 2310),
    (15, 1, "한 시트에 표가 2개면 좌우 2단 data2col 로 렌더"),
    (15, 2, 1230), (15, 3, 1420), (15, 4, 1620),
    (16, 1, "한 시트에 표가 3개 이상이면 시트 분리 권장"),
    (16, 2, 780),  (16, 3, 890),  (16, 4, 1020),
    (17, 1, "숫자 포맷(천단위 콤마·%·원)은 엑셀에 적힌 그대로 전달됨"),
    (17, 2, 2450), (17, 3, 2680), (17, 4, 2950),
    (18, 1, "블록 사이 구분은 반드시 빈 행 1줄 이상으로"),
    (18, 2, 1510), (18, 3, 1720), (18, 4, 1940),
    (19, 1, "표 마지막 행 뒤에 빈 행이 없으면 다음 [표] 제목과 붙어버림"),
    (19, 2, 920),  (19, 3, 1050), (19, 4, 1240),
    (20, 1, "합계"),
    (20, 2, 20470), (20, 3, 22720), (20, 4, 26720),

    # 21행 = 빈 행 (블록 구분)

    # ── 표 2 ──────────────────────────────────────────
    (22, 1, "[표] 두 번째 표 제목. 한 시트에 표 1~3개 권장 (4개 이상은 시트 분리)"),
    # 헤더 (3열)
    (23, 1, "항목"),
    (23, 2, "값"),
    (23, 3, "비고"),
    # 데이터 12행 (24~35)
    (24, 1, "시트 이름 규약"),
    (24, 2, "01_대목차_소목차"),
    (24, 3, "번호·대목차·소목차 3토큰 권장 (Ⅰ. / Ⅱ. 로마숫자 자동 변환)"),
    (25, 1, "단일 토큰 시트명"),
    (25, 2, "거래량"),
    (25, 3, "대목차만 생성됨 — fallback 으로 첫 블록 제목을 소목차로 보강"),
    (26, 1, "블록 프리픽스"),
    (26, 2, "[표] 또는 [차트]"),
    (26, 3, "이 스킬은 [표] 만 지원 — [차트] 는 거부됨"),
    (27, 1, "헤더 행"),
    (27, 2, "표 제목 바로 아래"),
    (27, 3, "첫 행은 자동으로 회색 배경 + Bold 처리됨"),
    (28, 1, "출처 표기"),
    (28, 2, "출처: 한국감정원"),
    (28, 3, "블록 마지막에 '출처' 로 시작하는 단일 셀이 있으면 자동 추출"),
    (29, 1, "메인 메시지 길이"),
    (29, 2, "한 줄 ≤ 40자 권장"),
    (29, 3, "길어지면 자동 줄바꿈되지만 가독성 떨어짐"),
    (30, 1, "서브 메시지 용도"),
    (30, 2, "범위·조건·단위"),
    (30, 3, "예: '서울 25개구 / 2022~2024 / 월별 평균' 같은 부가정보"),
    (31, 1, "auto 메시지"),
    (31, 2, "1·2행 모두 빈칸일 때"),
    (31, 3, "Claude 가 표 내용을 읽고 deck.cjs 에 mainMessage/subhead 를 직접 기입"),
    (32, 1, "좌측 라벨 배경 제거"),
    (32, 2, "deck.cjs 에서 leftColLabel: false"),
    (32, 3, "첫 열을 일반 데이터 열로 취급 (연파랑 배경 없음)"),
    (33, 1, "열 너비 조정"),
    (33, 2, "deck.cjs 의 colW 로 지정"),
    (33, 3, "합계 9.16 inch 권장 (풀폭 기준)"),
    (34, 1, "폰트 크기 조정"),
    (34, 2, "deck.cjs 의 fontSize"),
    (34, 3, "기본 9pt — 데이터 많으면 8pt 까지 축소 가능"),
    (35, 1, "강조 행 지정"),
    (35, 2, "emphRowIndices: [3, 7]"),
    (35, 3, "0-based 데이터 행 인덱스 (헤더 제외). 해당 행은 빨간 외곽 + 빨간 Bold"),

    # 36행 = 빈 행

    # 37행 = 출처
    (37, 1, "출처: 블록 마지막에 '출처' 로 시작하는 단일 셀로 입력합니다 (예: 출처: 국토교통부 실거래가)"),
]


def main():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    for (r, c, v) in CELLS:
        ws.cell(row=r, column=c, value=v)

    # 1·2행 (메시지) 강조 — 옅은 노란 배경
    msg_fill = PatternFill('solid', fgColor='FFF8DC')
    for r in (1, 2):
        ws.cell(row=r, column=1).fill = msg_fill
        ws.cell(row=r, column=1).font = Font(bold=True, size=11)

    # [표] 제목 행 (4, 22) — 옅은 네이비 배경 + 흰 글씨
    title_fill = PatternFill('solid', fgColor='00205F')
    title_font = Font(bold=True, color='FFFFFF', size=11)
    for r in (4, 22):
        ws.cell(row=r, column=1).fill = title_fill
        ws.cell(row=r, column=1).font = title_font

    # 헤더 행 (5, 23) — 회색 배경 + Bold
    header_fill = PatternFill('solid', fgColor='E8E8E8')
    header_font = Font(bold=True)
    for r in (5, 23):
        for c in range(1, 5):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                cell.fill = header_fill
                cell.font = header_font

    # 출처 행 (37) — 회색 이탤릭
    src_font = Font(italic=True, color='666666')
    ws.cell(row=37, column=1).font = src_font

    # 열 너비
    ws.column_dimensions['A'].width = 55
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18

    # 행 높이 (설명 긴 행 여유 있게)
    for r in range(1, 38):
        ws.row_dimensions[r].height = 22
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[4].height = 28
    ws.row_dimensions[22].height = 28
    ws.row_dimensions[37].height = 28

    # 정렬
    left_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    center = Alignment(horizontal='center', vertical='center')
    for r in range(1, 38):
        a_cell = ws.cell(row=r, column=1)
        if a_cell.value is not None:
            a_cell.alignment = left_wrap
        for c in (2, 3, 4):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                cell.alignment = center

    wb.save(OUT)
    print(f"✅ 생성 완료: {OUT}")


if __name__ == '__main__':
    main()
