"""廃棄物コネクト 案件量計算ツール - Excelテンプレート生成スクリプト。

再生成する場合: python3 generate_template.py
出力: waste_calculator_template.xlsx
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F3D24")
HEADER_FONT = Font(name=FONT_NAME, color="FFFFFF", bold=True, size=10)
INPUT_FONT = Font(name=FONT_NAME, color="0000FF", size=10)
FORMULA_FONT = Font(name=FONT_NAME, color="000000", size=10)
NOTE_FONT = Font(name=FONT_NAME, italic=True, color="6B6F65", size=9)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14, color="1F3D24")
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, color="6B6F65")
ASSUMPTION_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="D9DCD3")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TYPE_LIST = ["可燃ごみ", "生ごみ", "不燃ごみ", "瓶", "缶", "ペットボトル", "段ボール", "その他"]

wb = openpyxl.Workbook()

# ---------------------------------------------------------------- 使い方
ws_help = wb.active
ws_help.title = "使い方"
ws_help.sheet_view.showGridLines = False
ws_help.column_dimensions["A"].width = 2
ws_help.column_dimensions["B"].width = 100

r = 2
ws_help.cell(r, 2, "廃棄物コネクト 案件量計算ツール").font = TITLE_FONT
r += 1
ws_help.cell(r, 2, "ヒアリング内容から月間の袋数を自動計算するテンプレートです。").font = SUBTITLE_FONT
r += 2
lines = [
    ("① 入力のしかた", True),
    ("青字のセルのみ入力してください（黒字は自動計算の数式です）。", False),
    ("・案件名：案件ごとに分かる名前を入力（複数案件を1シートに並べてOK）", False),
    ("・種別：可燃ごみ／生ごみ／不燃ごみ／瓶／缶／ペットボトル／段ボール／その他 からドロップダウンで選択", False),
    ("・頻度(週◯回)：収集の頻度。例:「1日1〜2袋、週5回」→数量1〜2・頻度5、「週に1〜2袋」→数量1〜2・頻度1", False),
    ("・数量min／数量max：「1〜2袋」のように幅がある場合は両方入力。「2袋」のように1つだけの場合は数量maxを空欄にすればOK（数量minの値だけで計算されます）。", False),
    ("", False),
    ("② 前提", True),
    ("1ヶ月あたりの週数はC1セルで設定（既定値 4.35週 ＝ 365日 ÷ 12ヶ月 ÷ 7日）。", False),
    ("数量が明記されていない項目（例:「不燃物 週2日」のみで数量記載なし）は 1回1袋 と仮定しています。実数が分かり次第、数量欄を修正してください。", False),
]
for text, bold in lines:
    c = ws_help.cell(r, 2, text)
    c.font = Font(name=FONT_NAME, bold=True, size=11, color="1F3D24") if bold else Font(name=FONT_NAME, size=10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1


def add_type_validation(ws, col_letter, first_row, last_row):
    dv = DataValidation(type="list", formula1='"{}"'.format(",".join(TYPE_LIST)), allow_blank=True, showDropDown=False)
    dv.showErrorMessage = False
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{first_row}:{col_letter}{last_row}")


# ---------------------------------------------------------------- かんたん入力
ws_s = wb.create_sheet("かんたん入力")
ws_s.sheet_view.showGridLines = False

ws_s.cell(1, 1, "1ヶ月あたりの週数").font = Font(name=FONT_NAME, bold=True, size=10)
w = ws_s.cell(1, 3, 4.35)
w.font = INPUT_FONT
w.fill = ASSUMPTION_FILL
ws_s.cell(1, 4, "※365日÷12ヶ月÷7日で算出した平均週数。必要に応じて変更してください。").font = NOTE_FONT

s_headers = ["案件名", "種別", "数量min(1回あたり)", "数量max(1回あたり)", "頻度(週◯回)", "月間量(平均)"]
s_header_row = 3
for i, h in enumerate(s_headers, start=1):
    c = ws_s.cell(s_header_row, i, h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    c.border = BORDER
for i, wid in enumerate([26, 12, 16, 16, 12, 14], start=1):
    ws_s.column_dimensions[get_column_letter(i)].width = wid
ws_s.row_dimensions[s_header_row].height = 30

s_samples = [
    # site, type, qmin, qmax, freq
    ["新宿区の新規案件（laidback cafeの例）", "可燃ごみ", 1, 2, 7],
    ["新宿区の新規案件（laidback cafeの例）", "生ごみ", 1, 2, 7],
    ["新宿区の新規案件（laidback cafeの例）", "瓶", 1, 2, 1],
    ["新宿区の新規案件（laidback cafeの例）", "缶", 1, 2, 1],
    ["新宿区の新規案件（laidback cafeの例）", "ペットボトル", 1, 2, 1],
    ["新宿区の新規案件（laidback cafeの例）", "段ボール", 5, 10, 1],
    ["中野区の現行案件の例", "可燃ごみ", 1, 2, 5],
    ["中野区の現行案件の例", "不燃ごみ", 1, 1, 2],
]

s_first_row = s_header_row + 1
s_last_row = s_first_row + 39

for offset, row_data in enumerate(s_samples):
    r = s_first_row + offset
    site, typ, qmin, qmax, freq = row_data
    ws_s.cell(r, 1, site).font = INPUT_FONT
    ws_s.cell(r, 2, typ).font = INPUT_FONT
    ws_s.cell(r, 3, qmin).font = INPUT_FONT
    ws_s.cell(r, 4, qmax).font = INPUT_FONT
    ws_s.cell(r, 5, freq).font = INPUT_FONT

for r in range(s_first_row, s_last_row + 1):
    # F: 月間量(平均) = 数量(1つだけならその値、min〜maxならその平均) * 頻度(週) * 週数
    ws_s.cell(r, 6, f'=IF(C{r}="","",IF(D{r}="",C{r},(C{r}+D{r})/2)*E{r}*$C$1)')
    for col in range(1, 7):
        cell = ws_s.cell(r, col)
        cell.border = BORDER
        if col == 6:
            cell.font = FORMULA_FONT
            cell.number_format = "#,##0.0"

add_type_validation(ws_s, "B", s_first_row, s_last_row)
ws_s.freeze_panes = f"A{s_first_row}"

wb.save("waste_calculator_template.xlsx")
print("saved waste_calculator_template.xlsx")
