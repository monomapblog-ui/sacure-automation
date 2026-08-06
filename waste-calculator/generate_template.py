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
ws_help.cell(r, 2, "ヒアリング内容から月間の袋数・容量・概算料金を自動計算するテンプレートです。").font = SUBTITLE_FONT
r += 2
lines = [
    ("① 入力シート", True),
    ("青字のセルのみ入力してください（黒字は自動計算の数式です）。", False),
    ("・ラボ名：案件（ラボ）ごとに同じ表記で統一してください（比較表で名前を突き合わせます）", False),
    ("・種別：可燃ごみ／生ごみ／不燃ごみ／瓶／缶／ペットボトル／段ボール／その他 からセルのドロップダウンで選択できます", False),
    ("・頻度パターン：", False),
    ("　「日」＝1回（1日）あたりの数量 × 週の収集日数　（例：1〜2袋/日、週5回）", False),
    ("　「週」＝週あたりの数量をそのまま入力　（例：週に1〜2袋、瓶・缶・ペットボトルなど）", False),
    ("・週の収集日数：頻度パターンが「日」の場合のみ入力してください", False),
    ("・袋サイズ(L)：容量を計算したい場合のみ入力（段ボール等「枚」単位の項目は空欄でOK）", False),
    ("・単価(円)：概算料金を出したい場合のみ入力（未入力の項目は料金計算から除外されます）", False),
    ("", False),
    ("② 比較表シート", True),
    ("入力シートに登録したラボ名を1行ずつ入力すると、種別ごとの月間量・合計容量・概算料金が自動集計されます。", False),
    ("ラボ名は入力シートの表記と完全に一致させてください。", False),
    ("", False),
    ("③ 前提", True),
    ("1ヶ月あたりの週数は「入力」シート C1 セルで設定しています（既定値 4.35 週 ＝ 365日 ÷ 12ヶ月 ÷ 7日）。", False),
    ("数量が明記されていない項目（例：「不燃物 週2日」のみで数量記載なし）は 1回1袋 と仮定しています。実数が分かり次第、入力シートの数量を修正してください。", False),
]
for text, bold in lines:
    c = ws_help.cell(r, 2, text)
    c.font = Font(name=FONT_NAME, bold=True, size=11, color="1F3D24") if bold else Font(name=FONT_NAME, size=10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1

# ---------------------------------------------------------------- 入力
ws = wb.create_sheet("入力")
ws.sheet_view.showGridLines = False

ws.cell(1, 1, "1ヶ月あたりの週数").font = Font(name=FONT_NAME, bold=True, size=10)
weeks_cell = ws.cell(1, 3, 4.35)
weeks_cell.font = INPUT_FONT
weeks_cell.fill = ASSUMPTION_FILL
ws.cell(1, 4, "※365日÷12ヶ月÷7日で算出した平均週数。必要に応じて変更してください。").font = NOTE_FONT

headers = ["ラボ名", "種別", "単位(袋/枚)", "頻度パターン(日/週)", "数量min(1回あたり)",
           "数量max(1回あたり)", "週の収集日数(日パターンのみ)", "袋サイズ(L)", "単価(円)",
           "月間量(平均)", "月間容量(L)", "月間概算料金(円)"]
header_row = 3
for i, h in enumerate(headers, start=1):
    c = ws.cell(header_row, i, h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    c.border = BORDER

widths = [16, 12, 10, 16, 14, 14, 18, 10, 10, 12, 12, 14]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[header_row].height = 30

sample_rows = [
    # site, type, unit, freq, qmin, qmax, days, bagsize, price
    ["サンプルラボA（新宿区新規ラボ想定）", "可燃ごみ", "袋", "日", 1, 2, 7, 45, ""],
    ["サンプルラボA（新宿区新規ラボ想定）", "生ごみ", "袋", "日", 1, 2, 7, 45, ""],
    ["サンプルラボA（新宿区新規ラボ想定）", "瓶", "袋", "週", 1, 2, "", 45, ""],
    ["サンプルラボA（新宿区新規ラボ想定）", "缶", "袋", "週", 1, 2, "", 45, ""],
    ["サンプルラボA（新宿区新規ラボ想定）", "ペットボトル", "袋", "週", 1, 2, "", 45, ""],
    ["サンプルラボA（新宿区新規ラボ想定）", "段ボール", "枚", "週", 5, 10, "", "", ""],
    ["サンプルラボB（中野区の現ラボ想定）", "可燃ごみ", "袋", "日", 1, 2, 5, 45, 120],
    ["サンプルラボB（中野区の現ラボ想定）", "不燃ごみ", "袋", "日", 1, 1, 2, 45, 120],
]

first_data_row = header_row + 1
last_data_row = first_data_row + 39  # room for extra manual rows below samples

for offset, row_data in enumerate(sample_rows):
    r = first_data_row + offset
    site, typ, unit, freq, qmin, qmax, days, bagsize, price = row_data
    ws.cell(r, 1, site).font = INPUT_FONT
    ws.cell(r, 2, typ).font = INPUT_FONT
    ws.cell(r, 3, unit).font = INPUT_FONT
    ws.cell(r, 4, freq).font = INPUT_FONT
    ws.cell(r, 5, qmin).font = INPUT_FONT
    ws.cell(r, 6, qmax).font = INPUT_FONT
    ws.cell(r, 7, days).font = INPUT_FONT
    ws.cell(r, 8, bagsize).font = INPUT_FONT
    ws.cell(r, 9, price).font = INPUT_FONT

for r in range(first_data_row, last_data_row + 1):
    # J: 月間量(平均) = IF(freq="日", (min+max)/2*days, (min+max)/2) * weeks
    ws.cell(r, 10, f'=IF(E{r}="","",IF(D{r}="日",(E{r}+F{r})/2*G{r},(E{r}+F{r})/2)*$C$1)')
    # K: 月間容量(L) = 月間量 * 袋サイズ（袋サイズ未入力なら空欄）
    ws.cell(r, 11, f'=IF(OR(J{r}="",H{r}=""),"",J{r}*H{r})')
    # L: 月間概算料金 = 月間量 * 単価（単価未入力なら空欄）
    ws.cell(r, 12, f'=IF(OR(J{r}="",I{r}=""),"",J{r}*I{r})')
    for col in range(1, 13):
        cell = ws.cell(r, col)
        cell.border = BORDER
        if col >= 10:
            cell.font = FORMULA_FONT
            cell.number_format = "#,##0.0"
    ws.cell(r, 12).number_format = "#,##0"

# data validations
dv_type = DataValidation(type="list", formula1='"{}"'.format(",".join(TYPE_LIST)), allow_blank=True, showDropDown=False)
dv_type.showErrorMessage = False
ws.add_data_validation(dv_type)
dv_type.add(f"B{first_data_row}:B{last_data_row}")

dv_unit = DataValidation(type="list", formula1='"袋,枚"', allow_blank=True, showDropDown=False)
dv_unit.showErrorMessage = False
ws.add_data_validation(dv_unit)
dv_unit.add(f"C{first_data_row}:C{last_data_row}")

dv_freq = DataValidation(type="list", formula1='"日,週"', allow_blank=True, showDropDown=False)
dv_freq.showErrorMessage = False
ws.add_data_validation(dv_freq)
dv_freq.add(f"D{first_data_row}:D{last_data_row}")

ws.freeze_panes = f"A{first_data_row}"

# ---------------------------------------------------------------- 比較表
ws_cmp = wb.create_sheet("比較表")
ws_cmp.sheet_view.showGridLines = False

ws_cmp.cell(1, 1, "収集ラボ間の比較表").font = TITLE_FONT
ws_cmp.cell(2, 1, "「入力」シートに登録したラボ名を、下表のA列に入力シートと同じ表記で入力してください。").font = SUBTITLE_FONT

cmp_headers = ["ラボ名"] + TYPE_LIST[:-1] + ["合計個数", "合計容量(L)", "概算料金合計(円)"]
cmp_header_row = 4
for i, h in enumerate(cmp_headers, start=1):
    c = ws_cmp.cell(cmp_header_row, i, h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    c.border = BORDER
ws_cmp.row_dimensions[cmp_header_row].height = 28

for i in range(1, len(cmp_headers) + 1):
    ws_cmp.column_dimensions[get_column_letter(i)].width = 16
ws_cmp.column_dimensions["A"].width = 30

cmp_first_row = cmp_header_row + 1
cmp_last_row = cmp_first_row + 14
type_cols = TYPE_LIST[:-1]  # exclude "その他" catch-all label mismatch issues; 比較表 focuses on the 7 standard types

sample_sites = ["サンプルラボA（新宿区新規ラボ想定）", "サンプルラボB（中野区の現ラボ想定）"]
for offset, site in enumerate(sample_sites):
    r = cmp_first_row + offset
    ws_cmp.cell(r, 1, site).font = INPUT_FONT

for r in range(cmp_first_row, cmp_last_row + 1):
    for ci, typ in enumerate(type_cols, start=2):
        col_letter = get_column_letter(ci)
        formula = (f'=IF($A{r}="","",SUMIFS(入力!$J${first_data_row}:$J${last_data_row},'
                   f'入力!$A${first_data_row}:$A${last_data_row},$A{r},'
                   f'入力!$B${first_data_row}:$B${last_data_row},{col_letter}${cmp_header_row}))')
        cell = ws_cmp.cell(r, ci, formula)
        cell.font = FORMULA_FONT
        cell.number_format = "#,##0.0"
    total_col = len(type_cols) + 2
    vol_col = total_col + 1
    cost_col = vol_col + 1
    ws_cmp.cell(r, total_col,
                f'=IF($A{r}="","",SUMIF(入力!$A${first_data_row}:$A${last_data_row},$A{r},入力!$J${first_data_row}:$J${last_data_row}))')
    ws_cmp.cell(r, vol_col,
                f'=IF($A{r}="","",SUMIF(入力!$A${first_data_row}:$A${last_data_row},$A{r},入力!$K${first_data_row}:$K${last_data_row}))')
    ws_cmp.cell(r, cost_col,
                f'=IF($A{r}="","",SUMIF(入力!$A${first_data_row}:$A${last_data_row},$A{r},入力!$L${first_data_row}:$L${last_data_row}))')
    ws_cmp.cell(r, total_col).font = FORMULA_FONT
    ws_cmp.cell(r, vol_col).font = FORMULA_FONT
    ws_cmp.cell(r, cost_col).font = FORMULA_FONT
    for c in (total_col, vol_col, cost_col):
        ws_cmp.cell(r, c).number_format = "#,##0"
    for col in range(1, cost_col + 1):
        ws_cmp.cell(r, col).border = BORDER

ws_cmp.cell(cmp_last_row + 2, 1,
            "※「その他」で登録した独自種別は比較表の集計対象外です（入力シートの月間量・容量・料金列を直接ご確認ください）。").font = NOTE_FONT

wb.save("waste_calculator_template.xlsx")
print("saved waste_calculator_template.xlsx")
