# -*- coding: utf-8 -*-
"""
================================================================================
Верига супермаркети „Бакалия“ — Генератор на Опростена бланка за доставчици (2027)
================================================================================
Бланката съдържа само полетата, които доставчикът реално знае и попълва:
- Баркод (по желание)
- Наименование на артикула (по доставчик)
- Разфасовка / Мярка
- Редовна доставна цена без ДДС
- Промоционална доставна цена без ДДС (ЗКП)
- Търговска отстъпка ТО % (автоматична формула)
- Препоръчителна промо продажна цена с ДДС (по желание)
- Бележка / Специфика
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import datetime

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Цветове
NAVY_HEADER = "1B365D"
BLUE_ACCENT = "2563EB"
LIGHT_BLUE = "EFF6FF"
ICE_BLUE = "DBEAFE"
INPUT_FILL = "FFFBEB"       # Мек крем за клетките за въвеждане от доставчика
CALC_FILL = "F8FAFC"        # Светло сиво за формули
BORDER_GRAY = "CBD5E1"

font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
font_subtitle = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
font_sec_hdr = Font(name="Calibri", size=11, bold=True, color="1E293B")
font_tbl_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_regular = Font(name="Calibri", size=10, color="0F172A")
font_bold = Font(name="Calibri", size=10, bold=True, color="0F172A")
font_italic = Font(name="Calibri", size=9, italic=True, color="64748B")

fill_navy = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
fill_blue_hdr = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
fill_ice = PatternFill(start_color=ICE_BLUE, end_color=ICE_BLUE, fill_type="solid")
fill_light_blue = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
fill_input = PatternFill(start_color=INPUT_FILL, end_color=INPUT_FILL, fill_type="solid")
fill_calc = PatternFill(start_color=CALC_FILL, end_color=CALC_FILL, fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color=BORDER_GRAY),
    right=Side(style="thin", color=BORDER_GRAY),
    top=Side(style="thin", color=BORDER_GRAY),
    bottom=Side(style="thin", color=BORDER_GRAY)
)
header_border = Border(
    left=Side(style="thin", color="475569"),
    right=Side(style="thin", color="475569"),
    top=Side(style="medium", color=NAVY_HEADER),
    bottom=Side(style="medium", color=NAVY_HEADER)
)
double_bottom_border = Border(
    left=Side(style="thin", color=BORDER_GRAY),
    right=Side(style="thin", color=BORDER_GRAY),
    top=Side(style="thin", color=BORDER_GRAY),
    bottom=Side(style="double", color="1E293B")
)

# ------------------------------------------------------------------------------
# 1. ЛИСТ: Бланка_Предложения (Опростената форма за доставчици)
# ------------------------------------------------------------------------------
ws_form = wb.create_sheet(title="Бланка_Предложения")
ws_form.views.sheetView[0].showGridLines = True
ws_form.freeze_panes = "A13" # Замразяване на заглавния ред на таблицата

# Заглавен банер
ws_form.merge_cells("A1:I1")
ws_form["A1"] = "ВЕРИГА СУПЕРМАРКЕТИ „БАКАЛИЯ“ — ОФЕРТА ЗА ПРОМОЦИОНАЛНА БРОШУРА"
ws_form["A1"].font = font_title
ws_form["A1"].fill = fill_navy
ws_form["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_form.row_dimensions[1].height = 28

ws_form.merge_cells("A2:I2")
ws_form["A2"] = "Бланка за офериране на промоционални артикули (Кампания 2027 г.)"
ws_form["A2"].font = font_subtitle
ws_form["A2"].fill = fill_navy
ws_form["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws_form.row_dimensions[2].height = 18

ws_form.row_dimensions[3].height = 6

# Секция 1: Кампания и Секция 2: Доставчик (Редове 4 до 7)
ws_form.merge_cells("A4:D4")
ws_form["A4"] = "📅 1. ПАРАМЕТРИ НА БРОШУРАТА"
ws_form["A4"].font = font_sec_hdr
ws_form["A4"].fill = fill_ice
ws_form["A4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_form["A4"].border = thin_border

ws_form.merge_cells("E4:I4")
ws_form["E4"] = "🏢 2. ДАННИ ЗА ДОСТАВЧИКА"
ws_form["E4"].font = font_sec_hdr
ws_form["E4"].fill = fill_ice
ws_form["E4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_form["E4"].border = thin_border
ws_form.row_dimensions[4].height = 22

header_fields = [
    ("Целева Брошура №:", "Брошура № 01", "Фирма / Контрагент:", "Име на фирмата"),
    ("Период в магазините:", '=VLOOKUP(B5, График_2027!$A$3:$F$27, 4, FALSE)', "Бранд / Направление:", "Марка / Производител"),
    ("Промо доставна цена важи от:", '=VLOOKUP(B5, График_2027!$A$3:$F$27, 6, FALSE)', "Лице за контакт:", "Име и Фамилия"),
    ("Краен срок за оферти:", '=VLOOKUP(B5, График_2027!$A$3:$F$27, 7, FALSE) & " (до 17:00 ч.)"', "Телефон и Имейл:", "тел. / имейл")
]

for idx, (lbl_c, val_c, lbl_p, val_p) in enumerate(header_fields, 5):
    ws_form.row_dimensions[idx].height = 20
    # Кампания
    c_lbl = ws_form.cell(row=idx, column=1, value=lbl_c)
    c_lbl.font = font_bold
    c_lbl.alignment = Alignment(horizontal="right", vertical="center")
    c_lbl.border = thin_border
    
    ws_form.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=4)
    c_val = ws_form.cell(row=idx, column=2, value=val_c)
    c_val.font = Font(name="Calibri", size=11, bold=True, color="1B365D") if idx == 5 else font_bold
    c_val.alignment = Alignment(horizontal="center", vertical="center")
    c_val.border = thin_border
    c_val.fill = fill_input if idx == 5 else fill_calc

    # Доставчик
    ws_form.merge_cells(start_row=idx, start_column=5, end_row=idx, end_column=6)
    p_lbl = ws_form.cell(row=idx, column=5, value=lbl_p)
    p_lbl.font = font_bold
    p_lbl.alignment = Alignment(horizontal="right", vertical="center")
    p_lbl.border = thin_border

    ws_form.merge_cells(start_row=idx, start_column=7, end_row=idx, end_column=9)
    p_val = ws_form.cell(row=idx, column=7, value=val_p)
    p_val.font = font_regular
    p_val.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    p_val.border = thin_border
    p_val.fill = fill_input

ws_form.row_dimensions[9].height = 6

# Указателен банер
ws_form.merge_cells("A10:I10")
note_cell = ws_form["A10"]
note_cell.value = "⚠️ ВАЖНО: 1) Всички цени в бланката (доставни и продажни) се попълват С ВКЛЮЧЕН ДДС! 2) Промоционалната доставна цена (ЗКП) задължително важи 7 дни преди старта на брошурата! 3) Попълвайте само жълтите полета."
note_cell.font = Font(name="Calibri", size=9, bold=True, color="92400E")
note_cell.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
note_cell.alignment = Alignment(horizontal="center", vertical="center")
note_cell.border = thin_border
ws_form.row_dimensions[10].height = 20

ws_form.row_dimensions[11].height = 6

# Заглавия на таблицата с артикули (Ред 12)
table_cols = [
    ("№", 4, "center"),
    ("Баркод\n(по желание)", 16, "center"),
    ("Наименование на артикула\n(по доставчик)", 38, "left"),
    ("Разфасовка\n/ М.Ед.", 12, "center"),
    ("Редовна ДЦ\nс ДДС (€)", 15, "right"),
    ("ПРОМО ДЦ\nс ДДС (€)", 15, "right"),
    ("Търговска\nотстъпка ТО %", 14, "right"),
    ("Препоръчителна\nПромо ПЦ с ДДС (€)", 16, "right"),
    ("Бележка / Специфика\n(видове, акция, подарък)", 28, "left")
]

ws_form.row_dimensions[12].height = 32
for col_idx, (th, width, align) in enumerate(table_cols, 1):
    c = ws_form.cell(row=12, column=col_idx, value=th)
    c.font = font_tbl_hdr
    c.fill = fill_navy
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = header_border
    col_letter = get_column_letter(col_idx)
    ws_form.column_dimensions[col_letter].width = width

# Примерни редове и празни редове за попълване (13 до 42 -> 30 реда)
sample_rows = [
    ("380000000001", "Мляно месо 100% свинско", "кг.", 3.68, 3.24, 3.99, "Корица / ТОП цена"),
    ("380000000002", "Деликатес Житница хамбургски колбас", "кг.", 5.39, 5.12, 5.99, "Гастро тарелка"),
    ("380000000003", "Барила макарони избрани видове", "500гр.", 2.64, 2.22, 2.69, "Спагети №5, Пенне, Фусили"),
    ("380000000004", "Здравец кисело мляко 2%", "500гр.", 0.86, 0.73, 0.89, "БДС стандарт"),
    ("380000000005", "Airwick фреш матик пълнител", "250мл.", 6.43, 5.29, 5.99, "Асортимент аромати"),
]

for row_idx in range(13, 43):
    ws_form.row_dimensions[row_idx].height = 20
    i = row_idx - 12
    is_sample = (i <= len(sample_rows))
    s_data = sample_rows[i-1] if is_sample else None

    # Col A: №
    cA = ws_form.cell(row=row_idx, column=1, value=i)
    cA.font = font_bold
    cA.alignment = Alignment(horizontal="center", vertical="center")
    cA.fill = fill_calc
    cA.border = thin_border

    # Col B: Баркод (по желание)
    cB = ws_form.cell(row=row_idx, column=2, value=s_data[0] if is_sample else None)
    cB.font = font_regular
    cB.alignment = Alignment(horizontal="center", vertical="center")
    cB.fill = fill_input
    cB.border = thin_border
    cB.number_format = "@"

    # Col C: Име по доставчик (Задължително)
    cC = ws_form.cell(row=row_idx, column=3, value=s_data[1] if is_sample else None)
    cC.font = font_bold if is_sample else font_regular
    cC.alignment = Alignment(horizontal="left", vertical="center")
    cC.fill = fill_input
    cC.border = thin_border

    # Col D: Разфасовка / Мярка
    cD = ws_form.cell(row=row_idx, column=4, value=s_data[2] if is_sample else None)
    cD.font = font_regular
    cD.alignment = Alignment(horizontal="center", vertical="center")
    cD.fill = fill_input
    cD.border = thin_border

    # Col E: Редовна ДЦ без ДДС
    cE = ws_form.cell(row=row_idx, column=5, value=s_data[3] if is_sample else None)
    cE.font = font_regular
    cE.alignment = Alignment(horizontal="right", vertical="center")
    cE.fill = fill_input
    cE.border = thin_border
    cE.number_format = "#,##0.00"

    # Col F: ПРОМО ДЦ без ДДС
    cF = ws_form.cell(row=row_idx, column=6, value=s_data[4] if is_sample else None)
    cF.font = font_bold
    cF.alignment = Alignment(horizontal="right", vertical="center")
    cF.fill = fill_input
    cF.border = thin_border
    cF.number_format = "#,##0.00"

    # Col G: ТО % (Автоматична формула)
    cG = ws_form.cell(row=row_idx, column=7)
    cG.value = f"=IF(AND(ISNUMBER(E{row_idx}),ISNUMBER(F{row_idx}),E{row_idx}>0),(E{row_idx}-F{row_idx})/E{row_idx},\"\")"
    cG.font = font_bold
    cG.alignment = Alignment(horizontal="right", vertical="center")
    cG.fill = fill_calc
    cG.border = thin_border
    cG.number_format = "0.0%"

    # Col H: Препоръчителна Промо ПЦ (по желание)
    cH = ws_form.cell(row=row_idx, column=8, value=s_data[5] if is_sample else None)
    cH.font = font_regular
    cH.alignment = Alignment(horizontal="right", vertical="center")
    cH.fill = fill_input
    cH.border = thin_border
    cH.number_format = "#,##0.00"

    # Col I: Бележка / Специфика
    cI = ws_form.cell(row=row_idx, column=9, value=s_data[6] if is_sample else None)
    cI.font = font_italic
    cI.alignment = Alignment(horizontal="left", vertical="center")
    cI.fill = fill_input
    cI.border = thin_border

# Обобщаващ ред (Ред 43)
ws_form.row_dimensions[43].height = 24
ws_form.merge_cells("A43:D43")
t_lbl = ws_form["A43"]
t_lbl.value = "СРЕДНО / ОБЩО ЗА БЛАНКАТА:"
t_lbl.font = font_tbl_hdr
t_lbl.fill = fill_navy
t_lbl.alignment = Alignment(horizontal="right", vertical="center", indent=1)
t_lbl.border = double_bottom_border

ws_form.cell(row=43, column=5, value="").fill = fill_navy
ws_form.cell(row=43, column=6, value="").fill = fill_navy

t_to = ws_form.cell(row=43, column=7, value="=IFERROR(AVERAGE(G13:G42),0)")
t_to.font = font_tbl_hdr
t_to.fill = fill_navy
t_to.alignment = Alignment(horizontal="right", vertical="center")
t_to.border = double_bottom_border
t_to.number_format = "0.0%"

ws_form.cell(row=43, column=8, value="").fill = fill_navy
ws_form.cell(row=43, column=9, value="").fill = fill_navy

# Валидация на данни
dv_brochure = DataValidation(type="list", formula1="=График_2027!$A$3:$A$27", allow_blank=False)
ws_form.add_data_validation(dv_brochure)
dv_brochure.add("B5")

dv_units = DataValidation(type="list", formula1='"бр.,кг.,100гр.,200гр.,250гр.,400гр.,500гр.,700гр.,800гр.,1л.,1.5л.,2л.,оп.,пакет"', allow_blank=True)
ws_form.add_data_validation(dv_units)
dv_units.add("D13:D42")

# ------------------------------------------------------------------------------
# 2. ЛИСТ: График_2027 (Справочен лист за кампаниите)
# ------------------------------------------------------------------------------
ws_sched = wb.create_sheet(title="График_2027")
ws_sched.views.sheetView[0].showGridLines = True

schedule_data = [
    (1, '2027-01-09', '2027-01-17', 9),
    (2, '2027-01-23', '2027-01-31', 9),
    (3, '2027-02-06', '2027-02-14', 9),
    (4, '2027-02-20', '2027-02-28', 9),
    (5, '2027-03-06', '2027-03-14', 9),
    (6, '2027-03-20', '2027-03-28', 9),
    (7, '2027-04-03', '2027-04-11', 9),
    (8, '2027-04-17', '2027-04-25', 9),
    (9, '2027-04-29', '2027-05-09', 11),
    (10, '2027-05-22', '2027-05-30', 9),
    (11, '2027-06-05', '2027-06-13', 9),
    (12, '2027-06-19', '2027-06-27', 9),
    (13, '2027-07-03', '2027-07-11', 9),
    (14, '2027-07-17', '2027-07-25', 9),
    (15, '2027-08-07', '2027-08-15', 9),
    (16, '2027-08-21', '2027-08-29', 9),
    (17, '2027-09-04', '2027-09-12', 9),
    (18, '2027-09-18', '2027-09-26', 9),
    (19, '2027-10-02', '2027-10-10', 9),
    (20, '2027-10-23', '2027-10-31', 9),
    (21, '2027-11-06', '2027-11-14', 9),
    (22, '2027-11-20', '2027-11-28', 9),
    (23, '2027-12-04', '2027-12-12', 9),
    (24, '2027-12-18', '2027-12-26', 9),
    (25, '2027-12-27', '2028-01-02', 7),
]

sched_headers = [
    "Код на Брошурата", "Старт (магазини)", "Край (магазини)",
    "Период в супермаркетите", "Дни", "Период намалена доставна цена (ЗКП, X-7)",
    "Краен срок за оферти", "Краен срок за предпечат (X-10)"
]

ws_sched.merge_cells("A1:H1")
ws_sched["A1"] = "ОФИЦИАЛЕН ГРАФИК ЗА ПРОМОЦИОНАЛНИ БРОШУРИ — 2027 ГОДИНА"
ws_sched["A1"].font = font_title
ws_sched["A1"].fill = fill_navy
ws_sched["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_sched.row_dimensions[1].height = 30

for col_idx, h in enumerate(sched_headers, 1):
    cell = ws_sched.cell(row=2, column=col_idx, value=h)
    cell.font = font_tbl_hdr
    cell.fill = fill_blue_hdr
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border
ws_sched.row_dimensions[2].height = 24

for row_idx, (num, s_str, e_str, days) in enumerate(schedule_data, 3):
    s_dt = datetime.date.fromisoformat(s_str)
    e_dt = datetime.date.fromisoformat(e_str)
    zkp_start = s_dt - datetime.timedelta(days=7)
    sub_deadline = s_dt - datetime.timedelta(days=12)
    prep_deadline = s_dt - datetime.timedelta(days=10)

    b_code = f"Брошура № {num:02d}"
    b_period = f"{s_dt.strftime('%d.%m.%Y')} – {e_dt.strftime('%d.%m.%Y')}"
    zkp_period = f"{zkp_start.strftime('%d.%m.%Y')} – {e_dt.strftime('%d.%m.%Y')}"

    r_vals = [
        b_code, s_dt.strftime('%Y-%m-%d'), e_dt.strftime('%Y-%m-%d'),
        b_period, days, zkp_period,
        sub_deadline.strftime('%d.%m.%Y'), prep_deadline.strftime('%d.%m.%Y')
    ]
    for c_idx, val in enumerate(r_vals, 1):
        cell = ws_sched.cell(row=row_idx, column=c_idx, value=val)
        cell.font = font_regular
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if c_idx == 1:
            cell.font = font_bold
            cell.fill = fill_light_blue
    ws_sched.row_dimensions[row_idx].height = 18

for idx, w in enumerate([18, 15, 15, 26, 8, 30, 20, 20], 1):
    ws_sched.column_dimensions[get_column_letter(idx)].width = w

# ------------------------------------------------------------------------------
# 3. ЛИСТ: Указания_за_доставчици
# ------------------------------------------------------------------------------
ws_info = wb.create_sheet(title="Указания_за_доставчици")
ws_info.views.sheetView[0].showGridLines = True

ws_info.merge_cells("A1:G1")
ws_info["A1"] = "УКАЗАНИЯ ЗА ПОПЪЛВАНЕ НА БЛАНКАТА ЗА ПРОМОЦИОНАЛНА БРОШУРА"
ws_info["A1"].font = font_title
ws_info["A1"].fill = fill_navy
ws_info["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_info.row_dimensions[1].height = 28

rules = [
    ("1. Ценообразуване и ДДС (Задължително с ДДС)", [
        "• Всички цени в бланката (редовна доставна, промоционална доставна и продажна цена) задължително се попълват С ВКЛЮЧЕН ДДС (20%).",
        "• Всички стойности се изчисляват и попълват в ЕВРО (€) съгласно валутния режим за 2027 г."
    ]),
    ("2. Период на промоционалната доставна цена (ЗКП)", [
        "• Промоционалната доставна цена (ЗКП) влиза в сила задължително 7 дни преди старта на брошурата.",
        "• Това гарантира навременно зареждане на Централния склад и 11-те търговски обекта на веригата.",
        "• Точният период е автоматично посочен в клетка B7 на бланката след избор на Брошура №."
    ]),
    ("3. Попълване на артикулите", [
        "• Попълвайте само жълтите полета в лист „Бланка_Предложения“.",
        "• Не е необходимо да знаете вътрешните артикулни кодове на „Бакалия“.",
        "• Моля посочвайте баркод (EAN-13), ако разполагате с такъв — това ускорява обработката.",
        "• При оферти за групи артикули (асортимент/видове на една цена), посочете основните разновидности в колона „Бележка / Специфика“."
    ]),
    ("4. Връщане на попълнената бланка", [
        "• Изпратете попълнения файл обратно по имейл най-късно до крайния срок, посочен в клетка B8.",
        "• За контакти и въпроси: Търговски отдел на „Бакалия“ — имейл: s.mart.tonev@gmail.com."
    ])
]

curr_r = 3
for r_title, r_points in rules:
    ws_info.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
    sec_c = ws_info.cell(row=curr_r, column=1, value=r_title)
    sec_c.font = font_sec_hdr
    sec_c.fill = fill_ice
    sec_c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    sec_c.border = thin_border
    ws_info.row_dimensions[curr_r].height = 22
    curr_r += 1

    for pt in r_points:
        ws_info.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
        pt_c = ws_info.cell(row=curr_r, column=1, value=pt)
        pt_c.font = font_regular
        pt_c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        pt_c.border = thin_border
        ws_info.row_dimensions[curr_r].height = 18
        curr_r += 1
    curr_r += 1

ws_info.column_dimensions["A"].width = 110

# Запазване
output_path = r"g:\My Drive\AntiG_test\БЛАНКА_ДОСТАВЧИЦИ_2027.xlsx"
wb.save(output_path)
print(f"Опростената бланка за доставчици е успешно създадена: {output_path}")
