# -*- coding: utf-8 -*-
"""
Full generator for the consolidated Excel workbook:
СРАВНИТЕЛЕН_АНАЛИЗ_ЯСКО_РОЖЕН_ДЕКЕМВРИ_2025_И_ЯНУАРИ_2026.xlsx
Includes the scenario comparison table from the screenshot on every relevant sheet
in a 100% editable, formula-driven format.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_workbook():
    # 1. Load source data
    wb_dec = openpyxl.load_workbook('СРАВНИТЕЛЕН_АНАЛИЗ_ЯСКО_РОЖЕН_ДЕКЕМВРИ_2025.xlsx', data_only=True)
    wb_jan = openpyxl.load_workbook('СРАВНИТЕЛЕН_АНАЛИЗ_ЯСКО_РОЖЕН_ЯНУАРИ_2026.xlsx', data_only=True)

    ws_dec_art = wb_dec['По_Артикули']
    ws_jan_art = wb_jan['По_Артикули']
    ws_dec_stores = wb_dec['По_Обекти']
    ws_jan_stores = wb_jan['По_Обекти']

    # Read Dec articles into dict
    dec_art = {}
    for r in range(2, ws_dec_art.max_row + 1):
        code = str(ws_dec_art.cell(r, 1).value).strip()
        dec_art[code] = {
            'code': code,
            'brand': ws_dec_art.cell(r, 2).value,
            'name': ws_dec_art.cell(r, 3).value,
            'qty': float(ws_dec_art.cell(r, 4).value or 0.0),
            'revenue': float(ws_dec_art.cell(r, 5).value or 0.0),
            'erp_cost': float(ws_dec_art.cell(r, 6).value or 0.0),
            'erp_profit': float(ws_dec_art.cell(r, 7).value or 0.0),
            'p_rojen_base': float(ws_dec_art.cell(r, 8).value or 0.0),
            'p_rojen_net': float(ws_dec_art.cell(r, 9).value or 0.0),
            'p_yasko': float(ws_dec_art.cell(r, 10).value or 0.0),
        }

    # Read Jan articles into dict
    jan_art = {}
    for r in range(2, ws_jan_art.max_row + 1):
        code = str(ws_jan_art.cell(r, 1).value).strip()
        jan_art[code] = {
            'code': code,
            'brand': ws_jan_art.cell(r, 2).value,
            'name': ws_jan_art.cell(r, 3).value,
            'qty': float(ws_jan_art.cell(r, 4).value or 0.0),
            'revenue': float(ws_jan_art.cell(r, 5).value or 0.0),
            'erp_cost': float(ws_jan_art.cell(r, 6).value or 0.0),
            'erp_profit': float(ws_jan_art.cell(r, 7).value or 0.0),
            'p_rojen_base': float(ws_jan_art.cell(r, 8).value or 0.0),
            'p_rojen_net': float(ws_jan_art.cell(r, 9).value or 0.0),
            'p_yasko': float(ws_jan_art.cell(r, 10).value or 0.0),
        }

    # Common list of sorted codes (by Dec revenue desc)
    sorted_codes = sorted(dec_art.keys(), key=lambda c: dec_art[c]['revenue'], reverse=True)

    # Read stores
    stores_list = []
    for r in range(2, ws_dec_stores.max_row + 1):
        sname = str(ws_dec_stores.cell(r, 1).value).strip()
        if sname and sname != 'None':
            stores_list.append(sname)

    dec_store_data = {}
    for r in range(2, ws_dec_stores.max_row + 1):
        sname = str(ws_dec_stores.cell(r, 1).value).strip()
        dec_store_data[sname] = {
            'qty': float(ws_dec_stores.cell(r, 2).value or 0.0),
            'revenue': float(ws_dec_stores.cell(r, 3).value or 0.0),
            'erp_cost': float(ws_dec_stores.cell(r, 4).value or 0.0),
            'erp_profit': float(ws_dec_stores.cell(r, 5).value or 0.0),
            'rojen_base': float(ws_dec_stores.cell(r, 6).value or 0.0),
            'rojen_net': float(ws_dec_stores.cell(r, 7).value or 0.0),
            'yasko': float(ws_dec_stores.cell(r, 8).value or 0.0),
        }

    jan_store_data = {}
    for r in range(2, ws_jan_stores.max_row + 1):
        sname = str(ws_jan_stores.cell(r, 1).value).strip()
        jan_store_data[sname] = {
            'qty': float(ws_jan_stores.cell(r, 2).value or 0.0),
            'revenue': float(ws_jan_stores.cell(r, 3).value or 0.0),
            'erp_profit': float(ws_jan_stores.cell(r, 4).value or 0.0),
            'rojen_base': float(ws_jan_stores.cell(r, 5).value or 0.0),
            'rojen_net': float(ws_jan_stores.cell(r, 6).value or 0.0),
            'yasko': float(ws_jan_stores.cell(r, 7).value or 0.0),
        }

    # 2. Setup New Workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    font_family = "Segoe UI"
    font_title = Font(name=font_family, size=13, bold=True, color="1F4E78")
    font_subtitle = Font(name=font_family, size=11, bold=True, color="1F4E78")
    font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    font_data = Font(name=font_family, size=10, color="000000")
    font_data_bold = Font(name=font_family, size=10, bold=True, color="000000")
    font_data_italic = Font(name=font_family, size=10, italic=True, color="595959")
    font_total = Font(name=font_family, size=10, bold=True, color="000000")

    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_dark_gray = PatternFill(start_color="262626", end_color="262626", fill_type="solid")
    fill_sub_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_green_yasko = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_yellow_rojen = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_gray_erp = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border_side = Side(style='thin', color='D9D9D9')
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(left=Side(style='thin', color='FFFFFF'), right=Side(style='thin', color='FFFFFF'),
                           top=Side(style='medium', color='1F4E78'), bottom=Side(style='medium', color='1F4E78'))
    border_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'),
                          left=thin_border_side, right=thin_border_side)

    align_left = Alignment(horizontal='left', vertical='center')
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_header = Alignment(horizontal='center', vertical='center', wrap_text=True)

    fmt_currency = '#,##0.00 "€"'
    fmt_currency_diff = '+ #,##0.00 "€";- #,##0.00 "€";0.00 "€"'
    fmt_pct = '0.00%'
    fmt_pct_diff = '+ 0.00%;- 0.00%;0.00%'
    fmt_qty = '#,##0.00'
    fmt_price = '0.0000 "€"'

    # -------------------------------------------------------------
    # 1. MONTHLY DETAIL SHEETS: 'Декември 2025' and 'Януари 2026'
    # -------------------------------------------------------------
    def create_monthly_sheet(sheet_name, month_title, art_data_map):
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].showGridLines = True

        # Row 1: Title
        ws.merge_cells('A1:G1')
        ws['A1'] = f"СРАВНИТЕЛЕН АНАЛИЗ НА СЦЕНАРИИ — {month_title} (РЕЗЮМЕ)"
        ws['A1'].font = font_subtitle
        ws['A1'].alignment = align_left

        # Rows 2 to 6: SCENARIO TABLE FROM SCREENSHOT
        headers_scen = [
            "Показател / Сценарий",
            "Покупна стойност (€ с ДДС)",
            "Брутна печалба (€ с ДДС)",
            "Търговски марж (%)",
            "Разлика спрямо Яско (€)",
            "% разлика спрямо Яско",
            "Икономически ефект / Бележка"
        ]
        ws.row_dimensions[2].height = 28
        for c_idx, h in enumerate(headers_scen, start=1):
            cell = ws.cell(2, c_idx, h)
            cell.font = font_header
            cell.fill = fill_dark_gray
            cell.alignment = align_header
            cell.border = border_header

        # Data rows for scenario table
        # Total row is at row 105 (94 items at rows 11 to 104)
        scen_rows = [
            (3, "1. Рожен — Базова ценова листа", "=K105", "=E105-B3", "=C3/E105", "=B3-B5", "=IF(B3>0,(B3-B5)/B3,0)", '="Яско е по-изгоден с " & ROUND(F3*100, 1) & "%"', None),
            (4, "2. Рожен — Нето фактура (-6.8% / -4.0%)", "=L105", "=E105-B4", "=C4/E105", "=B4-B5", "=IF(B4>0,(B4-B5)/B4,0)", '="Яско е по-изгоден с " & ROUND(F4*100, 1) & "%"', fill_yellow_rojen),
            (5, "3. Яско — Директни цени с ДДС", "=M105", "=E105-B5", "=C5/E105", "=0.00", 0.0, '="0.00 € (База за сравнение)"', fill_green_yasko),
            (6, "Текущо отчетено в Мистрал ERP", "=F105", "=G105", "=C6/E105", "=B6-B5", "=IF(B6>0,(B6-B5)/B6,0)", '="Яско спестява +" & ROUND(E6, 2) & " € спрямо реалното в ERP"', fill_gray_erp)
        ]

        for r_idx, label, f_cost, f_prof, f_mar, f_diff, f_pct, f_note, r_fill in scen_rows:
            ws.row_dimensions[r_idx].height = 22
            c1 = ws.cell(r_idx, 1, label)
            c1.font = font_data_bold if r_idx in [4, 5] else (font_data_italic if r_idx == 6 else font_data)
            c1.alignment = align_left
            c1.border = border_cell
            if r_fill: c1.fill = r_fill

            c2 = ws.cell(r_idx, 2, f_cost)
            c2.font = font_data_bold if r_idx in [4, 5] else font_data
            c2.number_format = fmt_currency
            c2.alignment = align_right
            c2.border = border_cell
            if r_fill: c2.fill = r_fill

            c3 = ws.cell(r_idx, 3, f_prof)
            c3.font = font_data_bold if r_idx in [4, 5] else font_data
            c3.number_format = fmt_currency
            c3.alignment = align_right
            c3.border = border_cell
            if r_fill: c3.fill = r_fill

            c4 = ws.cell(r_idx, 4, f_mar)
            c4.font = font_data_bold if r_idx in [4, 5] else font_data
            c4.number_format = fmt_pct
            c4.alignment = align_right
            c4.border = border_cell
            if r_fill: c4.fill = r_fill

            c5 = ws.cell(r_idx, 5, f_diff)
            c5.font = font_data_bold if r_idx in [4, 5] else font_data
            c5.number_format = fmt_currency_diff
            c5.alignment = align_right
            c5.border = border_cell
            if r_fill: c5.fill = r_fill

            c6 = ws.cell(r_idx, 6, f_pct)
            c6.font = font_data_bold if r_idx in [4, 5] else font_data
            c6.number_format = fmt_pct
            c6.alignment = align_right
            c6.border = border_cell
            if r_fill: c6.fill = r_fill

            c7 = ws.cell(r_idx, 7, f_note)
            c7.font = font_data_bold if r_idx in [4, 5] else font_data
            c7.alignment = align_left
            c7.border = border_cell
            if r_fill: c7.fill = r_fill

        # Row 8: Title for Detailed Articles Table
        ws.merge_cells('A8:V8')
        ws['A8'] = f"ДЕТАЙЛНА СПРАВКА ПО АРТИКУЛИ — {month_title} (РЕАЛНИ ПРОДАЖБИ И НОВИ ЦЕНИ НА ДОСТАВЧИЦИ)"
        ws['A8'].font = font_subtitle
        ws['A8'].alignment = align_left

        # Row 10: Headers of Detailed Articles Table
        headers_detail = [
            "Мат. №", "Бранд", "Артикул (Ценова листа)", "Продадени К-во (бр)", "Оборот Сума с ДДС (€)",
            "Реална Себестойност ERP (€)", "Реална Печалба ERP (€)",
            "Рожен Ед. Цена Базова (€)", "Рожен Ед. Цена Нето Ф-ра (€)", "Яско Ед. Цена (€)",
            "Покупна Стойност Рожен Базова (€)", "Покупна Стойност Рожен Нето (€)", "Покупна Стойност Яско (€)",
            "Печалба Рожен Базова (€)", "Печалба Рожен Нето (€)", "Печалба Яско (€)",
            "Икономия спрямо Базова Рожен (€)", "Икономия спрямо Нето Рожен (€)",
            "Марж Реален ERP %", "Марж Рожен Базова %", "Марж Рожен Нето %", "Марж Яско %"
        ]
        ws.row_dimensions[10].height = 28
        for c_idx, h in enumerate(headers_detail, start=1):
            cell = ws.cell(10, c_idx, h)
            cell.font = font_header
            cell.fill = fill_navy
            cell.alignment = align_header
            cell.border = border_header

        # Rows 11 to 104: 94 articles
        row_cur = 11
        for code in sorted_codes:
            item = art_data_map[code]
            ws.row_dimensions[row_cur].height = 19

            ws.cell(row_cur, 1, item['code']).alignment = align_center
            ws.cell(row_cur, 2, item['brand']).alignment = align_center
            ws.cell(row_cur, 3, item['name']).alignment = align_left

            # D: Qty
            c_qty = ws.cell(row_cur, 4, item['qty'])
            c_qty.number_format = fmt_qty
            c_qty.alignment = align_right

            # E: Revenue
            c_rev = ws.cell(row_cur, 5, item['revenue'])
            c_rev.number_format = fmt_currency
            c_rev.alignment = align_right

            # F: ERP Cost
            c_ecost = ws.cell(row_cur, 6, item['erp_cost'])
            c_ecost.number_format = fmt_currency
            c_ecost.alignment = align_right

            # G: ERP Profit
            c_eprof = ws.cell(row_cur, 7, f"=E{row_cur}-F{row_cur}")
            c_eprof.number_format = fmt_currency
            c_eprof.alignment = align_right

            # H: Unit Base Rozhen
            c_ub = ws.cell(row_cur, 8, item['p_rojen_base'])
            c_ub.number_format = fmt_price
            c_ub.alignment = align_right

            # I: Unit Net Rozhen
            c_un = ws.cell(row_cur, 9, f'=IF(B{row_cur}="MDLZ", H{row_cur}*(1-0.068), H{row_cur}*(1-0.040))')
            c_un.number_format = fmt_price
            c_un.alignment = align_right

            # J: Unit Yasko
            c_uy = ws.cell(row_cur, 10, item['p_yasko'])
            c_uy.number_format = fmt_price
            c_uy.alignment = align_right

            # K: Cost Rozhen Base
            c_cb = ws.cell(row_cur, 11, f"=D{row_cur}*H{row_cur}")
            c_cb.number_format = fmt_currency
            c_cb.alignment = align_right

            # L: Cost Rozhen Net
            c_cn = ws.cell(row_cur, 12, f"=D{row_cur}*I{row_cur}")
            c_cn.number_format = fmt_currency
            c_cn.alignment = align_right

            # M: Cost Yasko
            c_cy = ws.cell(row_cur, 13, f"=D{row_cur}*J{row_cur}")
            c_cy.number_format = fmt_currency
            c_cy.alignment = align_right

            # N: Profit Rozhen Base
            c_pb = ws.cell(row_cur, 14, f"=E{row_cur}-K{row_cur}")
            c_pb.number_format = fmt_currency
            c_pb.alignment = align_right

            # O: Profit Rozhen Net
            c_pn = ws.cell(row_cur, 15, f"=E{row_cur}-L{row_cur}")
            c_pn.number_format = fmt_currency
            c_pn.alignment = align_right

            # P: Profit Yasko
            c_py = ws.cell(row_cur, 16, f"=E{row_cur}-M{row_cur}")
            c_py.number_format = fmt_currency
            c_py.alignment = align_right

            # Q: Savings vs Base
            c_sb = ws.cell(row_cur, 17, f"=K{row_cur}-M{row_cur}")
            c_sb.number_format = fmt_currency_diff
            c_sb.alignment = align_right

            # R: Savings vs Net
            c_sn = ws.cell(row_cur, 18, f"=L{row_cur}-M{row_cur}")
            c_sn.number_format = fmt_currency_diff
            c_sn.alignment = align_right

            # S: Margin ERP %
            c_me = ws.cell(row_cur, 19, f"=IF(E{row_cur}>0, G{row_cur}/E{row_cur}, 0)")
            c_me.number_format = fmt_pct
            c_me.alignment = align_right

            # T: Margin Rozhen Base %
            c_mb = ws.cell(row_cur, 20, f"=IF(E{row_cur}>0, N{row_cur}/E{row_cur}, 0)")
            c_mb.number_format = fmt_pct
            c_mb.alignment = align_right

            # U: Margin Rozhen Net %
            c_mn = ws.cell(row_cur, 21, f"=IF(E{row_cur}>0, O{row_cur}/E{row_cur}, 0)")
            c_mn.number_format = fmt_pct
            c_mn.alignment = align_right

            # V: Margin Yasko %
            c_my = ws.cell(row_cur, 22, f"=IF(E{row_cur}>0, P{row_cur}/E{row_cur}, 0)")
            c_my.number_format = fmt_pct
            c_my.alignment = align_right

            for c_i in range(1, 23):
                cell_item = ws.cell(row_cur, c_i)
                cell_item.border = border_cell
                cell_item.font = font_data

            row_cur += 1

        # Total Row at row 105
        total_row = row_cur
        ws.row_dimensions[total_row].height = 24
        ws.merge_cells(f'A{total_row}:C{total_row}')
        c_tot_label = ws.cell(total_row, 1, "ОБЩО / TOTAL")
        c_tot_label.font = font_total
        c_tot_label.alignment = align_center

        ws.cell(total_row, 4, f"=SUM(D11:D{total_row-1})").number_format = fmt_qty
        ws.cell(total_row, 5, f"=SUM(E11:E{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 6, f"=SUM(F11:F{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 7, f"=SUM(G11:G{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 8, "").alignment = align_center
        ws.cell(total_row, 9, "").alignment = align_center
        ws.cell(total_row, 10, "").alignment = align_center
        ws.cell(total_row, 11, f"=SUM(K11:K{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 12, f"=SUM(L11:L{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 13, f"=SUM(M11:M{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 14, f"=SUM(N11:N{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 15, f"=SUM(O11:O{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 16, f"=SUM(P11:P{total_row-1})").number_format = fmt_currency
        ws.cell(total_row, 17, f"=SUM(Q11:Q{total_row-1})").number_format = fmt_currency_diff
        ws.cell(total_row, 18, f"=SUM(R11:R{total_row-1})").number_format = fmt_currency_diff
        ws.cell(total_row, 19, f"=G{total_row}/E{total_row}").number_format = fmt_pct
        ws.cell(total_row, 20, f"=N{total_row}/E{total_row}").number_format = fmt_pct
        ws.cell(total_row, 21, f"=O{total_row}/E{total_row}").number_format = fmt_pct
        ws.cell(total_row, 22, f"=P{total_row}/E{total_row}").number_format = fmt_pct

        for c_i in range(1, 23):
            c_t = ws.cell(total_row, c_i)
            c_t.font = font_total
            c_t.border = border_total
            c_t.fill = fill_total
            if c_i >= 4:
                c_t.alignment = align_right

        # Column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.row in [1, 8]:
                    continue
                if '\n' in val_str:
                    lines = val_str.split('\n')
                    val_str = max(lines, key=len)
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 34
        ws.column_dimensions['G'].width = 32

    create_monthly_sheet('Декември 2025', 'ДЕКЕМВРИ 2025', dec_art)
    create_monthly_sheet('Януари 2026', 'ЯНУАРИ 2026', jan_art)

    # -------------------------------------------------------------
    # 2. CONSOLIDATED ARTICLES SHEET: 'Обобщено по артикули'
    # Has the 2-Month Scenario Table at rows 2-6, and detailed table at rows 11-105
    # -------------------------------------------------------------
    ws_art_all = wb.create_sheet(title='Обобщено по артикули')
    ws_art_all.views.sheetView[0].showGridLines = True

    # Row 1: Title
    ws_art_all.merge_cells('A1:G1')
    ws_art_all['A1'] = "СРАВНИТЕЛЕН АНАЛИЗ НА СЦЕНАРИИ — ОБЩО 2 МЕСЕЦА КУМУЛАТИВНО (РЕЗЮМЕ)"
    ws_art_all['A1'].font = font_subtitle
    ws_art_all['A1'].alignment = align_left

    # Rows 2 to 6: SCENARIO TABLE FROM SCREENSHOT (2-MONTH COMBINED)
    headers_scen = [
        "Показател / Сценарий",
        "Покупна стойност (€ с ДДС)",
        "Брутна печалба (€ с ДДС)",
        "Търговски марж (%)",
        "Разлика спрямо Яско (€)",
        "% разлика спрямо Яско",
        "Икономически ефект / Бележка"
    ]
    ws_art_all.row_dimensions[2].height = 28
    for c_idx, h in enumerate(headers_scen, start=1):
        cell = ws_art_all.cell(2, c_idx, h)
        cell.font = font_header
        cell.fill = fill_dark_gray
        cell.alignment = align_header
        cell.border = border_header

    # Total row is at row 105 (Revenue in I105, Base Cost in M105, Net Cost in N105, Yasko Cost in O105)
    scen_rows_2m = [
        (3, "1. Рожен — Базова ценова листа", "=M105", "=I105-B3", "=C3/I105", "=B3-B5", "=IF(B3>0,(B3-B5)/B3,0)", '="Яско е по-изгоден с " & ROUND(F3*100, 1) & "%"', None),
        (4, "2. Рожен — Нето фактура (-6.8% / -4.0%)", "=N105", "=I105-B4", "=C4/I105", "=B4-B5", "=IF(B4>0,(B4-B5)/B4,0)", '="Яско е по-изгоден с " & ROUND(F4*100, 1) & "%"', fill_yellow_rojen),
        (5, "3. Яско — Директни цени с ДДС", "=O105", "=I105-B5", "=C5/I105", "=0.00", 0.0, '="0.00 € (База за сравнение)"', fill_green_yasko),
        (6, "Текущо отчетено в Мистрал ERP", "='Декември 2025'!B6+'Януари 2026'!B6", "='Декември 2025'!C6+'Януари 2026'!C6", "=C6/I105", "=B6-B5", "=IF(B6>0,(B6-B5)/B6,0)", '="Яско спестява общо +" & ROUND(E6, 2) & " € спрямо реалното в ERP"', fill_gray_erp)
    ]

    for r_idx, label, f_cost, f_prof, f_mar, f_diff, f_pct, f_note, r_fill in scen_rows_2m:
        ws_art_all.row_dimensions[r_idx].height = 22
        c1 = ws_art_all.cell(r_idx, 1, label)
        c1.font = font_data_bold if r_idx in [4, 5] else (font_data_italic if r_idx == 6 else font_data)
        c1.alignment = align_left
        c1.border = border_cell
        if r_fill: c1.fill = r_fill

        c2 = ws_art_all.cell(r_idx, 2, f_cost)
        c2.font = font_data_bold if r_idx in [4, 5] else font_data
        c2.number_format = fmt_currency
        c2.alignment = align_right
        c2.border = border_cell
        if r_fill: c2.fill = r_fill

        c3 = ws_art_all.cell(r_idx, 3, f_prof)
        c3.font = font_data_bold if r_idx in [4, 5] else font_data
        c3.number_format = fmt_currency
        c3.alignment = align_right
        c3.border = border_cell
        if r_fill: c3.fill = r_fill

        c4 = ws_art_all.cell(r_idx, 4, f_mar)
        c4.font = font_data_bold if r_idx in [4, 5] else font_data
        c4.number_format = fmt_pct
        c4.alignment = align_right
        c4.border = border_cell
        if r_fill: c4.fill = r_fill

        c5 = ws_art_all.cell(r_idx, 5, f_diff)
        c5.font = font_data_bold if r_idx in [4, 5] else font_data
        c5.number_format = fmt_currency_diff
        c5.alignment = align_right
        c5.border = border_cell
        if r_fill: c5.fill = r_fill

        c6 = ws_art_all.cell(r_idx, 6, f_pct)
        c6.font = font_data_bold if r_idx in [4, 5] else font_data
        c6.number_format = fmt_pct
        c6.alignment = align_right
        c6.border = border_cell
        if r_fill: c6.fill = r_fill

        c7 = ws_art_all.cell(r_idx, 7, f_note)
        c7.font = font_data_bold if r_idx in [4, 5] else font_data
        c7.alignment = align_left
        c7.border = border_cell
        if r_fill: c7.fill = r_fill

    # Row 8: Title for Detailed Articles Table
    ws_art_all.merge_cells('A8:R8')
    ws_art_all['A8'] = "КУМУЛАТИВНА СПРАВКА ПО АРТИКУЛИ (2 МЕСЕЦА: ДЕКЕМВРИ 2025 И ЯНУАРИ 2026)"
    ws_art_all['A8'].font = font_subtitle
    ws_art_all['A8'].alignment = align_left

    # Row 10: Column Headers
    headers_all_art = [
        "Мат. №", "Бранд", "Артикул (Ценова листа)",
        "К-во Дек (бр)", "К-во Ян (бр)", "ОБЩО К-во (бр)",
        "Оборот Дек (€)", "Оборот Ян (€)", "ОБЩ Оборот с ДДС (€)",
        "Рожен Ед. Базова (€)", "Рожен Ед. Нето (€)", "Яско Ед. (€)",
        "Общо Рожен Базова (€)", "Общо Рожен Нето (€)", "Общо Яско (€)",
        "ЧИСТА ИКОНОМИЯ ЯСКО (€)", "Марж Рожен Нето %", "Марж Яско %"
    ]
    ws_art_all.row_dimensions[10].height = 28
    for c_i, h in enumerate(headers_all_art, start=1):
        cell = ws_art_all.cell(10, c_i, h)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_header
        cell.border = border_header

    # Rows 11 to 104 (94 articles)
    r_art = 11
    for idx, code in enumerate(sorted_codes):
        item = dec_art[code]
        src_row = 11 + idx

        ws_art_all.row_dimensions[r_art].height = 19
        ws_art_all.cell(r_art, 1, item['code']).alignment = align_center
        ws_art_all.cell(r_art, 2, item['brand']).alignment = align_center
        ws_art_all.cell(r_art, 3, item['name']).alignment = align_left

        # D: Qty Dec, E: Qty Jan, F: Total Qty
        ws_art_all.cell(r_art, 4, f"='Декември 2025'!D{src_row}").number_format = fmt_qty
        ws_art_all.cell(r_art, 5, f"='Януари 2026'!D{src_row}").number_format = fmt_qty
        ws_art_all.cell(r_art, 6, f"=D{r_art}+E{r_art}").number_format = fmt_qty

        # G: Rev Dec, H: Rev Jan, I: Total Rev
        ws_art_all.cell(r_art, 7, f"='Декември 2025'!E{src_row}").number_format = fmt_currency
        ws_art_all.cell(r_art, 8, f"='Януари 2026'!E{src_row}").number_format = fmt_currency
        ws_art_all.cell(r_art, 9, f"=G{r_art}+H{r_art}").number_format = fmt_currency

        # J: Unit Base Rozhen, K: Unit Net Rozhen, L: Unit Yasko
        ws_art_all.cell(r_art, 10, f"='Декември 2025'!H{src_row}").number_format = fmt_price
        ws_art_all.cell(r_art, 11, f"='Декември 2025'!I{src_row}").number_format = fmt_price
        ws_art_all.cell(r_art, 12, f"='Декември 2025'!J{src_row}").number_format = fmt_price

        # M: Total Cost Rozhen Base = F * J
        ws_art_all.cell(r_art, 13, f"=F{r_art}*J{r_art}").number_format = fmt_currency

        # N: Total Cost Rozhen Net = F * K
        ws_art_all.cell(r_art, 14, f"=F{r_art}*K{r_art}").number_format = fmt_currency

        # O: Total Cost Yasko = F * L
        ws_art_all.cell(r_art, 15, f"=F{r_art}*L{r_art}").number_format = fmt_currency

        # P: Net Savings Yasko vs Net Rozhen = N - O
        c_sav = ws_art_all.cell(r_art, 16, f"=N{r_art}-O{r_art}")
        c_sav.number_format = fmt_currency_diff
        c_sav.font = font_data_bold
        c_sav.fill = fill_green_yasko

        # Q: Margin Rozhen Net % = (I - N) / I
        ws_art_all.cell(r_art, 17, f"=IF(I{r_art}>0, (I{r_art}-N{r_art})/I{r_art}, 0)").number_format = fmt_pct

        # R: Margin Yasko % = (I - O) / I
        ws_art_all.cell(r_art, 18, f"=IF(I{r_art}>0, (I{r_art}-O{r_art})/I{r_art}, 0)").number_format = fmt_pct

        for c_i in range(1, 19):
            cell_item = ws_art_all.cell(r_art, c_i)
            cell_item.border = border_cell
            if c_i > 3:
                cell_item.alignment = align_right
                if c_i != 16:
                    cell_item.font = font_data

        r_art += 1

    # Total row at row 105
    tot_art_r = r_art
    ws_art_all.row_dimensions[tot_art_r].height = 24
    ws_art_all.merge_cells(f'A{tot_art_r}:C{tot_art_r}')
    ws_art_all.cell(tot_art_r, 1, "ОБЩО / TOTAL").font = font_total

    ws_art_all.cell(tot_art_r, 4, f"=SUM(D11:D{tot_art_r-1})").number_format = fmt_qty
    ws_art_all.cell(tot_art_r, 5, f"=SUM(E11:E{tot_art_r-1})").number_format = fmt_qty
    ws_art_all.cell(tot_art_r, 6, f"=SUM(F11:F{tot_art_r-1})").number_format = fmt_qty

    ws_art_all.cell(tot_art_r, 7, f"=SUM(G11:G{tot_art_r-1})").number_format = fmt_currency
    ws_art_all.cell(tot_art_r, 8, f"=SUM(H11:H{tot_art_r-1})").number_format = fmt_currency
    ws_art_all.cell(tot_art_r, 9, f"=SUM(I11:I{tot_art_r-1})").number_format = fmt_currency

    ws_art_all.cell(tot_art_r, 10, "").alignment = align_center
    ws_art_all.cell(tot_art_r, 11, "").alignment = align_center
    ws_art_all.cell(tot_art_r, 12, "").alignment = align_center

    ws_art_all.cell(tot_art_r, 13, f"=SUM(M11:M{tot_art_r-1})").number_format = fmt_currency
    ws_art_all.cell(tot_art_r, 14, f"=SUM(N11:N{tot_art_r-1})").number_format = fmt_currency
    ws_art_all.cell(tot_art_r, 15, f"=SUM(O11:O{tot_art_r-1})").number_format = fmt_currency
    ws_art_all.cell(tot_art_r, 16, f"=SUM(P11:P{tot_art_r-1})").number_format = fmt_currency_diff

    ws_art_all.cell(tot_art_r, 17, f"=(I{tot_art_r}-N{tot_art_r})/I{tot_art_r}").number_format = fmt_pct
    ws_art_all.cell(tot_art_r, 18, f"=(I{tot_art_r}-O{tot_art_r})/I{tot_art_r}").number_format = fmt_pct

    for c_i in range(1, 19):
        c_t = ws_art_all.cell(tot_art_r, c_i)
        c_t.font = font_total
        c_t.border = border_total
        c_t.fill = fill_total
        if c_i >= 4:
            c_t.alignment = align_right

    for col in ws_art_all.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row in [1, 8]:
                continue
            max_len = max(max_len, len(val_str))
        ws_art_all.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws_art_all.column_dimensions['A'].width = 12
    ws_art_all.column_dimensions['B'].width = 10
    ws_art_all.column_dimensions['C'].width = 34
    ws_art_all.column_dimensions['G'].width = 32

    # -------------------------------------------------------------
    # 3. MASTER SUMMARY SHEET: 'Обобщение (Резюме)'
    # -------------------------------------------------------------
    ws_sum = wb.create_sheet(title='Обобщение (Резюме)', index=0)
    ws_sum.views.sheetView[0].showGridLines = True

    # Title
    ws_sum.merge_cells('A1:G1')
    ws_sum['A1'] = "ВЕРИГА СУПЕРМАРКЕТИ „БАКАЛИЯ“ — ГЛАВЕН ТЪРГОВСКИ КОНТРОЛИНГ"
    ws_sum['A1'].font = font_title
    ws_sum['A1'].alignment = align_left

    ws_sum.merge_cells('A2:G2')
    ws_sum['A2'] = "СРАВНИТЕЛЕН АНАЛИЗ НА ДОСТАВЧИЦИ: ЯСКО vs. РОЖЕН (ДЕКЕМВРИ 2025 И ЯНУАРИ 2026)"
    ws_sum['A2'].font = font_subtitle
    ws_sum['A2'].alignment = align_left

    ws_sum.merge_cells('A3:G3')
    ws_sum['A3'] = "Обхват: 104 позиции (Mondelez & Chipita) | 11 магазина | Мистрал ERP Централен склад (CS.FDB)"
    ws_sum['A3'].font = font_data_italic
    ws_sum['A3'].alignment = align_left

    headers_sum_scen = [
        "Показател / Сценарий",
        "Покупна стойност (€ с ДДС)",
        "Брутна печалба (€ с ДДС)",
        "Търговски марж (%)",
        "Разлика спрямо Яско (€)",
        "% разлика спрямо Яско",
        "Икономически ефект / Бележка"
    ]

    def add_summary_scenario_block(start_r, title_text, sheet_ref, header_fill):
        ws_sum.merge_cells(f'A{start_r}:G{start_r}')
        ws_sum[f'A{start_r}'] = title_text
        ws_sum[f'A{start_r}'].font = font_subtitle
        ws_sum[f'A{start_r}'].alignment = align_left

        r_hdr = start_r + 1
        ws_sum.row_dimensions[r_hdr].height = 26
        for c_i, h in enumerate(headers_sum_scen, start=1):
            cell = ws_sum.cell(r_hdr, c_i, h)
            cell.font = font_header
            cell.fill = header_fill
            cell.alignment = align_header
            cell.border = border_header

        labels = [
            ("1. Рожен — Базова ценова листа", 3, None),
            ("2. Рожен — Нето фактура (-6.8% / -4.0%)", 4, fill_yellow_rojen),
            ("3. Яско — Директни цени с ДДС", 5, fill_green_yasko),
            ("Текущо отчетено в Мистрал ERP", 6, fill_gray_erp)
        ]

        for idx, (label, src_r, r_fill) in enumerate(labels, start=1):
            r = r_hdr + idx
            ws_sum.row_dimensions[r].height = 22

            c1 = ws_sum.cell(r, 1, label)
            c1.font = font_data_bold if idx in [2, 3] else (font_data_italic if idx == 4 else font_data)
            c1.alignment = align_left
            c1.border = border_cell
            if r_fill: c1.fill = r_fill

            # Cost
            c2 = ws_sum.cell(r, 2, f"='{sheet_ref}'!B{src_r}")
            c2.font = font_data_bold if idx in [2, 3] else font_data
            c2.number_format = fmt_currency
            c2.alignment = align_right
            c2.border = border_cell
            if r_fill: c2.fill = r_fill

            # Profit
            c3 = ws_sum.cell(r, 3, f"='{sheet_ref}'!C{src_r}")
            c3.font = font_data_bold if idx in [2, 3] else font_data
            c3.number_format = fmt_currency
            c3.alignment = align_right
            c3.border = border_cell
            if r_fill: c3.fill = r_fill

            # Margin
            c4 = ws_sum.cell(r, 4, f"='{sheet_ref}'!D{src_r}")
            c4.font = font_data_bold if idx in [2, 3] else font_data
            c4.number_format = fmt_pct
            c4.alignment = align_right
            c4.border = border_cell
            if r_fill: c4.fill = r_fill

            # Diff vs Yasko
            c5 = ws_sum.cell(r, 5, f"='{sheet_ref}'!E{src_r}")
            c5.font = font_data_bold if idx in [2, 3] else font_data
            c5.number_format = fmt_currency_diff
            c5.alignment = align_right
            c5.border = border_cell
            if r_fill: c5.fill = r_fill

            # % Diff
            c6 = ws_sum.cell(r, 6, f"='{sheet_ref}'!F{src_r}")
            c6.font = font_data_bold if idx in [2, 3] else font_data
            c6.number_format = fmt_pct
            c6.alignment = align_right
            c6.border = border_cell
            if r_fill: c6.fill = r_fill

            # Note
            c7 = ws_sum.cell(r, 7, f"='{sheet_ref}'!G{src_r}")
            c7.font = font_data_bold if idx in [2, 3] else font_data
            c7.alignment = align_left
            c7.border = border_cell
            if r_fill: c7.fill = r_fill

    # Block 1: December 2025
    add_summary_scenario_block(5, "1. ДЕКЕМВРИ 2025 — СРАВНЕНИЕ НА СЦЕНАРИИ (РЕЗЮМЕ)", 'Декември 2025', fill_dark_gray)

    # Block 2: January 2026
    add_summary_scenario_block(12, "2. ЯНУАРИ 2026 — СРАВНЕНИЕ НА СЦЕНАРИИ (РЕЗЮМЕ)", 'Януари 2026', fill_dark_gray)

    # Block 3: Combined 2 Months
    add_summary_scenario_block(19, "3. ОБЩО 2 МЕСЕЦА (КУМУЛАТИВНО) — СРАВНЕНИЕ НА СЦЕНАРИИ", 'Обобщено по артикули', fill_navy)

    # Block 4: Side-by-side Month Comparison
    ws_sum.merge_cells('A26:E26')
    ws_sum['A26'] = "4. ДИНАМИКА МЕСЕЦ ПО МЕСЕЦ (ДЕКЕМВРИ vs. ЯНУАРИ vs. ОБЩО)"
    ws_sum['A26'].font = font_subtitle
    ws_sum['A26'].alignment = align_left

    headers_mom = ["Показател", "Декември 2025 (€)", "Януари 2026 (€)", "Общо 2 Месеца (€)", "Анализ / Темп на промяна"]
    ws_sum.row_dimensions[27].height = 26
    for c_i, h in enumerate(headers_mom, start=1):
        cell = ws_sum.cell(27, c_i, h)
        cell.font = font_header
        cell.fill = fill_sub_header
        cell.alignment = align_header
        cell.border = border_header

    mom_metrics = [
        ("Продадени количества (бр.)", "='Декември 2025'!D105", "='Януари 2026'!D105", "=B28+C28", fmt_qty, '="Ръст Януари: " & ROUND(((C28-B28)/B28)*100, 1) & "%"'),
        ("Реализиран оборот с ДДС (€)", "='Декември 2025'!E105", "='Януари 2026'!E105", "=B29+C29", fmt_currency, '="Дял Декември: " & ROUND((B29/D29)*100, 1) & "%"'),
        ("Реална себестойност ERP (€)", "='Декември 2025'!F105", "='Януари 2026'!F105", "=B30+C30", fmt_currency, '=""'),
        ("Покупна стойност Рожен Базова (€)", "=B7", "=B14", "=B21", fmt_currency, '=""'),
        ("Покупна стойност Рожен Нето (€)", "=B8", "=B15", "=B22", fmt_currency, '=""'),
        ("Покупна стойност Яско (€)", "=B9", "=B16", "=B23", fmt_currency, '=""'),
        ("Чиста икономия от Яско спр. Рожен Базова (€)", "=B7-B9", "=B14-B16", "=B21-B23", fmt_currency, '="Общо спестени: " & ROUND(D34, 2) & " €"'),
        ("Чиста икономия от Яско спр. Рожен Нето (€)", "=B8-B9", "=B15-B16", "=B22-B23", fmt_currency, '="Общо спестени: " & ROUND(D35, 2) & " €"'),
        ("Търговски марж Рожен Нето (%)", "=D8", "=D15", "=D22", fmt_pct, '=""'),
        ("Търговски марж Яско (%)", "=D9", "=D16", "=D23", fmt_pct, '=""'),
        ("Допълнителен марж от Яско (+ п.п.)", "=D9-D8", "=D16-D15", "=D23-D22", fmt_pct_diff, '="Среден ръст: +" & ROUND(D38*100, 2) & " п.п."')
    ]

    for idx, (m_label, f_dec, f_jan, f_tot, m_fmt, f_comm) in enumerate(mom_metrics, start=1):
        r = 27 + idx
        ws_sum.row_dimensions[r].height = 20
        c1 = ws_sum.cell(r, 1, m_label)
        c1.font = font_data_bold if "икономия" in m_label.lower() or "марж" in m_label.lower() else font_data
        c1.alignment = align_left
        c1.border = border_cell

        c2 = ws_sum.cell(r, 2, f_dec)
        c2.font = font_data_bold if "икономия" in m_label.lower() else font_data
        c2.number_format = m_fmt
        c2.alignment = align_right
        c2.border = border_cell

        c3 = ws_sum.cell(r, 3, f_jan)
        c3.font = font_data_bold if "икономия" in m_label.lower() else font_data
        c3.number_format = m_fmt
        c3.alignment = align_right
        c3.border = border_cell

        c4 = ws_sum.cell(r, 4, f_tot)
        c4.font = font_data_bold
        c4.number_format = m_fmt
        c4.alignment = align_right
        c4.border = border_cell
        if "икономия" in m_label.lower():
            c4.fill = fill_green_yasko

        c5 = ws_sum.cell(r, 5, f_comm)
        c5.font = font_data_bold if "икономия" in m_label.lower() else font_data
        c5.alignment = align_left
        c5.border = border_cell

    # Block 5: Brands Summary
    r_br_start = 41
    ws_sum.merge_cells(f'A{r_br_start}:G{r_br_start}')
    ws_sum[f'A{r_br_start}'] = "5. РАЗПРЕДЕЛЕНИЕ ПО БРАНДОВЕ (MDLZ vs. CHIPITA — ОБЩО 2 МЕСЕЦА)"
    ws_sum[f'A{r_br_start}'].font = font_subtitle
    ws_sum[f'A{r_br_start}'].alignment = align_left

    headers_br = [
        "Бранд", "Общо К-во (бр)", "Общ Оборот (€)",
        "Себестойност Рожен Нето (€)", "Себестойност Яско (€)",
        "Чиста икономия от Яско (€)", "Марж Яско (%)"
    ]
    r_br_hdr = r_br_start + 1
    ws_sum.row_dimensions[r_br_hdr].height = 26
    for c_i, h in enumerate(headers_br, start=1):
        cell = ws_sum.cell(r_br_hdr, c_i, h)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_header
        cell.border = border_header

    # In 'Обобщено по артикули', Col B is Brand, Col F is Total Qty, Col I is Total Turnover, Col N is Total Rozhen Net, Col O is Total Yasko
    br_rows = [
        ("MDLZ (Mondelez)", '=SUMIF(\'Обобщено по артикули\'!B11:B104, "MDLZ", \'Обобщено по артикули\'!F11:F104)',
                           '=SUMIF(\'Обобщено по артикули\'!B11:B104, "MDLZ", \'Обобщено по артикули\'!I11:I104)',
                           '=SUMIF(\'Обобщено по артикули\'!B11:B104, "MDLZ", \'Обобщено по артикули\'!N11:N104)',
                           '=SUMIF(\'Обобщено по артикули\'!B11:B104, "MDLZ", \'Обобщено по артикули\'!O11:O104)',
                           '=D43-E43', '=(C43-E43)/C43'),
        ("CHIPITA",        '=SUMIF(\'Обобщено по артикули\'!B11:B104, "CHIPITA", \'Обобщено по артикули\'!F11:F104)',
                           '=SUMIF(\'Обобщено по артикули\'!B11:B104, "CHIPITA", \'Обобщено по артикули\'!I11:I104)',
                           '=SUMIF(\'Обобщено по артикули\'!B11:B104, "CHIPITA", \'Обобщено по артикули\'!N11:N104)',
                           '=SUMIF(\'Обобщено по артикули\'!B11:B104, "CHIPITA", \'Обобщено по артикули\'!O11:O104)',
                           '=D44-E44', '=(C44-E44)/C44'),
        ("ОБЩО",           '=SUM(B43:B44)', '=SUM(C43:C44)', '=SUM(D43:D44)', '=SUM(E43:E44)', '=D45-E45', '=(C45-E45)/C45')
    ]

    for idx, (b_label, f_bq, f_bo, f_brn, f_by, f_bs, f_bm) in enumerate(br_rows, start=1):
        r = r_br_hdr + idx
        ws_sum.row_dimensions[r].height = 22
        is_tot = (idx == 3)
        c1 = ws_sum.cell(r, 1, b_label)
        c1.font = font_total if is_tot else font_data_bold
        c1.alignment = align_left
        c1.border = border_total if is_tot else border_cell
        if is_tot: c1.fill = fill_total

        c2 = ws_sum.cell(r, 2, f_bq)
        c2.font = font_total if is_tot else font_data
        c2.number_format = fmt_qty
        c2.alignment = align_right
        c2.border = border_total if is_tot else border_cell
        if is_tot: c2.fill = fill_total

        c3 = ws_sum.cell(r, 3, f_bo)
        c3.font = font_total if is_tot else font_data
        c3.number_format = fmt_currency
        c3.alignment = align_right
        c3.border = border_total if is_tot else border_cell
        if is_tot: c3.fill = fill_total

        c4 = ws_sum.cell(r, 4, f_brn)
        c4.font = font_total if is_tot else font_data
        c4.number_format = fmt_currency
        c4.alignment = align_right
        c4.border = border_total if is_tot else border_cell
        if is_tot: c4.fill = fill_total

        c5 = ws_sum.cell(r, 5, f_by)
        c5.font = font_total if is_tot else font_data
        c5.number_format = fmt_currency
        c5.alignment = align_right
        c5.border = border_total if is_tot else border_cell
        if is_tot: c5.fill = fill_total

        c6 = ws_sum.cell(r, 6, f_bs)
        c6.font = font_total if is_tot else font_data_bold
        c6.number_format = fmt_currency
        c6.alignment = align_right
        c6.border = border_total if is_tot else border_cell
        if is_tot:
            c6.fill = fill_total
        else:
            c6.fill = fill_green_yasko

        c7 = ws_sum.cell(r, 7, f_bm)
        c7.font = font_total if is_tot else font_data_bold
        c7.number_format = fmt_pct
        c7.alignment = align_right
        c7.border = border_total if is_tot else border_cell
        if is_tot: c7.fill = fill_total

    # Adjust widths on summary sheet
    for col in ws_sum.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row in [1, 2, 3, 5, 12, 19, 26, 41]:
                continue
            max_len = max(max_len, len(val_str))
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 15)
    ws_sum.column_dimensions['A'].width = 38
    ws_sum.column_dimensions['B'].width = 24
    ws_sum.column_dimensions['C'].width = 24
    ws_sum.column_dimensions['D'].width = 18
    ws_sum.column_dimensions['E'].width = 24
    ws_sum.column_dimensions['F'].width = 20
    ws_sum.column_dimensions['G'].width = 38

    # -------------------------------------------------------------
    # 4. STORE COMPARISON SHEET: 'Сравнение по обекти'
    # -------------------------------------------------------------
    ws_st = wb.create_sheet(title='Сравнение по обекти')
    ws_st.views.sheetView[0].showGridLines = True

    ws_st.merge_cells('A1:Q1')
    ws_st['A1'] = "СРАВНИТЕЛЕН АНАЛИЗ ПО ТЪРГОВСКИ ОБЕКТИ (11 МАГАЗИНА) — ДЕКЕМВРИ 2025 vs. ЯНУАРИ 2026 vs. ОБЩО"
    ws_st['A1'].font = font_title
    ws_st['A1'].alignment = align_left

    headers_stores = [
        "Магазин / Обект",
        "К-во Дек (бр)", "К-во Ян (бр)", "ОБЩО К-во (бр)",
        "Оборот Дек (€)", "Оборот Ян (€)", "ОБЩ Оборот (€)",
        "Рожен Нето Дек (€)", "Рожен Нето Ян (€)", "ОБЩО Рожен Нето (€)",
        "Яско Дек (€)", "Яско Ян (€)", "ОБЩО Яско (€)",
        "Икономия Дек (€)", "Икономия Ян (€)", "ОБЩА ИКОНОМИЯ (€)",
        "Марж Яско 2м (%)"
    ]
    ws_st.row_dimensions[3].height = 28
    for c_i, h in enumerate(headers_stores, start=1):
        cell = ws_st.cell(3, c_i, h)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_header
        cell.border = border_header

    sorted_stores = sorted(stores_list, key=lambda s: dec_store_data.get(s, {}).get('revenue', 0.0), reverse=True)

    row_st = 4
    for sname in sorted_stores:
        ws_st.row_dimensions[row_st].height = 20
        d_st = dec_store_data.get(sname, {})
        j_st = jan_store_data.get(sname, {})

        ws_st.cell(row_st, 1, sname).alignment = align_left

        # B: Qty Dec, C: Qty Jan, D: Total Qty
        ws_st.cell(row_st, 2, d_st.get('qty', 0.0)).number_format = fmt_qty
        ws_st.cell(row_st, 3, j_st.get('qty', 0.0)).number_format = fmt_qty
        ws_st.cell(row_st, 4, f"=B{row_st}+C{row_st}").number_format = fmt_qty

        # E: Rev Dec, F: Rev Jan, G: Total Rev
        ws_st.cell(row_st, 5, d_st.get('revenue', 0.0)).number_format = fmt_currency
        ws_st.cell(row_st, 6, j_st.get('revenue', 0.0)).number_format = fmt_currency
        ws_st.cell(row_st, 7, f"=E{row_st}+F{row_st}").number_format = fmt_currency

        # H: Rojen Net Dec, I: Rojen Net Jan, J: Total Rojen Net
        ws_st.cell(row_st, 8, d_st.get('rojen_net', 0.0)).number_format = fmt_currency
        ws_st.cell(row_st, 9, j_st.get('rojen_net', 0.0)).number_format = fmt_currency
        ws_st.cell(row_st, 10, f"=H{row_st}+I{row_st}").number_format = fmt_currency

        # K: Yasko Dec, L: Yasko Jan, M: Total Yasko
        ws_st.cell(row_st, 11, d_st.get('yasko', 0.0)).number_format = fmt_currency
        ws_st.cell(row_st, 12, j_st.get('yasko', 0.0)).number_format = fmt_currency
        ws_st.cell(row_st, 13, f"=K{row_st}+L{row_st}").number_format = fmt_currency

        # N: Savings Dec, O: Savings Jan, P: Total Savings
        ws_st.cell(row_st, 14, f"=H{row_st}-K{row_st}").number_format = fmt_currency_diff
        ws_st.cell(row_st, 15, f"=I{row_st}-L{row_st}").number_format = fmt_currency_diff
        c_tot_sav = ws_st.cell(row_st, 16, f"=J{row_st}-M{row_st}")
        c_tot_sav.number_format = fmt_currency_diff
        c_tot_sav.font = font_data_bold
        c_tot_sav.fill = fill_green_yasko

        # Q: Margin Yasko 2m % = (Total Rev - Total Yasko) / Total Rev
        ws_st.cell(row_st, 17, f"=IF(G{row_st}>0, (G{row_st}-M{row_st})/G{row_st}, 0)").number_format = fmt_pct

        for c_i in range(1, 18):
            cell_item = ws_st.cell(row_st, c_i)
            cell_item.border = border_cell
            if c_i > 1:
                cell_item.alignment = align_right
                if c_i != 16:
                    cell_item.font = font_data

        row_st += 1

    # Total Row for stores
    tot_st_r = row_st
    ws_st.row_dimensions[tot_st_r].height = 24
    ws_st.cell(tot_st_r, 1, "ОБЩО / TOTAL").font = font_total

    ws_st.cell(tot_st_r, 2, f"=SUM(B4:B{tot_st_r-1})").number_format = fmt_qty
    ws_st.cell(tot_st_r, 3, f"=SUM(C4:C{tot_st_r-1})").number_format = fmt_qty
    ws_st.cell(tot_st_r, 4, f"=SUM(D4:D{tot_st_r-1})").number_format = fmt_qty

    ws_st.cell(tot_st_r, 5, f"=SUM(E4:E{tot_st_r-1})").number_format = fmt_currency
    ws_st.cell(tot_st_r, 6, f"=SUM(F4:F{tot_st_r-1})").number_format = fmt_currency
    ws_st.cell(tot_st_r, 7, f"=SUM(G4:G{tot_st_r-1})").number_format = fmt_currency

    ws_st.cell(tot_st_r, 8, f"=SUM(H4:H{tot_st_r-1})").number_format = fmt_currency
    ws_st.cell(tot_st_r, 9, f"=SUM(I4:I{tot_st_r-1})").number_format = fmt_currency
    ws_st.cell(tot_st_r, 10, f"=SUM(J4:J{tot_st_r-1})").number_format = fmt_currency

    ws_st.cell(tot_st_r, 11, f"=SUM(K4:K{tot_st_r-1})").number_format = fmt_currency
    ws_st.cell(tot_st_r, 12, f"=SUM(L4:L{tot_st_r-1})").number_format = fmt_currency
    ws_st.cell(tot_st_r, 13, f"=SUM(M4:M{tot_st_r-1})").number_format = fmt_currency

    ws_st.cell(tot_st_r, 14, f"=SUM(N4:N{tot_st_r-1})").number_format = fmt_currency_diff
    ws_st.cell(tot_st_r, 15, f"=SUM(O4:O{tot_st_r-1})").number_format = fmt_currency_diff
    ws_st.cell(tot_st_r, 16, f"=SUM(P4:P{tot_st_r-1})").number_format = fmt_currency_diff

    ws_st.cell(tot_st_r, 17, f"=(G{tot_st_r}-M{tot_st_r})/G{tot_st_r}").number_format = fmt_pct

    for c_i in range(1, 18):
        c_t = ws_st.cell(tot_st_r, c_i)
        c_t.font = font_total
        c_t.border = border_total
        c_t.fill = fill_total
        if c_i > 1:
            c_t.alignment = align_right

    for col in ws_st.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row == 1:
                continue
            max_len = max(max_len, len(val_str))
        ws_st.column_dimensions[col_letter].width = max(max_len + 3, 13)
    ws_st.column_dimensions['A'].width = 24

    # Save consolidated file
    out_filename = 'СРАВНИТЕЛЕН_АНАЛИЗ_ЯСКО_РОЖЕН_ДЕКЕМВРИ_2025_И_ЯНУАРИ_2026.xlsx'
    try:
        wb.save(out_filename)
        print(f"Successfully generated: {out_filename}")
    except PermissionError:
        out_v2 = 'СРАВНИТЕЛЕН_АНАЛИЗ_ЯСКО_РОЖЕН_ДЕКЕМВРИ_2025_И_ЯНУАРИ_2026_АКТУАЛИЗИРАН.xlsx'
        wb.save(out_v2)
        print(f"File was locked in Excel. Successfully saved updated version as: {out_v2}")

if __name__ == '__main__':
    build_workbook()
