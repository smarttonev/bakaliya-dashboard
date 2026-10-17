# -*- coding: utf-8 -*-
"""
Генериране на заявки по обекти за верига „Бакалия“:
- Захарни заводи: ~1200 €
- Победа: ~1200 €
- Итал фуудс: ~2000 €
- Престиж: покритие за около 1.5 месеца спрямо наличностите и продажбите

Резултат: 1 обединен Excel файл с отделни таблици и обобщение.
"""

import os
import sys
import shutil
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

# Пътища към файловете
INPUT_FILES = {
    'Захарни заводи': {
        'path': r'C:\Users\Tonev\Downloads\Захарни заводи.xls',
        'target_eur': 1200.0,
        'k': 1.9195,
        'has_price': True,
        'desc': 'Цел ~1200 € (~58 дни покритие)'
    },
    'Победа': {
        'path': r'C:\Users\Tonev\Downloads\Победа.xls',
        'target_eur': 1200.0,
        'k': 0.8850,
        'has_price': True,
        'desc': 'Цел ~1200 € (~27 дни покритие)'
    },
    'Итал фуудс': {
        'path': r'C:\Users\Tonev\Downloads\Италфуудс .xls',
        'target_eur': 2000.0,
        'k': 1.9155,
        'has_price': True,
        'desc': 'Цел ~2000 € (~58 дни покритие)'
    },
    'Престиж': {
        'path': r'C:\Users\Tonev\Downloads\Престиж (2).xls',
        'target_eur': None,
        'k': 1.5000,
        'has_price': False,
        'desc': 'Цел 1.5 месеца покритие (45 дни)'
    }
}

STANDARD_STORES = [
    'АРСЕНАЛСКИ',
    'БАНКЯ',
    'ВЕЛЧО АТАНАСОВ',
    'ЛОЗЕНЕЦ',
    'ЛЮЛИН 7',
    'МАНАСТИРСКИ ЛИВАДИ',
    'НАДЕЖДА',
    'ОБЕЛЯ 2',
    'РЕДУТА',
    'СВОБОДА',
    'ХАЛИ ОКРЪЖНА'
]


def load_supplier_raw(path):
    df = pd.read_excel(path, header=None)
    headers = [i for i in range(len(df)) if df.shape[1] > 1 and 'АРСЕНАЛСКИ' in str(df.iloc[i, 1]).upper()]
    h_stock, h_sales = headers[0], headers[1]
    
    # Store names from header
    raw_stores = [str(c).strip() for c in df.iloc[h_stock, 1:12].tolist()]
    
    stock_dict = {}
    for i in range(h_stock + 1, h_sales):
        item = str(df.iloc[i, 0]).strip()
        if item.lower() == 'grand total' or item == 'nan':
            break
        vals = [pd.to_numeric(df.iloc[i, c], errors='coerce') for c in range(1, 12)]
        stock_dict[item] = [0.0 if np.isnan(v) else float(v) for v in vals]
        
    sales_dict = {}
    price_dict = {}
    for i in range(h_sales + 1, len(df)):
        item = str(df.iloc[i, 0]).strip()
        if item.lower() == 'grand total' or item == 'nan':
            break
        vals = [pd.to_numeric(df.iloc[i, c], errors='coerce') for c in range(1, 12)]
        sales_dict[item] = [0.0 if np.isnan(v) else float(v) for v in vals]
        if df.shape[1] > 13:
            price_dict[item] = float(pd.to_numeric(df.iloc[i, 13], errors='coerce'))
        else:
            price_dict[item] = np.nan
            
    return raw_stores, stock_dict, sales_dict, price_dict


def calculate_orders(stock_dict, sales_dict, price_dict, k):
    order_dict = {}
    for item, s_list in sales_dict.items():
        n_list = stock_dict.get(item, [0.0] * 11)
        store_orders = []
        for s, n in zip(s_list, n_list):
            if s > 0:
                eff_n = max(0.0, n)
                q = max(0, round(k * s - eff_n))
            else:
                q = 0
            store_orders.append(q)
        order_dict[item] = store_orders
    return order_dict


def build_workbook():
    wb = openpyxl.Workbook()
    # remove default sheet later or use it
    ws_summary = wb.active
    ws_summary.title = 'Обобщение'
    
    # Styles
    font_family = 'Segoe UI'
    
    title_font = Font(name=font_family, size=15, bold=True, color='1F4E79')
    subtitle_font = Font(name=font_family, size=10, italic=True, color='595959')
    section_font = Font(name=font_family, size=12, bold=True, color='1F4E79')
    
    th_font = Font(name=font_family, size=10, bold=True, color='FFFFFF')
    th_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    
    subth_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
    
    cell_font = Font(name=font_family, size=10)
    bold_cell_font = Font(name=font_family, size=10, bold=True)
    
    total_fill = PatternFill(start_color='EAEEF3', end_color='EAEEF3', fill_type='solid')
    zebra_fill = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
    accent_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    
    thin_side = Side(style='thin', color='D3D3D3')
    medium_bottom = Side(style='medium', color='1F4E79')
    double_bottom = Side(style='double', color='1F4E79')
    
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    th_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=medium_bottom)
    total_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)
    
    align_left = Alignment(horizontal='left', vertical='center')
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_th = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    num_fmt_int = '#,##0'
    num_fmt_curr = '#,##0.00 "€"'
    num_fmt_price = '0.00 "€"'
    
    # Process all suppliers
    suppliers_data = {}
    for name, cfg in INPUT_FILES.items():
        raw_stores, stock_dict, sales_dict, price_dict = load_supplier_raw(cfg['path'])
        order_dict = calculate_orders(stock_dict, sales_dict, price_dict, cfg['k'])
        
        # Totals
        tot_qty = sum(sum(q) for q in order_dict.values())
        tot_val = sum(sum(order_dict[it]) * price_dict[it] for it in order_dict) if cfg['has_price'] else None
        
        # Store sums
        store_qtys = [sum(order_dict[it][c] for it in order_dict) for c in range(11)]
        store_vals = [sum(order_dict[it][c] * price_dict[it] for it in order_dict) for c in range(11)] if cfg['has_price'] else None
        
        # Sales & stock sums
        tot_sales_qty = sum(sum(v) for v in sales_dict.values())
        tot_stock_qty = sum(sum(stock_dict.get(k, [0]*11)) for k in sales_dict)
        tot_sales_val = sum(sum(sales_dict[k]) * price_dict[k] for k in sales_dict) if cfg['has_price'] else None
        tot_stock_val = sum(sum(stock_dict.get(k, [0]*11)) * price_dict[k] for k in sales_dict) if cfg['has_price'] else None
        
        suppliers_data[name] = {
            'cfg': cfg,
            'raw_stores': raw_stores,
            'stock_dict': stock_dict,
            'sales_dict': sales_dict,
            'price_dict': price_dict,
            'order_dict': order_dict,
            'tot_qty': tot_qty,
            'tot_val': tot_val,
            'store_qtys': store_qtys,
            'store_vals': store_vals,
            'tot_sales_qty': tot_sales_qty,
            'tot_stock_qty': tot_stock_qty,
            'tot_sales_val': tot_sales_val,
            'tot_stock_val': tot_stock_val,
        }

    # =========================================================================
    # 1. SUMMARY SHEET
    # =========================================================================
    ws_summary.views.sheetView[0].showGridLines = True
    
    ws_summary.cell(row=1, column=1, value='ВЕРИГА ЗА ХРАНИТЕЛНИ СТОКИ „БАКАЛИЯ“').font = title_font
    ws_summary.cell(row=2, column=1, value='СВОДНА СПРАВКА: Оптимизирани заявки по търговски обекти на база наличности и продажби').font = subtitle_font
    ws_summary.cell(row=3, column=1, value='Анализиран базов период: 1 месец продажби | Валутен режим: 100% суми в ЕВРО (€) без превалутиране').font = subtitle_font
    
    # Table 1: KPI by supplier
    ws_summary.cell(row=5, column=1, value='1. Обобщение по доставчици и постигнати целеви параметри').font = section_font
    
    kpi_headers = [
        'Доставчик',
        'Целеви параметър',
        'Изчислена стойност (€)',
        'Общо бройки (бр.)',
        'Месечни продажби (€)',
        'Текуща наличност (€)',
        'Целево покритие (месеци)',
        'Целево покритие (дни)'
    ]
    
    for col_idx, h in enumerate(kpi_headers, 1):
        c = ws_summary.cell(row=6, column=col_idx, value=h)
        c.font = th_font
        c.fill = th_fill
        c.alignment = align_th
        c.border = th_border
    ws_summary.row_dimensions[6].height = 28
    
    r_idx = 7
    for name in ['Захарни заводи', 'Победа', 'Итал фуудс', 'Престиж']:
        d = suppliers_data[name]
        cfg = d['cfg']
        ws_summary.cell(row=r_idx, column=1, value=name).font = bold_cell_font
        ws_summary.cell(row=r_idx, column=2, value=cfg['desc']).font = cell_font
        
        c3 = ws_summary.cell(row=r_idx, column=3, value=d['tot_val'] if d['tot_val'] else '-')
        c3.font = bold_cell_font
        c3.alignment = align_right
        if d['tot_val']:
            c3.number_format = num_fmt_curr
            
        c4 = ws_summary.cell(row=r_idx, column=4, value=d['tot_qty'])
        c4.font = bold_cell_font
        c4.alignment = align_right
        c4.number_format = num_fmt_int
        
        c5 = ws_summary.cell(row=r_idx, column=5, value=d['tot_sales_val'] if d['tot_sales_val'] else f"{d['tot_sales_qty']:,} бр.")
        c5.font = cell_font
        c5.alignment = align_right
        if d['tot_sales_val']:
            c5.number_format = num_fmt_curr
            
        c6 = ws_summary.cell(row=r_idx, column=6, value=d['tot_stock_val'] if d['tot_stock_val'] else f"{d['tot_stock_qty']:,} бр.")
        c6.font = cell_font
        c6.alignment = align_right
        if d['tot_stock_val']:
            c6.number_format = num_fmt_curr
            
        c7 = ws_summary.cell(row=r_idx, column=7, value=f"{cfg['k']:.2f} мес.")
        c7.font = cell_font
        c7.alignment = align_center
        
        c8 = ws_summary.cell(row=r_idx, column=8, value=f"~{round(cfg['k']*30)} дни")
        c8.font = cell_font
        c8.alignment = align_center
        
        for c in range(1, 9):
            ws_summary.cell(row=r_idx, column=c).border = cell_border
        r_idx += 1
        
    # Total row for Table 1
    ws_summary.cell(row=r_idx, column=1, value='ОБЩО / TOTAL').font = bold_cell_font
    ws_summary.cell(row=r_idx, column=2, value='4 доставчика').font = bold_cell_font
    
    tot_val_cell = ws_summary.cell(row=r_idx, column=3, value=f'=SUM(C7:C9)')
    tot_val_cell.font = bold_cell_font
    tot_val_cell.alignment = align_right
    tot_val_cell.number_format = num_fmt_curr
    
    tot_qty_cell = ws_summary.cell(row=r_idx, column=4, value=f'=SUM(D7:D10)')
    tot_qty_cell.font = bold_cell_font
    tot_qty_cell.alignment = align_right
    tot_qty_cell.number_format = num_fmt_int
    
    ws_summary.cell(row=r_idx, column=5, value='-').alignment = align_center
    ws_summary.cell(row=r_idx, column=6, value='-').alignment = align_center
    ws_summary.cell(row=r_idx, column=7, value='-').alignment = align_center
    ws_summary.cell(row=r_idx, column=8, value='-').alignment = align_center
    
    for c in range(1, 9):
        cell = ws_summary.cell(row=r_idx, column=c)
        cell.border = total_border
        cell.fill = total_fill
    r_idx += 3
    
    # Table 2: Store breakdown
    ws_summary.cell(row=r_idx, column=1, value='2. Разпределение на заявените количества и стойности по търговски обекти').font = section_font
    r_idx += 1
    
    store_headers = [
        'Търговски обект (Магазин)',
        'Захарни зав. (€)',
        'Захарни зав. (бр.)',
        'Победа (€)',
        'Победа (бр.)',
        'Итал фуудс (€)',
        'Итал фуудс (бр.)',
        'Престиж (бр.)',
        'ОБЩО СТОЙНОСТ (€)',
        'ОБЩО БРОЙКИ (бр.)'
    ]
    
    for col_idx, h in enumerate(store_headers, 1):
        c = ws_summary.cell(row=r_idx, column=col_idx, value=h)
        c.font = th_font
        c.fill = th_fill
        c.alignment = align_th
        c.border = th_border
    ws_summary.row_dimensions[r_idx].height = 28
    
    start_store_row = r_idx + 1
    for s_idx, st_name in enumerate(STANDARD_STORES):
        curr_row = start_store_row + s_idx
        ws_summary.cell(row=curr_row, column=1, value=st_name).font = bold_cell_font
        
        # Захарни заводи
        c_zz_v = ws_summary.cell(row=curr_row, column=2, value=suppliers_data['Захарни заводи']['store_vals'][s_idx])
        c_zz_v.font = cell_font
        c_zz_v.alignment = align_right
        c_zz_v.number_format = num_fmt_curr
        
        c_zz_q = ws_summary.cell(row=curr_row, column=3, value=suppliers_data['Захарни заводи']['store_qtys'][s_idx])
        c_zz_q.font = cell_font
        c_zz_q.alignment = align_right
        c_zz_q.number_format = num_fmt_int
        
        # Победа
        c_pb_v = ws_summary.cell(row=curr_row, column=4, value=suppliers_data['Победа']['store_vals'][s_idx])
        c_pb_v.font = cell_font
        c_pb_v.alignment = align_right
        c_pb_v.number_format = num_fmt_curr
        
        c_pb_q = ws_summary.cell(row=curr_row, column=5, value=suppliers_data['Победа']['store_qtys'][s_idx])
        c_pb_q.font = cell_font
        c_pb_q.alignment = align_right
        c_pb_q.number_format = num_fmt_int
        
        # Итал фуудс
        c_if_v = ws_summary.cell(row=curr_row, column=6, value=suppliers_data['Итал фуудс']['store_vals'][s_idx])
        c_if_v.font = cell_font
        c_if_v.alignment = align_right
        c_if_v.number_format = num_fmt_curr
        
        c_if_q = ws_summary.cell(row=curr_row, column=7, value=suppliers_data['Итал фуудс']['store_qtys'][s_idx])
        c_if_q.font = cell_font
        c_if_q.alignment = align_right
        c_if_q.number_format = num_fmt_int
        
        # Престиж
        c_pr_q = ws_summary.cell(row=curr_row, column=8, value=suppliers_data['Престиж']['store_qtys'][s_idx])
        c_pr_q.font = cell_font
        c_pr_q.alignment = align_right
        c_pr_q.number_format = num_fmt_int
        
        # Total € per store = ZZ + PB + IF
        c_tot_v = ws_summary.cell(row=curr_row, column=9, value=f'=B{curr_row}+D{curr_row}+F{curr_row}')
        c_tot_v.font = bold_cell_font
        c_tot_v.alignment = align_right
        c_tot_v.number_format = num_fmt_curr
        c_tot_v.fill = accent_fill
        
        # Total Qty per store = ZZ + PB + IF + PR
        c_tot_q = ws_summary.cell(row=curr_row, column=10, value=f'=C{curr_row}+E{curr_row}+G{curr_row}+H{curr_row}')
        c_tot_q.font = bold_cell_font
        c_tot_q.alignment = align_right
        c_tot_q.number_format = num_fmt_int
        c_tot_q.fill = accent_fill
        
        for c in range(1, 11):
            cell = ws_summary.cell(row=curr_row, column=c)
            cell.border = cell_border
            if s_idx % 2 == 1 and c not in (9, 10):
                cell.fill = zebra_fill
                
    end_store_row = start_store_row + len(STANDARD_STORES) - 1
    tot_store_row = end_store_row + 1
    
    ws_summary.cell(row=tot_store_row, column=1, value='ОБЩО / TOTAL').font = bold_cell_font
    
    for c_idx, letter in enumerate(['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'], 2):
        cell = ws_summary.cell(row=tot_store_row, column=c_idx, value=f'=SUM({letter}{start_store_row}:{letter}{end_store_row})')
        cell.font = bold_cell_font
        cell.alignment = align_right
        if letter in ['B', 'D', 'F', 'I']:
            cell.number_format = num_fmt_curr
        else:
            cell.number_format = num_fmt_int
            
    for c in range(1, 11):
        cell = ws_summary.cell(row=tot_store_row, column=c)
        cell.border = total_border
        cell.fill = total_fill

    # =========================================================================
    # 2. INDIVIDUAL SUPPLIER SHEETS
    # =========================================================================
    for name in ['Захарни заводи', 'Победа', 'Итал фуудс', 'Престиж']:
        ws = wb.create_sheet(title=name)
        ws.views.sheetView[0].showGridLines = True
        
        d = suppliers_data[name]
        cfg = d['cfg']
        has_price = cfg['has_price']
        
        # Title block
        ws.cell(row=1, column=1, value=f'ЗАЯВКА ПО ОБЕКТИ: {name.upper()}').font = title_font
        val_txt = f"Целева стойност: {cfg['target_eur']:,.2f} € | Изчислена: {d['tot_val']:,.2f} € | Покритие: {cfg['k']:.2f} мес." if has_price else f"Целево покритие: {cfg['k']:.2f} месеца спрямо продажби и наличности | Общо бройки: {d['tot_qty']:,} бр."
        ws.cell(row=2, column=1, value=val_txt).font = subtitle_font
        
        # Consistent headers across all sheets
        headers = ['№', 'Артикул / Номенклатура', 'Доставна цена (€)'] + STANDARD_STORES + ['ОБЩО (бр.)', 'ОБЩА СТОЙНОСТ (€)']
            
        header_row = 4
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=header_row, column=col_idx, value=h)
            c.font = th_font
            c.fill = th_fill
            c.alignment = align_th
            c.border = th_border
        ws.row_dimensions[header_row].height = 28
        
        # Data rows
        items = list(d['sales_dict'].keys())
        items.sort()
        
        first_store_col = 4
        last_store_col = 14
        total_qty_col = 15
        total_val_col = 16
        
        curr_row = 5
        for it_idx, item in enumerate(items, 1):
            ws.cell(row=curr_row, column=1, value=it_idx).font = cell_font
            ws.cell(row=curr_row, column=1).alignment = align_center
            
            ws.cell(row=curr_row, column=2, value=item).font = cell_font
            ws.cell(row=curr_row, column=2).alignment = align_left
            
            if has_price:
                price = d['price_dict'][item]
                cp = ws.cell(row=curr_row, column=3, value=price)
                cp.font = cell_font
                cp.alignment = align_right
                cp.number_format = num_fmt_price
            else:
                cp = ws.cell(row=curr_row, column=3, value='-')
                cp.font = cell_font
                cp.alignment = align_center
                
            orders = d['order_dict'][item]
            for s_idx, q in enumerate(orders):
                col_c = first_store_col + s_idx
                cq = ws.cell(row=curr_row, column=col_c, value=q)
                cq.font = bold_cell_font if q > 0 else cell_font
                cq.alignment = align_right
                cq.number_format = num_fmt_int
                if q == 0:
                    cq.font = Font(name=font_family, size=10, color='B0B0B0')
                    
            # Total Qty formula
            c_start = get_column_letter(first_store_col)
            c_end = get_column_letter(last_store_col)
            cq_tot = ws.cell(row=curr_row, column=total_qty_col, value=f'=SUM({c_start}{curr_row}:{c_end}{curr_row})')
            cq_tot.font = bold_cell_font
            cq_tot.alignment = align_right
            cq_tot.number_format = num_fmt_int
            cq_tot.fill = accent_fill
            
            # Total Value formula
            if has_price:
                c_p_let = get_column_letter(3)
                c_q_let = get_column_letter(total_qty_col)
                cv_tot = ws.cell(row=curr_row, column=total_val_col, value=f'={c_q_let}{curr_row}*{c_p_let}{curr_row}')
                cv_tot.font = bold_cell_font
                cv_tot.alignment = align_right
                cv_tot.number_format = num_fmt_curr
                cv_tot.fill = accent_fill
            else:
                cv_tot = ws.cell(row=curr_row, column=total_val_col, value='-')
                cv_tot.font = cell_font
                cv_tot.alignment = align_center
                
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = cell_border
                if it_idx % 2 == 0 and c not in (total_qty_col, total_val_col):
                    cell.fill = zebra_fill
                    
            curr_row += 1
            
        last_data_row = curr_row - 1
        
        # Summary Row 1: Total Qty per column
        ws.cell(row=curr_row, column=1, value='').font = bold_cell_font
        ws.cell(row=curr_row, column=2, value='ОБЩО КОЛИЧЕСТВО (бр.)').font = bold_cell_font
        ws.cell(row=curr_row, column=3, value='-').alignment = align_center
            
        for s_col in range(first_store_col, last_store_col + 1):
            col_let = get_column_letter(s_col)
            cq = ws.cell(row=curr_row, column=s_col, value=f'=SUM({col_let}5:{col_let}{last_data_row})')
            cq.font = bold_cell_font
            cq.alignment = align_right
            cq.number_format = num_fmt_int
            
        # Grand Total Qty
        q_tot_let = get_column_letter(total_qty_col)
        cgq = ws.cell(row=curr_row, column=total_qty_col, value=f'=SUM({q_tot_let}5:{q_tot_let}{last_data_row})')
        cgq.font = bold_cell_font
        cgq.alignment = align_right
        cgq.number_format = num_fmt_int
        
        # Grand Total Value
        if has_price:
            v_tot_let = get_column_letter(total_val_col)
            cgv = ws.cell(row=curr_row, column=total_val_col, value=f'=SUM({v_tot_let}5:{v_tot_let}{last_data_row})')
            cgv.font = bold_cell_font
            cgv.alignment = align_right
            cgv.number_format = num_fmt_curr
        else:
            ws.cell(row=curr_row, column=total_val_col, value='-').alignment = align_center
            
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=curr_row, column=c)
            cell.border = total_border
            cell.fill = total_fill
            
        # If has price, add Summary Row 2: Total Value per store
        if has_price:
            curr_row += 1
            ws.cell(row=curr_row, column=1, value='').font = bold_cell_font
            ws.cell(row=curr_row, column=2, value='ОБЩА СТОЙНОСТ (€)').font = bold_cell_font
            ws.cell(row=curr_row, column=3, value='-').alignment = align_center
            
            p_let = get_column_letter(3)
            for s_col in range(first_store_col, last_store_col + 1):
                col_let = get_column_letter(s_col)
                cv = ws.cell(row=curr_row, column=s_col, value=f'=SUMPRODUCT(${p_let}$5:${p_let}${last_data_row}, {col_let}$5:{col_let}${last_data_row})')
                cv.font = bold_cell_font
                cv.alignment = align_right
                cv.number_format = num_fmt_curr
                
            ws.cell(row=curr_row, column=total_qty_col, value='-').alignment = align_center
            cgv2 = ws.cell(row=curr_row, column=total_val_col, value=f'={get_column_letter(total_val_col)}{curr_row - 1}')
            cgv2.font = bold_cell_font
            cgv2.alignment = align_right
            cgv2.number_format = num_fmt_curr
            
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=curr_row, column=c)
                cell.border = total_border
                cell.fill = total_fill

    # =========================================================================
    # 3. COMBINED SHEET: ALL 4 TABLES IN 1 SHEET
    # =========================================================================
    ws_all = wb.create_sheet(title='Всички заявки')
    ws_all.views.sheetView[0].showGridLines = True
    
    ws_all.cell(row=1, column=1, value='ВЕРИГА „БАКАЛИЯ“ — ВСИЧКИ ЗАЯВКИ ПО ОБЕКТИ').font = title_font
    ws_all.cell(row=2, column=1, value='Сборен работен лист с пълните таблици за 4-те доставчика').font = subtitle_font
    
    comb_row = 4
    for name in ['Захарни заводи', 'Победа', 'Итал фуудс', 'Престиж']:
        d = suppliers_data[name]
        cfg = d['cfg']
        has_price = cfg['has_price']
        
        # Section title
        ws_all.cell(row=comb_row, column=1, value=f'ТАБЛИЦА: {name.upper()} ({cfg["desc"]})').font = section_font
        comb_row += 1
        
        headers = ['№', 'Артикул / Номенклатура', 'Доставна цена (€)'] + STANDARD_STORES + ['ОБЩО (бр.)', 'ОБЩА СТОЙНОСТ (€)']
            
        header_row = comb_row
        for col_idx, h in enumerate(headers, 1):
            c = ws_all.cell(row=header_row, column=col_idx, value=h)
            c.font = th_font
            c.fill = subth_fill
            c.alignment = align_th
            c.border = th_border
        ws_all.row_dimensions[header_row].height = 26
        comb_row += 1
        
        first_store_col = 4
        last_store_col = 14
        total_qty_col = 15
        total_val_col = 16
        
        items = list(d['sales_dict'].keys())
        items.sort()
        start_data_row = comb_row
        
        for it_idx, item in enumerate(items, 1):
            ws_all.cell(row=comb_row, column=1, value=it_idx).font = cell_font
            ws_all.cell(row=comb_row, column=1).alignment = align_center
            ws_all.cell(row=comb_row, column=2, value=item).font = cell_font
            
            if has_price:
                price = d['price_dict'][item]
                cp = ws_all.cell(row=comb_row, column=3, value=price)
                cp.font = cell_font
                cp.alignment = align_right
                cp.number_format = num_fmt_price
            else:
                cp = ws_all.cell(row=comb_row, column=3, value='-')
                cp.font = cell_font
                cp.alignment = align_center
                
            orders = d['order_dict'][item]
            for s_idx, q in enumerate(orders):
                col_c = first_store_col + s_idx
                cq = ws_all.cell(row=comb_row, column=col_c, value=q)
                cq.font = bold_cell_font if q > 0 else cell_font
                cq.alignment = align_right
                cq.number_format = num_fmt_int
                if q == 0:
                    cq.font = Font(name=font_family, size=10, color='B0B0B0')
                    
            c_start = get_column_letter(first_store_col)
            c_end = get_column_letter(last_store_col)
            cq_tot = ws_all.cell(row=comb_row, column=total_qty_col, value=f'=SUM({c_start}{comb_row}:{c_end}{comb_row})')
            cq_tot.font = bold_cell_font
            cq_tot.alignment = align_right
            cq_tot.number_format = num_fmt_int
            cq_tot.fill = accent_fill
            
            if has_price:
                c_p_let = get_column_letter(3)
                c_q_let = get_column_letter(total_qty_col)
                cv_tot = ws_all.cell(row=comb_row, column=total_val_col, value=f'={c_q_let}{comb_row}*{c_p_let}{comb_row}')
                cv_tot.font = bold_cell_font
                cv_tot.alignment = align_right
                cv_tot.number_format = num_fmt_curr
                cv_tot.fill = accent_fill
            else:
                cv_tot = ws_all.cell(row=comb_row, column=total_val_col, value='-')
                cv_tot.font = cell_font
                cv_tot.alignment = align_center
                
            for c in range(1, len(headers) + 1):
                cell = ws_all.cell(row=comb_row, column=c)
                cell.border = cell_border
                if it_idx % 2 == 0 and c not in (total_qty_col, total_val_col):
                    cell.fill = zebra_fill
            comb_row += 1
            
        last_data_row = comb_row - 1
        
        # Total Qty row
        ws_all.cell(row=comb_row, column=1, value='').font = bold_cell_font
        ws_all.cell(row=comb_row, column=2, value='ОБЩО (бр.)').font = bold_cell_font
        ws_all.cell(row=comb_row, column=3, value='-').alignment = align_center
            
        for s_col in range(first_store_col, last_store_col + 1):
            col_let = get_column_letter(s_col)
            cq = ws_all.cell(row=comb_row, column=s_col, value=f'=SUM({col_let}{start_data_row}:{col_let}{last_data_row})')
            cq.font = bold_cell_font
            cq.alignment = align_right
            cq.number_format = num_fmt_int
            
        q_tot_let = get_column_letter(total_qty_col)
        cgq = ws_all.cell(row=comb_row, column=total_qty_col, value=f'=SUM({q_tot_let}{start_data_row}:{q_tot_let}{last_data_row})')
        cgq.font = bold_cell_font
        cgq.alignment = align_right
        cgq.number_format = num_fmt_int
        
        if has_price:
            v_tot_let = get_column_letter(total_val_col)
            cgv = ws_all.cell(row=comb_row, column=total_val_col, value=f'=SUM({v_tot_let}{start_data_row}:{v_tot_let}{last_data_row})')
            cgv.font = bold_cell_font
            cgv.alignment = align_right
            cgv.number_format = num_fmt_curr
        else:
            ws_all.cell(row=comb_row, column=total_val_col, value='-').alignment = align_center
            
        for c in range(1, len(headers) + 1):
            cell = ws_all.cell(row=comb_row, column=c)
            cell.border = total_border
            cell.fill = total_fill
            
        comb_row += 3  # spacing between tables

    # =========================================================================
    # AUTO-FIT COLUMN WIDTHS ACROSS ALL SHEETS
    # =========================================================================
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None:
                    # ignore long title in row 1-3 for width
                    if cell.row in [1, 2, 3] and col[0].column == 1:
                        continue
                    # remove formula prefix if formula
                    txt = str(val)
                    if txt.startswith('='):
                        # estimate formatted len
                        txt = "123,456.78 €"
                    max_len = max(max_len, len(txt))
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 11)
            
        # specifically ensure column B (Артикул) is wide enough
        if 'B' in sheet.column_dimensions:
            sheet.column_dimensions['B'].width = max(sheet.column_dimensions['B'].width, 38)
        if 'A' in sheet.column_dimensions:
            sheet.column_dimensions['A'].width = max(sheet.column_dimensions['A'].width, 6)

    # Save to workspace
    out_workspace = r'g:\My Drive\AntiG_test\Заявки_по_обекти.xlsx'
    wb.save(out_workspace)
    print(f'Successfully saved: {out_workspace}')
    
    # Save copy to Downloads
    out_downloads = r'C:\Users\Tonev\Downloads\Заявки_по_обекти.xlsx'
    shutil.copyfile(out_workspace, out_downloads)
    print(f'Successfully copied: {out_downloads}')

if __name__ == '__main__':
    build_workbook()
