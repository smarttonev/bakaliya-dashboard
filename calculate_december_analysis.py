# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
import firebirdsql
import time

# 1. Load price list
p1 = r'C:\Users\Tonev\.gemini\antigravity\brain\d872b9fb-1db0-4f76-abd7-263b17bbdcb9\.user_uploaded\media_1791894563117_a1485000.xlsx'
wb1 = openpyxl.load_workbook(p1)
ws1 = wb1.active
data1 = list(ws1.iter_rows(values_only=True))

price_items = []
codes_set = set()
for idx, r in enumerate(data1[1:], start=2):
    if not any(r):
        continue
    weight, name, p_rojen, p_yasko, art_no = r[:5]
    if art_no is not None:
        art_clean = str(art_no).strip().rstrip(',').strip()
        is_chipita = ('7D' in str(name).upper() or '7ДЕЙ' in str(name).upper() or 
                      'БЕЙК' in str(name).upper() or 'КРОАСАН' in str(name).upper() or idx >= 102)
        brand = 'CHIPITA' if is_chipita else 'MDLZ'
        price_items.append({
            'row_idx': idx,
            'weight': weight,
            'name_pl': str(name).strip(),
            'p_rojen_base': float(p_rojen) if p_rojen is not None else 0.0,
            'p_yasko': float(p_yasko) if p_yasko is not None else 0.0,
            'art_no': art_clean,
            'brand': brand
        })
        if art_clean.isdigit():
            codes_set.add(int(art_clean))

df_pl = pd.DataFrame(price_items)
pl_distinct = df_pl.drop_duplicates(subset=['art_no']).copy()
code_list_str = ','.join(str(c) for c in codes_set)

# 2. Query December 2025 sales from CS.FDB
CS_CONFIG = {
    'host': '192.168.5.254',
    'database': 'E:/db/CS.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

conn = firebirdsql.connect(**CS_CONFIG)
cur = conn.cursor()

# Detailed query by material and store
sql_dec = f'''
    SELECT 
        sc.MATERIALCODE,
        m.MATERIAL,
        ln.LOCATION,
        SUM(sc.QTY),
        SUM(sc.QTY * sc.SALEPRICE),
        SUM(sc.QTY * sc.LASTDELIVERYPRICE),
        SUM(sc.QTY * (sc.SALEPRICE - sc.LASTDELIVERYPRICE))
    FROM SALECONTENT sc
    JOIN SALE s ON s.NUM = sc.NUM AND s.LOCATIONID = sc.LOCATIONID
    JOIN LOCATIONNAME ln ON ln.ID = s.LOCATIONNAMEID
    LEFT JOIN MATERIAL m ON m.MATERIALCODE = sc.MATERIALCODE AND m.LOCATIONID = s.LOCATIONNAMEID
    WHERE sc.MATERIALCODE IN ({code_list_str})
      AND s.SALEDATE >= '2025-12-01' AND s.SALEDATE < '2026-01-01'
    GROUP BY sc.MATERIALCODE, m.MATERIAL, ln.LOCATION
'''
cur.execute(sql_dec)
raw_sales = cur.fetchall()
conn.close()

BGN_EUR_RATE = 1.95583

sales_records = []
for r in raw_sales:
    sales_records.append({
        'art_no': str(r[0]),
        'mat_name': r[1] if r[1] else '',
        'store': r[2] if r[2] else '',
        'qty': float(r[3]) if r[3] else 0.0,
        'revenue': (float(r[4]) / BGN_EUR_RATE) if r[4] else 0.0,
        'cost_erp': (float(r[5]) / BGN_EUR_RATE) if r[5] else 0.0,
        'profit_erp': (float(r[6]) / BGN_EUR_RATE) if r[6] else 0.0
    })

df_sales_dec = pd.DataFrame(sales_records)

# Aggregate by art_no
art_summary = df_sales_dec.groupby('art_no').agg({
    'mat_name': 'first',
    'qty': 'sum',
    'revenue': 'sum',
    'cost_erp': 'sum',
    'profit_erp': 'sum'
}).reset_index()

merged_dec = pd.merge(art_summary, pl_distinct, on='art_no', how='inner')

# Calculations for December 2025
merged_dec['cost_rojen_base'] = merged_dec['qty'] * merged_dec['p_rojen_base']

# Discounts: MDLZ 6.8%, Chipita 4.0%
merged_dec['disc_pct_rojen'] = merged_dec['brand'].map({'MDLZ': 0.068, 'CHIPITA': 0.040})
merged_dec['p_rojen_net_inv'] = merged_dec['p_rojen_base'] * (1 - merged_dec['disc_pct_rojen'])
merged_dec['cost_rojen_net_inv'] = merged_dec['qty'] * merged_dec['p_rojen_net_inv']

# Rozhen with flat 6.8%
merged_dec['p_rojen_net_68'] = merged_dec['p_rojen_base'] * (1 - 0.068)
merged_dec['cost_rojen_net_68'] = merged_dec['qty'] * merged_dec['p_rojen_net_68']

# Yasko
merged_dec['cost_yasko'] = merged_dec['qty'] * merged_dec['p_yasko']

# Profits
merged_dec['profit_rojen_base'] = merged_dec['revenue'] - merged_dec['cost_rojen_base']
merged_dec['profit_rojen_net_inv'] = merged_dec['revenue'] - merged_dec['cost_rojen_net_inv']
merged_dec['profit_rojen_net_68'] = merged_dec['revenue'] - merged_dec['cost_rojen_net_68']
merged_dec['profit_yasko'] = merged_dec['revenue'] - merged_dec['cost_yasko']

merged_dec['diff_cost_base'] = merged_dec['cost_rojen_base'] - merged_dec['cost_yasko']
merged_dec['diff_cost_net'] = merged_dec['cost_rojen_net_inv'] - merged_dec['cost_yasko']

# Store-level analysis
store_merged = pd.merge(df_sales_dec, pl_distinct[['art_no', 'brand', 'p_rojen_base', 'p_yasko']], on='art_no', how='inner')
store_merged['disc_pct_rojen'] = store_merged['brand'].map({'MDLZ': 0.068, 'CHIPITA': 0.040})
store_merged['p_rojen_net_inv'] = store_merged['p_rojen_base'] * (1 - store_merged['disc_pct_rojen'])

store_merged['cost_rojen_base'] = store_merged['qty'] * store_merged['p_rojen_base']
store_merged['cost_rojen_net_inv'] = store_merged['qty'] * store_merged['p_rojen_net_inv']
store_merged['cost_yasko'] = store_merged['qty'] * store_merged['p_yasko']

store_summary = store_merged.groupby('store').agg({
    'qty': 'sum',
    'revenue': 'sum',
    'cost_erp': 'sum',
    'profit_erp': 'sum',
    'cost_rojen_base': 'sum',
    'cost_rojen_net_inv': 'sum',
    'cost_yasko': 'sum'
}).reset_index()

store_summary['profit_rojen_base'] = store_summary['revenue'] - store_summary['cost_rojen_base']
store_summary['profit_rojen_net'] = store_summary['revenue'] - store_summary['cost_rojen_net_inv']
store_summary['profit_yasko'] = store_summary['revenue'] - store_summary['cost_yasko']
store_summary['diff_cost_base'] = store_summary['cost_rojen_base'] - store_summary['cost_yasko']
store_summary['diff_cost_net'] = store_summary['cost_rojen_net_inv'] - store_summary['cost_yasko']
store_summary['margin_rojen_net_%'] = (store_summary['profit_rojen_net'] / store_summary['revenue']) * 100
store_summary['margin_yasko_%'] = (store_summary['profit_yasko'] / store_summary['revenue']) * 100
store_summary['margin_erp_%'] = (store_summary['profit_erp'] / store_summary['revenue']) * 100

# Brand summary
brand_summary = merged_dec.groupby('brand').agg({
    'qty': 'sum',
    'revenue': 'sum',
    'cost_erp': 'sum',
    'profit_erp': 'sum',
    'cost_rojen_base': 'sum',
    'cost_rojen_net_inv': 'sum',
    'cost_yasko': 'sum',
    'profit_rojen_base': 'sum',
    'profit_rojen_net_inv': 'sum',
    'profit_yasko': 'sum'
}).reset_index()

brand_summary['diff_cost_base'] = brand_summary['cost_rojen_base'] - brand_summary['cost_yasko']
brand_summary['diff_cost_net'] = brand_summary['cost_rojen_net_inv'] - brand_summary['cost_yasko']
brand_summary['margin_rojen_net_%'] = (brand_summary['profit_rojen_net_inv'] / brand_summary['revenue']) * 100
brand_summary['margin_yasko_%'] = (brand_summary['profit_yasko'] / brand_summary['revenue']) * 100
brand_summary['margin_erp_%'] = (brand_summary['profit_erp'] / brand_summary['revenue']) * 100

# Write report
with open('december_analysis_report.txt', 'w', encoding='utf-8') as out:
    tot_qty = merged_dec['qty'].sum()
    tot_rev = merged_dec['revenue'].sum()
    tot_cost_erp = merged_dec['cost_erp'].sum()
    tot_prof_erp = merged_dec['profit_erp'].sum()
    tot_mar_erp = (tot_prof_erp / tot_rev) * 100
    
    cost_rb = merged_dec['cost_rojen_base'].sum()
    cost_rn = merged_dec['cost_rojen_net_inv'].sum()
    cost_rn68 = merged_dec['cost_rojen_net_68'].sum()
    cost_y = merged_dec['cost_yasko'].sum()
    
    prof_rb = merged_dec['profit_rojen_base'].sum()
    prof_rn = merged_dec['profit_rojen_net_inv'].sum()
    prof_rn68 = merged_dec['profit_rojen_net_68'].sum()
    prof_y = merged_dec['profit_yasko'].sum()
    
    mar_rb = (prof_rb / tot_rev) * 100
    mar_rn = (prof_rn / tot_rev) * 100
    mar_rn68 = (prof_rn68 / tot_rev) * 100
    mar_y = (prof_y / tot_rev) * 100
    
    diff_base = cost_rb - cost_y
    diff_net = cost_rn - cost_y
    
    out.write('=== ОБЩ СУМАРЕН АНАЛИЗ ЗА ДЕКЕМВРИ 2025 ===\n')
    out.write(f'Артикули с продажби: {len(merged_dec)} бр. (от 104 в листата)\n')
    out.write(f'Общо продадени бройки: {tot_qty:,.2f} бр.\n')
    out.write(f'Общ реализиран оборот (Сума с ДДС): {tot_rev:,.2f} €\n')
    out.write(f'Реална себестойност в ERP: {tot_cost_erp:,.2f} €\n')
    out.write(f'Реална печалба в ERP: {tot_prof_erp:,.2f} € (Марж: {tot_mar_erp:.2f}%)\n')
    out.write('\n--- СЦЕНАРИЙ 1: РОЖЕН БАЗОВА ЦЕНА С ДДС ---\n')
    out.write(f'Покупна стойност: {cost_rb:,.2f} €\n')
    out.write(f'Печалба при тези продажби: {prof_rb:,.2f} €\n')
    out.write(f'Търговски марж: {mar_rb:.2f}%\n')
    out.write('\n--- СЦЕНАРИЙ 2: РОЖЕН НЕТО С ТО НА ФАКТУРА (-6.8% MDLZ / -4.0% CHIP) ---\n')
    out.write(f'Покупна стойност: {cost_rn:,.2f} €\n')
    out.write(f'Печалба при тези продажби: {prof_rn:,.2f} €\n')
    out.write(f'Търговски марж: {mar_rn:.2f}%\n')
    out.write('\n--- СЦЕНАРИЙ 2Б: РОЖЕН НЕТО С -6.8% ЗА ВСИЧКИ ---\n')
    out.write(f'Покупна стойност: {cost_rn68:,.2f} €\n')
    out.write(f'Печалба: {prof_rn68:,.2f} € (Марж: {mar_rn68:.2f}%)\n')
    out.write('\n--- СЦЕНАРИЙ 3: ЯСКО (ЦЕНИ С ДДС) ---\n')
    out.write(f'Покупна стойност: {cost_y:,.2f} €\n')
    out.write(f'Печалба при тези продажби: {prof_y:,.2f} €\n')
    out.write(f'Търговски марж: {mar_y:.2f}%\n')
    out.write('\n=== РАЗЛИКИ И ИКОНОМИЧЕСКИ ЕФЕКТ (ДЕКЕМВРИ 2025) ===\n')
    out.write(f'Разлика в покупна стойност (Рожен Базова vs Яско): {diff_base:+,.2f} € (Яско е по-изгоден с {diff_base:,.2f} € или {diff_base/cost_rb*100:.2f}%)\n')
    out.write(f'Разлика в покупна стойност (Рожен Нето фактура vs Яско): {diff_net:+,.2f} € (Яско е по-изгоден с {diff_net:,.2f} € или {diff_net/cost_rn*100:.2f}%)\n')
    out.write(f'Разлика в печалбата (Яско vs Рожен Базова): {prof_y - prof_rb:+,.2f} € (Марж: +{mar_y - mar_rb:.2f} процентни пункта)\n')
    out.write(f'Разлика в печалбата (Яско vs Рожен Нето): {prof_y - prof_rn:+,.2f} € (Марж: +{mar_y - mar_rn:.2f} процентни пункта)\n')
    out.write(f'Разлика в печалбата (Яско vs Текуща реална в ERP): {prof_y - tot_prof_erp:+,.2f} € (Марж: +{mar_y - tot_mar_erp:.2f} процентни пункта)\n')
    
    out.write('\n\n=== АНАЛИЗ ПО БРАНДОВЕ ЗА ДЕКЕМВРИ 2025 ===\n')
    for _, br in brand_summary.iterrows():
        b_name = br['brand']
        b_qty = br['qty']
        b_rev = br['revenue']
        b_cerp = br['cost_erp']
        b_perp = br['profit_erp']
        b_merp = br['margin_erp_%']
        b_crb = br['cost_rojen_base']
        b_crn = br['cost_rojen_net_inv']
        b_cy = br['cost_yasko']
        b_prb = br['profit_rojen_base']
        b_prn = br['profit_rojen_net_inv']
        b_py = br['profit_yasko']
        out.write(f'Бранд: {b_name}\n')
        out.write(f'  Продадени бройки: {b_qty:,.2f} бр.\n')
        out.write(f'  Приход (Сума): {b_rev:,.2f} €\n')
        out.write(f'  Реална себестойност ERP: {b_cerp:,.2f} € (Печалба: {b_perp:,.2f} €, Марж: {b_merp:.2f}%)\n')
        out.write(f'  Покупна Рожен Базова: {b_crb:,.2f} € (Печалба: {b_prb:,.2f} €, Марж: {b_prb/b_rev*100:.2f}%)\n')
        out.write(f'  Покупна Рожен Нето фактура: {b_crn:,.2f} € (Печалба: {b_prn:,.2f} €, Марж: {b_prn/b_rev*100:.2f}%)\n')
        out.write(f'  Покупна Яско: {b_cy:,.2f} € (Печалба: {b_py:,.2f} €, Марж: {b_py/b_rev*100:.2f}%)\n')
        out.write(f'  Икономия от Яско спрямо Рожен Базова: {b_crb - b_cy:,.2f} €\n')
        out.write(f'  Икономия от Яско спрямо Рожен Нето: {b_crn - b_cy:,.2f} € (Марж: +{(b_py/b_rev - b_prn/b_rev)*100:.2f}%)\n\n')

    out.write('\n=== АНАЛИЗ ПО ОБЕКТИ (МАГАЗИНИ) ЗА ДЕКЕМВРИ 2025 ===\n')
    out.write('Обект\tК-во\tПриход\tРеална Печалба\tРеален Марж %\tРожен Базова\tРожен Нето\tЯско\tИкономия Яско (спр. Нето)\tМарж Рожен Нето %\tМарж Яско %\n')
    for _, st in store_summary.sort_values(by='qty', ascending=False).iterrows():
        s_name = st['store']
        s_qty = st['qty']
        s_rev = st['revenue']
        s_perp = st['profit_erp']
        s_merp = st['margin_erp_%']
        s_crb = st['cost_rojen_base']
        s_crn = st['cost_rojen_net_inv']
        s_cy = st['cost_yasko']
        s_diff = st['diff_cost_net']
        s_mrn = st['margin_rojen_net_%']
        s_my = st['margin_yasko_%']
        out.write(f'{s_name}\t{s_qty:,.1f}\t{s_rev:,.2f}\t{s_perp:,.2f}\t{s_merp:.2f}%\t{s_crb:,.2f}\t{s_crn:,.2f}\t{s_cy:,.2f}\t{s_diff:,.2f}\t{s_mrn:.2f}%\t{s_my:.2f}%\n')

    out.write('\n=== ТОП 20 АРТИКУЛА ПО ИКОНОМИЯ ОТ ЯСКО (ДЕКЕМВРИ 2025) ===\n')
    merged_sorted = merged_dec.sort_values(by='diff_cost_net', ascending=False)
    out.write('Мат №\tАртикул\tБранд\tК-во\tОборот €\tРожен Нето\tЯско\tОбщо Рожен Нето\tОбщо Яско\tИкономия Яско\tМарж Рожен Нето %\tМарж Яско %\n')
    for _, a in merged_sorted.head(20).iterrows():
        a_id = a['art_no']
        a_name = a['name_pl']
        a_brand = a['brand']
        a_qty = a['qty']
        a_rev = a['revenue']
        a_rn = a['p_rojen_net_inv']
        a_y = a['p_yasko']
        a_tot_rn = a['cost_rojen_net_inv']
        a_tot_y = a['cost_yasko']
        a_sav = a['diff_cost_net']
        a_mrn = ((a_rev - a_tot_rn) / a_rev) * 100 if a_rev > 0 else 0
        a_my = ((a_rev - a_tot_y) / a_rev) * 100 if a_rev > 0 else 0
        out.write(f'{a_id}\t{a_name}\t{a_brand}\t{a_qty:,.0f}\t{a_rev:,.2f}\t{a_rn:.3f}\t{a_y:.3f}\t{a_tot_rn:,.2f}\t{a_tot_y:,.2f}\t{a_sav:,.2f}\t{a_mrn:.2f}%\t{a_my:.2f}%\n')

# 3. Export to Excel
with pd.ExcelWriter('СРАВНИТЕЛЕН_АНАЛИЗ_ЯСКО_РОЖЕН_ДЕКЕМВРИ_2025.xlsx', engine='openpyxl') as writer:
    # Sheet 1: Articles
    cols_art = [
        'art_no', 'brand', 'name_pl', 'qty', 'revenue', 'cost_erp', 'profit_erp',
        'p_rojen_base', 'p_rojen_net_inv', 'p_yasko',
        'cost_rojen_base', 'cost_rojen_net_inv', 'cost_yasko',
        'profit_rojen_base', 'profit_rojen_net_inv', 'profit_yasko',
        'diff_cost_base', 'diff_cost_net'
    ]
    art_export = merged_dec[cols_art].copy()
    art_export['Марж_Реален_ERP_%'] = (art_export['profit_erp'] / art_export['revenue']) * 100
    art_export['Марж_Рожен_Базова_%'] = (art_export['profit_rojen_base'] / art_export['revenue']) * 100
    art_export['Марж_Рожен_Нето_%'] = (art_export['profit_rojen_net_inv'] / art_export['revenue']) * 100
    art_export['Марж_Яско_%'] = (art_export['profit_yasko'] / art_export['revenue']) * 100
    art_export.rename(columns={
        'art_no': 'Мат. №',
        'brand': 'Бранд',
        'name_pl': 'Артикул (Ценова листа)',
        'qty': 'Продадени К-во (бр)',
        'revenue': 'Оборот Сума с ДДС (€)',
        'cost_erp': 'Реална Себестойност ERP (€)',
        'profit_erp': 'Реална Печалба ERP (€)',
        'p_rojen_base': 'Рожен Ед. Цена Базова (€)',
        'p_rojen_net_inv': 'Рожен Ед. Цена Нето Ф-ра (€)',
        'p_yasko': 'Яско Ед. Цена (€)',
        'cost_rojen_base': 'Покупна Стойност Рожен Базова (€)',
        'cost_rojen_net_inv': 'Покупна Стойност Рожен Нето (€)',
        'cost_yasko': 'Покупна Стойност Яско (€)',
        'profit_rojen_base': 'Печалба Рожен Базова (€)',
        'profit_rojen_net_inv': 'Печалба Рожен Нето (€)',
        'profit_yasko': 'Печалба Яско (€)',
        'diff_cost_base': 'Икономия спрямо Базова Рожен (€)',
        'diff_cost_net': 'Икономия спрямо Нето Рожен (€)'
    }, inplace=True)
    art_export.sort_values(by='Икономия спрямо Нето Рожен (€)', ascending=False).to_excel(writer, sheet_name='По_Артикули', index=False)
    
    # Sheet 2: Stores
    store_summary.rename(columns={
        'store': 'Магазин / Обект',
        'qty': 'Продадени К-во (бр)',
        'revenue': 'Оборот Сума (€)',
        'cost_erp': 'Реална Себестойност ERP (€)',
        'profit_erp': 'Реална Печалба ERP (€)',
        'cost_rojen_base': 'Покупна Рожен Базова (€)',
        'cost_rojen_net_inv': 'Покупна Рожен Нето (€)',
        'cost_yasko': 'Покупна Яско (€)',
        'diff_cost_base': 'Разлика Рожен Базова - Яско (€)',
        'diff_cost_net': 'Разлика Рожен Нето - Яско (€)',
        'profit_rojen_base': 'Печалба Рожен Базова (€)',
        'profit_rojen_net': 'Печалба Рожен Нето (€)',
        'profit_yasko': 'Печалба Яско (€)',
        'margin_erp_%': 'Марж Реално ERP %',
        'margin_rojen_net_%': 'Марж Рожен Нето %',
        'margin_yasko_%': 'Марж Яско %'
    }, inplace=True)
    store_summary.sort_values(by='Продадени К-во (бр)', ascending=False).to_excel(writer, sheet_name='По_Обекти', index=False)
    
    # Sheet 3: Brands
    brand_summary.rename(columns={
        'brand': 'Бранд',
        'qty': 'Продадени К-во (бр)',
        'revenue': 'Оборот Сума (€)',
        'cost_erp': 'Реална Себестойност ERP (€)',
        'profit_erp': 'Реална Печалба ERP (€)',
        'cost_rojen_base': 'Покупна Рожен Базова (€)',
        'cost_rojen_net_inv': 'Покупна Рожен Нето (€)',
        'cost_yasko': 'Покупна Яско (€)',
        'profit_rojen_base': 'Печалба Рожен Базова (€)',
        'profit_rojen_net_inv': 'Печалба Рожен Нето (€)',
        'profit_yasko': 'Печалба Яско (€)',
        'diff_cost_base': 'Икономия Яско спр. Рожен Базова (€)',
        'diff_cost_net': 'Икономия Яско спр. Рожен Нето (€)',
        'margin_erp_%': 'Марж Реално ERP %',
        'margin_rojen_net_%': 'Марж Рожен Нето %',
        'margin_yasko_%': 'Марж Яско %'
    }, inplace=True)
    brand_summary.to_excel(writer, sheet_name='По_Брандове', index=False)

print('December 2025 calculation complete!')
