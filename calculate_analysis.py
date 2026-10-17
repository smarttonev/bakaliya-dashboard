# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl

p1 = r'C:\Users\Tonev\.gemini\antigravity\brain\d872b9fb-1db0-4f76-abd7-263b17bbdcb9\.user_uploaded\media_1791894563117_a1485000.xlsx'
wb1 = openpyxl.load_workbook(p1)
ws1 = wb1.active
data1 = list(ws1.iter_rows(values_only=True))

price_items = []
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

df_pl = pd.DataFrame(price_items)
pl_distinct = df_pl.drop_duplicates(subset=['art_no']).copy()

# Sales data
p2 = r'C:\Users\Tonev\.gemini\antigravity\brain\d872b9fb-1db0-4f76-abd7-263b17bbdcb9\.user_uploaded\media_1791894543129_7f137ae1.xls'
df_sales = pd.read_csv(p2, sep='\t', encoding='cp1251')
df_sales['art_no'] = df_sales['Мат. №'].astype(str).str.strip()

# Aggregate sales by art_no
sales_by_art = df_sales.groupby('art_no').agg({
    'Име на материал': 'first',
    'К-во': 'sum',
    'Сума': 'sum',
    'Печалба': 'sum'
}).reset_index()

merged = pd.merge(sales_by_art, pl_distinct, on='art_no', how='inner')

# Calculations
merged['cost_rojen_base'] = merged['К-во'] * merged['p_rojen_base']

# Discounts: MDLZ 6.8%, Chipita 4.0%
merged['disc_pct_rojen'] = merged['brand'].map({'MDLZ': 0.068, 'CHIPITA': 0.040})
merged['p_rojen_net_inv'] = merged['p_rojen_base'] * (1 - merged['disc_pct_rojen'])
merged['cost_rojen_net_inv'] = merged['К-во'] * merged['p_rojen_net_inv']

# Rozhen with flat 6.8%
merged['p_rojen_net_68'] = merged['p_rojen_base'] * (1 - 0.068)
merged['cost_rojen_net_68'] = merged['К-во'] * merged['p_rojen_net_68']

# Yasko
merged['cost_yasko'] = merged['К-во'] * merged['p_yasko']

# Revenue and profit
merged['revenue'] = merged['Сума']
merged['current_cost'] = merged['Сума'] - merged['Печалба']
merged['current_profit'] = merged['Печалба']

merged['profit_rojen_base'] = merged['revenue'] - merged['cost_rojen_base']
merged['profit_rojen_net_inv'] = merged['revenue'] - merged['cost_rojen_net_inv']
merged['profit_rojen_net_68'] = merged['revenue'] - merged['cost_rojen_net_68']
merged['profit_yasko'] = merged['revenue'] - merged['cost_yasko']

# Store-level analysis
sales_matched_store = df_sales[df_sales['art_no'].isin(pl_distinct['art_no'])].copy()
sales_matched_store = pd.merge(sales_matched_store, pl_distinct[['art_no', 'brand', 'p_rojen_base', 'p_yasko']], on='art_no', how='left')

sales_matched_store['disc_pct_rojen'] = sales_matched_store['brand'].map({'MDLZ': 0.068, 'CHIPITA': 0.040})
sales_matched_store['p_rojen_net_inv'] = sales_matched_store['p_rojen_base'] * (1 - sales_matched_store['disc_pct_rojen'])

sales_matched_store['cost_rojen_base'] = sales_matched_store['К-во'] * sales_matched_store['p_rojen_base']
sales_matched_store['cost_rojen_net_inv'] = sales_matched_store['К-во'] * sales_matched_store['p_rojen_net_inv']
sales_matched_store['cost_yasko'] = sales_matched_store['К-во'] * sales_matched_store['p_yasko']

store_summary = sales_matched_store.groupby('Oбект').agg({
    'К-во': 'sum',
    'Сума': 'sum',
    'Печалба': 'sum',
    'cost_rojen_base': 'sum',
    'cost_rojen_net_inv': 'sum',
    'cost_yasko': 'sum'
}).reset_index()

store_summary['diff_base'] = store_summary['cost_rojen_base'] - store_summary['cost_yasko']
store_summary['diff_net'] = store_summary['cost_rojen_net_inv'] - store_summary['cost_yasko']
store_summary['profit_rojen_base'] = store_summary['Сума'] - store_summary['cost_rojen_base']
store_summary['profit_rojen_net'] = store_summary['Сума'] - store_summary['cost_rojen_net_inv']
store_summary['profit_yasko'] = store_summary['Сума'] - store_summary['cost_yasko']

# Brand summary
brand_summary = merged.groupby('brand').agg({
    'К-во': 'sum',
    'revenue': 'sum',
    'current_cost': 'sum',
    'current_profit': 'sum',
    'cost_rojen_base': 'sum',
    'cost_rojen_net_inv': 'sum',
    'cost_yasko': 'sum',
    'profit_rojen_base': 'sum',
    'profit_rojen_net_inv': 'sum',
    'profit_yasko': 'sum'
}).reset_index()

with open('analysis_output.txt', 'w', encoding='utf-8') as out:
    out.write('=== ОБЩ СУМАРЕН АНАЛИЗ (ЯНУАРИ 2026) ===\n')
    tot_qty = merged['К-во'].sum()
    tot_rev = merged['revenue'].sum()
    tot_cur_cost = merged['current_cost'].sum()
    tot_cur_profit = merged['current_profit'].sum()
    tot_cur_margin = (tot_cur_profit / tot_rev) * 100
    
    cost_rb = merged['cost_rojen_base'].sum()
    cost_rn = merged['cost_rojen_net_inv'].sum()
    cost_rn68 = merged['cost_rojen_net_68'].sum()
    cost_y = merged['cost_yasko'].sum()
    
    prof_rb = merged['profit_rojen_base'].sum()
    prof_rn = merged['profit_rojen_net_inv'].sum()
    prof_rn68 = merged['profit_rojen_net_68'].sum()
    prof_y = merged['profit_yasko'].sum()
    
    mar_rb = (prof_rb / tot_rev) * 100
    mar_rn = (prof_rn / tot_rev) * 100
    mar_rn68 = (prof_rn68 / tot_rev) * 100
    mar_y = (prof_y / tot_rev) * 100
    
    diff_base = cost_rb - cost_y
    diff_net = cost_rn - cost_y
    diff_net68 = cost_rn68 - cost_y
    
    out.write(f'Общо продадени бройки: {tot_qty:,.2f} бр.\n')
    out.write(f'Общ реализиран оборот (Сума с ДДС): {tot_rev:,.2f} €\n')
    out.write(f'Текуща себестойност (ERP): {tot_cur_cost:,.2f} €\n')
    out.write(f'Текуща реализирана печалба: {tot_cur_profit:,.2f} € (Марж: {tot_cur_margin:.2f}%)\n')
    out.write('\n--- СЦЕНАРИЙ 1: РОЖЕН БАЗОВА ЦЕНА С ДДС ---\n')
    out.write(f'Покупна стойност: {cost_rb:,.2f} €\n')
    out.write(f'Печалба при тези продажби: {prof_rb:,.2f} €\n')
    out.write(f'Търговски марж: {mar_rb:.2f}%\n')
    out.write('\n--- СЦЕНАРИЙ 2: РОЖЕН НЕТО С ОТСТЪПКА НА ФАКТУРА (MDLZ -6.8%, CHIPITA -4.0%) ---\n')
    out.write(f'Покупна стойност: {cost_rn:,.2f} €\n')
    out.write(f'Печалба при тези продажби: {prof_rn:,.2f} €\n')
    out.write(f'Търговски марж: {mar_rn:.2f}%\n')
    out.write('\n--- СЦЕНАРИЙ 2Б: РОЖЕН НЕТО С -6.8% И ЗА ЧИПИТА ---\n')
    out.write(f'Покупна стойност: {cost_rn68:,.2f} €\n')
    out.write(f'Печалба при тези продажби: {prof_rn68:,.2f} €\n')
    out.write(f'Търговски марж: {mar_rn68:.2f}%\n')
    out.write('\n--- СЦЕНАРИЙ 3: ЯСКО (ЦЕНИ С ДДС) ---\n')
    out.write(f'Покупна стойност: {cost_y:,.2f} €\n')
    out.write(f'Печалба при тези продажби: {prof_y:,.2f} €\n')
    out.write(f'Търговски марж: {mar_y:.2f}%\n')
    out.write('\n=== РАЗЛИКИ И ИКОНОМИЧЕСКИ ЕФЕКТ ===\n')
    out.write(f'Разлика в покупна стойност (Рожен Базова vs Яско): {diff_base:+,.2f} € (Яско е по-изгоден с {diff_base:,.2f} € или {diff_base/cost_rb*100:.2f}%)\n')
    out.write(f'Разлика в покупна стойност (Рожен Нето фактура vs Яско): {diff_net:+,.2f} € (Яско е по-изгоден с {diff_net:,.2f} € или {diff_net/cost_rn*100:.2f}%)\n')
    out.write(f'Разлика в печалбата (Яско vs Рожен Базова): {prof_y - prof_rb:+,.2f} € (Марж: +{mar_y - mar_rb:.2f} процентни пункта)\n')
    out.write(f'Разлика в печалбата (Яско vs Рожен Нето): {prof_y - prof_rn:+,.2f} € (Марж: +{mar_y - mar_rn:.2f} процентни пункта)\n')
    out.write(f'Разлика в печалбата (Яско vs Текуща в ERP): {prof_y - tot_cur_profit:+,.2f} € (Марж: +{mar_y - tot_cur_margin:.2f} процентни пункта)\n')
    
    out.write('\n\n=== АНАЛИЗ ПО БРАНДОВЕ (MDLZ vs CHIPITA) ===\n')
    for _, br in brand_summary.iterrows():
        b_name = br['brand']
        b_qty = br['К-во']
        b_rev = br['revenue']
        b_crb = br['cost_rojen_base']
        b_crn = br['cost_rojen_net_inv']
        b_cy = br['cost_yasko']
        b_prb = br['profit_rojen_base']
        b_prn = br['profit_rojen_net_inv']
        b_py = br['profit_yasko']
        out.write(f'Бранд: {b_name}\n')
        out.write(f'  Продадени бройки: {b_qty:,.2f} бр.\n')
        out.write(f'  Приход (Сума): {b_rev:,.2f} €\n')
        out.write(f'  Покупна Рожен Базова: {b_crb:,.2f} € (Печалба: {b_prb:,.2f} €, Марж: {b_prb/b_rev*100:.2f}%)\n')
        out.write(f'  Покупна Рожен Нето фактура: {b_crn:,.2f} € (Печалба: {b_prn:,.2f} €, Марж: {b_prn/b_rev*100:.2f}%)\n')
        out.write(f'  Покупна Яско: {b_cy:,.2f} € (Печалба: {b_py:,.2f} €, Марж: {b_py/b_rev*100:.2f}%)\n')
        out.write(f'  Икономия от Яско спрямо Рожен Базова: {b_crb - b_cy:,.2f} € (Марж: +{(b_py/b_rev - b_prb/b_rev)*100:.2f}%)\n')
        out.write(f'  Икономия от Яско спрямо Рожен Нето: {b_crn - b_cy:,.2f} € (Марж: +{(b_py/b_rev - b_prn/b_rev)*100:.2f}%)\n')
        out.write('\n')

    out.write('\n=== АНАЛИЗ ПО ОБЕКТИ (МАГАЗИНИ) ===\n')
    out.write('Обект\tК-во\tПриход\tРожен Базова\tРожен Нето\tЯско\tИкономия Яско (спр. Нето)\tМарж Рожен Нето %\tМарж Яско %\n')
    for _, st in store_summary.sort_values(by='К-во', ascending=False).iterrows():
        st_name = st['Oбект']
        st_qty = st['К-во']
        st_rev = st['Сума']
        st_crb = st['cost_rojen_base']
        st_crn = st['cost_rojen_net_inv']
        st_cy = st['cost_yasko']
        st_econ = st_crn - st_cy
        st_mr_net = (st['profit_rojen_net'] / st_rev) * 100
        st_my = (st['profit_yasko'] / st_rev) * 100
        out.write(f'{st_name}\t{st_qty:,.1f}\t{st_rev:,.2f}\t{st_crb:,.2f}\t{st_crn:,.2f}\t{st_cy:,.2f}\t{st_econ:,.2f}\t{st_mr_net:.2f}%\t{st_my:.2f}%\n')

    out.write('\n=== ТОП 20 АРТИКУЛА ПО ИКОНОМИЯ ОТ ЯСКО ===\n')
    merged['saving_vs_net'] = merged['cost_rojen_net_inv'] - merged['cost_yasko']
    merged_sorted = merged.sort_values(by='saving_vs_net', ascending=False)
    out.write('Мат №\tАртикул\tБранд\tК-во\tРожен Базова\tРожен Нето\tЯско\tОбщо Рожен Нето\tОбщо Яско\tИкономия Яско\tМарж Рожен Нето %\tМарж Яско %\n')
    for _, a in merged_sorted.head(20).iterrows():
        a_id = a['art_no']
        a_name = a['name_pl']
        a_brand = a['brand']
        a_qty = a['К-во']
        a_rb = a['p_rojen_base']
        a_rn = a['p_rojen_net_inv']
        a_y = a['p_yasko']
        a_tot_rn = a['cost_rojen_net_inv']
        a_tot_y = a['cost_yasko']
        a_sav = a['saving_vs_net']
        a_rev = a['revenue']
        a_mrn = ((a_rev - a_tot_rn) / a_rev) * 100 if a_rev > 0 else 0
        a_my = ((a_rev - a_tot_y) / a_rev) * 100 if a_rev > 0 else 0
        out.write(f'{a_id}\t{a_name}\t{a_brand}\t{a_qty:,.0f}\t{a_rb:.3f}\t{a_rn:.3f}\t{a_y:.3f}\t{a_tot_rn:,.2f}\t{a_tot_y:,.2f}\t{a_sav:,.2f}\t{a_mrn:.2f}%\t{a_my:.2f}%\n')

# Export full detailed comparison to Excel
with pd.ExcelWriter('СРАВНИТЕЛЕН_АНАЛИЗ_ЯСКО_РОЖЕН_ЯНУАРИ_2026.xlsx', engine='openpyxl') as writer:
    # Sheet 1: Articles Comparison
    cols_art = [
        'art_no', 'brand', 'name_pl', 'К-во', 'revenue', 'current_cost', 'current_profit',
        'p_rojen_base', 'p_rojen_net_inv', 'p_yasko',
        'cost_rojen_base', 'cost_rojen_net_inv', 'cost_yasko',
        'profit_rojen_base', 'profit_rojen_net_inv', 'profit_yasko'
    ]
    art_export = merged[cols_art].copy()
    art_export['Икономия_спрямо_Базова_Рожен'] = art_export['cost_rojen_base'] - art_export['cost_yasko']
    art_export['Икономия_спрямо_Нето_Рожен'] = art_export['cost_rojen_net_inv'] - art_export['cost_yasko']
    art_export['Марж_Текущ_%'] = (art_export['current_profit'] / art_export['revenue']) * 100
    art_export['Марж_Рожен_Базова_%'] = (art_export['profit_rojen_base'] / art_export['revenue']) * 100
    art_export['Марж_Рожен_Нето_%'] = (art_export['profit_rojen_net_inv'] / art_export['revenue']) * 100
    art_export['Марж_Яско_%'] = (art_export['profit_yasko'] / art_export['revenue']) * 100
    art_export.rename(columns={
        'art_no': 'Мат. №',
        'brand': 'Бранд',
        'name_pl': 'Артикул (Ценова листа)',
        'revenue': 'Оборот Сума с ДДС (€)',
        'current_cost': 'Себестойност Текуща (€)',
        'current_profit': 'Печалба Текуща (€)',
        'p_rojen_base': 'Рожен Ед. Цена Базова (€)',
        'p_rojen_net_inv': 'Рожен Ед. Цена Нето Ф-ра (€)',
        'p_yasko': 'Яско Ед. Цена (€)',
        'cost_rojen_base': 'Покупна Стойност Рожен Базова (€)',
        'cost_rojen_net_inv': 'Покупна Стойност Рожен Нето (€)',
        'cost_yasko': 'Покупна Стойност Яско (€)',
        'profit_rojen_base': 'Печалба Рожен Базова (€)',
        'profit_rojen_net_inv': 'Печалба Рожен Нето (€)',
        'profit_yasko': 'Печалба Яско (€)'
    }, inplace=True)
    art_export.sort_values(by='Икономия_спрямо_Нето_Рожен', ascending=False).to_excel(writer, sheet_name='По_Артикули', index=False)
    
    # Sheet 2: Stores
    store_summary.rename(columns={
        'Oбект': 'Магазин / Обект',
        'Сума': 'Оборот Сума (€)',
        'Печалба': 'Печалба Текуща (€)',
        'cost_rojen_base': 'Покупна Рожен Базова (€)',
        'cost_rojen_net_inv': 'Покупна Рожен Нето (€)',
        'cost_yasko': 'Покупна Яско (€)',
        'diff_base': 'Разлика Рожен Базова - Яско (€)',
        'diff_net': 'Разлика Рожен Нето - Яско (€)',
        'profit_rojen_base': 'Печалба Рожен Базова (€)',
        'profit_rojen_net': 'Печалба Рожен Нето (€)',
        'profit_yasko': 'Печалба Яско (€)'
    }, inplace=True)
    store_summary['Марж_Рожен_Нето_%'] = (store_summary['Печалба Рожен Нето (€)'] / store_summary['Оборот Сума (€)']) * 100
    store_summary['Марж_Яско_%'] = (store_summary['Печалба Яско (€)'] / store_summary['Оборот Сума (€)']) * 100
    store_summary.sort_values(by='К-во', ascending=False).to_excel(writer, sheet_name='По_Обекти', index=False)
    
    # Sheet 3: Brands
    brand_summary.rename(columns={
        'brand': 'Бранд',
        'revenue': 'Оборот Сума (€)',
        'current_cost': 'Себестойност Текуща (€)',
        'current_profit': 'Печалба Текуща (€)',
        'cost_rojen_base': 'Покупна Рожен Базова (€)',
        'cost_rojen_net_inv': 'Покупна Рожен Нето (€)',
        'cost_yasko': 'Покупна Яско (€)',
        'profit_rojen_base': 'Печалба Рожен Базова (€)',
        'profit_rojen_net_inv': 'Печалба Рожен Нето (€)',
        'profit_yasko': 'Печалба Яско (€)'
    }, inplace=True)
    brand_summary['Икономия_Яско_спрямо_Рожен_Нето'] = brand_summary['Покупна Рожен Нето (€)'] - brand_summary['Покупна Яско (€)']
    brand_summary['Марж_Рожен_Нето_%'] = (brand_summary['Печалба Рожен Нето (€)'] / brand_summary['Оборот Сума (€)']) * 100
    brand_summary['Марж_Яско_%'] = (brand_summary['Печалба Яско (€)'] / brand_summary['Оборот Сума (€)']) * 100
    brand_summary.to_excel(writer, sheet_name='По_Брандове', index=False)

print('Calculation complete and saved!')
