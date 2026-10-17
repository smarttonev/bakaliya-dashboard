import openpyxl, xlrd, os, re, sys, glob
import pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
from brochure_mapping import BrochureMappingManager
from brochure_rules import audit_brochure_item

def parse_any_supplier_blank(file_path):
    engine = 'openpyxl' if file_path.endswith('.xlsx') else 'xlrd'
    xl = pd.ExcelFile(file_path, engine=engine)
    
    if 'Бланка_Предложения' in xl.sheet_names:
        df = xl.parse('Бланка_Предложения', header=None)
        brochure = str(df.iloc[4, 1] if pd.notna(df.iloc[4, 1]) else df.iloc[4, 3])
        supplier = str(df.iloc[4, 6] if pd.notna(df.iloc[4, 6]) else df.iloc[4, 11])
        items = []
        for r in range(12, len(df)):
            name = df.iloc[r, 2] if pd.notna(df.iloc[r, 2]) else df.iloc[r, 3]
            if pd.notna(name) and str(name).strip() and not str(name).startswith('Пример'):
                barcode = str(df.iloc[r, 1]).strip() if pd.notna(df.iloc[r, 1]) else ''
                unit = str(df.iloc[r, 3]).strip() if pd.notna(df.iloc[r, 3]) else ''
                def to_flt(v):
                    if pd.isna(v): return 0.0
                    try: return float(str(v).replace(',', '.').strip())
                    except: return 0.0
                reg_dc = to_flt(df.iloc[r, 4])
                promo_dc = to_flt(df.iloc[r, 5])
                rec_pc = to_flt(df.iloc[r, 7])
                note = str(df.iloc[r, 8]).strip() if pd.notna(df.iloc[r, 8]) else ''
                items.append({
                    'raw_name': str(name).strip(), 'barcode': barcode, 'unit': unit,
                    'reg_dc': reg_dc, 'promo_dc': promo_dc, 'promo_pc': rec_pc, 'note': note
                })
        return {'format': '2027_TEMPLATE', 'brochure': brochure, 'supplier': supplier, 'items_count': len(items), 'items': items}

    # Legacy blank
    df = xl.parse(xl.sheet_names[0], header=None)
    hdr_row = None
    for r in range(min(30, len(df))):
        cells = [str(x).lower().strip() for x in df.iloc[r].dropna().values]
        has_art = any('артикул' in c or 'наименование' in c for c in cells)
        has_price = any('цена' in c or 'дц' in c for c in cells)
        if has_art and has_price:
            hdr_row = r
            break
            
    if hdr_row is None:
        return {'error': 'Не е намерен ред със заглавки'}

    hdr = df.iloc[hdr_row]
    col_art, col_reg_dc, col_promo_dc, col_promo_pc, col_barcode = None, None, None, None, None
    for c in range(df.shape[1]):
        val = str(hdr[c]).lower().strip()
        if 'артикул' in val or 'наименование' in val:
            col_art = c
        elif 'базова' in val or ('редовна' in val and 'цена' in val):
            col_reg_dc = c
        elif 'промоционална' in val or ('промо' in val and 'цена' in val):
            col_promo_dc = c
        elif 'препоръчителна' in val or 'продажна' in val:
            col_promo_pc = c
        elif 'баркод' in val or 'код' in val:
            col_barcode = c

    # Extract supplier name from legal text or filename
    supplier = os.path.basename(file_path).replace('.xls', '').replace('.xlsx', '')
    for r in range(min(14, len(df))):
        txt = str(df.iloc[r, 0])
        if '2.' in txt or '2 .' in txt or 'със седалище' in txt:
            clean_txt = txt.replace('2.', '').replace('2 .', '').strip()
            parts = clean_txt.split('със седалище')
            cand_supp = parts[0].strip(' .–-„"')
            if len(cand_supp) > 2 and not cand_supp.startswith('...'):
                supplier = cand_supp
                break

    items = []
    curr_brochure = 'Брошура'
    for r in range(hdr_row + 1, len(df)):
        c0_val = str(df.iloc[r, 0]).strip() if pd.notna(df.iloc[r, 0]) else ''
        if 'брошура' in c0_val.lower():
            curr_brochure = c0_val
            
        if col_art is not None and col_art < df.shape[1]:
            name = df.iloc[r, col_art]
            if pd.notna(name) and str(name).strip():
                s_name = str(name).strip()
                if not any(w in s_name.lower() for w in ['моля', 'подпишете', 'период на']):
                    def to_flt(v):
                        if pd.isna(v): return 0.0
                        try: return float(str(v).replace(',', '.').strip())
                        except: return 0.0
                    reg_dc = to_flt(df.iloc[r, col_reg_dc]) if col_reg_dc is not None else 0.0
                    promo_dc = to_flt(df.iloc[r, col_promo_dc]) if col_promo_dc is not None else 0.0
                    rec_pc = to_flt(df.iloc[r, col_promo_pc]) if col_promo_pc is not None else 0.0
                    
                    barcode = ''
                    if col_barcode is not None and pd.notna(df.iloc[r, col_barcode]):
                        barcode = str(df.iloc[r, col_barcode]).replace('.0', '').strip()
                    else:
                        for c in range(df.shape[1]):
                            cand = str(df.iloc[r, c]).replace('.0', '').strip()
                            if cand.isdigit() and len(cand) >= 6 and c != col_reg_dc and c != col_promo_dc and c != col_promo_pc:
                                barcode = cand
                                break
                    items.append({
                        'brochure': curr_brochure,
                        'raw_name': s_name,
                        'barcode': barcode,
                        'reg_dc': reg_dc,
                        'promo_dc': promo_dc,
                        'promo_pc': rec_pc
                    })

    return {'format': 'LEGACY_BLANK', 'brochure': curr_brochure, 'supplier': supplier, 'items_count': len(items), 'items': items}

if __name__ == '__main__':
    mgr = BrochureMappingManager()
    test_file = 'incoming_brochures/email_5_Бланка брошура- Ноември Бакалия.xls'
    parsed = parse_any_supplier_blank(test_file)
    print('Testing End-to-End on:', test_file)
    print('Format:', parsed['format'], '| Items count:', len(parsed['items']))

    resolved_items = []
    for it in parsed['items']:
        res = mgr.resolve_supplier_item(
            'Белла България АД',
            it['raw_name'],
            barcode=it.get('barcode'),
            promo_dc=it.get('promo_dc'),
            reg_dc=it.get('reg_dc'),
            rec_promo_pc=it.get('promo_pc')
        )
        res['promo_dc'] = it.get('promo_dc', 0.0)
        res['reg_dc'] = it.get('reg_dc', 0.0)
        res['promo_pc'] = it.get('promo_pc') or (res.get('reg_pc', 0.0) * 0.85)
        
        # Audit rules
        audit_in = {
            'name': res['marketing_name'],
            'promo_dc': res['promo_dc'],
            'promo_pc': res['promo_pc'],
            'reg_dc': res['reg_dc'],
            'code': res['mistral_codes'][0] if res.get('mistral_codes') else None
        }
        audit_res = audit_brochure_item(audit_in, resolved_items, [])
        res['audit'] = audit_res
        resolved_items.append(res)

    for idx, r in enumerate(resolved_items, 1):
        m_name = r['marketing_name']
        codes = r.get('mistral_codes', [])
        grp = r.get('group_name')
        st = r.get('status')
        pdc = r.get('promo_dc', 0.0)
        ppc = r.get('promo_pc', 0.0)
        st_audit = r['audit']['status']
        margin = r['audit']['margin_pct']
        print(f"{idx}. {m_name} | Grp: {grp} | Code: {codes} | Status: {st} | PromoDC: {pdc:.2f} | PromoPC: {ppc:.2f} | Margin: {margin:.1f}% | Audit: {st_audit}")
