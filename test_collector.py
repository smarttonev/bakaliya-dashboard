import glob, os, sys
sys.stdout.reconfigure(encoding='utf-8')
from brochure_mapping import BrochureMappingManager, PRODUCT_GROUPS

mgr = BrochureMappingManager()
all_files = sorted(glob.glob('incoming_brochures/*.*'))

master_pool = []
for f in all_files:
    if not f.endswith(('.xls', '.xlsx')) or 'temp_' in f:
        continue
    parsed, err = mgr.parse_any_supplier_blank(f)
    if parsed and parsed.get('items'):
        fn = os.path.basename(f)
        for it in parsed['items']:
            raw_l = it['raw_name'].lower()
            if any(w in raw_l for w in ['орехите', 'белла', 'тракийски', 'перелик']):
                it['supplier'] = 'Белла България АД'
            elif any(w in raw_l for w in ['чио', 'пом бер']):
                it['supplier'] = 'Интерснак България ЕООД'
            elif any(w in raw_l for w in ['еверест', 'fitspo']):
                it['supplier'] = 'Победа АД / Ефект 2002'
            elif any(w in raw_l for w in ['фарел']):
                it['supplier'] = 'Бикс / Фарел (Боди ЕООД)'
            elif any(w in raw_l for w in ['троко', 'кирилов']):
                it['supplier'] = 'Сириус ООД'
            it['source_file'] = fn
            master_pool.append(it)

print(f"Master Pool Total Items: {len(master_pool)}")
grp_counts = {}
for it in master_pool:
    gname = it.get('group_name', 'масови')
    grp_counts[gname] = grp_counts.get(gname, 0) + 1

print("\nGroup distribution:")
for g, c in sorted(grp_counts.items(), key=lambda x: -x[1]):
    print(f"  {g}: {c} items")

print("\nSample items from different suppliers in Master Pool:")
for it in master_pool[::20]:
    s = it['supplier']
    m = it['marketing_name']
    g = it['group_name']
    dc = it.get('promo_dc', 0.0)
    pc = it.get('promo_pc', 0.0)
    print(f"  [{s}] {m} -> {g} | DC: {dc:.2f} | PC: {pc:.2f}")
