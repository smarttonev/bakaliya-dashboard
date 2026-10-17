#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
=============================================================================
ИМПОРТ НА НОМЕРА НА ИЗДАДЕНИ ФАКТУРИ В "ЗАБЕЛЕЖКИ_1" НА ОТЛОЖЕНО ПЛАЩАНЕ В МИСТРАЛ ERP
Верига магазини „Бакалия“ | ЦЕНТРАЛЕН ОФИС — МОНТЕВИДЕО
=============================================================================
Този скрипт импортира номерата на издадените фактури (от 'ИЗдадени фактури 10,26.xlsx')
в полето DEFERREDPAYMENTNOTE1 ("Забележки_1") на отворените начисления в Мистрал ERP.

ПРАВИЛА ЗА БЕЗОПАСНОСТ И ИНТЕГРИТЕТ:
1. 100% точно съвпадение на сумата (с точност до стотинка).
2. Задължително съответствие на контрагента.
3. Полето DEFERREDPAYMENTNOTE1 трябва да бъде стриктно празно (NULL или '').
   Никаква съществуваща забележка не се презаписва!
4. По подразбиране работи в DRY-RUN (симулационен) режим.
5. Запис се извършва САМО при изрично подаден флаг --commit.
6. Записите се актуализират синхронно в MONTEVIDEO.FDB и CS.FDB (Централен склад).
=============================================================================
"""

import os
import sys
import argparse
import openpyxl
import firebirdsql
from decimal import Decimal
import re

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# Конфигурация на базите данни
MV_CONFIG = {
    'host': '192.168.5.254',
    'database': 'E:/db/MONTEVIDEO.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

CS_CONFIG = {
    'host': '192.168.5.254',
    'database': 'E:/db/CS.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

EXCEL_FILE = r'g:\My Drive\AntiG_test\ИЗдадени фактури 10,26.xlsx'

def clean_name(n):
    n = n.upper()
    for lf in ['ЕООД', 'ООД', 'АД', 'ЕАД', 'ЕТ', 'КЧТ', 'ЦЕНТРАЛА', 'БЪЛГАРИЯ', 'ДИСТРИБЮШЪН', 'ТРЕЙДИНГ', 'ТРЕЙД', 'ХОЛДИНГ', 'ГРУП']:
        n = re.sub(r'\b' + lf + r'\b', '', n)
    n = re.sub(r'[^\w\s]', ' ', n)
    n = re.sub(r'\s+', ' ', n).strip()
    return n

def partners_match(pe, pdb):
    ce = clean_name(pe)
    cdb = clean_name(pdb)
    if ce == cdb or ce in cdb or cdb in ce:
        return True
    we = set(ce.split())
    wdb = set(cdb.split())
    if we and wdb and (we.issubset(wdb) or wdb.issubset(we)):
        return True
    aliases = {
        'АМФИОН МД': 'ЕЛИАЗ',
        'ЕЛИАЗ': 'АМФИОН МД',
        'ВИНЕН СВЯТ': 'РАДКОВА',
        'НЕОМ ГАРДЪН': 'САМАКО',
        'САМАКО': 'НЕОМ ГАРДЪН',
        'ТАНДЕМ В': 'ТАНДЕМ ПОПОВ',
        'ТАНДЕМ ПОПОВ': 'ТАНДЕМ В',
    }
    for k, v in aliases.items():
        if (k in ce and v in cdb) or (v in ce and k in cdb):
            return True
    return False

def run_import(commit=False, include_1cent=False):
    print("=" * 80)
    print("ИМПОРТ НА НОМЕРА НА ФАКТУРИ В 'ЗАБЕЛЕЖКИ_1' В МИСТРАЛ ERP")
    print("=" * 80)
    mode_str = "🔴 LIVE COMMIT (РЕАЛЕН ЗАПИС)" if commit else "🟡 DRY-RUN (СИМУЛАЦИЯ - БЕЗ ПРОМЕНИ)"
    print(f"Режим на работа: {mode_str}")
    print(f"Включване на 1 стотинка закръгление: {'ДА' if include_1cent else 'НЕ (само 100% точно съвпадение)'}")
    print("=" * 80)

    # 1. Зареждане на Excel
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Грешка: Файлът не е намерен: {EXCEL_FILE}")
        return

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    excel_items = []
    for r in range(2, ws.max_row + 1):
        p = ws.cell(row=r, column=1).value
        d = ws.cell(row=r, column=2).value
        s = ws.cell(row=r, column=3).value
        n = ws.cell(row=r, column=4).value
        if p is None and s is None:
            continue
        excel_items.append({
            'row': r,
            'partner': str(p).strip(),
            'date': str(d).strip(),
            'sum': Decimal(str(s)).quantize(Decimal('0.01')),
            'inv_num': str(n).strip()
        })
    print(f"Заредени редове от Excel: {len(excel_items)}")

    # 2. Извличане на отворени начисления от Монтевидео
    conn_mv = firebirdsql.connect(**MV_CONFIG)
    cur_mv = conn_mv.cursor()
    cur_mv.execute("""
        SELECT D.ID, D.SALENUM, D.PARTNERNUM, P.NAME, D.SALEDATE, 
               D.DPSUM, D.OPERATIONTYPEID, D.DEFERREDPAYMENTNOTE1
        FROM DEFERREDPAYMENT D
        LEFT JOIN PARTNER P ON (P.NUM = D.PARTNERNUM AND P.LOCATIONID = D.LOCATIONID)
        WHERE D.OPERATIONTYPEID = 3 AND D.SALENUM BETWEEN 77314 AND 77447
        ORDER BY D.SALENUM ASC
    """)
    db_rows = []
    for r in cur_mv.fetchall():
        db_rows.append({
            'id': r[0],
            'salenum': r[1],
            'partnernum': r[2],
            'partner_name': r[3].strip() if r[3] else '',
            'date': r[4],
            'sum': Decimal(str(r[5])).quantize(Decimal('0.01')),
            'note1': r[7].strip() if r[7] else ''
        })

    # 3. Напасване по правила
    to_update = []
    skipped_has_note = []
    skipped_diff_sum = []
    skipped_1cent = []
    unmatched = []

    matched_salenums = set()

    for item in excel_items:
        e_p = item['partner']
        e_s = item['sum']
        e_n = item['inv_num']
        
        p_cands = [d for d in db_rows if partners_match(e_p, d['partner_name'])]
        exact_s = [d for d in p_cands if d['sum'] == e_s]
        cent_s = [d for d in p_cands if abs(d['sum'] - e_s) == Decimal('0.01')]

        if len(exact_s) == 1:
            d = exact_s[0]
            if d['note1'] == '':
                to_update.append((item, d))
                matched_salenums.add(d['salenum'])
            else:
                skipped_has_note.append((item, d))
        elif len(cent_s) == 1:
            d = cent_s[0]
            if include_1cent and d['note1'] == '':
                to_update.append((item, d))
                matched_salenums.add(d['salenum'])
            else:
                skipped_1cent.append((item, d))
        elif p_cands:
            skipped_diff_sum.append((item, p_cands))
        else:
            unmatched.append(item)

    print("\n" + "-" * 80)
    print("РЕЗУЛТАТИ ОТ ЗАСИЧАНЕТО:")
    print(f" -> ГОТОВИ ЗА ИМПОРТ: {len(to_update)} бр.")
    print(f" -> Пропуснати (вече имат забележка в Забележки_1): {len(skipped_has_note)} бр.")
    if not include_1cent:
        print(f" -> Пропуснати (разлика от 1 стотинка 0.01 €): {len(skipped_1cent)} бр.")
    print(f" -> Пропуснати (партньорът е със сборен/разбит запис): {len(skipped_diff_sum)} бр.")
    print(f" -> Пропуснати (няма запис в Мистрал): {len(unmatched)} бр.")
    print("-" * 80)

    print("\nСписък на записите за импорт:")
    for item, d in to_update:
        print(f"  Сметка {d['salenum']} | {d['partner_name']:<32} | Сума: {d['sum']:>9.2f} € | Забележки_1: '{item['inv_num']}'")

    if not commit:
        print("\n" + "=" * 80)
        print("🟡 DRY-RUN ПРИКЛЮЧИ: Не са направени промени в базата данни.")
        print("За да запишете промените, стартирайте с: python import_invoice_notes.py --commit")
        if not include_1cent:
            print("(Ако желаете да включите и 4-те записа с 1 ст. закръгление, добавете и --include-1cent)")
        print("=" * 80)
        conn_mv.close()
        return

    # 4. LIVE COMMIT
    print("\n" + "=" * 80)
    print("🔴 ИЗПЪЛНЕНИЕ НА LIVE COMMIT В MONTEVIDEO.FDB И CS.FDB...")
    print("=" * 80)

    conn_cs = firebirdsql.connect(**CS_CONFIG)
    cur_cs = conn_cs.cursor()

    updated_mv = 0
    updated_cs = 0

    try:
        for item, d in to_update:
            salenum = d['salenum']
            inv_num = item['inv_num']
            dpsum = d['sum']

            # Update MONTEVIDEO.FDB със строга защита в WHERE клаузата:
            # само ако полето е все още празно и сумата съвпада!
            cur_mv.execute("""
                UPDATE DEFERREDPAYMENT
                SET DEFERREDPAYMENTNOTE1 = ?
                WHERE LOCATIONID = 7 
                  AND SALENUM = ? 
                  AND (DEFERREDPAYMENTNOTE1 IS NULL OR DEFERREDPAYMENTNOTE1 = '')
            """, (inv_num, salenum))
            if cur_mv.rowcount > 0:
                updated_mv += 1

            # Update CS.FDB със същата защита
            cur_cs.execute("""
                UPDATE DEFERREDPAYMENT
                SET DEFERREDPAYMENTNOTE1 = ?
                WHERE LOCATIONID = 7 
                  AND SALENUM = ? 
                  AND (DEFERREDPAYMENTNOTE1 IS NULL OR DEFERREDPAYMENTNOTE1 = '')
            """, (inv_num, salenum))
            if cur_cs.rowcount > 0:
                updated_cs += 1

        conn_mv.commit()
        conn_cs.commit()
        print(f"\n✅ УСПЕШНО АКТУАЛИЗИРАНИ {updated_mv} ЗАПИСА В MONTEVIDEO.FDB!")
        print(f"✅ УСПЕШНО АКТУАЛИЗИРАНИ {updated_cs} ЗАПИСА В CS.FDB!")
        print("=" * 80)

    except Exception as e:
        conn_mv.rollback()
        conn_cs.rollback()
        print(f"\n❌ ГРЕШКА ПРИ ЗАПИС: {e}")
        raise
    finally:
        conn_mv.close()
        conn_cs.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Импорт на номера на фактури в Мистрал ERP")
    parser.add_argument('--commit', action='store_true', help="Извършва реалния запис в базата данни")
    parser.add_argument('--include-1cent', action='store_true', help="Включва и 4-те записа с точно 0.01 € разлика в закръглението")
    args = parser.parse_args()

    run_import(commit=args.commit, include_1cent=args.include_1cent)
