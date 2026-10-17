#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
=============================================================================
ИМПОРТ НА НОМЕРА НА ФАКТУРИ ЗА СБОРНИ / РАЗБИТИ ЗАПИСИ В МИСТРАЛ ERP
Верига магазини „Бакалия“ | ЦЕНТРАЛЕН ОФИС — МОНТЕВИДЕО
=============================================================================
Този скрипт въвежда номерата на фактурите в полето DEFERREDPAYMENTNOTE1 ("Забележки_1")
за всички 23 сметки със сборни и разбити записи в Мистрал, при които сумата съвпада.

Запазва съществуващите брандове (напр. 'К1', 'Храна', 'Девин', 'Лактима', 'Чио', 
'Джей ди', 'Лейс', 'ТАТ', 'Фереро', 'Рошен', 'Нестле') и добавя точните номера към тях!
=============================================================================
"""

import sys
import argparse
import firebirdsql
from decimal import Decimal

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

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

CONSOLIDATED_UPDATES = [
    # 1. Експрес Дистрибюшън ЕООД (Сборен: 1014.11 + 330 + 220 + 110 = 1674.11 €)
    {
        'salenum': 77358,
        'partner': 'Експрес Дистрибюшън ЕООД',
        'sum': Decimal('1674.11'),
        'new_note': '9000006085-6088 (4 бр.)',
        'desc': '4 фактури в Ексел (сбор 1674.11 €)'
    },
    # 2. Елит -П ЕООД (Сборен: 459.11 + 220 = 679.11 €)
    {
        'salenum': 77359,
        'partner': 'Елит -П ЕООД',
        'sum': Decimal('679.11'),
        'new_note': '9000006089, 9000006090',
        'desc': '2 фактури в Ексел (сбор 679.11 €)'
    },
    # 3. Маргарит ЕООД (Сборен: 6 ф-ри по 220/110 = 1100.00 €)
    {
        'salenum': 77381,
        'partner': 'Маргарит ЕООД',
        'sum': Decimal('1100.00'),
        'new_note': '9000006123-6128 (6 бр.)',
        'desc': '6 фактури в Ексел (сбор 1100.00 €)'
    },
    # 4. Мургаш 88 ООД (Сборен: 18 ф-ри = 2479.55 €)
    {
        'salenum': 77389,
        'partner': 'Мургаш 88 ООД',
        'sum': Decimal('2479.55'),
        'new_note': '9000006136-6153 (18 бр.)',
        'desc': '18 фактури в Ексел (сбор 2479.55 €)'
    },
    # 5. Пауър Брандс ЕООД (Сборен: 8 ф-ри = 2445.39 €)
    {
        'salenum': 77402,
        'partner': 'Пауър Брандс ЕООД',
        'sum': Decimal('2445.39'),
        'new_note': '9000006167-6174 (8 бр.)',
        'desc': '8 фактури в Ексел (сбор 2445.39 €)'
    },
    # 6. Орбико България ЕООД (Сборен: 110 + 110 + 550 = 770.00 € vs 770.01 €)
    {
        'salenum': 77400,
        'partner': 'Орбико България ЕООД',
        'sum': Decimal('770.01'),
        'new_note': '9000006164-6166 (3 бр.)',
        'desc': '3 фактури в Ексел (сбор 770.00 €)'
    },
    # 7. См Трейд 2000 ЕООД (Сборен: 4 ф-ри по 110 = 440.00 € vs 440.01 €)
    {
        'salenum': 77412,
        'partner': 'См Трейд 2000 ЕООД',
        'sum': Decimal('440.01'),
        'new_note': '9000006190-6193 (4 бр.)',
        'desc': '4 фактури в Ексел (сбор 440.00 €)'
    },
    # 8. СИРИУС 2005 ООД (Сборен: 8 ф-ри = 1232.52 €)
    {
        'salenum': 77410,
        'partner': 'СИРИУС 2005 ООД',
        'sum': Decimal('1232.52'),
        'new_note': '9000006181-6188 (8 бр.)',
        'desc': '8 фактури в Ексел (сбор 1232.52 €)'
    },
    # 9. СИРИУС 2005 ООД (Лактима: 1 ф-ра за 410.69 €)
    {
        'salenum': 77403,
        'partner': 'СИРИУС 2005 ООД',
        'sum': Decimal('410.69'),
        'new_note': 'Лактима: 9000006175',
        'desc': '1 фактура за Лактима (410.69 €)'
    },
    # 10. Бел Групс ЕООД (Разбити записи: част 1, 720.01 €)
    {
        'salenum': 77325,
        'partner': 'Бел Групс ЕООД',
        'sum': Decimal('720.01'),
        'new_note': '9000006045 (част 1)',
        'desc': 'Фактура 9000006045 за 1945.66 € (част 1 от 2)'
    },
    # 11. Бел Групс ЕООД (Разбити записи: част 2, 1225.66 €)
    {
        'salenum': 77432,
        'partner': 'Бел Групс ЕООД',
        'sum': Decimal('1225.66'),
        'new_note': '9000006045 (част 2)',
        'desc': 'Фактура 9000006045 за 1945.66 € (част 2 от 2)'
    },
    # 12. Ефект 2002 ЕООД (Основен: 14 ф-ри = 2444.42 €)
    {
        'salenum': 77362,
        'partner': 'Ефект 2002 ЕООД',
        'sum': Decimal('2444.42'),
        'new_note': '9000006093-6106 (14 бр.)',
        'desc': '14 фактури за основен оборот (сбор 2444.42 €)'
    },
    # 13. Ефект 2002 ЕООД (Рошен: 142.63 + 660.00 = 802.63 €)
    {
        'salenum': 77363,
        'partner': 'Ефект 2002 ЕООД',
        'sum': Decimal('802.63'),
        'new_note': 'Рошен: 9000006107, 6108',
        'desc': '2 фактури за Рошен (сбор 802.63 €)'
    },
    # 14. Ефект 2002 ЕООД (Нестле: 2113.08 + 410.00 = 2523.08 €)
    {
        'salenum': 77393,
        'partner': 'Ефект 2002 ЕООД',
        'sum': Decimal('2523.08'),
        'new_note': 'Нестле: 9000006156, 6157',
        'desc': '2 фактури за Нестле (сбор 2523.08 €)'
    },
    # 15. Виста АВТ АД (Друго: 5.04 + 22.45 + 110.00 = 137.49 €)
    {
        'salenum': 77343,
        'partner': 'Виста АВТ АД',
        'sum': Decimal('137.49'),
        'new_note': '9000006061-6063 (3 бр.)',
        'desc': '3 фактури Виста друго (сбор 137.49 €)'
    },
    # 16. Виста АВТ АД (Джей ди: 134.06 + 440 + 110 + 245 = 929.06 €)
    {
        'salenum': 77344,
        'partner': 'Виста АВТ АД',
        'sum': Decimal('929.06'),
        'new_note': 'Джей ди: 6064-6067 (4 бр.)',
        'desc': '4 фактури за Джей ди (сбор 929.06 €)'
    },
    # 17. Виста АВТ АД (Лейс: 190.37 + 110 + 660 + 330 = 1290.37 €)
    {
        'salenum': 77345,
        'partner': 'Виста АВТ АД',
        'sum': Decimal('1290.37'),
        'new_note': 'Лейс: 6068,6069,6072,6073',
        'desc': '4 фактури за Лейс (сбор 1290.37 €)'
    },
    # 18. Виста АВТ АД (ТАТ: 990.00 + 331.53 = 1321.53 €)
    {
        'salenum': 77346,
        'partner': 'Виста АВТ АД',
        'sum': Decimal('1321.53'),
        'new_note': 'ТАТ: 9000006070, 6071',
        'desc': '2 фактури за ТАТ (сбор 1321.53 €)'
    },
    # 19. Виста АВТ АД (Фереро: 448.41 + 110.00 = 558.41 €)
    {
        'salenum': 77347,
        'partner': 'Виста АВТ АД',
        'sum': Decimal('558.41'),
        'new_note': 'Фереро: 9000006074, 6075',
        'desc': '2 фактури за Фереро (сбор 558.41 €)'
    },
    # 20. Амперел ООД (К1: 180.72 €)
    {
        'salenum': 77317,
        'partner': 'Амперел ООД',
        'sum': Decimal('180.72'),
        'new_note': 'К1: 9000006038',
        'desc': '1 фактура за К1 (180.72 €)'
    },
    # 21. Амперел ООД (Храна: 759.77 €)
    {
        'salenum': 77318,
        'partner': 'Амперел ООД',
        'sum': Decimal('759.77'),
        'new_note': 'Храна: 9000006039',
        'desc': '1 фактура за Храна (759.77 €)'
    },
    # 22. ЕМГ Комерсиал ООД (Девин: 750.82 €)
    {
        'salenum': 77352,
        'partner': 'ЕМГ Комерсиал ООД',
        'sum': Decimal('750.82'),
        'new_note': 'Девин: 9000006079',
        'desc': '1 фактура за Девин (750.82 €)'
    },
    # 23. Яско Комерс ЕООД (Чио: 492.54 €)
    {
        'salenum': 77431,
        'partner': 'Яско Комерс ЕООД',
        'sum': Decimal('492.54'),
        'new_note': 'Чио: 9000006207',
        'desc': '1 фактура за Чио (492.54 €)'
    }
]

def run_consolidated_import(commit=False):
    print("=" * 80)
    print("ИМПОРТ НА НОМЕРА НА ФАКТУРИ ЗА СБОРНИ / РАЗБИТИ ЗАПИСИ В МИСТРАЛ ERP")
    print("=" * 80)
    mode_str = "🔴 LIVE COMMIT (РЕАЛЕН ЗАПИС)" if commit else "🟡 DRY-RUN (СИМУЛАЦИЯ - БЕЗ ПРОМЕНИ)"
    print(f"Режим на работа: {mode_str}")
    print(f"Брой сметки за актуализация: {len(CONSOLIDATED_UPDATES)} бр.")
    print("=" * 80)

    for idx, u in enumerate(CONSOLIDATED_UPDATES, 1):
        print(f"{idx:>2}. Сметка {u['salenum']} | {u['partner']:<25} | {u['sum']:>9.2f} € | Нова: '{u['new_note']}'")

    if not commit:
        print("\n" + "=" * 80)
        print("🟡 DRY-RUN ПРИКЛЮЧИ: Не са направени промени в базите данни.")
        print("За да запишете промените, стартирайте с: python import_consolidated_notes.py --commit")
        print("=" * 80)
        return

    # LIVE COMMIT
    print("\n" + "=" * 80)
    print("🔴 ИЗПЪЛНЕНИЕ НА LIVE COMMIT В MONTEVIDEO.FDB И CS.FDB...")
    print("=" * 80)

    conn_mv = firebirdsql.connect(**MV_CONFIG)
    cur_mv = conn_mv.cursor()

    conn_cs = firebirdsql.connect(**CS_CONFIG)
    cur_cs = conn_cs.cursor()

    updated_mv = 0
    updated_cs = 0

    try:
        for u in CONSOLIDATED_UPDATES:
            salenum = u['salenum']
            new_note = u['new_note']

            cur_mv.execute("""
                UPDATE DEFERREDPAYMENT
                SET DEFERREDPAYMENTNOTE1 = ?
                WHERE LOCATIONID = 7 AND SALENUM = ?
            """, (new_note, salenum))
            if cur_mv.rowcount > 0:
                updated_mv += 1

            cur_cs.execute("""
                UPDATE DEFERREDPAYMENT
                SET DEFERREDPAYMENTNOTE1 = ?
                WHERE LOCATIONID = 7 AND SALENUM = ?
            """, (new_note, salenum))
            if cur_cs.rowcount > 0:
                updated_cs += 1

        conn_mv.commit()
        conn_cs.commit()
        print(f"\n✅ УСПЕШНО АКТУАЛИЗИРАНИ {updated_mv} СБОРНИ ЗАПИСА В MONTEVIDEO.FDB!")
        print(f"✅ УСПЕШНО АКТУАЛИЗИРАНИ {updated_cs} СБОРНИ ЗАПИСА В CS.FDB!")
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
    parser = argparse.ArgumentParser(description="Импорт на номера на фактури за сборни записи в Мистрал ERP")
    parser.add_argument('--commit', action='store_true', help="Извършва реалния запис в базата данни")
    args = parser.parse_args()

    run_consolidated_import(commit=args.commit)
