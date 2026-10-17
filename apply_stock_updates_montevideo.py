#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Актуализиране на складовите наличности (MATERIAL.QTY / GETARTIKULNAL) в Монтевидео (ЦО)
за доставените артикули от импортираните доставки № 112252 до № 112271.
"""

import sys
import datetime
from decimal import Decimal
import firebirdsql

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG_MV = {
    'host': '192.168.5.254',
    'database': 'E:/db/MONTEVIDEO.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

LOCATION_ID = 7
STORAGE_ID = 1
MIN_OP = 112252
MAX_OP = 112271

def main():
    conn = firebirdsql.connect(**DB_CONFIG_MV)
    cur = conn.cursor()
    now = datetime.datetime.now()

    print("=" * 80)
    print("АКТУАЛИЗИРАНЕ НА НАЛИЧНОСТИТЕ В БАЗАТА ДАННИ НА МОНТЕВИДЕО (ОБЕКТ 7, СКЛАД 1)")
    print(f"Период на операции: № {MIN_OP} - № {MAX_OP} | Дата/час: {now}")
    print("=" * 80)

    # 1. Извличане на доставените количества от MATERIALQTYLOG
    cur.execute("""
        SELECT m.MATERIALCODE, m.MATERIAL, m.QTY AS CURRENT_QTY,
               SUM(log.QTY) AS DELIVERED_QTY
        FROM MATERIALQTYLOG log
        JOIN MATERIAL m ON log.LOCATIONID = m.LOCATIONID AND log.STORAGEID = m.STORAGEID AND log.MATERIALCODE = m.MATERIALCODE
        WHERE log.LOCATIONID = ? AND log.STORAGEID = ? 
          AND log.OPERAIONNUM >= ? AND log.OPERAIONNUM <= ?
        GROUP BY m.MATERIALCODE, m.MATERIAL, m.QTY
        ORDER BY m.MATERIALCODE
    """, (LOCATION_ID, STORAGE_ID, MIN_OP, MAX_OP))

    items = cur.fetchall()
    if not items:
        print("❌ Не са намерени записи в MATERIALQTYLOG за посочените операции!")
        return

    print(f"\nНамерени {len(items)} артикула за актуализиране:")
    print(f"{'Код':<8} | {'Артикул':<28} | {'Преди':<12} | {'Доставено':<14} | {'Нова наличност':<14}")
    print("-" * 84)

    updates = []
    for r in items:
        code = r[0]
        name = r[1]
        current_qty = Decimal(str(r[2]))
        delivered_qty = Decimal(str(r[3]))
        new_qty = current_qty + delivered_qty
        updates.append((code, name, current_qty, delivered_qty, new_qty))
        print(f"{code:<8} | {name:<28} | {current_qty:>12.3f} | {delivered_qty:>14.3f} | {new_qty:>14.3f}")

    # 2. Изпълнение на актуализацията в MATERIAL
    print("\nИзпълнение на UPDATE MATERIAL...")
    for code, name, curr, deliv, new_q in updates:
        cur.execute("""
            UPDATE MATERIAL
            SET QTY = ?,
                LASTDELIVERYDATE = ?,
                LASTOPERATIONID = 2,
                LASTEDITDATE = ?
            WHERE LOCATIONID = ? AND STORAGEID = ? AND MATERIALCODE = ?
        """, (float(new_q), now, now, LOCATION_ID, STORAGE_ID, code))
        print(f"  ✅ Код {code} ({name}): QTY {curr:.3f} -> {new_q:.3f}")

    conn.commit()
    print("\n🎉 УСПЕШЕН COMMIT! Складовите наличности са актуализирани в Монтевидео.")

    # 3. Верификация чрез GETARTIKULNAL и SELECT от MATERIAL
    print("\n" + "=" * 80)
    print("ВЕРИФИКАЦИЯ НА НАЛИЧНОСТИТЕ В СИСТЕМАТА (GETARTIKULNAL / MATERIAL.QTY):")
    print("=" * 80)
    print(f"{'Код':<8} | {'Артикул':<28} | {'MATERIAL.QTY':<16} | {'GETARTIKULNAL':<16} | {'Статус':<10}")
    print("-" * 84)

    all_ok = True
    for code, name, _, _, expected_q in updates:
        cur.execute("SELECT QTY FROM MATERIAL WHERE LOCATIONID = ? AND STORAGEID = ? AND MATERIALCODE = ?", (LOCATION_ID, STORAGE_ID, code))
        mat_qty = cur.fetchone()[0]

        cur.execute("SELECT NAL FROM GETARTIKULNAL(?, ?)", (LOCATION_ID, code))
        sys_nal = cur.fetchone()[0]

        status = "OK ✅" if abs(Decimal(str(mat_qty)) - expected_q) < Decimal('0.001') and abs(Decimal(str(sys_nal)) - expected_q) < Decimal('0.001') else "MISMATCH ❌"
        if "MISMATCH" in status:
            all_ok = False
        print(f"{code:<8} | {name:<28} | {mat_qty:>16.3f} | {sys_nal:>16.3f} | {status:<10}")

    conn.close()
    if all_ok:
        print("\n✅ ВСИЧКИ НАЛИЧНОСТИ СА 100% ВЕРИФИЦИРАНИ И СЪОТВЕТСТВАТ НА ДОСТАВКИТЕ!")
    else:
        print("\n⚠️ ВНИМАНИЕ: Открити са разминавания при верификацията!")

if __name__ == '__main__':
    main()
