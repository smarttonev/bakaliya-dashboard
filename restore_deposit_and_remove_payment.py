#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
1. Възстановяване на отворения запис от 500 € в DEFERREDPAYMENT (Депозит Ботев 2-ри ап).
2. Премахване на плащането от -500 € (DEFERREDPAYMENTTRANSACTION ID 77034 и MONEYFLOW ID 98434).
3. Корекция на касовата наличност в MONEYFLOW и MONEYFLOWUNIT с +500.00 €.
"""

import sys
import datetime
from decimal import Decimal
import firebirdsql

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

DB_CONFIG = {
    'host': '192.168.5.254',
    'database': 'E:/db/MONTEVIDEO.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

def execute_restoration():
    conn = firebirdsql.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        # -------------------------------------------------------------
        # 1. Възстановяване в DEFERREDPAYMENT
        # -------------------------------------------------------------
        print("1. Генериране на ново ID и възстановяване в DEFERREDPAYMENT...")
        cur.execute("SELECT GEN_ID(G_DEFERREDPAYMENT, 1) FROM RDB$DATABASE")
        new_dp_id = cur.fetchone()[0]

        cur.execute("""
            INSERT INTO DEFERREDPAYMENT (
                ID, LOCATIONID, PARTNERNUM, PARTNERLOCATIONNUM, PARTNERCONTACTNUM,
                SALENUM, SALEDATE, DPSUM, DATERETURN, USERNAME,
                OPERATIONTYPEID, DOCUMENTNUM, DOCUMENTDATE, DOCUMENTTYPE, DBTYPE,
                LASTDELIVERYPRICESUM, AVGDELIVERYPRICESUM,
                DEFERREDPAYMENTNOTE1, DEFERREDPAYMENTNOTE2
            ) VALUES (
                ?, 7, 392, -1, -1,
                76896, ?, ?, ?, 'Тонев',
                3, NULL, ?, 'без документ', 2,
                ?, ?,
                'Депозит Ботев 2-ри ап', ''
            )
        """, (
            new_dp_id,
            datetime.datetime(2026, 7, 31, 14, 31, 40),
            Decimal('500.000'),
            datetime.date(2026, 7, 31),
            datetime.date(2026, 7, 31),
            Decimal('500.00'),
            Decimal('500.00')
        ))
        print(f"   -> Успешно възстановен запис в DEFERREDPAYMENT с ID {new_dp_id} за 500.00 € ('Депозит Ботев 2-ри ап')")

        # -------------------------------------------------------------
        # 2. Премахване на плащането от 500 €
        # -------------------------------------------------------------
        print("\n2. Изтриване на транзакцията от DEFERREDPAYMENTTRANSACTION (ID 77034)...")
        cur.execute("DELETE FROM DEFERREDPAYMENTTRANSACTION WHERE ID = 77034")
        print("   -> Изтрит DEFERREDPAYMENTTRANSACTION ID 77034")

        print("\n3. Изтриване на касовото движение от MONEYFLOW (ID 98434)...")
        cur.execute("DELETE FROM MONEYFLOW WHERE ID = 98434")
        print("   -> Изтрит MONEYFLOW ID 98434")

        # -------------------------------------------------------------
        # 3. Корекция на касовата наличност в MONEYFLOW и MONEYFLOWUNIT
        # -------------------------------------------------------------
        print("\n4. Корекция на AVAILABILITY за последващите записи в MONEYFLOW...")
        cur.execute("""
            SELECT ID, AVAILABILITY
            FROM MONEYFLOW
            WHERE ID > 98434 AND MONEYFLOWUNITID = 52
            ORDER BY ID ASC
        """)
        mf_rows = cur.fetchall()
        for r in mf_rows:
            mf_id = r[0]
            old_avail = r[1]
            new_avail = old_avail + Decimal('500.00')
            cur.execute("UPDATE MONEYFLOW SET AVAILABILITY = ? WHERE ID = ?", (new_avail, mf_id))
            print(f"   MF ID {mf_id}: {old_avail} -> {new_avail}")

        print("\n5. Корекция на наличност в MONEYFLOWUNIT (Каса ДСК Бакалия)...")
        cur.execute("SELECT ID, NAME, AVAILABILITY FROM MONEYFLOWUNIT WHERE ID = 179607")
        desk_row = cur.fetchone()
        old_desk_avail = desk_row[2]
        new_desk_avail = old_desk_avail + Decimal('500.00')
        cur.execute("UPDATE MONEYFLOWUNIT SET AVAILABILITY = ? WHERE ID = 179607", (new_desk_avail,))
        print(f"   MONEYFLOWUNIT 179607: {old_desk_avail} -> {new_desk_avail}")

        conn.commit()
        print("\n✅ УСПЕШЕН COMMIT В MONTEVIDEO.FDB!")

    except Exception as e:
        conn.rollback()
        print(f"\n❌ ГРЕШКА: {e}")
        raise
    finally:
        conn.close()

if __name__ == '__main__':
    execute_restoration()
