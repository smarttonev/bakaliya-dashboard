#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
Корекция на транзакция от 05.10.2026 за Наем чужденци (Мартин Петев Ботев)
- Смяна на DEFERREDPAYMENTTRANSACTION ID 77034 от Внасяне (OPERATIONTYPE=8, OPERATIONTYPEID=3) на Теглене (OPERATIONTYPE=7, OPERATIONTYPEID=2)
- Смяна на MONEYFLOW ID 98434 от Внасяне (OPERATIONID=10, OPERATIONTYPE=8) на Теглене (OPERATIONID=9, OPERATIONTYPE=7)
- Корекция на касовата наличност в MONEYFLOW и MONEYFLOWUNIT с -1000.00 € (отстраняване на +500 вместо -500)
"""

import sys
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

def apply_fix():
    conn = firebirdsql.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        print("1. Проверка на DEFERREDPAYMENTTRANSACTION ID 77034...")
        cur.execute("SELECT ID, SUMPAYED, OPERATIONTYPE, OPERATIONTYPEID FROM DEFERREDPAYMENTTRANSACTION WHERE ID = 77034")
        tx_row = cur.fetchone()
        print(f"   Преди: {tx_row}")

        cur.execute("""
            UPDATE DEFERREDPAYMENTTRANSACTION
            SET OPERATIONTYPE = 7, OPERATIONTYPEID = 2
            WHERE ID = 77034
        """)
        print("   -> Обновено: OPERATIONTYPE = 7, OPERATIONTYPEID = 2")

        print("\n2. Проверка на MONEYFLOW ID 98434...")
        cur.execute("SELECT ID, OPERATIONSUM, OPERATIONID, OPERATIONTYPE, AVAILABILITY FROM MONEYFLOW WHERE ID = 98434")
        mf_row = cur.fetchone()
        print(f"   Преди: {mf_row}")

        cur.execute("""
            UPDATE MONEYFLOW
            SET OPERATIONID = 9, OPERATIONTYPE = 7
            WHERE ID = 98434
        """)
        print("   -> Обновено: OPERATIONID = 9, OPERATIONTYPE = 7")

        print("\n3. Корекция на AVAILABILITY за редовете в MONEYFLOW след 98434...")
        cur.execute("""
            SELECT ID, AVAILABILITY
            FROM MONEYFLOW
            WHERE ID >= 98434 AND MONEYFLOWUNITID = 52
            ORDER BY ID ASC
        """)
        mf_rows = cur.fetchall()
        for r in mf_rows:
            mf_id = r[0]
            old_avail = r[1]
            new_avail = old_avail - Decimal('1000.00')
            cur.execute("UPDATE MONEYFLOW SET AVAILABILITY = ? WHERE ID = ?", (new_avail, mf_id))
            print(f"   MF ID {mf_id}: {old_avail} -> {new_avail}")

        print("\n4. Корекция на текуща наличност в MONEYFLOWUNIT (Каса ДСК Бакалия)...")
        cur.execute("SELECT ID, NAME, AVAILABILITY FROM MONEYFLOWUNIT WHERE ID = 179607")
        desk_row = cur.fetchone()
        old_desk_avail = desk_row[2]
        new_desk_avail = old_desk_avail - Decimal('1000.00')
        cur.execute("UPDATE MONEYFLOWUNIT SET AVAILABILITY = ? WHERE ID = 179607", (new_desk_avail,))
        print(f"   MONEYFLOWUNIT 179607: {old_desk_avail} -> {new_desk_avail}")

        conn.commit()
        print("\n✅ УСПЕШЕН COMMIT В MONTEVIDEO.FDB!")

    except Exception as e:
        conn.rollback()
        print(f"\n❌ ГРЕШКА при корекцията: {e}")
        raise
    finally:
        conn.close()

if __name__ == '__main__':
    apply_fix()
