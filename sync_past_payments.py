#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
=============================================================================
СИНХРОНИЗАЦИЯ НА МИНАЛИ АВТОМАТИЧНИ РАЗПЛАЩАНИЯ КЪМ CS.FDB (ЦЕНТРАЛЕН СКЛАД)
Верига магазини „Бакалия“ | ЦЕНТРАЛЕН ОФИС — МОНТЕВИДЕО
=============================================================================
Този скрипт намира всички разплащания, направени автоматично в МОНТЕВИДЕО,
които все още липсват в опашката за синхронизация (SYNCHRONIZATIONLOG),
и генерира репликационни команди SAVEPAYRAZCR.
При следващата синхронизация на SyncManager те автоматично се изпълняват в CS.FDB
и актуализират салдата на доставчиците!
=============================================================================
"""

import sys
import datetime
from decimal import Decimal
import firebirdsql

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


def sync_past_payments(dry_run=False):
    print("=" * 80)
    print("СТАРТИРАНЕ НА ПРОЦЕС ПО СИНХРОНИЗАЦИЯ НА РАЗПЛАЩАНИЯТА КЪМ CS.FDB")
    print("=" * 80)

    conn_mv = firebirdsql.connect(**MV_CONFIG)
    cur_mv = conn_mv.cursor()

    conn_cs = firebirdsql.connect(**CS_CONFIG)
    cur_cs = conn_cs.cursor()

    # 1. Извличане на всички транзакции от потребител Тонев (ID 33) от 14.09 насам
    cur_mv.execute("""
        SELECT t.ID, t.SALENUM, t.SUMPAYED, t.DATEPAYMENT, p.PARTNERNAME,
               t.DOCUMENTNUM, t.DOCUMENTDATE, t.SALEDATE,
               t.LASTDELIVERYPRICESUM, t.AVGDELIVERYPRICESUM, t.REPORTINGDATE
        FROM DEFERREDPAYMENTTRANSACTION t
        LEFT JOIN PARTNERNAME p ON t.PARTNERNAMEID = p.ID
        WHERE t.DATEPAYMENT >= '2026-09-14' AND t.USERID = 33
        ORDER BY t.ID
    """)
    mv_txs = cur_mv.fetchall()
    print(f"Общ брой транзакции в Монтевидео (потребител Тонев от 14.09.2026): {len(mv_txs)}")

    to_sync = []
    for tx in mv_txs:
        salenum = tx[1]
        amount = tx[2]
        tx_date = tx[3]
        partner_name = tx[4] if tx[4] else ''

        # Проверка дали вече е отразено в CS.FDB
        cur_cs.execute("""
            SELECT ID FROM DEFERREDPAYMENTTRANSACTION
            WHERE SALENUM = ? AND DATEPAYMENT >= '2026-09-14'
        """, (salenum,))
        cs_done = cur_cs.fetchall()

        if cs_done:
            continue  # Вече е отразено в CS.FDB

        # Проверка дали вече чака в SYNCHRONIZATIONLOG в Монтевидео
        cur_mv.execute("""
            SELECT ID FROM SYNCHRONIZATIONLOG
            WHERE SYNCHROLOG LIKE ?
        """, (f'%SmetkaNo\x1d{salenum}\x1d%',))
        in_sync_queue = cur_mv.fetchall()

        if in_sync_queue:
            print(f"  ⏳ Сметка {salenum} ({amount:>8.2f} €) | {partner_name:<30} -> ВЕЧЕ ЧАКА В ОПАШКАТА (ID: {[q[0] for q in in_sync_queue]})")
            continue

        to_sync.append(tx)

    print(f"\nНамерени {len(to_sync)} несинхронизирани разплащания за добавяне в SYNCHRONIZATIONLOG.")

    if not to_sync:
        print("\nНяма чакащи за добавяне разплащания. Всичко е синхронизирано или вече чака в опашката!")
        conn_mv.close()
        conn_cs.close()
        return

    if dry_run:
        print("\n[DRY RUN] Списък на разплащанията, които ще бъдат добавени:")
        for tx in to_sync:
            print(f"  -> Tx {tx[0]} | Сметка {tx[1]} | {tx[2]:>8.2f} € | {tx[4]} | Дата: {tx[3]}")
        print("\n[DRY RUN] Запис не е извършен.")
        conn_mv.close()
        conn_cs.close()
        return

    # 2. Генериране и вмъкване на SAVEPAYRAZCR в SYNCHRONIZATIONLOG
    inserted_count = 0
    total_amount = Decimal('0.00')

    for tx in to_sync:
        tx_id, salenum, amount, tx_date, partner_name, docnum, docdate, saledate, lastprice, avgprice, rep_date = tx
        
        cur_mv.execute("SELECT GEN_ID(G_SYNCHRONIZATIONLOG, 1) FROM RDB$DATABASE")
        next_id = cur_mv.fetchone()[0]

        reporting_date_str = rep_date.strftime('%Y-%m-%d') if rep_date else tx_date.strftime('%Y-%m-%d')
        dt_pay_str = tx_date.strftime('%Y-%m-%d %H:%M:%S.0000')
        smetka_date_str = saledate.strftime('%Y-%m-%d %H:%M:%S.0000') if saledate else '\x1e'
        doc_date_str = docdate.strftime('%Y-%m-%d') if docdate else '\x1e'
        doc_no_str = str(docnum) if docnum else '\x1e'
        last_price_str = f"{lastprice:.3f}" if lastprice else f"{amount:.3f}"
        avg_price_str = f"{avgprice:.3f}" if avgprice else "1.000"

        fields = [
            ('ObektId', '7'),
            ('SmetkaNo', str(salenum)),
            ('ActionId', '3'),
            ('TypeDb', '2'),
            ('PaySuma', f"{amount:.3f}"),
            ('ObektName', 'МОНТЕВИДЕО'),
            ('dtPay', dt_pay_str),
            ('PartnerName', partner_name),
            ('PartnerStoreName', '\x1e'),
            ('PartnerEmployeeName', '\x1e'),
            ('Oper', 'Тонев'),
            ('SumRazcr', '0.000'),
            ('DocNo', doc_no_str),
            ('DocDate', doc_date_str),
            ('DocType', 'Фактура'),
            ('SmetkaDate', smetka_date_str),
            ('chrStorno', '0'),
            ('MoneyInName', 'ДСК Бакалия'),
            ('Log', '0'),
            ('Note', '\x1e'),
            ('PayDate', '\x1e'),
            ('PayDocNo', '\x1e'),
            ('PayDocSuma', '\x1e'),
            ('SumLastDoctPrice', last_price_str),
            ('SumAvgDoctPrice', avg_price_str),
            ('Kaca', '0'),
            ('RazcrNote1', '\x1c'),
            ('RazcrNote2', '\x1c'),
            ('ReportingDate', reporting_date_str),
            ('MoneyInObektId', '7'),
            ('CalcNal', '1')
        ]

        payload = 'SAVEPAYRAZCR'
        for k, v in fields:
            payload += f'\x1d{k}\x1d{v}'
        payload += '\x1d'

        now = datetime.datetime.now()
        cur_mv.execute("""
            INSERT INTO SYNCHRONIZATIONLOG (ID, LOCATIONID, DATESAVE, SYNCHROLOG)
            VALUES (?, 7, ?, ?)
        """, (next_id, now, payload))

        inserted_count += 1
        total_amount += amount
        print(f"  ✅ Добавено в SYNCHRONIZATIONLOG: ID {next_id} | Сметка {salenum} ({amount:>8.2f} €) | {partner_name}")

    conn_mv.commit()
    print("\n" + "=" * 80)
    print(f"🎉 УСПЕШНО ДОБАВЕНИ {inserted_count} РАЗПЛАЩАНИЯ ЗА ОБЩО {total_amount:.2f} € В ОПАШКАТА НА МОНТЕВИДЕО!")
    print("При следващата синхронизация те ще се репликират към CS.FDB и ще изчистят салдата!")
    print("=" * 80)

    conn_mv.close()
    conn_cs.close()


if __name__ == '__main__':
    dry_run = '--dry-run' in sys.argv
    sync_past_payments(dry_run=dry_run)
