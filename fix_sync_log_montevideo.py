#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Корекция на SYNCHRONIZATIONLOG в Монтевидео (ЦО).
Премахва микросекундите от TimeDoct, DateDoct и DateSaved, които причиняват грешка
в Mistral SyncManager (Delphi StrToTime / EConvertError).
"""

import sys
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

def main():
    conn = firebirdsql.connect(**DB_CONFIG_MV)
    cur = conn.cursor()

    print("=" * 80)
    print("КОРЕКЦИЯ НА SYNCHRONIZATIONLOG В МОНТЕВИДЕО ЗА ИЗЧИСТВАНЕ НА СИНХРОНИЗАЦИЯТА")
    print("=" * 80)

    cur.execute("""
        SELECT ID, SYNCHROLOG
        FROM SYNCHRONIZATIONLOG
        WHERE ID >= 317400
        ORDER BY ID
    """)

    rows = cur.fetchall()
    updated_count = 0

    for r in rows:
        log_id = r[0]
        txt = r[1]
        if hasattr(txt, 'read'):
            txt = txt.read()
        if isinstance(txt, bytes):
            txt = txt.decode('windows-1251', errors='replace')

        if '11:36:53.4237' in txt:
            # 1. За TimeDoct: заменяме невалидния час с празна стойност (\x1e - NULL token в протокола на Мистрал)
            # точно както е в нативните записи на Mistral (напр. Lulin7)
            new_txt = txt.replace('TimeDoct\x1d11:36:53.4237\x1d', 'TimeDoct\x1d\x1e\x1d')

            # 2. За DateDoct и DateSaved: форматираме до стандартен timestamp (с .0000 вместо .4237)
            new_txt = new_txt.replace('11:36:53.4237', '11:36:53.0000')

            cur.execute("""
                UPDATE SYNCHRONIZATIONLOG
                SET SYNCHROLOG = ?
                WHERE ID = ?
            """, (new_txt, log_id))

            updated_count += 1
            print(f"  ✓ Коригиран LogId {log_id}")

    conn.commit()
    print(f"\n🎉 УСПЕШЕН COMMIT! Общо коригирани записи: {updated_count} бр.")

    # Проверка дали са останали записи с грешния формат
    cur.execute("""
        SELECT ID, SYNCHROLOG
        FROM SYNCHRONIZATIONLOG
        WHERE ID >= 317400
        ORDER BY ID
    """)
    remaining = 0
    for r in cur.fetchall():
        txt = r[1]
        if hasattr(txt, 'read'):
            txt = txt.read()
        if isinstance(txt, bytes):
            txt = txt.decode('windows-1251', errors='replace')
        if '11:36:53.4237' in txt:
            remaining += 1

    if remaining == 0:
        print("✅ Всички записи в SYNCHRONIZATIONLOG са напълно изчистени и съвместими с Mistral SyncManager!")
    else:
        print(f"⚠️ Остават {remaining} записа с проблемния формат!")

    conn.close()

if __name__ == '__main__':
    main()
