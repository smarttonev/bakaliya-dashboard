#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
================================================================================
Mistral ERP — Web Импортьор на Промоционални продажни цени (ПЦ)
             и Заключени договорени покупни цени (ЗКП)
================================================================================
Интерфейс: Streamlit Web App
База данни: 192.168.5.254:E:\DB\CS.FDB (Централен склад)
Обекти: Централен склад (1) + 11 магазина:
        5 (Надежда), 6 (Арсеналски), 18 (Лозенец), 19 (Свобода),
        22 (Обеля 2), 27 (Манастирски ливади), 30 (Люлин 7), 31 (Банкя),
        32 (Редута), 33 (Хали Окръжна), 35 (Велчо Атанасов).
Изключен обект: 7 (Монтевидео)
Склад: Склад 1 (STORAGEID = 1)
"""

import os
import sys

# Ensure UTF-8 output on Windows terminal
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import io
import math
import tempfile
import datetime
import subprocess
import pandas as pd
import streamlit as st

# Опит за зареждане на firebirdsql
try:
    import firebirdsql
    HAS_FIREBIRDSQL = True
except ImportError:
    HAS_FIREBIRDSQL = False

# ------------------------------------------------------------------------------
# Конфигурация на приложението и базата данни
# ------------------------------------------------------------------------------
DB_CONFIG = {
    'host': '192.168.5.254',
    'database': r'E:\DB\CS.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

ISQL_PATH = r"C:\Program Files\Firebird\Firebird_2_5\bin\isql.exe"
DB_ISQL_PATH = "192.168.5.254:E:\\DB\\CS.FDB"

CENTRAL_LOCATION = 1
REMOTE_LOCATIONS = [5, 6, 18, 19, 22, 27, 30, 31, 32, 33, 35]
TARGET_LOCATIONS = [CENTRAL_LOCATION] + REMOTE_LOCATIONS

LOCATION_NAMES = {
    1: "ЦЕНТРАЛЕН СКЛАД",
    5: "НАДЕЖДА",
    6: "АРСЕНАЛСКИ",
    18: "ЛОЗЕНЕЦ",
    19: "СВОБОДА",
    22: "ОБЕЛЯ 2",
    27: "МАНАСТИРСКИ ЛИВАДИ",
    30: "ЛЮЛИН 7",
    31: "БАНКЯ",
    32: "РЕДУТА",
    33: "ХАЛИ ОКРЪЖНА",
    35: "ВЕЛЧО АТАНАСОВ"
}

# ------------------------------------------------------------------------------
# Помощни функции за Firebird
# ------------------------------------------------------------------------------
def get_db_connection():
    """Създава връзка към Firebird базата през firebirdsql с автоматично тестване на локален и мрежов адрес."""
    if not HAS_FIREBIRDSQL:
        return None
    hosts = [DB_CONFIG['host']]
    for alt in ['127.0.0.1', 'localhost']:
        if alt not in hosts:
            hosts.append(alt)
            
    for h in hosts:
        try:
            cfg = dict(DB_CONFIG)
            cfg['host'] = h
            conn = firebirdsql.connect(**cfg)
            return conn
        except Exception:
            continue
    return None

def run_isql_script(sql_script_text, timeout=60):
    """Изпълнява SQL скрипт през Firebird isql.exe при необходимост."""
    temp_dir = tempfile.gettempdir()
    temp_sql_path = os.path.join(temp_dir, f"mistral_web_{os.getpid()}_{int(datetime.datetime.now().timestamp())}.sql")
    try:
        with open(temp_sql_path, "w", encoding="cp1251", errors="replace") as f:
            f.write(sql_script_text)
        
        cmd = [ISQL_PATH, "-u", DB_CONFIG['user'], "-p", DB_CONFIG['password'], "-ch", "NONE", DB_ISQL_PATH, "-q", "-m", "-i", temp_sql_path]
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout)
        stdout = proc.stdout.decode("cp1251", errors="replace")
        stderr = proc.stderr.decode("cp1251", errors="replace")
        
        if proc.returncode != 0 and stderr.strip():
            raise RuntimeError(f"Firebird ISQL грешка:\n{stderr}")
        return stdout
    finally:
        if os.path.exists(temp_sql_path):
            try:
                os.remove(temp_sql_path)
            except Exception:
                pass

def fetch_db_article_info(codes, mode):
    """
    Извлича за всеки код на артикул:
    - Име в базата (GETARTIKULFIELDS / ARTIKUL)
    - Текуща базова цена (доставна за ЗКП, продажна за ПЦ)
    - Всички активни и съществуващи договорени / промо цени (all_existing) за засичане на дубликати и застъпвания
    """
    if not codes:
        return {}
    
    unique_codes = sorted(list(set(codes)))
    info_map = {}
    
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            
            # Четене на порции по 100 кода за максимална скорост и надеждност
            batch_size = 100
            for i in range(0, len(unique_codes), batch_size):
                batch = unique_codes[i:i + batch_size]
                in_clause = ",".join(str(c) for c in batch)
                
                # 1. Номенклатура и последна доставна цена през GETARTIKULFIELDS
                unions = " UNION ALL ".join(f"SELECT {c} AS CODE, ARTIKUL, LASTDOCTPRICE FROM GETARTIKULFIELDS(1, {c})" for c in batch)
                try:
                    cur.execute(unions)
                    for r in cur.fetchall():
                        aid = int(r[0])
                        art_name = r[1].strip() if r[1] else None
                        base_del = float(r[2]) if r[2] is not None else 0.0
                        if aid not in info_map:
                            info_map[aid] = {
                                "name": art_name,
                                "base_delivery": base_del,
                                "base_sale": None,
                                "last_price": None,
                                "last_start": None,
                                "last_end": None,
                                "all_existing": []
                            }
                except Exception:
                    # Fallback за единично извличане ако някой код в пакета даде грешка
                    for c in batch:
                        try:
                            cur.execute("SELECT ARTIKUL, LASTDOCTPRICE FROM GETARTIKULFIELDS(1, ?)", (c,))
                            row = cur.fetchone()
                            if row:
                                info_map[c] = {
                                    "name": row[0].strip() if row[0] else None,
                                    "base_delivery": float(row[1]) if row[1] is not None else 0.0,
                                    "base_sale": None,
                                    "last_price": None,
                                    "last_start": None,
                                    "last_end": None,
                                    "all_existing": []
                                }
                        except Exception:
                            pass
                
                # 2. Базова продажна цена от PRICELIST
                try:
                    cur.execute(f"SELECT MATERIALCODE, PRICE FROM PRICELIST WHERE LOCATIONID = 1 AND MATERIALCODE IN ({in_clause})")
                    for r in cur.fetchall():
                        aid = int(r[0])
                        if aid in info_map and r[1] is not None:
                            info_map[aid]["base_sale"] = float(r[1])
                except Exception:
                    pass
                
                # 3. Съществуващи договорени / промо цени
                if mode == "ZKP":
                    try:
                        cur.execute(f"""
                            SELECT MATERIALCODE, AGREEDDELIVERYPRICE, ACTIVEFROM, ACTIVETO 
                            FROM AGREEDDELIVERYPRICE 
                            WHERE LOCATIONID = 1 AND STORAGEID = 1 AND MATERIALCODE IN ({in_clause})
                            ORDER BY ACTIVEFROM DESC, ID DESC
                        """)
                        for r in cur.fetchall():
                            aid = int(r[0])
                            if aid in info_map:
                                p_val = float(r[1]) if r[1] is not None else None
                                s_val = str(r[2])[:19] if r[2] is not None else ""
                                e_val = str(r[3])[:19] if r[3] is not None else ""
                                if info_map[aid]["last_price"] is None:
                                    info_map[aid]["last_price"] = p_val
                                    info_map[aid]["last_start"] = s_val
                                    info_map[aid]["last_end"] = e_val
                                info_map[aid]["all_existing"].append({
                                    "price": p_val,
                                    "start": s_val,
                                    "end": e_val,
                                    "pct": None
                                })
                    except Exception:
                        pass
                else: # PC
                    try:
                        cur.execute(f"""
                            SELECT SALEMATERIALCODE, SALEPRICE, STARTDATE, ENDDATE, PERCENTAGEDISCOUNT 
                            FROM PROMOTIONSALEPRICE 
                            WHERE LOCATIONID = 1 AND SALEMATERIALCODE IN ({in_clause})
                            ORDER BY STARTDATE DESC, ID DESC
                        """)
                        for r in cur.fetchall():
                            aid = int(r[0])
                            if aid in info_map:
                                p_val = float(r[1]) if r[1] is not None else None
                                s_val = str(r[2])[:19] if r[2] is not None else ""
                                e_val = str(r[3])[:19] if r[3] is not None else ""
                                pct_val = float(r[4]) if r[4] is not None else None
                                if info_map[aid]["last_price"] is None:
                                    info_map[aid]["last_price"] = p_val
                                    info_map[aid]["last_start"] = s_val
                                    info_map[aid]["last_end"] = e_val
                                info_map[aid]["all_existing"].append({
                                    "price": p_val,
                                    "start": s_val,
                                    "end": e_val,
                                    "pct": pct_val
                                })
                    except Exception:
                        pass
                        
            cur.close()
            conn.close()
            return info_map
        except Exception:
            if conn:
                try: conn.close()
                except Exception: pass
                
    # Fallback към isql.exe ако firebirdsql не е наличен или е върнал грешка
    if os.path.exists(ISQL_PATH):
        try:
            batch_size = 50
            for i in range(0, len(unique_codes), batch_size):
                batch = unique_codes[i:i + batch_size]
                in_clause = ",".join(str(c) for c in batch)
                
                # 1. Номенклатура през GETARTIKULFIELDS
                unions = " UNION ALL ".join(f"SELECT {c} AS CODE, ARTIKUL, LASTDOCTPRICE FROM GETARTIKULFIELDS(1, {c})" for c in batch)
                sql_art = f"SET HEADING OFF;\n{unions};\nQUIT;\n"
                out_art = run_isql_script(sql_art)
                for line in out_art.splitlines():
                    parts = line.strip().split()
                    if parts and parts[0].isdigit():
                        aid = int(parts[0])
                        if len(parts) >= 2:
                            last_price_candidate = None
                            try:
                                last_price_candidate = float(parts[-1])
                                name = " ".join(parts[1:-1])
                            except ValueError:
                                name = " ".join(parts[1:])
                            if name and name.upper() != "<NULL>":
                                if aid not in info_map:
                                    info_map[aid] = {
                                        "name": name,
                                        "base_delivery": last_price_candidate,
                                        "base_sale": None,
                                        "last_price": None,
                                        "last_start": None,
                                        "last_end": None,
                                        "all_existing": []
                                    }
                                    
                # 2. Базови продажни цени от PRICELIST
                sql_price = f"SET HEADING OFF;\nSELECT MATERIALCODE, PRICE FROM PRICELIST WHERE LOCATIONID = 1 AND MATERIALCODE IN ({in_clause});\nQUIT;\n"
                try:
                    out_p = run_isql_script(sql_price)
                    for line in out_p.splitlines():
                        parts = line.strip().split()
                        if len(parts) >= 2 and parts[0].isdigit():
                            aid = int(parts[0])
                            if aid in info_map:
                                try: info_map[aid]["base_sale"] = float(parts[1])
                                except ValueError: pass
                except Exception:
                    pass
                    
                # 3. Съществуващи промоции/договорени цени за засичане на дубликати и застъпвания
                if mode == "ZKP":
                    sql_zkp = f"""
                    SET HEADING OFF;
                    SELECT MATERIALCODE, AGREEDDELIVERYPRICE, ACTIVEFROM, ACTIVETO 
                    FROM AGREEDDELIVERYPRICE 
                    WHERE LOCATIONID = 1 AND STORAGEID = 1 AND MATERIALCODE IN ({in_clause})
                    ORDER BY ACTIVEFROM DESC;
                    QUIT;
                    """
                    try:
                        out_zkp = run_isql_script(sql_zkp)
                        for line in out_zkp.splitlines():
                            parts = line.strip().split()
                            if len(parts) >= 6 and parts[0].isdigit():
                                aid = int(parts[0])
                                if aid in info_map:
                                    p_val = float(parts[1])
                                    s_val = f"{parts[2]} {parts[3]}"[:19]
                                    e_val = f"{parts[4]} {parts[5]}"[:19]
                                    if info_map[aid]["last_price"] is None:
                                        info_map[aid]["last_price"] = p_val
                                        info_map[aid]["last_start"] = s_val
                                        info_map[aid]["last_end"] = e_val
                                    info_map[aid]["all_existing"].append({
                                        "price": p_val,
                                        "start": s_val,
                                        "end": e_val,
                                        "pct": None
                                    })
                    except Exception:
                        pass
                else: # PC
                    sql_pc = f"""
                    SET HEADING OFF;
                    SELECT SALEMATERIALCODE, SALEPRICE, STARTDATE, ENDDATE, COALESCE(PERCENTAGEDISCOUNT, 0)
                    FROM PROMOTIONSALEPRICE 
                    WHERE LOCATIONID = 1 AND SALEMATERIALCODE IN ({in_clause})
                    ORDER BY STARTDATE DESC;
                    QUIT;
                    """
                    try:
                        out_pc = run_isql_script(sql_pc)
                        for line in out_pc.splitlines():
                            parts = line.strip().split()
                            if len(parts) >= 6 and parts[0].isdigit():
                                aid = int(parts[0])
                                if aid in info_map:
                                    p_val = float(parts[1])
                                    s_val = f"{parts[2]} {parts[3]}"[:19]
                                    e_val = f"{parts[4]} {parts[5]}"[:19]
                                    pct_val = float(parts[6]) if len(parts) > 6 else 0.0
                                    if info_map[aid]["last_price"] is None:
                                        info_map[aid]["last_price"] = p_val
                                        info_map[aid]["last_start"] = s_val
                                        info_map[aid]["last_end"] = e_val
                                    info_map[aid]["all_existing"].append({
                                        "price": p_val,
                                        "start": s_val,
                                        "end": e_val,
                                        "pct": pct_val if pct_val > 0 else None
                                    })
                    except Exception:
                        pass
            return info_map
        except Exception:
            pass
            
    return info_map

def parse_date(val, is_end=False):
    """Парсира дати от Excel (текст или datetime) и връща низ с час."""
    if pd.isna(val) or val is None or str(val).strip() == "":
        raise ValueError("Липсваща дата")
    
    time_part = "23:59:59" if is_end else "00:00:00"
    
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        d_str = val.strftime("%Y-%m-%d")
        return f"{d_str} {time_part}"
    
    s = str(val).strip()
    for fmt in ["%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y.%m.%d", "%d-%m-%Y"]:
        try:
            dt = datetime.datetime.strptime(s.split()[0], fmt)
            return f"{dt.strftime('%Y-%m-%d')} {time_part}"
        except ValueError:
            pass
            
    try:
        dt = pd.to_datetime(s, dayfirst=True)
        return f"{dt.strftime('%Y-%m-%d')} {time_part}"
    except Exception:
        raise ValueError(f"Неразпознат формат: '{s}'")

def check_period_overlap(start1_dt, end1_dt, start2_dt, end2_dt):
    """
    Проверява дали два интервала от дати се застъпват (припокриват).
    Периодите се застъпват ако: start1 <= end2 AND end1 >= start2
    """
    s1 = str(start1_dt)[:10]
    e1 = str(end1_dt)[:10]
    s2 = str(start2_dt)[:10]
    e2 = str(end2_dt)[:10]
    return (s1 <= e2) and (e1 >= s2)

def execute_db_import(records, mode):
    """
    Изпълнява реалния импорт в базата данни на Централен склад за 12 локации със синхронизация.
    Гарантира стриктно спазване на 3-те правила:
    - Ситуация А: 100% съвпадение -> пропуска се автоматично (0 нови записа, без дубликати).
    - Ситуация Б: Същият период, различна цена -> първо изтрива старата промоция през DELETEPROMOTIONSALESPRICE с репликация.
    - Ситуация В: Застъпен период -> АБСОЛЮТНО БЛОКИРАН (хвърля грешка и прави пълен ROLLBACK).
    """
    records_to_process = [r for r in records if not r.get('skip', False)]
    if not records_to_process:
        return True, "Всички артикули вече са активни в Mistral с идентични параметри. Няма дублирани записи."

    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            inserted_count = 0
            updated_count = 0
            skipped_count = 0
            
            for r in records_to_process:
                code = r['code']
                if pd.isna(r.get('price')) or r.get('price') is None:
                    raise ValueError(f"Код {code} ({r.get('db_name', '')}) има липсваща цена в Excel файла!")
                price = round(float(r['price']), 2)
                if math.isnan(price) or math.isinf(price) or price <= 0:
                    raise ValueError(f"Код {code} ({r.get('db_name', '')}) има невалидна цена ({price} €)!")
                s_dt = r['start_dt']
                e_dt = r['end_dt']
                s_date_str = str(s_dt)[:10]
                e_date_str = str(e_dt)[:10]
                
                if mode == "ZKP":
                    # Засичане в AGREEDDELIVERYPRICE
                    cur.execute("""
                        SELECT ID, ACTIVEFROM, ACTIVETO, AGREEDDELIVERYPRICE
                        FROM AGREEDDELIVERYPRICE
                        WHERE LOCATIONID = 1 AND STORAGEID = 1 AND MATERIALCODE = ?
                    """, (code,))
                    existing_zkp = cur.fetchall()
                    
                    is_exact_match = False
                    overlap_conflict = None
                    exact_period_old_ids = []
                    
                    for ex_id, ex_s, ex_e, ex_p in existing_zkp:
                        if not ex_s or not ex_e: continue
                        ex_s_str = str(ex_s)[:10]
                        ex_e_str = str(ex_e)[:10]
                        if check_period_overlap(s_date_str, e_date_str, ex_s_str, ex_e_str):
                            if ex_s_str == s_date_str and ex_e_str == e_date_str:
                                if ex_p is not None and abs(float(ex_p) - price) < 0.001:
                                    is_exact_match = True
                                else:
                                    exact_period_old_ids.append(ex_id)
                            else:
                                overlap_conflict = f"{ex_s_str} - {ex_e_str} ({ex_p} €)"
                                break
                                
                    if overlap_conflict:
                        raise RuntimeError(f"Код {code} ({r.get('db_name', '')}) има застъпен период с договорена цена {overlap_conflict}! Записът е прекратен.")
                    if is_exact_match:
                        skipped_count += 1
                        continue
                        
                    # Изтриване на стари записи за същия период
                    for old_id in exact_period_old_ids:
                        for loc_id in TARGET_LOCATIONS:
                            cur.execute("""
                                DELETE FROM AGREEDDELIVERYPRICE 
                                WHERE LOCATIONID = ? AND STORAGEID = 1 AND ID = ?
                            """, (loc_id, old_id))
                        updated_count += 1
                            
                    # Нов запис
                    cur.execute("SELECT GEN_ID(G_AGREEDDELIVERYPRICE, 1) FROM RDB$DATABASE")
                    new_id = cur.fetchone()[0]
                    # Location 1 (LOG = '0')
                    cur.execute("EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES(1, 1, ?, ?, ?, ?, '0', ?, -1, -1, -1, 0.00)", (code, price, s_dt, e_dt, new_id))
                    # Remote locations (LOG = '1')
                    for loc_id in REMOTE_LOCATIONS:
                        cur.execute("EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES(?, 1, ?, ?, ?, ?, '1', ?, -1, -1, -1, 0.00)", (loc_id, code, price, s_dt, e_dt, new_id))
                    inserted_count += 1
                else: # PC
                    # Засичане в PROMOTIONSALEPRICE
                    cur.execute("""
                        SELECT ID, STARTDATE, ENDDATE, SALEPRICE, PERCENTAGEDISCOUNT
                        FROM PROMOTIONSALEPRICE
                        WHERE LOCATIONID = 1 AND SALEMATERIALCODE = ?
                    """, (code,))
                    existing_promos = cur.fetchall()
                    
                    is_exact_match = False
                    overlap_conflict = None
                    exact_period_old_ids = []
                    
                    for ex_id, ex_s, ex_e, ex_p, ex_pct in existing_promos:
                        if not ex_s or not ex_e: continue
                        ex_s_str = str(ex_s)[:10]
                        ex_e_str = str(ex_e)[:10]
                        if check_period_overlap(s_date_str, e_date_str, ex_s_str, ex_e_str):
                            if ex_s_str == s_date_str and ex_e_str == e_date_str:
                                if ex_p is not None and abs(float(ex_p) - price) < 0.001 and (ex_pct is None or ex_pct == 0):
                                    is_exact_match = True
                                else:
                                    exact_period_old_ids.append(ex_id)
                            else:
                                overlap_conflict = f"{ex_s_str} - {ex_e_str} ({ex_p} €)"
                                break
                                
                    if overlap_conflict:
                        raise RuntimeError(f"Код {code} ({r.get('db_name', '')}) има застъпен период с активна промоция {overlap_conflict}! Записът е блокиран.")
                    if is_exact_match:
                        skipped_count += 1
                        continue
                        
                    # Изчистване на стари промоции за същия период през DELETEPROMOTIONSALESPRICE
                    for old_id in exact_period_old_ids:
                        cur.execute("EXECUTE PROCEDURE DELETEPROMOTIONSALESPRICE(1, '0', '0', ?)", (old_id,))
                        for loc_id in REMOTE_LOCATIONS:
                            cur.execute("EXECUTE PROCEDURE DELETEPROMOTIONSALESPRICE(?, '1', '0', ?)", (loc_id, old_id))
                        updated_count += 1
                        
                    # Генериране на ново ID и запис
                    cur.execute("SELECT GEN_ID(G_PROMOTIONSALEPRICE, 1) FROM RDB$DATABASE")
                    new_id = cur.fetchone()[0]
                    # Location 1 (LOG = '0')
                    cur.execute("""
                        EXECUTE PROCEDURE INSERTPROMOTIONSALESPRICE(
                            1, ?, ?, ?, ?,
                            '0', '0', NULL, NULL, ?, NULL, '0',
                            '0', '0', '0', '0', '0', '0', '0',
                            '0', '0', '0', '0', '', 0
                        )
                    """, (code, s_dt, e_dt, price, new_id))
                    # Remote locations (LOG = '1')
                    for loc_id in REMOTE_LOCATIONS:
                        cur.execute("""
                            EXECUTE PROCEDURE INSERTPROMOTIONSALESPRICE(
                                ?, ?, ?, ?, ?,
                                '1', '0', NULL, NULL, ?, NULL, '0',
                                '0', '0', '0', '0', '0', '0', '0',
                                '0', '0', '0', '0', '', 0
                            )
                        """, (loc_id, code, s_dt, e_dt, price, new_id))
                    inserted_count += 1
                    
            conn.commit()
            cur.close()
            conn.close()
            
            total_db = (inserted_count + updated_count) * len(TARGET_LOCATIONS)
            total_sync = (inserted_count + updated_count) * len(REMOTE_LOCATIONS)
            return True, f"Успешно обработени: {inserted_count} нови, {updated_count} актуализирани (изчистена стара цена), {skipped_count} пропуснати (без промяна). Генерирани {total_db} записа в базата и {total_sync} синхронизационни пакета към 11-те магазина!"
        except Exception as e:
            if conn:
                try: conn.rollback(); conn.close()
                except Exception: pass
            return False, f"Грешка при запис в базата: {e}"
            
    # Fallback към isql.exe
    if os.path.exists(ISQL_PATH):
        try:
            sql_lines = ["SET AUTODDL OFF;"]
            if mode == "ZKP":
                for r in records_to_process:
                    code = r["code"]
                    price = r["price"]
                    s_dt = r["start_dt"]
                    e_dt = r["end_dt"]
                    sql_lines.append(
                        f"DELETE FROM AGREEDDELIVERYPRICE WHERE LOCATIONID = {CENTRAL_LOCATION} AND STORAGEID = 1 AND MATERIALCODE = {code} AND ACTIVEFROM = '{s_dt}' AND ACTIVETO = '{e_dt}';"
                    )
                    sql_lines.append(
                        f"EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES({CENTRAL_LOCATION}, 1, {code}, {price:.2f}, '{s_dt}', '{e_dt}', '0', GEN_ID(G_AGREEDDELIVERYPRICE, 1), -1, -1, -1, 0.00);"
                    )
                    for loc_id in REMOTE_LOCATIONS:
                        sql_lines.append(
                            f"DELETE FROM AGREEDDELIVERYPRICE WHERE LOCATIONID = {loc_id} AND STORAGEID = 1 AND MATERIALCODE = {code} AND ACTIVEFROM = '{s_dt}' AND ACTIVETO = '{e_dt}';"
                        )
                        sql_lines.append(
                            f"EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES({loc_id}, 1, {code}, {price:.2f}, '{s_dt}', '{e_dt}', '1', GEN_ID(G_AGREEDDELIVERYPRICE, 0), -1, -1, -1, 0.00);"
                        )
            else:
                for r in records_to_process:
                    code = r["code"]
                    price = r["price"]
                    s_dt = r["start_dt"]
                    e_dt = r["end_dt"]
                    sql_lines.append(
                        f"DELETE FROM PROMOTIONSALEPRICE WHERE LOCATIONID = {CENTRAL_LOCATION} AND SALEMATERIALCODE = {code} AND STARTDATE = '{s_dt}' AND ENDDATE = '{e_dt}';"
                    )
                    sql_lines.append(
                        f"EXECUTE PROCEDURE INSERTPROMOTIONSALESPRICE({CENTRAL_LOCATION}, {code}, '{s_dt}', '{e_dt}', {price:.2f}, '0', '0', NULL, NULL, GEN_ID(G_PROMOTIONSALEPRICE, 1), NULL, '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', 0);"
                    )
                    for loc_id in REMOTE_LOCATIONS:
                        sql_lines.append(
                            f"DELETE FROM PROMOTIONSALEPRICE WHERE LOCATIONID = {loc_id} AND SALEMATERIALCODE = {code} AND STARTDATE = '{s_dt}' AND ENDDATE = '{e_dt}';"
                        )
                        sql_lines.append(
                            f"EXECUTE PROCEDURE INSERTPROMOTIONSALESPRICE({loc_id}, {code}, '{s_dt}', '{e_dt}', {price:.2f}, '1', '0', NULL, NULL, GEN_ID(G_PROMOTIONSALEPRICE, 0), NULL, '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', 0);"
                        )
            sql_lines.append("COMMIT;")
            sql_lines.append("QUIT;")
            run_isql_script("\n".join(sql_lines))
            total_db_records = len(records_to_process) * len(TARGET_LOCATIONS)
            total_sync_records = len(records_to_process) * len(REMOTE_LOCATIONS)
            return True, f"Успешно записани/актуализирани {total_db_records} позиции през ISQL и генерирани {total_sync_records} пакета за синхронизация към 11-те магазина!"
        except Exception as e:
            return False, f"Грешка при ISQL запис: {e}"

    return False, "Няма достъпна връзка към базата данни (проверете мрежата към 192.168.5.254)!"

# ------------------------------------------------------------------------------
# Streamlit UI
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Mistral ERP — Импорт на цени",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Персонализирани стилове
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 5px;
    }
    .sub-header {
        font-size: 15px;
        color: #4B5563;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Заглавие
st.markdown('<div class="main-header">🏪 ВЕРИГА „БАКАЛИЯ“ — ИМПОРТ НА ЦЕНИ В MISTRAL ERP</div>', unsafe_allow_html=True)

# Проверка на връзката за сайдбара
test_conn = get_db_connection()
if test_conn:
    db_status_text = "🟢 FirebirdSQL Driver (Директна връзка)"
    try: test_conn.close()
    except Exception: pass
elif os.path.exists(ISQL_PATH):
    db_status_text = "🟡 Firebird ISQL Fallback (isql.exe)"
else:
    db_status_text = "🔴 Няма връзка към Firebird"

# Странична лента
with st.sidebar:
    st.title("⚙️ Режим на работа")
    
    mode_choice = st.radio(
        "Изберете тип операция:",
        ["ЗКП — Заключени покупни цени (ДТУ)", "ПЦ — Промоционални продажни цени"],
        index=1
    )
    mode = "ZKP" if "ЗКП" in mode_choice else "PC"
    
    st.markdown("---")
    st.caption(f"🔌 **Статус база:** {db_status_text}")
    st.caption("🔒 **База данни:** `192.168.5.254:E:\\DB\\CS.FDB`")
    st.caption("📦 **Склад:** Склад 1 (`STORAGEID = 1`)")
    st.caption("🌐 **Мрежов адрес:** `http://192.168.5.254:8501` или `http://192.168.5.99:8501`")
    st.caption("🔄 **Синхронизация:** 11 магазина (през съхранени процедури)")

# Избор на файл
st.subheader("📁 Избор и зареждане на Excel файл")
uploaded_file = st.file_uploader(
    "Качете Excel файл с цени (.xlsx, .xls)",
    type=["xlsx", "xls"],
    help="Таблицата трябва да съдържа колони: Код на артикул, Цена, Период от, Период до."
)

df_raw = None
file_source_name = ""

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        file_source_name = uploaded_file.name
        st.success(f"Зареден файл: **{file_source_name}** ({len(df_raw)} реда)")
    except Exception as e:
        st.error(f"Грешка при четене на качения файл: {e}")

if df_raw is not None:
    # Разпознаване на колони
    col_code = None
    col_name = None
    col_price = None
    col_from = None
    col_to = None
    
    for c in df_raw.columns:
        c_clean = str(c).strip().lower()
        if "код" in c_clean:
            col_code = c
        elif "име" in c_clean or "артикул" in c_clean or "наименование" in c_clean:
            col_name = c
        elif any(p in c_clean for p in ["цена", "покупна", "продажна", "договорена", "промо"]):
            col_price = c
        elif "от" in c_clean or "начало" in c_clean:
            col_from = c
        elif "до" in c_clean or "край" in c_clean:
            col_to = c
            
    missing_cols = []
    if not col_code: missing_cols.append("Код на артикул")
    if not col_price: missing_cols.append("Цена")
    if not col_from: missing_cols.append("Период от")
    if not col_to: missing_cols.append("Период до")
    
    if missing_cols:
        st.error(f"❌ В таблицата липсват задължителни колони: {', '.join(missing_cols)}")
        st.write("Намерени колони в Excel:", list(df_raw.columns))
    else:
        # Парсиране на редовете
        rows_to_check = []
        parse_errors = []
        
        for idx, row in df_raw.iterrows():
            row_num = idx + 2
            raw_code = row[col_code]
            raw_name = str(row[col_name]).strip() if col_name and pd.notna(row[col_name]) else ""
            raw_price = row[col_price]
            raw_from = row[col_from]
            raw_to = row[col_to]
            
            if pd.isna(raw_code) and pd.isna(raw_price) and pd.isna(raw_from):
                continue
                
            try:
                art_code = int(raw_code)
                if art_code <= 0: raise ValueError()
            except Exception:
                parse_errors.append(f"Ред {row_num}: Невалиден код '{raw_code}'")
                continue
                
            try:
                if pd.isna(raw_price) or str(raw_price).strip() == "":
                    raise ValueError("Липсваща цена")
                price_val = float(raw_price)
                if math.isnan(price_val) or math.isinf(price_val) or price_val <= 0:
                    raise ValueError("Цената трябва да е положително число")
            except Exception:
                parse_errors.append(f"Ред {row_num} (Код {art_code}): Липсваща или невалидна цена '{raw_price}'")
                continue
                
            try:
                start_dt = parse_date(raw_from, is_end=False)
                end_dt = parse_date(raw_to, is_end=True)
                if start_dt > end_dt:
                    parse_errors.append(f"Ред {row_num} (Код {art_code}): Начална дата ({start_dt[:10]}) е след крайна ({end_dt[:10]})!")
                    continue
            except Exception as e:
                parse_errors.append(f"Ред {row_num} (Код {art_code}): Грешка в датите: {e}")
                continue
                
            rows_to_check.append({
                "row_num": row_num,
                "code": art_code,
                "input_name": raw_name,
                "price": price_val,
                "start_dt": start_dt,
                "end_dt": end_dt
            })
            
        if parse_errors:
            st.error(f"""
            ### ⚠️ ВНИМАНИЕ: Открити са {len(parse_errors)} реда с липсващи или невалидни данни!
            Тези {len(parse_errors)} артикула са **пропуснати** и **НЕ СА** включени в импорта:
            """)
            with st.expander(f"🔍 Списък на пропуснатите {len(parse_errors)} реда (липсваща цена / невалиден код)", expanded=True):
                for pe in parse_errors:
                    st.warning(pe)
                    
        # 1. Засичане на застъпвания вътре в самия Excel файл
        code_to_file_rows = {}
        for r in rows_to_check:
            code_to_file_rows.setdefault(r["code"], []).append(r)

        for c_code, r_list in code_to_file_rows.items():
            if len(r_list) > 1:
                for i in range(len(r_list)):
                    for j in range(i + 1, len(r_list)):
                        r1 = r_list[i]
                        r2 = r_list[j]
                        if check_period_overlap(r1["start_dt"], r1["end_dt"], r2["start_dt"], r2["end_dt"]):
                            r1.setdefault("file_overlap_conflicts", []).append(
                                f"ред {r2['row_num']} ({r2['start_dt'][:10]} - {r2['end_dt'][:10]})"
                            )
                            r2.setdefault("file_overlap_conflicts", []).append(
                                f"ред {r1['row_num']} ({r1['start_dt'][:10]} - {r1['end_dt'][:10]})"
                            )

        # 2. Извличане на информация от базата данни
        codes_list = [r["code"] for r in rows_to_check]
        with st.spinner("Сверяване на артикулите и извличане на активните цени и периоди от Mistral CS..."):
            db_info = fetch_db_article_info(codes_list, mode)
            
        # 3. Формиране на сравнителната таблица с проверка за 3-те търговски ситуации:
        # А: 100% съвпадение на период и цена -> ⏭️ Вече активен (без промяна)
        # Б: Същият период, различна цена -> ⚠️ Промяна на цена за активен период
        # В: Застъпващи се периоди (различни дати) -> ⛔ ГРЕШКА: Застъпен период! (ПЪЛЕН БЛОКАЖ)
        comparison_rows = []
        valid_records = []
        
        for r in rows_to_check:
            code = r["code"]
            info = db_info.get(code)
            
            if not info:
                comparison_rows.append({
                    "Ред": r["row_num"],
                    "Код": code,
                    "Име в Mistral": "❌ НЕ СЪЩЕСТВУВА В БАЗАТА",
                    "Име в Excel": r["input_name"],
                    "Цена преди (€)": None,
                    "Нова цена (€)": r["price"],
                    "Разлика (€)": None,
                    "Промяна (%)": None,
                    "Предишен период": "-",
                    "Нов период": f"{r['start_dt'][:10]} -> {r['end_dt'][:10]}",
                    "Статус": "❌ Невалиден код",
                    "Бележка": "Артикулът не е намерен в номенклатурата на Mistral"
                })
                continue
                
            db_name = info["name"]
            r["db_name"] = db_name
            valid_records.append(r)
            
            # Определяне на базова/предишна цена
            if mode == "ZKP":
                prev_price = info["last_price"] if info["last_price"] is not None else info["base_delivery"]
                prev_period = f"{str(info['last_start'])[:10]} -> {str(info['last_end'])[:10]}" if info["last_start"] else "Базова доставна"
            else:
                prev_price = info["last_price"] if info["last_price"] is not None else info["base_sale"]
                prev_period = f"{str(info['last_start'])[:10]} -> {str(info['last_end'])[:10]}" if info["last_start"] else "Базова продажна"
            
            diff_eur = (r["price"] - prev_price) if prev_price is not None else None
            diff_pct = ((r["price"] - prev_price) / prev_price * 100) if prev_price and prev_price > 0 else None
            
            # А. Проверка за вътрешен конфликт във файла
            file_conflicts = r.get("file_overlap_conflicts", [])
            if file_conflicts:
                r["overlap_error"] = True
                status_text = "⛔ ГРЕШКА: Застъпен период във файла!"
                note_text = f"Застъпване с {', '.join(file_conflicts)}"
                comparison_rows.append({
                    "Ред": r["row_num"],
                    "Код": code,
                    "Име в Mistral": db_name,
                    "Име в Excel": r["input_name"] if r["input_name"] else db_name,
                    "Цена преди (€)": prev_price,
                    "Нова цена (€)": r["price"],
                    "Разлика (€)": diff_eur,
                    "Промяна (%)": diff_pct,
                    "Предишен период": prev_period,
                    "Нов период": f"{r['start_dt'][:10]} -> {r['end_dt'][:10]}",
                    "Статус": status_text,
                    "Бележка": note_text
                })
                continue
                
            # Б. Проверка спрямо съществуващи промоции/цени в Mistral
            overlap_conflicts = []
            exact_match_found = False
            price_change_candidates = []
            
            r_start_str = r["start_dt"][:10]
            r_end_str = r["end_dt"][:10]
            r_price = round(float(r["price"]), 2)
            
            for ex in info.get("all_existing", []):
                ex_p = ex.get("price")
                ex_s = ex.get("start", "")[:10]
                ex_e = ex.get("end", "")[:10]
                ex_pct = ex.get("pct")
                
                if not ex_s or not ex_e:
                    continue
                    
                if check_period_overlap(r_start_str, r_end_str, ex_s, ex_e):
                    if ex_s == r_start_str and ex_e == r_end_str:
                        # Напълно идентичен интервал от дати
                        if ex_p is not None and abs(ex_p - r_price) < 0.001 and (mode != "PC" or ex_pct is None or ex_pct == 0):
                            exact_match_found = True
                        else:
                            price_change_candidates.append((ex_p, ex_s, ex_e))
                    else:
                        # Различни дати, но застъпен интервал -> СИТУАЦИЯ В!
                        overlap_conflicts.append(f"{ex_s} - {ex_e} ({ex_p:.2f} €)" if ex_p is not None else f"{ex_s} - {ex_e}")
                        
            # Йерархия на правилата съгласно фирмената политика:
            # 1. Ако има застъпване на период с различни дати -> ПЪЛЕН БЛОКАЖ (Ситуация В)
            if overlap_conflicts:
                r["overlap_error"] = True
                status_text = "⛔ ГРЕШКА: Застъпен период!"
                overlap_note = f"Застъпване с активна промоция в Mistral: {', '.join(overlap_conflicts)}"
            elif price_change_candidates:
                # 2. Същият период, но различна цена -> Ситуация Б
                r["modifying_active"] = True
                old_p, s_str, e_str = price_change_candidates[0]
                status_text = "⚠️ Промяна на цена за активен период"
                overlap_note = f"Промяна от {old_p:.2f} € на {r_price:.2f} € за {s_str} - {e_str}"
            elif exact_match_found:
                # 3. 100% съвпадение на период и цена -> Ситуация А
                r["skip"] = True
                status_text = "⏭️ Вече активен (без промяна)"
                overlap_note = f"Идентичен активен запис в Mistral ({r_price:.2f} €)"
            else:
                if prev_period != "Базова доставна" and prev_period != "Базова продажна":
                    status_text = "🔄 Нов период"
                    overlap_note = "Планиране на нов период"
                else:
                    status_text = "🆕 Нова позиция"
                    overlap_note = "Първоначално въвеждане"
                    
            if r["input_name"] and r["input_name"].lower()[:8] != db_name.lower()[:8]:
                overlap_note += " [⚠️ разл. име в Excel]"
                
            comparison_rows.append({
                "Ред": r["row_num"],
                "Код": code,
                "Име в Mistral": db_name,
                "Име в Excel": r["input_name"] if r["input_name"] else db_name,
                "Цена преди (€)": prev_price,
                "Нова цена (€)": r["price"],
                "Разлика (€)": diff_eur,
                "Промяна (%)": diff_pct,
                "Предишен период": prev_period,
                "Нов период": f"{r['start_dt'][:10]} -> {r['end_dt'][:10]}",
                "Статус": status_text,
                "Бележка": overlap_note
            })
                
        df_comp = pd.DataFrame(comparison_rows)
        
        st.markdown("---")
        
        # Обобщени метрики
        overlap_errors = [r for r in rows_to_check if r.get("overlap_error", False)]
        modified_active_records = [r for r in valid_records if r.get("modifying_active", False) and not r.get("skip", False) and not r.get("overlap_error", False)]
        items_to_import = [r for r in valid_records if not r.get("skip", False) and not r.get("overlap_error", False)]
        skipped_items = [r for r in valid_records if r.get("skip", False)]
        invalid_items = len(rows_to_check) - len(valid_records)
        total_target_records = len(items_to_import) * len(TARGET_LOCATIONS)
        
        cols_metric = st.columns(5 if parse_errors else 4)
        with cols_metric[0]:
            st.metric("📦 Валидни артикули", f"{len(df_comp)} бр.")
        with cols_metric[1]:
            if overlap_errors:
                st.metric("⛔ Застъпени периоди", f"{len(overlap_errors)} бр.", delta="ГРЕШКА", delta_color="inverse")
            else:
                st.metric("⚡ За запис/корекция", f"{len(items_to_import)} бр.")
        with cols_metric[2]:
            st.metric("⏭️ Без промяна", f"{len(skipped_items)} бр.")
        if parse_errors:
            with cols_metric[3]:
                st.metric("⚠️ С липсваща цена", f"{len(parse_errors)} бр.", delta="ПРОПУСНАТИ", delta_color="inverse")
            with cols_metric[4]:
                st.metric("🏢 Целеви записи", f"{total_target_records} бр.")
        else:
            with cols_metric[3]:
                if invalid_items > 0:
                    st.metric("❌ Ненамерени в базата", f"{invalid_items} бр.", delta_color="inverse")
                else:
                    st.metric("🏢 Целеви записи", f"{total_target_records} бр.")

        # 1. СЪОБЩЕНИЕ ЗА ГРЕШКА ПРИ ЗАСТЪПЕНИ ПЕРИОДИ (Правило 3 / Ситуация В)
        if overlap_errors:
            st.error(f"""
            ### ⛔ ОТКРИТИ СА {len(overlap_errors)} ЗАСТЪПВАЩИ СЕ ПЕРИОДА! ИМПОРТЪТ Е НАПЪЛНО БЛОКИРАН.
            **Търговско правило на „Бакалия“:** Когато има застъпващи се периоди за един и същи артикул, системата **НЕ коригира автоматично промоцията**, за да се предотвратят ценови колизии на касите и етикетите по обектите.
            
            👉 **Действие:** Прегледайте редовете със статус `⛔ ГРЕШКА: Застъпен период!` в таблицата по-долу. Коригирайте датите във входящия Excel файл и го качете отново.
            """)
            
        # 2. ИЗИСКВАНЕ ЗА ПОТВЪРЖДЕНИЕ ПРИ МОДИФИКАЦИЯ НА АКТИВНИ ПРОМОЦИИ (Правило 1 / Ситуация Б)
        confirm_active_mod = False
        if modified_active_records and not overlap_errors:
            st.warning(f"""
            ### ⚠️ Внимание: Открити са {len(modified_active_records)} бр. активни позиции за промяна на цена!
            Във файла се съдържат нови цени за периоди, които вече са активирани в Mistral ERP. Системата първо ще изчисти старите записи през `DELETEPROMOTIONSALESPRICE` със синхронизация към 11-те магазина и ще запише новите цени.
            """)
            confirm_active_mod = st.checkbox(
                f"✅ Потвърждавам изрично промяната на цената за {len(modified_active_records)} бр. вече активирани промоции в Mistral ERP",
                value=False,
                help="Съгласно правилата на Бакалия се изисква изрично потвърждение преди промяна на активни промоции."
            )

        # Интерактивна сравнителна таблица
        st.markdown("### 📋 Преглед и одит на цените")
        st.dataframe(
            df_comp,
            use_container_width=True
        )

        # Бутони за експорт и действия
        st.markdown("---")
        st.subheader("🚀 Изпълнение на операцията")
        
        can_execute = True
        block_msg = ""
        
        if overlap_errors:
            can_execute = False
            block_msg = f"Импортът е напълно блокиран: открити са {len(overlap_errors)} застъпващи се периода!"
        elif modified_active_records and not confirm_active_mod:
            can_execute = False
            block_msg = "Изисква се потвърждение за промяна на активните промоции!"
        elif len(items_to_import) == 0:
            can_execute = False
            block_msg = "Всички артикули вече са активни с идентични параметри. Няма позиции за запис."
            
        act_col1, act_col2, act_col3 = st.columns([1, 1, 2])
        
        with act_col1:
            run_simulation = st.button("🧪 Тест / Симулация (Dry-Run)", use_container_width=True)
            
        with act_col2:
            if can_execute:
                run_execute = st.button("⚡ ИЗПЪЛНИ ИМПОРТА В MISTRAL (RUN)", type="primary", use_container_width=True)
            else:
                st.button("⚡ ИЗПЪЛНИ ИМПОРТА В MISTRAL (RUN)", type="primary", use_container_width=True, disabled=True, help=block_msg)
                run_execute = False
                
        with act_col3:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_comp.to_excel(writer, index=False, sheet_name="Сравнение_Цени")
            st.download_button(
                label="📥 Изтегли сравнителния отчет (Excel)",
                data=buffer.getvalue(),
                file_name=f"Одит_Цени_{mode}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        if run_simulation:
            st.info("🧪 **Резултат от симулацията (DRY-RUN):**")
            if overlap_errors:
                st.error(f"⛔ БЛОКИРАН! Намерени са {len(overlap_errors)} застъпващи се периода. Операцията не може да бъде изпълнена.")
            else:
                st.write(f"- Позиции за запис/корекция: **{len(items_to_import)} бр.**")
                st.write(f"- От тях активни за промяна на цена: **{len(modified_active_records)} бр.**")
                st.write(f"- Пропуснати (вече активни без промяна): **{len(skipped_items)} бр.**")
                st.write(f"- Локации за обновяване: **{len(TARGET_LOCATIONS)} бр.** (Централен склад + 11 магазина)")
                st.write(f"- Общ брой нови записи за генериране: **{total_target_records} бр.**")
                if len(items_to_import) == 0:
                    st.info("ℹ️ Всички артикули от файла вече са активни в базата с идентични цени и периоди. Не се налагат промени.")
                else:
                    if modified_active_records and not confirm_active_mod:
                        st.warning("⚠️ Забележка: За реален импорт трябва да поставите отметката за потвърждение на активните промоции.")
                    else:
                        st.success("✅ Всички валидации преминаха успешно! Готови сте да натиснете бутона 'ИЗПЪЛНИ ИМПОРТА'.")
            
        if run_execute:
            if not items_to_import:
                st.error("❌ Няма валидни артикули за запис!")
            else:
                with st.spinner("Записване на транзакциите в базата данни на Централен склад и генериране на пакети..."):
                    success, msg = execute_db_import(items_to_import, mode)
                    if success:
                        st.balloons()
                        st.success(f"🎉 **{msg}**")
                        if parse_errors:
                            st.warning(f"""
                            ⚠️ **Внимание: {len(parse_errors)} артикула от файла НЕ бяха записани поради липсваща цена:**
                            """ + "\n".join([f"- {pe}" for pe in parse_errors]))
                        st.info("💡 При следващата синхронизация новите стойности автоматично ще бъдат разпратени към 11-те магазина.")
                    else:
                        st.error(f"❌ {msg}")
else:
    st.info("👆 Моля, качете Excel файл с цени от полето горе.")
