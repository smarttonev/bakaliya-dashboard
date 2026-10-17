#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
================================================================================
Mistral ERP — Автоматизиран импортьор на Промоционални продажни цени (ПЦ)
             и Заключени договорени покупни цени (ЗКП)
================================================================================
Автор: Главен търговски контролер / Antigravity AI
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
import math
import argparse
import subprocess
import datetime
import pandas as pd

# Осигуряване на съвместимост с Windows терминал
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ------------------------------------------------------------------------------
# Конфигурация на базата данни и файловата среда
# ------------------------------------------------------------------------------
ISQL_PATH = r"C:\Program Files\Firebird\Firebird_2_5\bin\isql.exe"
DB_PATH = r"192.168.5.254:E:\DB\CS.FDB"
DB_USER = "SYSDBA"
DB_PASS = "masterkey"
DB_CHARSET = "NONE"

# Списък на целевите локации: Централен склад (1) + 11-те магазина (БЕЗ Монтевидео 7)
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
# Помощни функции за работа с Firebird през isql.exe
# ------------------------------------------------------------------------------
def run_isql(sql_script_text, timeout=30):
    """Изпълнява SQL скрипт през официалния Firebird isql.exe клиент."""
    temp_sql_path = os.path.join(os.environ.get("TEMP", "C:\\Temp"), f"mistral_imp_{os.getpid()}_{int(datetime.datetime.now().timestamp())}.sql")
    try:
        with open(temp_sql_path, "w", encoding="cp1251", errors="replace") as f:
            f.write(sql_script_text)
        
        cmd = [ISQL_PATH, "-u", DB_USER, "-p", DB_PASS, "-ch", DB_CHARSET, DB_PATH, "-q", "-m", "-i", temp_sql_path]
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

def get_articles_from_db(article_codes):
    """Извлича актуалните имена на артикулите от Mistral през GETARTIKULFIELDS."""
    if not article_codes:
        return {}
    
    unique_codes = list(set(article_codes))
    db_articles = {}
    
    # Порции по 50 артикула за UNION ALL
    batch_size = 50
    for i in range(0, len(unique_codes), batch_size):
        batch = unique_codes[i:i + batch_size]
        unions = " UNION ALL ".join(f"SELECT {c} AS CODE, ARTIKUL FROM GETARTIKULFIELDS(1, {c})" for c in batch)
        sql = f"""
        {unions};
        QUIT;
        """
        out = run_isql(sql)
        for line in out.splitlines():
            if len(line) >= 15 and line[:12].strip().isdigit():
                aid = int(line[:12].strip())
                name = line[13:].strip()
                if name and name.upper() != "<NULL>":
                    db_articles[aid] = name
                    
    return db_articles

def get_next_id(generator_name):
    """Извлича следващ валиден пореден ID от съответния Firebird генератор."""
    sql = f"SELECT GEN_ID({generator_name}, 1) FROM RDB$DATABASE; QUIT;"
    out = run_isql(sql)
    for line in out.splitlines():
        line = line.strip()
        if line.isdigit():
            return int(line)
    raise RuntimeError(f"Не може да бъде извлечен ID от генератор {generator_name}")

def parse_date(val, is_end=False):
    """Парсира дата от Excel (текст или datetime) и я форматира с час 00:00:00 или 23:59:59."""
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
        raise ValueError(f"Неразпознат формат на дата: '{s}'")

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

def filter_duplicate_records(records, mode):
    """
    Проверява за:
    1. Вътрешни застъпвания във файла
    2. Застъпвания със съществуващи промоции в базата (overlap_errors)
    3. Промени по вече активни промоции (modified_active)
    4. Идентични съществуващи записи без промяна (skipped)
    5. Чисто нови валидни позиции (filtered)
    """
    if not records:
        return [], [], [], []
        
    # 1. Засичане на застъпвания вътре в самия файл
    code_groups = {}
    for r in records:
        code_groups.setdefault(r["code"], []).append(r)
        
    for code, group in code_groups.items():
        if len(group) > 1:
            for i in range(len(group)):
                for j in range(i + 1, len(group)):
                    r1 = group[i]
                    r2 = group[j]
                    if check_period_overlap(r1["start_dt"], r1["end_dt"], r2["start_dt"], r2["end_dt"]):
                        r1.setdefault("overlap_error_list", []).append(f"застъпване във файла с ред {r2['row_num']}")
                        r2.setdefault("overlap_error_list", []).append(f"застъпване във файла с ред {r1['row_num']}")

    filtered = []
    skipped = []
    modified_active = []
    overlap_errors = []
    
    conn = None
    try:
        import firebirdsql
        parts = DB_PATH.split(":", 1)
        db_host = parts[0]
        db_file = parts[1] if len(parts) > 1 else ""
        
        hosts = [db_host]
        for alt in ['127.0.0.1', 'localhost']:
            if alt not in hosts:
                hosts.append(alt)
                
        for h in hosts:
            try:
                conn = firebirdsql.connect(
                    host=h,
                    database=db_file,
                    port=3050,
                    user=DB_USER,
                    password=DB_PASS,
                    charset='WIN1251'
                )
                break
            except Exception:
                continue
                
        if not conn:
            raise RuntimeError(f"Не може да бъде осъществена връзка към {DB_PATH}")
            
        cur = conn.cursor()
        for r in records:
            if "overlap_error_list" in r:
                r["overlap_error"] = "; ".join(r["overlap_error_list"])
                overlap_errors.append(r)
                continue
                
            code = r["code"]
            price = round(float(r["price"]), 2)
            s_dt = r["start_dt"]
            e_dt = r["end_dt"]
            s_date_str = str(s_dt)[:10]
            e_date_str = str(e_dt)[:10]
            
            if mode == "ZKP":
                cur.execute("""
                    SELECT AGREEDDELIVERYPRICE, ACTIVEFROM, ACTIVETO, ID 
                    FROM AGREEDDELIVERYPRICE 
                    WHERE MATERIALCODE = ? AND STORAGEID = 1 AND LOCATIONID = 1
                    ORDER BY ACTIVEFROM DESC, ID DESC
                """, (code,))
            else:
                cur.execute("""
                    SELECT SALEPRICE, STARTDATE, ENDDATE, PERCENTAGEDISCOUNT, ID 
                    FROM PROMOTIONSALEPRICE 
                    WHERE SALEMATERIALCODE = ? AND LOCATIONID = 1
                    ORDER BY STARTDATE DESC, ID DESC
                """, (code,))
                
            existing_rows = cur.fetchall()
            overlap_conflicts = []
            exact_match_found = False
            price_change_candidates = []
            old_ids = []
            
            for ex in existing_rows:
                ex_p = float(ex[0]) if ex[0] is not None else None
                ex_s = str(ex[1])[:10] if ex[1] is not None else ""
                ex_e = str(ex[2])[:10] if ex[2] is not None else ""
                ex_pct = ex[3] if len(ex) > 3 else None
                ex_id = ex[4] if len(ex) > 4 else None
                
                if not ex_s or not ex_e:
                    continue
                    
                if check_period_overlap(s_date_str, e_date_str, ex_s, ex_e):
                    if ex_s == s_date_str and ex_e == e_date_str:
                        if ex_p is not None and abs(ex_p - price) < 0.001 and (mode != "PC" or ex_pct is None or ex_pct == 0):
                            exact_match_found = True
                        else:
                            price_change_candidates.append((ex_p, ex_s, ex_e))
                            if ex_id: old_ids.append(ex_id)
                    else:
                        overlap_conflicts.append(f"{ex_s} - {ex_e} ({ex_p:.2f} €)" if ex_p is not None else f"{ex_s} - {ex_e}")
                        
            if overlap_conflicts:
                r["overlap_error"] = f"застъпване със съществуваща активна промоция: {', '.join(overlap_conflicts)}"
                overlap_errors.append(r)
            elif price_change_candidates:
                r["mod_note"] = f"промяна на цена от {price_change_candidates[0][0]:.2f} € на {price:.2f} € за {price_change_candidates[0][1]} - {price_change_candidates[0][2]}"
                r["old_ids"] = old_ids
                modified_active.append(r)
                filtered.append(r)
            elif exact_match_found:
                skipped.append(r)
            else:
                filtered.append(r)
                
        conn.close()
        return filtered, skipped, modified_active, overlap_errors
    except Exception as e:
        if conn:
            try: conn.close()
            except Exception: pass
        print(f"[ГРЕШКА при засичане на съществуващи промоции/дубликати в базата]: {e}")
        raise RuntimeError(f"Неуспешна проверка за дубликати в базата данни на Mistral: {e}")


# ------------------------------------------------------------------------------
# Обработка и импорт
# ------------------------------------------------------------------------------
def process_import(file_path, mode, dry_run=False):
    """
    mode: 'ZKP' (Заключени покупни цени) или 'PC' (Промоционални продажни цени)
    """
    mode_title = "ЗАКЛЮЧЕНИ ПОКУПНИ ЦЕНИ (ЗКП)" if mode == "ZKP" else "ПРОМОЦИОНАЛНИ ПРОДАЖНИ ЦЕНИ (ПЦ)"
    print("=" * 80)
    print(f" МИСТРАЛ CS — {mode_title}")
    print("=" * 80)
    print(f"[>] Файл за импорт: {file_path}")
    print(f"[>] Режим: {'[СИМУЛАЦИЯ - DRY-RUN (БЕЗ ЗАПИС В БАЗАТА)]' if dry_run else '[РЕАЛЕН ЗАПИС В БАЗАТА]'}")
    print(f"[>] Локации: {len(TARGET_LOCATIONS)} бр. (Централен склад + 11 магазина; БЕЗ Монтевидео)")
    print(f"[>] Склад: Склад 1 (STORAGEID = 1)")
    print("-" * 80)
    
    if not os.path.exists(file_path):
        print(f"[ГРЕШКА] Файлът '{file_path}' не беше намерен!")
        return False
        
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[ГРЕШКА] Грешка при отваряне на Excel файла: {e}")
        return False
        
    # Разпознаване на колоните
    col_code = None
    col_name = None
    col_price = None
    col_from = None
    col_to = None
    
    for c in df.columns:
        c_clean = str(c).strip().lower()
        if "код" in c_clean:
            col_code = c
        elif "име" in c_clean or "артикул" in c_clean:
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
        print(f"[ГРЕШКА] В таблицата липсват задължителни колони: {', '.join(missing_cols)}")
        print(f"         Намерени колони в Excel файла: {list(df.columns)}")
        return False

    # Извличане на валидни редове
    rows_to_import = []
    errors = []
    
    for idx, row in df.iterrows():
        row_num = idx + 2  # Excel ред
        raw_code = row[col_code]
        raw_name = str(row[col_name]).strip() if col_name and pd.notna(row[col_name]) else ""
        raw_price = row[col_price]
        raw_from = row[col_from]
        raw_to = row[col_to]
        
        # Проверка за празен ред
        if pd.isna(raw_code) and pd.isna(raw_price) and pd.isna(raw_from):
            continue
            
        # Валидация на код
        try:
            art_code = int(raw_code)
            if art_code <= 0:
                raise ValueError()
        except Exception:
            errors.append(f"Ред {row_num}: Невалиден код на артикул '{raw_code}'")
            continue
            
        # Валидация на цена
        try:
            if pd.isna(raw_price) or str(raw_price).strip() == "":
                raise ValueError("Липсваща цена")
            price = float(raw_price)
            if math.isnan(price) or math.isinf(price) or price <= 0:
                raise ValueError("Цената трябва да е положително число")
        except Exception:
            errors.append(f"Ред {row_num} (Код {art_code}): Липсваща или невалидна цена '{raw_price}'")
            continue
            
        # Валидация на дати
        try:
            start_dt = parse_date(raw_from, is_end=False)
            end_dt = parse_date(raw_to, is_end=True)
            if start_dt > end_dt:
                errors.append(f"Ред {row_num} (Код {art_code}): Началната дата ({start_dt}) е след крайната дата ({end_dt})!")
                continue
        except Exception as e:
            errors.append(f"Ред {row_num} (Код {art_code}): Грешка в датите: {e}")
            continue
            
        rows_to_import.append({
            "row_num": row_num,
            "code": art_code,
            "input_name": raw_name,
            "price": price,
            "start_dt": start_dt,
            "end_dt": end_dt
        })
        
    if not rows_to_import:
        print("[!] Не са намерени валидни редове за импорт.")
        if errors:
            print("\nГрешки:")
            for err in errors:
                print(f"  * {err}")
        return False

    # Сверка срещу базата данни през GETARTIKULFIELDS
    codes = [r["code"] for r in rows_to_import]
    print(f"[*] Сверяване на {len(codes)} артикула с номенклатурата в базата на Централен склад...")
    db_names = get_articles_from_db(codes)
    
    valid_records = []
    print("\n" + "=" * 80)
    print(f"{'РЕД':<5} | {'КОД':<7} | {'ИМЕ В MISTRAL':<35} | {'ЦЕНА (€)':<9} | {'ПЕРИОД'}")
    print("-" * 80)
    
    for r in rows_to_import:
        code = r["code"]
        if code not in db_names:
            errors.append(f"Ред {r['row_num']}: Код {code} НЕ съществува в базата данни на Mistral!")
            continue
            
        db_name = db_names[code]
        r["db_name"] = db_name
        valid_records.append(r)
        
        warning_flag = ""
        if r["input_name"] and r["input_name"].lower()[:8] != db_name.lower()[:8]:
            warning_flag = " [!] различно име в Excel"
            
        period_str = f"{r['start_dt'][:10]} -> {r['end_dt'][:10]}"
        print(f"{r['row_num']:<5} | {code:<7} | {db_name[:35]:<35} | {r['price']:<9.2f} | {period_str}{warning_flag}")
        
    print("-" * 80)
    
    if errors:
        print("\n[!] Предупреждения и грешки:")
        for err in errors:
            print(f"  * {err}")
            
    if not valid_records:
        print("\n[ГРЕШКА] Няма валидни артикули за запис след проверката в базата данни!")
        return False
        
    # Проверка за застъпвания, дубликати и активни позиции
    records_to_process, skipped_records, modified_active, overlap_errors = filter_duplicate_records(valid_records, mode)
    
    # 1. Застъпени периоди -> незабавно прекратяване с грешка (Правило 3)
    if overlap_errors:
        print("\n" + "!" * 80)
        print(" ⛔ ГРЕШКА: ОТКРИТИ СА ЗАСТЪПВАЩИ СЕ ПЕРИОДИ! ИМПОРТЪТ Е ПРЕКРАТЕН.")
        print("!" * 80)
        print("Съгласно търговските правила на „Бакалия“, при застъпващи се периоди промоцията")
        print("НЕ се коригира автоматично, за да се предотвратят ценови колизии в обектите.\n")
        for oe in overlap_errors:
            print(f"  ❌ Код {oe['code']} ({oe.get('db_name', '')[:30]}): {oe['overlap_error']}")
        print("\n[!] Моля, коригирайте датите във входящия Excel файл преди импорт.")
        print("!" * 80)
        return False

    # 2. Модификация на активни промоции -> изисква се изрично потвърждение (Правило 1)
    if modified_active:
        print(f"\n⚠️ ВНИМАНИЕ: {len(modified_active)} бр. активни/планирани промоции ще бъдат модифицирани:")
        for ma in modified_active:
            print(f"  * Код {ma['code']} ({ma.get('db_name', '')[:30]}): {ma.get('mod_note', 'промяна на активна позиция')}")
        if not dry_run and not confirm_modify:
            print("\n[!] Защита на активни промоции (Правило 1): Изисква се изрично потвърждение.")
            ans = input("Потвърждавате ли промяната на вече активните промоции? (ДА/НЕ): ").strip().upper()
            if ans not in ["ДА", "YES", "Y", "D"]:
                print("[x] Операцията е отменена от оператора. Няма направени промени.")
                return False

    if skipped_records:
        print(f"\n[i] Пропуснати артикули (вече активни със същите параметри без промяна): {len(skipped_records)} бр.")
        for sr in skipped_records:
            print(f"    ⏭️ Код {sr['code']}: {sr['db_name'][:35]} | {sr['price']:.2f} € ({sr['start_dt'][:10]} -> {sr['end_dt'][:10]})")
            
    if not records_to_process:
        print("\n[OK] Всички валидирани позиции вече са активни в Mistral без промяна. Не се налага запис и няма дублирани записи.")
        return True
        
    total_db_records = len(records_to_process) * len(TARGET_LOCATIONS)
    print(f"\n[i] Обобщение:")
    print(f"  * Артикули за запис/корекция: {len(records_to_process)} бр.")
    if skipped_records:
        print(f"  * Пропуснати (вече активни без промяна): {len(skipped_records)} бр.")
    print(f"  * Целеви обекти за всеки артикул: {len(TARGET_LOCATIONS)} бр.")
    print(f"  * Общ брой записи за създаване/обновяване в базата: {total_db_records} бр.")
    
    if dry_run:
        print("\n[OK] Симулацията (DRY-RUN) приключи успешно. Данни НЕ са записвани в базата.")
        return True


    print("\n[*] Подготовка на SQL транзакцията и пакетите за синхронизация...")
    sql_lines = ["SET AUTODDL OFF;"]
    
    if mode == "ZKP":
        for r in records_to_process:
            code = r["code"]
            price = r["price"]
            s_dt = r["start_dt"]
            e_dt = r["end_dt"]
            new_id = get_next_id("G_AGREEDDELIVERYPRICE")
            
            # 1. Централен склад (LOCATIONID = 1) -> локален запис (LOG = '0')
            sql_lines.append(
                f"DELETE FROM AGREEDDELIVERYPRICE WHERE LOCATIONID = {CENTRAL_LOCATION} AND STORAGEID = 1 AND MATERIALCODE = {code} AND ACTIVEFROM = '{s_dt}' AND ACTIVETO = '{e_dt}';"
            )
            sql_lines.append(
                f"EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES({CENTRAL_LOCATION}, 1, {code}, {price:.2f}, '{s_dt}', '{e_dt}', '0', {new_id}, -1, -1, -1, 0.00);"
            )
            
            # 2. 11-те магазина -> запис в базата И автоматично пакетиране в SYNCHRONIZATIONLOG (LOG = '1')
            for loc_id in REMOTE_LOCATIONS:
                sql_lines.append(
                    f"DELETE FROM AGREEDDELIVERYPRICE WHERE LOCATIONID = {loc_id} AND STORAGEID = 1 AND MATERIALCODE = {code} AND ACTIVEFROM = '{s_dt}' AND ACTIVETO = '{e_dt}';"
                )
                sql_lines.append(
                    f"EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES({loc_id}, 1, {code}, {price:.2f}, '{s_dt}', '{e_dt}', '1', {new_id}, -1, -1, -1, 0.00);"
                )
    else:
        for r in records_to_process:
            code = r["code"]
            price = r["price"]
            s_dt = r["start_dt"]
            e_dt = r["end_dt"]
            new_id = get_next_id("G_PROMOTIONSALEPRICE")
            
            # Изчистване на стари промоции за същия период със синхронизация през DELETEPROMOTIONSALESPRICE
            for old_id in r.get("old_ids", []):
                sql_lines.append(f"EXECUTE PROCEDURE DELETEPROMOTIONSALESPRICE({CENTRAL_LOCATION}, '0', '0', {old_id});")
                for loc_id in REMOTE_LOCATIONS:
                    sql_lines.append(f"EXECUTE PROCEDURE DELETEPROMOTIONSALESPRICE({loc_id}, '1', '0', {old_id});")
            
            # Локално изчистване при липса на old_id
            if not r.get("old_ids"):
                sql_lines.append(
                    f"DELETE FROM PROMOTIONSALEPRICE WHERE LOCATIONID = {CENTRAL_LOCATION} AND SALEMATERIALCODE = {code} AND STARTDATE = '{s_dt}' AND ENDDATE = '{e_dt}';"
                )
                for loc_id in REMOTE_LOCATIONS:
                    sql_lines.append(
                        f"DELETE FROM PROMOTIONSALEPRICE WHERE LOCATIONID = {loc_id} AND SALEMATERIALCODE = {code} AND STARTDATE = '{s_dt}' AND ENDDATE = '{e_dt}';"
                    )
            
            # 1. Централен склад (LOCATIONID = 1) -> локален запис (LOG = '0', часове и % отстъпка строго NULL)
            sql_lines.append(
                f"EXECUTE PROCEDURE INSERTPROMOTIONSALESPRICE({CENTRAL_LOCATION}, {code}, '{s_dt}', '{e_dt}', {price:.2f}, '0', '0', NULL, NULL, {new_id}, NULL, '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', 0);"
            )
            
            # 2. 11-те магазина -> запис в базата И автоматично пакетиране в SYNCHRONIZATIONLOG (LOG = '1', часове и % отстъпка строго NULL)
            for loc_id in REMOTE_LOCATIONS:
                sql_lines.append(
                    f"EXECUTE PROCEDURE INSERTPROMOTIONSALESPRICE({loc_id}, {code}, '{s_dt}', '{e_dt}', {price:.2f}, '1', '0', NULL, NULL, {new_id}, NULL, '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', 0);"
                )
                
    sql_lines.append("COMMIT;")
    sql_lines.append("QUIT;")
    
    full_sql = "\n".join(sql_lines)
    
    print("[*] Изпълнение на транзакцията в базата данни на Централен склад...")
    try:
        run_isql(full_sql, timeout=120)
        total_sync_records = len(records_to_process) * len(REMOTE_LOCATIONS)
        print("\n[OK] УСПЕШНО ЗАВЪРШВАНЕ!")
        print(f"     Създадени/обновени {total_db_records} записа в Mistral CS.")
        print(f"     Генерирани {total_sync_records} пакета в SYNCHRONIZATIONLOG за 11-те магазина.")
        print("     При следващата синхронизация новите стойности автоматично ще бъдат получени и отразени по всички обекти.")
        return True
    except Exception as e:
        print(f"[ГРЕШКА] Грешка при запис в базата данни:\n{e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Импорт на ЗКП и ПЦ в Mistral ERP")
    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--zkp", action="store_true", help="Импорт на заключени покупни цени")
    group.add_argument("--pc", action="store_true", help="Импорт на промоционални продажни цени")
    parser.add_argument("--file", type=str, help="Път към Excel файла")
    parser.add_argument("--dry-run", action="store_true", help="Симулация без запис в базата")
    parser.add_argument("--execute", action="store_true", help="Директно изпълнение без питане")
    
    args = parser.parse_args()
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    mode = None
    if args.zkp:
        mode = "ZKP"
    elif args.pc:
        mode = "PC"
    else:
        print("=" * 60)
        print(" MISTRAL ERP — ИМПОРТ НА ЦЕНИ ЗА 11 ОБЕКТА + ЦС")
        print("=" * 60)
        print("1. Заключени договорени покупни цени (ЗКП)")
        print("2. Промоционални продажни цени (ПЦ)")
        choice = input("Изберете режим (1 или 2): ").strip()
        if choice == "1":
            mode = "ZKP"
        elif choice == "2":
            mode = "PC"
        else:
            print("[x] Невалиден избор. Прекратяване.")
            return

    file_path = args.file
    if not file_path:
        default_name = "ЗКП_Импорт.xlsx" if mode == "ZKP" else "ПЦ_Импорт.xlsx"
        file_path = os.path.join(base_dir, default_name)

    is_simulation = args.dry_run or (not args.execute)
    success = process_import(file_path, mode, dry_run=is_simulation)
    
    if success and not args.dry_run and not args.execute:
        print("\n" + "=" * 80)
        confirm = input("[?] Желаете ли да запишете тези данни в реалната база данни? (ДА/НЕ): ").strip().upper()
        if confirm in ["ДА", "YES", "Y", "D"]:
            process_import(file_path, mode, dry_run=False)
        else:
            print("[i] Операцията беше отменена от потребителя. Няма направени промени.")

if __name__ == "__main__":
    main()
