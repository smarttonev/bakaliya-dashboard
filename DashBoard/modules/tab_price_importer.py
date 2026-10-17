r"""
================================================================================
Модул: Импорт на промоционални продажни цени (ПЦ) и договорени покупни цени (ЗКП)
================================================================================
Интерфейс за Единния Dashboard на верига „Бакалия“.
База данни: 192.168.5.254:E:\DB\CS.FDB (Централен склад) + 11 магазина.
"""

import os
import sys
import io
import math
import tempfile
import datetime
import subprocess
import pandas as pd
import streamlit as st

try:
    import firebirdsql
    HAS_FIREBIRDSQL = True
except ImportError:
    HAS_FIREBIRDSQL = False

# Конфигурация на базата данни
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
        
        if proc.returncode != 0:
            return False, f"Грешка при изпълнение на ISQL: {stderr or stdout}"
        return True, stdout
    except Exception as e:
        return False, str(e)
    finally:
        if os.path.exists(temp_sql_path):
            try:
                os.remove(temp_sql_path)
            except Exception:
                pass


def fetch_articles_bulk(material_codes, mode):
    """Извлича имена, базови цени и съществуващи промоции/ДТУ за пакет от артикули."""
    if not material_codes:
        return {}
        
    info_map = {}
    conn = get_db_connection()
    if conn:
        try:
            cur = conn.cursor()
            batch_size = 500
            for i in range(0, len(material_codes), batch_size):
                batch = material_codes[i:i + batch_size]
                in_clause = ",".join(str(c) for c in batch)
                
                # 1. Име и базова доставна цена от GETARTIKULFIELDS
                try:
                    cur.execute(f"SELECT CODE, ARTIKUL, LASTDOCTPRICE FROM GETARTIKULFIELDS(1, NULL) WHERE CODE IN ({in_clause})")
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
    """Проверява дали два интервала от дати се застъпват: start1 <= end2 AND end1 >= start2"""
    s1 = str(start1_dt)[:10]
    e1 = str(end1_dt)[:10]
    s2 = str(start2_dt)[:10]
    e2 = str(end2_dt)[:10]
    return (s1 <= e2) and (e1 >= s2)


def execute_db_import(records, mode):
    """Изпълнява реалния импорт в базата данни на Централен склад за 12 локации със синхронизация."""
    records_to_process = [r for r in records if not r.get('skip', False)]
    if not records_to_process:
        return True, "Всички артикули вече са активни в Mistral с идентични параметри. Няма нови записи за добавяне."

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
                    raise ValueError(f"Код {code} ({r.get('db_name', '')}) има липсваща цена!")
                price = round(float(r['price']), 2)
                if math.isnan(price) or math.isinf(price) or price <= 0:
                    raise ValueError(f"Код {code} ({r.get('db_name', '')}) има невалидна цена ({price} €)!")
                s_dt = r['start_dt']
                e_dt = r['end_dt']
                s_date_str = str(s_dt)[:10]
                e_date_str = str(e_dt)[:10]
                
                if mode == "ZKP":
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
                        
                    for old_id in exact_period_old_ids:
                        for loc_id in TARGET_LOCATIONS:
                            cur.execute("""
                                DELETE FROM AGREEDDELIVERYPRICE 
                                WHERE LOCATIONID = ? AND STORAGEID = 1 AND ID = ?
                            """, (loc_id, old_id))
                        updated_count += 1
                            
                    cur.execute("SELECT GEN_ID(G_AGREEDDELIVERYPRICE, 1) FROM RDB$DATABASE")
                    new_id = cur.fetchone()[0]
                    cur.execute("EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES(1, 1, ?, ?, ?, ?, '0', ?, -1, -1, -1, 0.00)", (code, price, s_dt, e_dt, new_id))
                    for loc_id in REMOTE_LOCATIONS:
                        cur.execute("EXECUTE PROCEDURE INSERTINTODOGDOCTPRICES(?, 1, ?, ?, ?, ?, '1', ?, -1, -1, -1, 0.00)", (loc_id, code, price, s_dt, e_dt, new_id))
                    inserted_count += 1
                else: # PC
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
                        
                    for old_id in exact_period_old_ids:
                        cur.execute("EXECUTE PROCEDURE DELETEPROMOTIONSALESPRICE(1, '0', '0', ?)", (old_id,))
                        for loc_id in REMOTE_LOCATIONS:
                            cur.execute("EXECUTE PROCEDURE DELETEPROMOTIONSALESPRICE(?, '1', '0', ?)", (loc_id, old_id))
                        updated_count += 1
                        
                    cur.execute("SELECT GEN_ID(G_PROMOTIONSALEPRICE, 1) FROM RDB$DATABASE")
                    new_id = cur.fetchone()[0]
                    cur.execute("""
                        EXECUTE PROCEDURE INSERTPROMOTIONSALESPRICE(
                            1, ?, ?, ?, ?,
                            '0', '0', NULL, NULL, ?, NULL, '0',
                            '0', '0', '0', '0', '0', '0', '0',
                            '0', '0', '0', '0', '', 0
                        )
                    """, (code, s_dt, e_dt, price, new_id))
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
            return True, f"Успешен запис! Въведени нови: {inserted_count}, Актуализирани: {updated_count}, Пропуснати (идентични): {skipped_count} за всички 12 локации."
        except Exception as e:
            if conn:
                try: conn.rollback()
                except Exception: pass
            return False, f"Грешка при запис: {e}"
            
    return False, "Няма връзка към Firebird базата данни (CS.FDB)."


def render_price_importer():
    """Главен рендер за Таб: Импорт на цени в Mistral ERP."""
    st.subheader("🏷️ Импорт на Промоционални продажни цени (ПЦ) и Договорени покупни цени (ЗКП)")
    st.caption("Автоматизиран контрол за 11 обекта + Централен склад (CS.FDB) с репликация")
    
    col_mode, col_info = st.columns([1, 2])
    with col_mode:
        mode_choice = st.radio(
            "Изберете тип операция:",
            ["ПЦ — Промоционални продажни цени", "ЗКП — Заключени покупни цени (ДТУ)"],
            index=0,
            key="price_importer_mode"
        )
        mode = "PC" if "ПЦ" in mode_choice else "ZKP"
        
    with col_info:
        test_conn = get_db_connection()
        if test_conn:
            st.success("🟢 Свързан към Централен склад: `192.168.5.254:E:\\DB\\CS.FDB`")
            try: test_conn.close()
            except Exception: pass
        else:
            st.warning("🟡 Връзката към CS.FDB се осъществява при поискване / ISQL fallback")
        st.info("📦 **Обхват:** 11 магазина + Централен склад | **Валута:** Евро (€) | **Склад:** Склад 1")

    st.markdown("---")
    
    uploaded_file = st.file_uploader(
        "Качете Excel файл с цени (.xlsx, .xls)",
        type=["xlsx", "xls"],
        key="price_importer_file_uploader",
        help="Таблицата трябва да съдържа колони: Код на артикул, Цена, Период от, Период до."
    )
    
    if uploaded_file is not None:
        try:
            df_raw = pd.read_excel(uploaded_file)
            st.success(f"Зареден файл: **{uploaded_file.name}** ({len(df_raw)} реда)")
        except Exception as e:
            st.error(f"Грешка при четене на качения файл: {e}")
            return

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
            st.write("Намерени колони:", list(df_raw.columns))
            return

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
                price = float(raw_price)
                if price <= 0: raise ValueError()
            except Exception:
                parse_errors.append(f"Ред {row_num}: Невалидна цена '{raw_price}'")
                continue
                
            try:
                start_dt = parse_date(raw_from, is_end=False)
                end_dt = parse_date(raw_to, is_end=True)
                if start_dt > end_dt:
                    parse_errors.append(f"Ред {row_num}: Началната дата е след крайната ({raw_from} -> {raw_to})")
                    continue
            except Exception as e:
                parse_errors.append(f"Ред {row_num}: Невалидна дата ({e})")
                continue
                
            rows_to_check.append({
                "row_num": row_num,
                "code": art_code,
                "input_name": raw_name,
                "price": price,
                "start_dt": start_dt,
                "end_dt": end_dt
            })
            
        if parse_errors:
            with st.expander(f"⚠️ Грешки при парсиране ({len(parse_errors)} реда)", expanded=True):
                for err in parse_errors[:10]:
                    st.write(f"- {err}")
                    
        if not rows_to_check:
            st.warning("Няма валидни редове за обработка във файла.")
            return

        # Зареждане на данни от базата
        with st.spinner("Зареждане на съществуващи цени и проверка за застъпвания в CS.FDB..."):
            all_codes = list(set(r["code"] for r in rows_to_check))
            db_info_map = fetch_articles_bulk(all_codes, mode)
            
        comparison_rows = []
        has_overlap_errors = False
        
        for r in rows_to_check:
            code = r["code"]
            info = db_info_map.get(code, {
                "name": None, "base_delivery": None, "base_sale": None,
                "last_price": None, "last_start": None, "last_end": None, "all_existing": []
            })
            db_name = info["name"] or "— Не е открит в базата —"
            
            prev_price = info["last_price"] if info["last_price"] is not None else (info["base_delivery"] if mode == "ZKP" else info["base_sale"])
            prev_period = f"{str(info['last_start'])[:10]} -> {str(info['last_end'])[:10]}" if info["last_start"] else ("Базова доставна" if mode == "ZKP" else "Базова продажна")
            
            diff_eur = (r["price"] - prev_price) if prev_price is not None else None
            diff_pct = ((r["price"] - prev_price) / prev_price * 100) if prev_price and prev_price > 0 else None
            
            # Проверка за застъпване
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
                if not ex_s or not ex_e: continue
                
                if check_period_overlap(r_start_str, r_end_str, ex_s, ex_e):
                    if ex_s == r_start_str and ex_e == r_end_str:
                        if ex_p is not None and abs(ex_p - r_price) < 0.001:
                            exact_match_found = True
                        else:
                            price_change_candidates.append((ex_p, ex_s, ex_e))
                    else:
                        overlap_conflicts.append(f"{ex_s} - {ex_e} ({ex_p:.2f} €)" if ex_p is not None else f"{ex_s} - {ex_e}")
                        
            if overlap_conflicts:
                has_overlap_errors = True
                r["overlap_error"] = True
                status_text = "⛔ ГРЕШКА: Застъпен период!"
                note = f"Конфликт с {', '.join(overlap_conflicts)}"
            elif price_change_candidates:
                r["modifying_active"] = True
                status_text = "⚠️ Промяна за активен период"
                note = f"Промяна от {price_change_candidates[0][0]:.2f} € на {r_price:.2f} €"
            elif exact_match_found:
                r["skip"] = True
                status_text = "⏭️ Идентичен запис (пропуска се)"
                note = "Вече активен с тази цена"
            else:
                status_text = "🆕 Нов запис"
                note = "Първоначално въвеждане / Нов период"
                
            comparison_rows.append({
                "Ред": r["row_num"],
                "Код": code,
                "Име в Mistral": db_name,
                "Цена преди (€)": f"{prev_price:.2f} €" if prev_price is not None else "—",
                "Нова цена (€)": f"{r['price']:.2f} €",
                "Разлика (€)": f"{diff_eur:+.2f} €" if diff_eur is not None else "—",
                "Предишен период": prev_period,
                "Нов период": f"{r_start_str} -> {r_end_str}",
                "Статус": status_text,
                "Бележка": note
            })
            
        df_comp = pd.DataFrame(comparison_rows)
        
        # KPI карти
        total_items = len(comparison_rows)
        new_items = sum(1 for r in comparison_rows if "Нов" in r["Статус"])
        updated_items = sum(1 for r in comparison_rows if "Промяна" in r["Статус"])
        skipped_items = sum(1 for r in comparison_rows if "Идентичен" in r["Статус"])
        error_items = sum(1 for r in comparison_rows if "⛔" in r["Статус"])
        
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Общо артикули", total_items)
        k2.metric("🆕 Нови записи", new_items)
        k3.metric("⚠️ Промени", updated_items)
        k4.metric("⏭️ Без промяна", skipped_items)
        k5.metric("⛔ Застъпени", error_items)
        
        st.dataframe(df_comp, use_container_width=True)
        
        if has_overlap_errors:
            st.error("⛔ ВНИМАНИЕ: Открити са застъпени периоди (Ситуация В)! Импортът е блокиран, докато конфликтите не бъдат изчистени.")
        else:
            col_btn, col_check = st.columns([1, 2])
            with col_check:
                confirm_import = st.checkbox(
                    f"Потвърждавам записа в CS.FDB и синхронизацията към 11 магазина ({new_items + updated_items} артикула)",
                    key="confirm_price_import"
                )
            with col_btn:
                if st.button("🚀 Изпълни запис в Mistral ERP", type="primary", disabled=not confirm_import):
                    with st.spinner("Записване в базата и регистриране на синхронизация..."):
                        success, msg = execute_db_import(rows_to_check, mode)
                        if success:
                            st.success(f"✅ {msg}")
                            st.balloons()
                        else:
                            st.error(f"❌ {msg}")
