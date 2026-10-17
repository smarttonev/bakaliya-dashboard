# -*- coding: utf-8 -*-
r"""
================================================================================
Верига супермаркети „Бакалия“ — Трислоен Мапинг Мениджър за Брошури (2027)
================================================================================
Архитектура на Трислойния мапинг:
  (А) Слой 1: Входящ ред от доставчика [Име по доставчик, Фирма, Баркод, ДЦ, Мярка]
  (Б) Слой 2: Връзка към Мистрал ERP (CS.FDB) [Списък от 1..N Арт. кода, Редовна ПЦ, Базова ДЦ]
  (В) Слой 3: Маркетингов предпечат [Маркетингово име за печат, М.Ед, Продуктова група 1..8]

База данни: 192.168.5.254:E:\DB\CS.FDB (Централен склад)
Локална база съответствия: brochure_item_mapping.json
"""

import os
import sys
import re
import json
import difflib
import datetime
from decimal import Decimal
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Опит за зареждане на firebirdsql
try:
    import firebirdsql
    HAS_FIREBIRD = True
except ImportError:
    HAS_FIREBIRD = False

# ------------------------------------------------------------------------------
# 8-ТЕ ПРОДУКТОВИ ГРУПИ ЗА ПРЕДПЕЧАТ
# ------------------------------------------------------------------------------
PRODUCT_GROUPS = {
    1: {"name": "напитки", "color": "00CCFF", "hex": "#00CCFF", "desc": "Безалкохолни, бира, вино, алкохол, вода, сок, кафе, чай"},
    2: {"name": "ядки", "color": "0066FF", "hex": "#0066FF", "desc": "Ядки, сушени плодове, чипс, снаксове"},
    3: {"name": "захарни", "color": "333399", "hex": "#333399", "desc": "Шоколад, вафли, бисквити, бонбони, кроасани"},
    4: {"name": "куче/коте", "color": "993300", "hex": "#993300", "desc": "Храни и аксесоари за домашни любимци"},
    5: {"name": "препарати", "color": "FFFF00", "hex": "#FFFF00", "desc": "Битова химия, перилни, омекотители, хигиена, козметика"},
    6: {"name": "масови", "color": "FFCC00", "hex": "#FFCC00", "desc": "Олио, брашно, захар, ориз, консерви, макарони, подправки"},
    7: {"name": "месни", "color": "FF0000", "hex": "#FF0000", "desc": "Колбаси, пресни меса, кайма, деликатеси, риба"},
    8: {"name": "млечни", "color": "33CC33", "hex": "#33CC33", "desc": "Сирене, кашкавал, кисело/прясно мляко, масло, извара"}
}

DB_CONFIG = {
    'host': '192.168.5.254',
    'database': r'E:\DB\CS.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

MAPPING_FILE = os.path.join(os.path.dirname(__file__), "brochure_item_mapping.json")
DELIVERIES_CACHE_FILE = os.path.join(os.path.dirname(__file__), "123-2608.xlsx")


def normalize_text(text):
    """Почиства и нормализира текст за търсене и ключове."""
    if not isinstance(text, str):
        return ""
    t = text.lower().strip()
    t = re.sub(r'[\.,/\\\(\)\-\*\+\"\';:!_]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def extract_weight_info(text):
    """Извлича мерки и грамажи от текста (напр. 500г, 1л, 250мл, 8бр)."""
    if not isinstance(text, str):
        return None
    patterns = [
        r'(\d+[\.,]?\d*)\s*(?:кг|kg)',
        r'(\d+[\.,]?\d*)\s*(?:гр|г|g|gr)',
        r'(\d+[\.,]?\d*)\s*(?:мл|ml)',
        r'(\d+[\.,]?\d*)\s*(?:л|l|литра)',
        r'(\d+)\s*(?:бр|бр\.|pcs)',
    ]
    for p in patterns:
        m = re.search(p, text.lower())
        if m:
            return m.group(0).replace(' ', '')
    return None


class BrochureMappingManager:
    def __init__(self, mapping_file=None, deliveries_file=None):
        self.mapping_file = mapping_file or MAPPING_FILE
        self.deliveries_file = deliveries_file or DELIVERIES_CACHE_FILE
        self.data = self._load_storage()
        self.db_conn = None
        self._partner_cache = {}
        self._deliveries_df = None
        self._materials_cache = None
        self._materials_cache_file = os.path.join(os.path.dirname(__file__), "mistral_materials_cache.json")

    # --------------------------------------------------------------------------
    # Зареждане и персистентност на речника
    # --------------------------------------------------------------------------
    def _load_storage(self):
        if os.path.exists(self.mapping_file):
            try:
                with open(self.mapping_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "version": "1.0",
            "last_updated": datetime.datetime.now().isoformat(),
            "mappings": {},
            "barcode_index": {}
        }

    def save_storage(self):
        self.data["last_updated"] = datetime.datetime.now().isoformat()
        with open(self.mapping_file, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    # --------------------------------------------------------------------------
    # Връзка към Firebird CS.FDB
    # --------------------------------------------------------------------------
    def get_db(self):
        if not HAS_FIREBIRD:
            return None
        if self.db_conn is not None:
            try:
                # Проверка за жива връзка
                cur = self.db_conn.cursor()
                cur.execute("SELECT 1 FROM RDB$DATABASE")
                cur.fetchall()
                return self.db_conn
            except Exception:
                self.db_conn = None

        hosts = [DB_CONFIG['host'], '127.0.0.1', 'localhost']
        for h in hosts:
            try:
                cfg = dict(DB_CONFIG)
                cfg['host'] = h
                conn = firebirdsql.connect(**cfg)
                self.db_conn = conn
                return self.db_conn
            except Exception:
                continue
        return None

    # --------------------------------------------------------------------------
    # Автоматично определяне на Продуктова група (1 до 8)
    # --------------------------------------------------------------------------
    def detect_product_group(self, text):
        """Определя най-вероятната продуктова група (1..8) по ключови думи."""
        t = normalize_text(text)
        
        # 1. Напитки
        if any(w in t for w in ['бира', 'вино', 'водка', 'уиски', 'ракия', 'сок', 'сокове', 'вода', 'напитка', 'напитки', 'кафе', 'чай', 'джин', 'ром', 'енергийна', 'сайдер', 'пепси', 'кола', 'фанта', 'спрайт']):
            return 1, PRODUCT_GROUPS[1]["name"]
        
        # 4. Куче/коте (с предимство пред масови)
        if any(w in t for w in ['куче', 'коте', 'котка', 'кучета', 'котки', 'pedigree', 'whiskas', 'purina', 'friskies', 'chappi', 'felix', 'gourmet', 'пауч за коте', 'гранули за куче']):
            return 4, PRODUCT_GROUPS[4]["name"]
            
        # 7. Месни (свинско, пиле, колбаси, орехите, деликатеси)
        if any(w in t for w in ['месо', 'свинск', 'пиле', 'телешк', 'кайма', 'надениц', 'салам', 'колбас', 'филе', 'кренвирш', 'шунка', 'бекон', 'риба', 'деликатес', 'суджук', 'луканка', 'шпек', 'роле', 'бабек', 'пастърма', 'старец', 'орехите', 'перелик', 'тандем', 'градус', 'кент', 'карначе', 'кебапч', 'кюфте']):
            return 7, PRODUCT_GROUPS[7]["name"]
            
        # 8. Млечни
        if any(w in t for w in ['сирене', 'кашкавал', 'кисело мляко', 'прясно мляко', 'масло', 'извара', 'сметана', 'айрян', 'катък', 'топено сирене']):
            return 8, PRODUCT_GROUPS[8]["name"]
            
        # 2. Ядки / чипс (внимание: 'орех' само ако не е марката 'орехите')
        has_nuts = any(w in t for w in ['ядки', 'фъстък', 'бадем', 'лешник', 'чипс', 'снакс', 'пуканки', 'брускети', 'пелети', 'семки', 'слънчоглед', 'pom bar', 'пом бер', 'попкорн'])
        if has_nuts or ('орех' in t and 'орехите' not in t):
            return 2, PRODUCT_GROUPS[2]["name"]
            
        # 3. Захарни
        if any(w in t for w in ['шоколад', 'вафла', 'вафли', 'бисквит', 'бисквити', 'бонбон', 'бонбони', 'десерт', 'кроасан', 'кекс', 'локум', 'халва', 'мед', 'нутела', 'крем', 'донът', 'еверест']):
            return 3, PRODUCT_GROUPS[3]["name"]
            
        # 5. Препарати / хигиена
        if any(w in t for w in ['прах за пране', 'гел за пране', 'капсули за пране', 'омекотител', 'веро', 'сапун', 'шампоан', 'паста за зъби', 'тоалетна хартия', 'салфетки', 'кърпи', 'препарат', 'доместос', 'ароматизатор', 'памперс', 'пелени', 'душ гел', 'дезодорант', 'airwick', 'glade', 'освежител', 'пълнител', 'белина', 'чист', 'фарел', 'пръчици']):
            return 5, PRODUCT_GROUPS[5]["name"]
            
        # 6. Масови (олио, брашно, консерви, плодове, зеленчуци и др.)
        if any(w in t for w in ['олио', 'зехтин', 'брашно', 'захар', 'ориз', 'макарон', 'спагет', 'лютеница', 'консерва', 'домат', 'краставиц', 'картоф', 'плод', 'зеленчук', 'боб', 'леща', 'сол', 'оцет', 'грах', 'царевица']):
            return 6, PRODUCT_GROUPS[6]["name"]
            
        # По подразбиране група 6 (масови)
        return 6, PRODUCT_GROUPS[6]["name"]

    # --------------------------------------------------------------------------
    # Почистване на маркетингово име за брошура
    # --------------------------------------------------------------------------
    def clean_marketing_name(self, raw_name):
        """Преобразува сурово име от доставчик в стегнато маркетингово наименование."""
        if not raw_name:
            return ""
        name = raw_name.strip()
        # Премахване на вътрешни кодове в скоби или началото (напр. "[12345] Продукт")
        name = re.sub(r'^\[\d+\]\s*', '', name)
        name = re.sub(r'^\d+\s*[-/]\s*', '', name)
        # Стандартизиране на думи за видове
        name = re.sub(r'(?i)\b(различни видове|избрани видове|асортимент|видове асортимент)\b', 'видове', name)
        # Премахване на излишни дублирани интервали
        name = re.sub(r'\s+', ' ', name).strip()
        return name

    # --------------------------------------------------------------------------
    # Резолвиране на входящ артикул от доставчик (Layer A -> Layer B + C)
    # --------------------------------------------------------------------------
    def resolve_supplier_item(self, supplier, raw_name, barcode=None, unit=None, promo_dc=None, reg_dc=None, rec_promo_pc=None):
        """
        Решава съответствието за даден ред от бланката:
        1. Точно съвпадение в Речника (Mappings)
        2. Съвпадение по Баркод (в речника или в CS.FDB)
        3. Размито търсене (Fuzzy Match) срещу номенклатурата на доставчика в Мистрал
        """
        supplier_norm = normalize_text(supplier)
        name_norm = normalize_text(raw_name)
        barcode_clean = str(barcode).strip() if barcode else ""
        if barcode_clean.endswith('.0'):
            barcode_clean = barcode_clean[:-2]

        map_key = f"{supplier_norm}|||{name_norm}"

        # ----------------------------------------------------------------------
        # 1. Проверка в персистентния речник на съответствията
        # ----------------------------------------------------------------------
        # 1.1 Точно съвпадение за конкретния доставчик
        if map_key in self.data["mappings"]:
            item = dict(self.data["mappings"][map_key])
            item["status"] = "EXACT_MAPPING"
            item["confidence"] = 1.0
            item["source"] = "потвърден речник (памет)"
            return item

        # 1.2 Проверка в общия речник ("общи")
        general_key = f"общи|||{name_norm}"
        if general_key in self.data["mappings"]:
            item = dict(self.data["mappings"][general_key])
            item["supplier"] = supplier
            item["status"] = "EXACT_MAPPING"
            item["confidence"] = 0.95
            item["source"] = "общ речник (памет)"
            return item

        # ----------------------------------------------------------------------
        # 2. Проверка по Баркод
        # ----------------------------------------------------------------------
        if barcode_clean:
            # 2.1 Баркод индекс в локалния речник
            if barcode_clean in self.data["barcode_index"]:
                target_key = self.data["barcode_index"][barcode_clean]
                if target_key in self.data["mappings"]:
                    item = dict(self.data["mappings"][target_key])
                    item["status"] = "BARCODE_MAPPING"
                    item["confidence"] = 1.0
                    item["source"] = "баркод от речника"
                    return item

            # 2.2 Търсене на баркода директно в CS.FDB
            db_match = self._lookup_barcode_in_db(barcode_clean)
            if db_match:
                gid, gname = self.detect_product_group(db_match["mistral_name"])
                mkt_name = self.clean_marketing_name(raw_name or db_match["mistral_name"])
                return {
                    "status": "BARCODE_MISTRAL",
                    "confidence": 0.98,
                    "source": "баркод в Мистрал (CS.FDB)",
                    "supplier": supplier,
                    "raw_name": raw_name,
                    "barcode": barcode_clean,
                    "mistral_codes": [db_match["code"]],
                    "mistral_names": [db_match["mistral_name"]],
                    "marketing_name": mkt_name,
                    "unit": unit or extract_weight_info(raw_name) or "бр.",
                    "group_id": gid,
                    "group_name": gname,
                    "reg_pc": db_match.get("reg_pc") or rec_promo_pc or 0.0,
                    "reg_dc": db_match.get("reg_dc") or reg_dc or 0.0,
                    "promo_dc": promo_dc or 0.0,
                    "confirmed": False
                }

        # ----------------------------------------------------------------------
        # 3. Размито търсене (Fuzzy Match) срещу каталога на доставчика
        # ----------------------------------------------------------------------
        candidates = self._fuzzy_search_candidates(supplier, raw_name)
        gid, gname = self.detect_product_group(raw_name)
        mkt_name = self.clean_marketing_name(raw_name)

        if candidates:
            best = candidates[0]
            # Проверка дали артикулът е асортимент от видове
            is_assortment = any(w in name_norm for w in ['видове', 'асортимент', 'избрани'])
            assortment_codes = [best["code"]]
            assortment_names = [best["mistral_name"]]

            if is_assortment and len(candidates) > 1:
                # Вземаме всички близки кандидати от същия бранд/група със сходство над 0.50
                for c in candidates[1:]:
                    if c["score"] >= 0.50 and c["code"] not in assortment_codes:
                        assortment_codes.append(c["code"])
                        assortment_names.append(c["mistral_name"])

            return {
                "status": "FUZZY_PROPOSED",
                "confidence": round(best["score"], 2),
                "source": "размито напасване (Fuzzy Match)",
                "supplier": supplier,
                "raw_name": raw_name,
                "barcode": barcode_clean,
                "mistral_codes": assortment_codes,
                "mistral_names": assortment_names,
                "marketing_name": mkt_name,
                "unit": unit or extract_weight_info(raw_name) or "бр.",
                "group_id": gid,
                "group_name": gname,
                "reg_pc": best.get("reg_pc") or rec_promo_pc or 0.0,
                "reg_dc": best.get("reg_dc") or reg_dc or 0.0,
                "promo_dc": promo_dc or 0.0,
                "all_candidates": candidates[:5],
                "confirmed": False
            }

        # Няма намерени кандидати
        return {
            "status": "NO_MATCH",
            "confidence": 0.0,
            "source": "неразпознат",
            "supplier": supplier,
            "raw_name": raw_name,
            "barcode": barcode_clean,
            "mistral_codes": [],
            "mistral_names": [],
            "marketing_name": mkt_name,
            "unit": unit or extract_weight_info(raw_name) or "бр.",
            "group_id": gid,
            "group_name": gname,
            "reg_pc": rec_promo_pc or 0.0,
            "reg_dc": reg_dc or 0.0,
            "promo_dc": promo_dc or 0.0,
            "all_candidates": [],
            "confirmed": False
        }

    # --------------------------------------------------------------------------
    # Записване / Потвърждение на мапинг (1 клик от потребителя)
    # --------------------------------------------------------------------------
    def confirm_mapping(self, supplier, raw_name, mistral_codes, mistral_names, marketing_name, unit, group_id, barcode=None, reg_pc=None, reg_dc=None):
        """Записва потвърдената връзка в локалния персистентен речник."""
        supplier_norm = normalize_text(supplier)
        name_norm = normalize_text(raw_name)
        map_key = f"{supplier_norm}|||{name_norm}"

        if isinstance(mistral_codes, (int, str)):
            mistral_codes = [int(mistral_codes)]
        else:
            mistral_codes = [int(c) for c in mistral_codes if str(c).strip()]

        if isinstance(mistral_names, str):
            mistral_names = [mistral_names]

        gid = int(group_id)
        gname = PRODUCT_GROUPS.get(gid, {}).get("name", "масови")

        barcode_clean = str(barcode).strip() if barcode else ""
        if barcode_clean.endswith('.0'):
            barcode_clean = barcode_clean[:-2]

        record = {
            "supplier": supplier.strip(),
            "raw_name": raw_name.strip(),
            "barcode": barcode_clean,
            "mistral_codes": mistral_codes,
            "mistral_names": mistral_names,
            "marketing_name": marketing_name.strip(),
            "unit": unit.strip(),
            "group_id": gid,
            "group_name": gname,
            "reg_pc": float(reg_pc or 0.0),
            "reg_dc": float(reg_dc or 0.0),
            "confirmed": True,
            "updated_at": datetime.datetime.now().isoformat()
        }

        self.data["mappings"][map_key] = record

        if barcode_clean:
            self.data["barcode_index"][barcode_clean] = map_key

        self.save_storage()
        return record

    # --------------------------------------------------------------------------
    # Вътрешни методи: Търсене в CS.FDB и Каталог
    # --------------------------------------------------------------------------
    def _lookup_barcode_in_db(self, barcode):
        conn = self.get_db()
        if not conn:
            return None
        try:
            cur = conn.cursor()
            sql = """
                SELECT FIRST 1 b.STORAGEMATERIALCODE, m.MATERIAL, p.PRICE, m.LASTDELIVERYPRICE
                FROM BARCODE b
                JOIN MATERIAL m ON b.STORAGEMATERIALCODE = m.MATERIALCODE AND m.LOCATIONID = 1
                LEFT JOIN PRICELIST p ON p.MATERIALCODE = m.MATERIALCODE AND p.LOCATIONID = 1
                WHERE b.CODE = ?
            """
            cur.execute(sql, (str(barcode),))
            row = cur.fetchone()
            if row:
                return {
                    "code": int(row[0]),
                    "mistral_name": row[1].strip() if row[1] else "",
                    "reg_pc": float(row[2]) if row[2] is not None else 0.0,
                    "reg_dc": float(row[3]) if row[3] is not None else 0.0
                }
        except Exception:
            pass
        return None

    def _load_all_materials(self):
        """Зарежда пълния кеш от материали от mistral_materials_cache.json за мигновен офлайн достъп."""
        if self._materials_cache is not None:
            return self._materials_cache
        if os.path.exists(self._materials_cache_file):
            try:
                with open(self._materials_cache_file, "r", encoding="utf-8") as f:
                    self._materials_cache = json.load(f)
                    return self._materials_cache
            except Exception:
                pass
        self._materials_cache = []
        return self._materials_cache

    def _load_partner_catalog(self, supplier):
        """Зарежда номенклатурата за даден доставчик от CS.FDB или локалния кеш."""
        supplier_norm = normalize_text(supplier)
        if supplier_norm in self._partner_cache:
            return self._partner_cache[supplier_norm]

        items = []
        conn = self.get_db()
        if conn:
            try:
                cur = conn.cursor()
                # 1. Търсим partnernum по име на доставчика
                cur.execute("SELECT NUM, NAME FROM PARTNER WHERE LOWER(NAME) LIKE ?", (f"%{supplier_norm[:6]}%",))
                p_rows = cur.fetchall()
                p_nums = [r[0] for r in p_rows]

                if p_nums:
                    in_p = ",".join(str(p) for p in p_nums)
                    sql = f"""
                        SELECT m.MATERIALCODE, m.MATERIAL, p.PRICE, m.LASTDELIVERYPRICE
                        FROM MATERIAL m
                        LEFT JOIN PRICELIST p ON p.MATERIALCODE = m.MATERIALCODE AND p.LOCATIONID = 1
                        WHERE m.LOCATIONID = 1 AND (m.PARTNERNUM IN ({in_p}) OR m.DEFAULTPARTNERNUM IN ({in_p}))
                    """
                    cur.execute(sql)
                    for r in cur.fetchall():
                        items.append({
                            "code": int(r[0]),
                            "mistral_name": r[1].strip() if r[1] else "",
                            "reg_pc": float(r[2]) if r[2] is not None else 0.0,
                            "reg_dc": float(r[3]) if r[3] is not None else 0.0
                        })
            except Exception:
                pass

        # 2. Опит през бързия кеш mistral_materials_cache.json
        if not items:
            all_mats = self._load_all_materials()
            if all_mats and supplier_norm:
                s_key = supplier_norm[:6]
                for m in all_mats:
                    p_norm = normalize_text(m.get("partner", ""))
                    if s_key and s_key in p_norm:
                        items.append({
                            "code": int(m["code"]),
                            "mistral_name": m["name"],
                            "reg_dc": float(m.get("reg_dc", 0.0)),
                            "reg_pc": float(m.get("reg_pc", 0.0)),
                            "partner": m.get("partner", "")
                        })

        # 3. Резервен опит от 123-2608.xlsx
        if not items and os.path.exists(self.deliveries_file) and not self._materials_cache:
            try:
                if self._deliveries_df is None:
                    import pandas as pd
                    self._deliveries_df = pd.read_excel(self.deliveries_file, usecols=['Партньор', 'Номер', 'Име на материал', 'Ед. цена', 'Продажна цена'])
                
                df = self._deliveries_df
                mask = df['Партньор'].astype(str).str.contains(supplier[:6], case=False, na=False)
                df_p = df[mask].drop_duplicates(subset=['Номер', 'Име на материал'])
                for _, r in df_p.iterrows():
                    items.append({
                        "code": int(r['Номер']),
                        "mistral_name": str(r['Име на материал']).strip(),
                        "reg_dc": float(r['Ед. цена']) if pd.notna(r['Ед. цена']) else 0.0,
                        "reg_pc": float(r['Продажна цена']) if pd.notna(r['Продажна цена']) else 0.0
                    })
            except Exception:
                pass

        self._partner_cache[supplier_norm] = items
        return items

    def _fuzzy_search_candidates(self, supplier, raw_name):
        """Изпълнява претеглено размито търсене с отчет на токени, грамаж и сходство."""
        items = self._load_partner_catalog(supplier)
        is_global = False
        if not items:
            # Ако доставчикът е бранд (напр. Победа, Chio, Farel, Pom-bar, Тандем),
            # чиито доставки са заведени през дистрибутор, търсим в целия кеш:
            items = self._load_all_materials()
            is_global = True

        if not items:
            return []

        q_norm = normalize_text(raw_name)
        stop_words = {'видове', 'асортимент', 'избрани', 'вида', 'грам', 'грама', 'бр', 'броя', 'цена', 'промо'}
        q_tokens = set([w for w in q_norm.split() if len(w) > 2 and w not in stop_words])
        q_weight = extract_weight_info(raw_name)

        scored = []
        for it in items:
            db_name = it.get("mistral_name") or it.get("name", "")
            db_norm = normalize_text(db_name)
            db_tokens = set([w for w in db_norm.split() if len(w) > 2 and w not in stop_words])

            # Бърз префилтър: изисква се поне 1 общ ключов токен
            common_tokens = q_tokens.intersection(db_tokens)
            if not common_tokens:
                continue

            overlap = len(common_tokens) / max(len(q_tokens), 1)
            # Ако съвпадението на токени е твърде ниско, пропускаме бавния SequenceMatcher
            if overlap < 0.20:
                continue

            db_weight = extract_weight_info(db_name)
            seq_score = difflib.SequenceMatcher(None, q_norm, db_norm).ratio()
            score = (overlap * 0.6) + (seq_score * 0.4)

            # Бонус при съвпадение на грамаж
            if q_weight and db_weight and q_weight == db_weight:
                score += 0.15

            if score >= 0.35:
                cand = {
                    "code": int(it["code"]),
                    "mistral_name": db_name,
                    "reg_dc": float(it.get("reg_dc", 0.0)),
                    "reg_pc": float(it.get("reg_pc", 0.0)),
                    "partner": it.get("partner", ""),
                    "score": min(score, 0.95)
                }
                scored.append(cand)

        scored.sort(key=lambda x: x["score"], reverse=True)
        return scored[:5]

    # --------------------------------------------------------------------------
    # Универсален четец на входящи бланки (2027 шаблон + Legacy бланки)
    # --------------------------------------------------------------------------
    def parse_any_supplier_blank(self, file_path):
        """
        Универсален парсер за попълнени бланки:
        1. Поддържа .xlsx (openpyxl) и .xls (xlrd)
        2. Разпознава Новия опростен шаблон 2027 (лист 'Бланка_Предложения')
        3. Разпознава Legacy бланки (Sheet1, заглавки на ред 15, текст на споразумението)
        """
        import pandas as pd
        engine = 'openpyxl' if file_path.endswith('.xlsx') else 'xlrd'
        try:
            xl = pd.ExcelFile(file_path, engine=engine)
        except Exception as e:
            return None, f"Грешка при зареждане на Excel: {str(e)}"

        # Формат А: Нов шаблон 2027
        if 'Бланка_Предложения' in xl.sheet_names:
            df = xl.parse('Бланка_Предложения', header=None)
            brochure = str(df.iloc[4, 1] if pd.notna(df.iloc[4, 1]) else df.iloc[4, 3])
            supplier = str(df.iloc[4, 6] if pd.notna(df.iloc[4, 6]) else df.iloc[4, 11])
            items = []
            for r in range(12, len(df)):
                name = df.iloc[r, 2] if pd.notna(df.iloc[r, 2]) else df.iloc[r, 3]
                if pd.notna(name) and str(name).strip() and not str(name).startswith('Пример'):
                    barcode = str(df.iloc[r, 1]).strip() if pd.notna(df.iloc[r, 1]) else ''
                    unit = str(df.iloc[r, 3]).strip() if pd.notna(df.iloc[r, 3]) else ''
                    def to_flt(v):
                        if pd.isna(v): return 0.0
                        try: return float(str(v).replace(',', '.').strip())
                        except: return 0.0
                    reg_dc = to_flt(df.iloc[r, 4])
                    promo_dc = to_flt(df.iloc[r, 5])
                    rec_pc = to_flt(df.iloc[r, 7])
                    note = str(df.iloc[r, 8]).strip() if pd.notna(df.iloc[r, 8]) else ''

                    res_m = self.resolve_supplier_item(
                        supplier, str(name).strip(),
                        barcode=barcode, unit=unit,
                        promo_dc=promo_dc, reg_dc=reg_dc, rec_promo_pc=rec_pc
                    )

                    items.append({
                        "raw_name": str(name).strip(),
                        "marketing_name": res_m.get("marketing_name") or str(name).strip(),
                        "supplier": supplier,
                        "barcode": barcode or "",
                        "unit": res_m.get("unit") or unit or "бр.",
                        "reg_dc": reg_dc,
                        "promo_dc": promo_dc,
                        "promo_pc": rec_pc or res_m.get("reg_pc", 0.0),
                        "mistral_codes": res_m.get("mistral_codes", []),
                        "group_id": res_m.get("group_id", 6),
                        "group_name": res_m.get("group_name", "масови"),
                        "mapping_status": res_m.get("status"),
                        "confidence": res_m.get("confidence", 0.0),
                        "note": note
                    })

            return {
                "format": "2027_TEMPLATE",
                "brochure": brochure,
                "supplier": supplier,
                "items_count": len(items),
                "items": items
            }, None

        # Формат Б: Legacy бланка
        df = xl.parse(xl.sheet_names[0], header=None)
        hdr_row = None
        for r in range(min(30, len(df))):
            cells = [str(x).lower().strip() for x in df.iloc[r].dropna().values]
            has_art = any('артикул' in c or 'наименование' in c for c in cells)
            has_price = any('цена' in c or 'дц' in c for c in cells)
            if has_art and has_price:
                hdr_row = r
                break
                
        if hdr_row is None:
            return None, "Не е намерен ред със заглавки (колони Артикул и Доставна цена)"

        hdr = df.iloc[hdr_row]
        col_art, col_reg_dc, col_promo_dc, col_promo_pc, col_barcode = None, None, None, None, None
        for c in range(df.shape[1]):
            val = str(hdr[c]).lower().strip()
            if 'артикул' in val or 'наименование' in val:
                col_art = c
            elif 'базова' in val or ('редовна' in val and 'цена' in val):
                col_reg_dc = c
            elif 'промоционална' in val or ('промо' in val and 'цена' in val):
                col_promo_dc = c
            elif 'препоръчителна' in val or 'продажна' in val:
                col_promo_pc = c
            elif 'баркод' in val or 'код' in val:
                col_barcode = c

        # Извличане на име на доставчик от текста на споразумението или името на файла
        supplier = os.path.basename(file_path).replace('.xls', '').replace('.xlsx', '')
        for r in range(min(14, len(df))):
            txt = str(df.iloc[r, 0])
            if ('2.' in txt or '2 .' in txt) and 'бакалия' not in txt.lower():
                clean_txt = txt.replace('2.', '').replace('2 .', '').strip()
                parts = clean_txt.split('със седалище')
                cand_supp = parts[0].strip(' .–-„"')
                if len(cand_supp) > 2 and not cand_supp.startswith('...') and 'бакалия' not in cand_supp.lower():
                    supplier = cand_supp
                    break

        # Интелигентно разпознаване на доставчика по брандове в таблицата (Белла, Интерснак, Победа, Фарел, Сириус)
        if not supplier or 'email_' in supplier or 'Бакалия' in supplier:
            sample_txt = " ".join([str(df.iloc[r, c]).lower() for r in range(hdr_row, min(hdr_row+15, len(df))) for c in range(df.shape[1]) if pd.notna(df.iloc[r, c])])
            if any(w in sample_txt for w in ['орехите', 'белла', 'тракийски', 'перелик']):
                supplier = 'Белла България АД'
            elif any(w in sample_txt for w in ['чио', 'пом бер', 'chio', 'помбер']):
                supplier = 'Интерснак България ЕООД'
            elif any(w in sample_txt for w in ['еверест', 'победа', 'fitspo']):
                supplier = 'Победа АД / Ефект 2002'
            elif any(w in sample_txt for w in ['фарел', 'farel']):
                supplier = 'Бикс / Фарел (Боди ЕООД)'
            elif any(w in sample_txt for w in ['троко', 'кирилов']):
                supplier = 'Сириус ООД'

        items = []
        curr_brochure = "Брошура"
        for r in range(hdr_row + 1, len(df)):
            c0_val = str(df.iloc[r, 0]).strip() if pd.notna(df.iloc[r, 0]) else ''
            if 'брошура' in c0_val.lower():
                curr_brochure = c0_val
                
            if col_art is not None and col_art < df.shape[1]:
                name = df.iloc[r, col_art]
                if pd.notna(name) and str(name).strip():
                    s_name = str(name).strip()
                    if not any(w in s_name.lower() for w in ['моля', 'подпишете', 'период на']):
                        def to_flt(v):
                            if pd.isna(v): return 0.0
                            try: return float(str(v).replace(',', '.').strip())
                            except: return 0.0
                        reg_dc = to_flt(df.iloc[r, col_reg_dc]) if col_reg_dc is not None else 0.0
                        promo_dc = to_flt(df.iloc[r, col_promo_dc]) if col_promo_dc is not None else 0.0
                        rec_pc = to_flt(df.iloc[r, col_promo_pc]) if col_promo_pc is not None else 0.0
                        
                        barcode = ''
                        if col_barcode is not None and pd.notna(df.iloc[r, col_barcode]):
                            barcode = str(df.iloc[r, col_barcode]).replace('.0', '').strip()
                        else:
                            for c in range(df.shape[1]):
                                cand = str(df.iloc[r, c]).replace('.0', '').strip()
                                if cand.isdigit() and len(cand) >= 6 and c != col_reg_dc and c != col_promo_dc and c != col_promo_pc:
                                    barcode = cand
                                    break

                        # Резолвиране през трислойния мапинг
                        res_m = self.resolve_supplier_item(
                            supplier, s_name,
                            barcode=barcode,
                            promo_dc=promo_dc, reg_dc=reg_dc, rec_promo_pc=rec_pc
                        )

                        items.append({
                            "raw_name": s_name,
                            "marketing_name": res_m.get("marketing_name") or s_name,
                            "supplier": supplier,
                            "barcode": barcode,
                            "unit": res_m.get("unit") or extract_weight_info(s_name) or "бр.",
                            "reg_dc": reg_dc,
                            "promo_dc": promo_dc,
                            "promo_pc": rec_pc or res_m.get("reg_pc", 0.0),
                            "mistral_codes": res_m.get("mistral_codes", []),
                            "group_id": res_m.get("group_id", 6),
                            "group_name": res_m.get("group_name", "масови"),
                            "mapping_status": res_m.get("status"),
                            "confidence": res_m.get("confidence", 0.0),
                            "note": ""
                        })

        return {
            "format": "LEGACY_BLANK",
            "brochure": curr_brochure,
            "supplier": supplier,
            "items_count": len(items),
            "items": items
        }, None

    # --------------------------------------------------------------------------
    # ЕКСПОРТ 1: Двулистов файл за ПРЕДПЕЧАТ (образец 10-18 October-пълен вариант 1.xls)
    # --------------------------------------------------------------------------
    def export_prepress_excel(self, approved_items, output_path, brochure_title="Брошура"):
        """
        Генерира точния двулистов Excel файл за предпечат:
        - Sheet2: Артикули по 8-те цветни групи [Група, Наименование, М.Ед, Стара ПЦ, Промо ПЦ]
        - Sheet1: Грид решетка (1..12 корица + страници по категории) и номериран списък (1..201)
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active) # Remove default sheet

        font_hdr = Font(name="Arial", size=10, bold=True, color="000000")
        font_data = Font(name="Arial", size=10, color="000000")
        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB")
        )

        # ----------------------------------------------------------------------
        # ЛИСТ 2: Sheet2 — Одобрените артикули, оцветени по 8-те групи
        # ----------------------------------------------------------------------
        ws2 = wb.create_sheet(title="Sheet2")
        ws2.views.sheetView[0].showGridLines = True

        headers_s2 = ["", "Наименование артикул", "М. Ед", "Пр.ц.", "Промо пр.ц."]
        for col_idx, h in enumerate(headers_s2, 1):
            c = ws2.cell(row=1, column=col_idx, value=h)
            c.font = font_hdr
            c.alignment = Alignment(horizontal="center", vertical="center")

        ws2.column_dimensions["A"].width = 6
        ws2.column_dimensions["B"].width = 45
        ws2.column_dimensions["C"].width = 12
        ws2.column_dimensions["D"].width = 12
        ws2.column_dimensions["E"].width = 14

        # Сортиране по група (1..8)
        sorted_items = sorted(approved_items, key=lambda x: int(x.get("group_id", 6)))

        curr_row = 2
        for it in sorted_items:
            gid = int(it.get("group_id", 6))
            g_info = PRODUCT_GROUPS.get(gid, PRODUCT_GROUPS[6])
            fill_grp = PatternFill(start_color=g_info["color"], end_color=g_info["color"], fill_type="solid")

            # Col A: Група (1..8)
            cA = ws2.cell(row=curr_row, column=1, value=gid)
            cA.font = font_data
            cA.alignment = Alignment(horizontal="center", vertical="center")
            cA.fill = fill_grp
            cA.border = thin_border

            # Col B: Наименование
            cB = ws2.cell(row=curr_row, column=2, value=it.get("marketing_name") or it.get("raw_name"))
            cB.font = font_data
            cB.alignment = Alignment(horizontal="left", vertical="center")
            cB.fill = fill_grp
            cB.border = thin_border

            # Col C: М. Ед
            cC = ws2.cell(row=curr_row, column=3, value=it.get("unit", "бр."))
            cC.font = font_data
            cC.alignment = Alignment(horizontal="center", vertical="center")
            cC.fill = fill_grp
            cC.border = thin_border

            # Col D: Стара ПЦ
            cD = ws2.cell(row=curr_row, column=4, value=float(it.get("reg_pc", 0.0)))
            cD.font = font_data
            cD.alignment = Alignment(horizontal="right", vertical="center")
            cD.number_format = "#,##0.00"
            cD.fill = fill_grp
            cD.border = thin_border

            # Col E: Промо ПЦ
            cE = ws2.cell(row=curr_row, column=5, value=float(it.get("promo_pc", 0.0)))
            cE.font = font_data
            cE.alignment = Alignment(horizontal="right", vertical="center")
            cE.number_format = "#,##0.00"
            cE.fill = fill_grp
            cE.border = thin_border

            curr_row += 1

        # ----------------------------------------------------------------------
        # ЛИСТ 1: Sheet1 — Решетка на страниците и номериран списък (1..201)
        # ----------------------------------------------------------------------
        ws1 = wb.create_sheet(title="Sheet1", index=0)
        ws1.views.sheetView[0].showGridLines = True

        # Секция А: Грид макет на страниците (Редове 1 до 92)
        # Заглавия на категориите на ред 1
        grid_headers = [
            (2, "топ предложения S_mart"),
            (6, "напитки"),
            (11, "ядки/чипс"),
            (12, "захарни"),
            (16, "масови"),
            (18, "препарати"),
            (21, "куче/коте"),
            (23, "месни / млечни")
        ]
        for col_idx, g_title in grid_headers:
            c = ws1.cell(row=1, column=col_idx, value=g_title)
            c.font = font_hdr

        # Позиции 1..12 на първа страница (Корица S_mart)
        top_positions = [
            (7, 2, 1), (7, 3, 2), (7, 4, 3),
            (13, 2, 4), (13, 3, 5), (13, 4, 6),
            (19, 2, 7), (19, 3, 8), (19, 4, 9),
            (25, 2, 10), (25, 3, 11), (25, 4, 12)
        ]
        for r, c, pos_num in top_positions:
            cell = ws1.cell(row=r, column=c, value=float(pos_num))
            cell.font = font_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Секция Б: Номериран списък от ред 99 надолу (1..201)
        ws1.cell(row=99, column=1, value="номер").font = font_hdr
        ws1.cell(row=99, column=2, value="артикул").font = font_hdr
        ws1.cell(row=99, column=3, value="разф").font = font_hdr
        ws1.cell(row=99, column=4, value="стара цена").font = font_hdr
        ws1.cell(row=99, column=5, value="промо").font = font_hdr

        for col_i, w in enumerate([10, 45, 12, 14, 14], 1):
            ws1.column_dimensions[get_column_letter(col_i)].width = w

        for idx, it in enumerate(approved_items[:201], 1):
            r_idx = 99 + idx
            ws1.cell(row=r_idx, column=1, value=float(idx)).font = font_data
            ws1.cell(row=r_idx, column=2, value=it.get("marketing_name") or it.get("raw_name")).font = font_data
            ws1.cell(row=r_idx, column=3, value=it.get("unit", "бр.")).font = font_data
            
            p_old = ws1.cell(row=r_idx, column=4, value=float(it.get("reg_pc", 0.0)))
            p_old.font = font_data
            p_old.number_format = "#,##0.00"

            p_promo = ws1.cell(row=r_idx, column=5, value=float(it.get("promo_pc", 0.0)))
            p_promo.font = font_data
            p_promo.number_format = "#,##0.00"

        wb.save(output_path)
        return output_path

    # --------------------------------------------------------------------------
    # ЕКСПОРТ 2: Файлове за импорт в Мистрал ERP (ПЦ и ЗКП)
    # --------------------------------------------------------------------------
    def export_mistral_import_files(self, approved_items, pc_start, pc_end, zkp_start, zkp_end, output_dir):
        """
        Генерира ПЦ_Импорт.xlsx и ЗКП_Импорт.xlsx за Мистрал ERP.
        При асортименти (1-to-many), разпъва ВСЕКИ вътрешен код от mistral_codes!
        """
        import pandas as pd
        os.makedirs(output_dir, exist_ok=True)

        rows_pc = []
        rows_zkp = []

        for it in approved_items:
            codes = it.get("mistral_codes", [])
            promo_pc = float(it.get("promo_pc", 0.0))
            promo_dc = float(it.get("promo_dc", 0.0))
            promo_dc_wo_vat = round(promo_dc / 1.20, 2)

            for code in codes:
                rows_pc.append({
                    "Код на артикул": int(code),
                    "Промо продажна цена": promo_pc,
                    "Начална дата": pc_start,
                    "Крайна дата": pc_end
                })
                rows_zkp.append({
                    "Код на артикул": int(code),
                    "Промо доставна цена (ЗКП)": promo_dc_wo_vat,
                    "Доставна цена с ДДС": promo_dc,
                    "Начална дата": zkp_start,
                    "Крайна дата": zkp_end
                })

        df_pc = pd.DataFrame(rows_pc)
        df_zkp = pd.DataFrame(rows_zkp)

        pc_file = os.path.join(output_dir, "ПЦ_Импорт_Мистрал.xlsx")
        zkp_file = os.path.join(output_dir, "ЗКП_Импорт_Мистрал.xlsx")

        df_pc.to_excel(pc_file, index=False)
        df_zkp.to_excel(zkp_file, index=False)

        return pc_file, zkp_file
