# -*- coding: utf-8 -*-
"""
================================================================================
Верига „Бакалия“ — Модул за интелигентно напасване на артикули за брошури
================================================================================
"""

import os
import sys
import re
import json
import difflib
import datetime
import pandas as pd
import openpyxl

# Път към речника със синоними
ALIAS_FILE = os.path.join(os.path.dirname(__file__), "supplier_item_aliases.json")

def load_aliases():
    if os.path.exists(ALIAS_FILE):
        try:
            with open(ALIAS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_alias(supplier, raw_name, matched_code, matched_name):
    aliases = load_aliases()
    key = f"{supplier.strip()}|||{raw_name.strip()}"
    aliases[key] = {
        "supplier": supplier.strip(),
        "raw_name": raw_name.strip(),
        "code": int(matched_code),
        "mistral_name": matched_name.strip(),
        "updated_at": datetime.datetime.now().isoformat()
    }
    with open(ALIAS_FILE, "w", encoding="utf-8") as f:
        json.dump(aliases, f, ensure_ascii=False, indent=2)

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    t = text.lower()
    t = re.sub(r'[\.,/\\\(\)\-\*\+]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

class BrochureItemMatcher:
    def __init__(self, deliveries_file=None):
        self.deliveries_file = deliveries_file or "123-2608.xlsx"
        self.df_db = None
        self.aliases = load_aliases()
        self._load_catalog()

    def _load_catalog(self):
        if os.path.exists(self.deliveries_file):
            df = pd.read_excel(self.deliveries_file, usecols=['Партньор', 'Номер', 'Име на материал', 'Ед. цена', 'Продажна цена'])
            self.df_db = df.drop_duplicates(subset=['Партньор', 'Номер', 'Име на материал'])
        else:
            self.df_db = pd.DataFrame(columns=['Партньор', 'Номер', 'Име на материал', 'Ед. цена', 'Продажна цена'])

    def match_item(self, supplier, raw_name):
        key = f"{supplier.strip()}|||{raw_name.strip()}"
        # 1. Проверка в Речника на синонимите (Паметта)
        if key in self.aliases:
            item_info = self.aliases[key]
            return {
                "status": "EXACT_ALIAS",
                "code": item_info["code"],
                "mistral_name": item_info["mistral_name"],
                "score": 1.0,
                "source": "памет (синоним)"
            }

        norm_query = normalize_text(raw_name)
        q_tokens = set(norm_query.split())

        # 2. Филтриране САМО на артикулите на този доставчик
        mask = self.df_db['Партньор'].str.contains(supplier, case=False, na=False)
        df_partner = self.df_db[mask]
        if len(df_partner) == 0:
            df_partner = self.df_db

        candidates = []
        for _, row in df_partner.iterrows():
            db_name = row['Име на материал']
            db_norm = normalize_text(db_name)
            db_tokens = set(db_norm.split())

            token_overlap = len(q_tokens.intersection(db_tokens)) / max(len(q_tokens), 1)
            seq_ratio = difflib.SequenceMatcher(None, norm_query, db_norm).ratio()
            combined_score = (token_overlap * 0.6) + (seq_ratio * 0.4)

            if combined_score > 0.40:
                candidates.append({
                    "code": int(row['Номер']),
                    "mistral_name": db_name,
                    "score": combined_score,
                    "reg_dc": float(row['Ед. цена']),
                    "reg_pc": float(row['Продажна цена']),
                    "partner": row['Партньор']
                })

        candidates.sort(key=lambda x: x['score'], reverse=True)

        if candidates:
            best = candidates[0]
            if best["score"] >= 0.75:
                status = "AUTO_MATCH"
            else:
                status = "PROPOSED_MATCH"
            best["status"] = status
            best["all_candidates"] = candidates[:5]
            return best

        return {
            "status": "NOT_FOUND",
            "code": None,
            "mistral_name": None,
            "score": 0.0,
            "all_candidates": []
        }
