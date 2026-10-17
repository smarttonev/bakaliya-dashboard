# -*- coding: utf-8 -*-
r"""
================================================================================
ВЕРИГА ЗА БЪРЗООБОРОТНИ СТОКИ „БАКАЛИЯ“ — ЕДИНЕН ОПЕРАТИВЕН DASHBOARD
================================================================================
Главно уеб табло за управление на финансовите, ценови и търговски процеси.
Обединява всички разработени инструменти:
  1. 🏷️ Импорт на цени (ПЦ / ЗКП) в Mistral ERP (CS.FDB за 11 обекта)
  2. 🏦 Банково засичане на Банка ДСК в Монтевидео (MONTEVIDEO.FDB)
  3. 🛒 Хъб за промоционални брошури 2027 (Одит по 4-те правила, Matcher, Бланки)
  4. 📦 Импорт на складови доставки в Монтевидео
  5. 📊 Засичане на обороти по обекти и Колона G в Бюджета
  6. 🐙 GitHub Синхронизация & Системна диагностика
================================================================================
"""

import os
import sys
import datetime
import pandas as pd
import streamlit as st

# Осигуряване на правилни пътища за модулите
current_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = os.path.dirname(current_dir)
for p in [current_dir, root_dir]:
    if p not in sys.path:
        sys.path.insert(0, p)

# Импорт на модулите за отделните раздели
from modules.tab_price_importer import render_price_importer
from modules.tab_bank_recon import render_bank_recon
from modules.tab_brochure_hub import render_brochure_hub
from modules.tab_delivery_importer import render_delivery_importer
from modules.tab_turnover_budget import render_turnover_budget
from modules.tab_github_sync import render_github_sync

# Конфигурация на страницата
st.set_page_config(
    page_title="Бакалия — Главен оперативен Dashboard",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Корпоративен дизайн и визуално форматиране
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        padding: 22px 28px;
        border-radius: 12px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    .main-header h1 {
        margin: 0;
        font-size: 26px;
        font-weight: 700;
        letter-spacing: 0.5px;
        color: #FFFFFF !important;
    }
    .main-header p {
        margin: 6px 0 0 0;
        font-size: 14px;
        opacity: 0.9;
        color: #E2E8F0;
    }
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 600;
    }
    .badge-ok { background-color: #DEF7EC; color: #03543F; }
    .badge-warn { background-color: #FEF08A; color: #854D0E; }
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        background-color: #F8FAFC;
        border-radius: 8px 8px 0px 0px;
        padding: 10px 16px;
        font-weight: 600;
        font-size: 14px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFFFFF !important;
        border-bottom: 3px solid #2563EB !important;
        color: #1E3A8A !important;
    }
</style>
""", unsafe_allow_html=True)

# Горен фирмен банер
st.markdown("""
<div class="main-header">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h1>🏪 ВЕРИГА СУПЕРМАРКЕТИ „БАКАЛИЯ“</h1>
            <p>Централизиран контролен панел за търговски одит, ценови импорти, банкови засичания и репликация</p>
        </div>
        <div style="text-align: right;">
            <span class="status-badge badge-ok">🟢 11 Магазина + ЦО</span><br>
            <span style="font-size: 12px; color: #CBD5E1;">Валутен режим: 100% EUR (€)</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Странична лента (Sidebar)
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/shop.png", width=64)
    st.title("Бакалия ERP")
    st.caption("Финансов и търговски контрол")
    
    st.markdown("---")
    st.markdown("### 🏢 Обекти на веригата:")
    st.markdown("""
    - **1:** Централен склад (CS.FDB)
    - **7:** Монтевидео (ЦО)
    - **5:** Надежда | **6:** Арсеналски
    - **18:** Лозенец | **19:** Свобода
    - **22:** Обеля 2 | **27:** Манастирски ливади
    - **30:** Люлин 7 | **31:** Банкя
    - **32:** Редута | **33:** Хали Окръжна
    - **35:** Велчо Атанасов
    """)
    
    st.markdown("---")
    st.caption("🔒 **Директива за интегритет:** Всички модификации се изпълняват първо в Dry-Run симулатор.")
    st.caption(f"🕒 **Час на системата:** {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}")

# Организиране на главните функционални раздели чрез табове
tabs = st.tabs([
    "🏷️ Импорт на цени (ПЦ/ЗКП)",
    "🏦 Банково засичане (ДСК)",
    "🛒 Промоционални брошури 2027",
    "📦 Доставки в Монтевидео",
    "📊 Засичане на обороти & Бюджет",
    "🐙 GitHub & Система"
])

# 1. ТАБ: Импорт на цени
with tabs[0]:
    render_price_importer()

# 2. ТАБ: Банково засичане
with tabs[1]:
    render_bank_recon()

# 3. ТАБ: Промоционални брошури 2027
with tabs[2]:
    render_brochure_hub()

# 4. ТАБ: Доставки в Монтевидео
with tabs[3]:
    render_delivery_importer()

# 5. ТАБ: Засичане на обороти & Бюджет
with tabs[4]:
    render_turnover_budget()

# 6. ТАБ: GitHub Синхронизация & Диагностика
with tabs[5]:
    render_github_sync()
