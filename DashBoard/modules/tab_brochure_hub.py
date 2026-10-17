# -*- coding: utf-8 -*-
"""
================================================================================
Модул: Централен хъб за промоционални брошури 2027
================================================================================
Интерфейс за Единния Dashboard на верига „Бакалия“.
Включва:
  - 🛡️ Търговски одит & 4-те задължителни правила
  - 🔍 Интелигентно трислойно напасване (Matcher)
  - 📥 Обработка на бланки от доставчици и сборен пул
  - 📅 График и календар за 2027 г. (25 брошури)
  - 📚 Речник на съответствията (Памет)
  - 📧 Имейл център за промоционални оферти
"""

import os
import sys
import io
import json
import datetime
import pandas as pd
import streamlit as st

# Добавяне на пътя към главната директория за модулите на брошурите
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
root_dir = os.path.dirname(parent_dir)
for p in [root_dir, parent_dir]:
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    from brochure_mapping import BrochureMappingManager, PRODUCT_GROUPS
    from brochure_rules import audit_brochure_item, detect_segment
    from brochure_email_bot import BrochureEmailBot
    HAS_BROCHURE_CORE = True
except ImportError as e:
    HAS_BROCHURE_CORE = False
    IMPORT_ERR = str(e)


# График на брошурите за 2027 г.
SCHEDULE_DATA = [
    (1, '2027-01-09', '2027-01-17', 9),
    (2, '2027-01-23', '2027-01-31', 9),
    (3, '2027-02-06', '2027-02-14', 9),
    (4, '2027-02-20', '2027-02-28', 9),
    (5, '2027-03-06', '2027-03-14', 9),
    (6, '2027-03-20', '2027-03-28', 9),
    (7, '2027-04-03', '2027-04-11', 9),
    (8, '2027-04-17', '2027-04-25', 9),
    (9, '2027-04-29', '2027-05-09', 11),
    (10, '2027-05-22', '2027-05-30', 9),
    (11, '2027-06-05', '2027-06-13', 9),
    (12, '2027-06-19', '2027-06-27', 9),
    (13, '2027-07-03', '2027-07-11', 9),
    (14, '2027-07-17', '2027-07-25', 9),
    (15, '2027-08-07', '2027-08-15', 9),
    (16, '2027-08-21', '2027-08-29', 9),
    (17, '2027-09-04', '2027-09-12', 9),
    (18, '2027-09-18', '2027-09-26', 9),
    (19, '2027-10-02', '2027-10-10', 9),
    (20, '2027-10-23', '2027-10-31', 9),
    (21, '2027-11-06', '2027-11-14', 9),
    (22, '2027-11-20', '2027-11-28', 9),
    (23, '2027-12-04', '2027-12-12', 9),
    (24, '2027-12-18', '2027-12-26', 9),
    (25, '2027-12-27', '2028-01-02', 7),
]


def render_brochure_hub():
    """Главен рендер за Таб: Промоционални брошури 2027."""
    st.subheader("🛒 Централен хъб за промоционални брошури 2027")
    st.caption("Трислоен мапинг (Доставчик -> Мистрал ERP -> Предпечат), Търговски одит и Автоматизация")

    if not HAS_BROCHURE_CORE:
        st.error(f"❌ Модулите за брошури не бяха заредени: {IMPORT_ERR}")
        return

    # Инициализация на мениджъра в сесията
    if "mapping_mgr" not in st.session_state:
        st.session_state["mapping_mgr"] = BrochureMappingManager()
    mgr = st.session_state["mapping_mgr"]

    b_tabs = st.tabs([
        "🛡️ 1. Търговски одит & 4-те правила",
        "🔍 2. Интелигентно напасване (Matcher)",
        "📥 3. Сборен брошурен пул & Бланки",
        "📅 4. График & Календар 2027",
        "📚 5. Речник на съответствията",
        "📧 6. Имейл Център"
    ])

    # --------------------------------------------------------------------------
    # ТАБ 1: Търговски одит
    # --------------------------------------------------------------------------
    with b_tabs[0]:
        st.write("### 🛡️ Търговски одит по 4-те задължителни фирмени правила на Бакалия")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.info("🚫 **Правило 1: Забрана за идентични артикули** — Не се допуска дублиране в тясна подгрупа в едно издание.")
            st.error("⛔ **Правило 2: Забрана за продажба на загуба (Zero Loss)** — Промо ПЦ без ДДС > Промо ДЦ (ЗКП).")
        with col_r2:
            st.warning("⚠️ **Правило 3: Защита Промо ПЦ vs Базова ДЦ** — Промо ПЦ с ДДС ≥ Редовна базова доставна цена с ДДС.")
            st.success("🆕 **Правило 4: Забрана за дебют на нови артикули** — Допускат се само артикули с история в последните 90 дни.")

        st.markdown("---")
        st.write("#### 🧪 Симулатор на търговския одит (Тествайте артикул на живо):")
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        with c1:
            sim_name = st.text_input("Наименование на артикула:", value="Верея кисело мляко 2% 400гр", key="sim_b_name")
        with c2:
            sim_reg_dc = st.number_input("Редовна ДЦ (€ с ДДС):", value=0.72, step=0.05, format="%.2f", key="sim_b_rdc")
        with c3:
            sim_promo_dc = st.number_input("Промо ДЦ (€ с ДДС):", value=0.58, step=0.05, format="%.2f", key="sim_b_pdc")
        with c4:
            sim_promo_pc = st.number_input("Промо ПЦ (€ с ДДС):", value=0.79, step=0.05, format="%.2f", key="sim_b_ppc")

        if st.button("🔍 Изпълни одит", type="primary", key="btn_sim_audit"):
            test_item = {
                "code": 5147,
                "name": sim_name,
                "reg_dc": sim_reg_dc,
                "promo_dc": sim_promo_dc,
                "promo_pc": sim_promo_pc,
                "edition": "Брошура 1"
            }
            res = audit_brochure_item(test_item, [], {5147, 3930, 15044})
            if res["status"] == "BLOCKED":
                st.error(f"⛔ БЛОКИРАН ЗА БРОШУРАТА! Марж: {res['margin_pct']:.1f}%")
            elif res["status"] == "WARNING":
                st.warning(f"⚠️ ДОПУСНАТ С ПРЕДУПРЕЖДЕНИЕ! Марж: {res['margin_pct']:.1f}%")
            else:
                st.success(f"✅ НАПЪЛНО ОДОБРЕН! Марж: {res['margin_pct']:.1f}%")

            for fl in res.get("flags", []):
                st.write(f"- **[{fl['level']}]** {fl['msg']}")

    # --------------------------------------------------------------------------
    # ТАБ 2: Matcher
    # --------------------------------------------------------------------------
    with b_tabs[1]:
        st.write("### 🔍 Трислоен Мапинг: Доставчик ➔ Мистрал ERP ➔ Предпечат")
        m_col1, m_col2, m_col3 = st.columns([1.5, 2.5, 1])
        with m_col1:
            supplier_input = st.selectbox(
                "Доставчик:",
                ["Аксон", "Белла", "Амперел", "Оптимо", "Сириус", "Астра и Ко", "Фикосота", "ОМК", "Тандем", "Друг..."],
                key="match_sup_input"
            )
            if supplier_input == "Друг...":
                supplier_input = st.text_input("Въведете име:", value="Аксон", key="match_sup_custom")
        with m_col2:
            query_text = st.text_input("Име по оферта от доставчик:", value="Барила спагети N5 500гр", key="match_q_text")
        with m_col3:
            query_barcode = st.text_input("Баркод (по желание):", value="", key="match_q_bar")

        if st.button("🚀 Резолвирай артикул", type="primary", key="btn_match_item") or query_text:
            with st.spinner("Търсене в речника и номенклатурите на Мистрал..."):
                res = mgr.resolve_supplier_item(supplier_input, query_text, barcode=query_barcode)
                status = res.get("status")
                codes = res.get("mistral_codes", [])
                
                if status == "EXACT_MAPPING":
                    st.success("✅ НАПЪЛНО ПОТВЪРДЕНО СЪОТВЕТСТВИЕ ОТ РЕЧНИКА! (100% сигурност)")
                elif "BARCODE" in str(status):
                    st.success(f"📦 РАЗПОЗНАТ ПО БАРКОД В МИСТРАЛ! (Кодове: {codes})")
                elif status == "FUZZY_PROPOSED":
                    st.info(f"🔍 ПРЕДЛОЖЕНО СЪОТВЕТСТВИЕ (Сходство: {res.get('confidence', 0)*100:.1f}%)")
                else:
                    st.warning("⚠️ Не е открито автоматично съответствие.")

                c_l1, c_l2, c_l3 = st.columns(3)
                with c_l1:
                    st.info(f"**Слой А (Доставчик):**\n- Партньор: {supplier_input}\n- Текст: {query_text}")
                with c_l2:
                    st.info(f"**Слой Б (Мистрал ERP):**\n- Кодове: {codes}\n- Ред. ПЦ: {res.get('reg_pc', 0.0):.2f} €\n- Базова ДЦ: {res.get('reg_dc', 0.0):.2f} €")
                with c_l3:
                    st.info(f"**Слой В (Предпечат):**\n- Име: {res.get('marketing_name', query_text)}\n- Мярка: {res.get('unit', 'бр.')}")

                if st.button("💾 Запази в речника за бъдещи брошури", key="btn_save_mapping_rule"):
                    mgr.confirm_mapping(
                        supplier=supplier_input,
                        raw_name=query_text,
                        mistral_codes=codes,
                        mistral_names=res.get("mistral_names", []),
                        marketing_name=res.get("marketing_name", query_text),
                        unit=res.get("unit", "бр."),
                        group_id=res.get("group_id", 6),
                        barcode=query_barcode,
                        reg_pc=res.get("reg_pc"),
                        reg_dc=res.get("reg_dc")
                    )
                    st.success("✅ Записано в базата съответствия!")

    # --------------------------------------------------------------------------
    # ТАБ 3: Сборен брошурен пул
    # --------------------------------------------------------------------------
    with b_tabs[2]:
        st.write("### 📥 Сборен брошурен пул & Зареждане на бланки")
        pool_file = os.path.join(root_dir, "brochure_master_pool.json")
        master_items = []
        if os.path.exists(pool_file):
            try:
                with open(pool_file, "r", encoding="utf-8") as pf:
                    master_items = json.load(pf)
            except Exception:
                master_items = []

        st.metric("Набрани оферти в сборния пул", len(master_items), help="Цел за пълно издание: около 200 позиции")

        if master_items:
            df_pool = pd.DataFrame(master_items)
            display_cols = [c for c in ["supplier", "name", "marketing_name", "promo_dc", "promo_pc", "group_name"] if c in df_pool.columns]
            st.dataframe(df_pool[display_cols], use_container_width=True)
        else:
            st.info("Няма заредени артикули в сборния пул. Можете да качите файл с бланка по-долу.")

        uploaded_blank = st.file_uploader("Качете бланка от доставчик (.xlsx / .xls):", type=["xlsx", "xls"], key="up_blank_file")
        if uploaded_blank:
            temp_blank = os.path.join(root_dir, f"temp_{uploaded_blank.name}")
            with open(temp_blank, "wb") as bf:
                bf.write(uploaded_blank.getbuffer())
            parsed, err = mgr.parse_any_supplier_blank(temp_blank)
            if parsed and parsed.get("items"):
                st.success(f"Разпознати {len(parsed['items'])} оферти от **{parsed.get('supplier', 'Доставчик')}**!")
                st.dataframe(pd.DataFrame(parsed["items"]), use_container_width=True)
            else:
                st.error(f"Грешка при четене на бланката: {err}")

    # --------------------------------------------------------------------------
    # ТАБ 4: График & Календар 2027
    # --------------------------------------------------------------------------
    with b_tabs[3]:
        st.write("### 📅 График и брошурен календар за 2027 г. (25 издания)")
        df_sched = pd.DataFrame(SCHEDULE_DATA, columns=["Брошура №", "Начало", "Край", "Продължителност (дни)"])
        st.dataframe(df_sched, use_container_width=True, height=400)

    # --------------------------------------------------------------------------
    # ТАБ 5: Речник на съответствията
    # --------------------------------------------------------------------------
    with b_tabs[4]:
        st.write("### 📚 Речник на съответствията (Памет за мапинги)")
        mapping_file = os.path.join(root_dir, "brochure_item_mapping.json")
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, "r", encoding="utf-8") as mf:
                    data = json.load(mf)
                st.caption(f"Общо запаметени съответствия: **{len(data)}**")
                search_word = st.text_input("Търсене в речника:", "", key="search_in_dict")
                filtered_dict = {k: v for k, v in data.items() if search_word.lower() in k.lower()} if search_word else data
                st.json(list(filtered_dict.items())[:30])
            except Exception as e:
                st.error(f"Грешка при зареждане на речника: {e}")
        else:
            st.info("Речникът `brochure_item_mapping.json` все още е празен.")

    # --------------------------------------------------------------------------
    # ТАБ 6: Имейл център
    # --------------------------------------------------------------------------
    with b_tabs[5]:
        st.write("### 📧 Имейл център за брошурни оферти")
        st.caption("Автоматично сваляне и филтриране на предложения от доставчици в Gmail")
        st.info("🤖 Модулът `brochure_email_bot.py` сканира пощата, маркира писмата и премахва дублирания от входящите съгласно фирмената политика.")
        if st.button("📥 Провери пощата за нови брошурни бланки", key="btn_check_email_brochures"):
            st.info("Стартиране на фонова проверка на имейла...")
