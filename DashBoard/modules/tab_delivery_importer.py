# -*- coding: utf-8 -*-
"""
================================================================================
Модул: Импорт на доставки в Мистрал ERP — Монтевидео (ЦО)
================================================================================
Интерфейс за Единния Dashboard на верига „Бакалия“.
База данни: 192.168.5.254:E:/db/MONTEVIDEO.FDB (Обект 7 Монтевидео, Склад 1)
"""

import os
import sys
import datetime
from decimal import Decimal
import pandas as pd
import streamlit as st

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
root_dir = os.path.dirname(parent_dir)
for p in [root_dir, parent_dir]:
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    from import_deliveries_montevideo import MontevideoDeliveryImporter, DB_CONFIG_MV
    HAS_DELIVERY_MODULE = True
except ImportError as e:
    HAS_DELIVERY_MODULE = False
    IMPORT_ERR = str(e)


def render_delivery_importer():
    """Главен рендер за Таб: Импорт на доставки в Монтевидео."""
    st.subheader("📦 Импорт на складови доставки в Мистрал ERP — Монтевидео")
    st.caption("Обединяване по партньори, проверка на номенклатури, запис на салда и репликационен лог")

    if not HAS_DELIVERY_MODULE:
        st.error(f"❌ Модулът `import_deliveries_montevideo.py` не беше зареден: {IMPORT_ERR}")
        return

    st.info("🏢 **Обект:** МОНТЕВИДЕО (ID = 7) | **Склад:** Склад 1 | **Оператор:** Тонев (33) | **Цел:** CS.FDB репликация")

    col_up, col_date = st.columns([2, 1])
    with col_up:
        uploaded_file = st.file_uploader(
            "Качете файл с доставки (CSV или Excel):",
            type=["csv", "xlsx", "xls"],
            key="deliv_file_uploader"
        )
    with col_date:
        doc_date = st.date_input("Дата на документите:", value=datetime.date.today(), key="deliv_doc_date")

    if uploaded_file is None:
        st.info("👆 Моля, качете входящ файл с доставки (напр. `доставки_октомври_2026.csv`).")
        return

    temp_path = os.path.join(root_dir, f"temp_{uploaded_file.name}")
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success(f"Зареден файл: **{uploaded_file.name}**")

    # Предварителен анализ
    if st.button("🔍 Стартирай валидация и Dry-Run анализ", type="primary", key="btn_run_deliv_val"):
        with st.spinner("Проверка на партньори и артикули в базата данни на Монтевидео..."):
            importer = MontevideoDeliveryImporter(temp_path, dry_run=True, doc_date=doc_date)
            try:
                importer.load_file()
                importer.connect_db()
                importer.validate_partners_and_materials()
                importer.group_deliveries()
                st.session_state["deliv_importer"] = importer
            except Exception as e:
                st.error(f"Грешка при валидация: {e}")
                return

    if "deliv_importer" in st.session_state:
        imp = st.session_state["deliv_importer"]

        st.markdown("### 📊 Обобщение на доставките за импорт:")
        c1, c2, c3 = st.columns(3)
        c1.metric("Брой контрагенти", len(imp.deliveries))
        total_items_count = sum(len(items) for items in imp.deliveries.values())
        c2.metric("Общо артикулни редове", total_items_count)
        c3.metric("Обща стойност", f"{imp.total_sum:.2f} €")

        # Таблица по партньори
        partner_rows = []
        for p_name, items in imp.deliveries.items():
            p_val = sum(Decimal(str(it.get('total_price', 0))) for it in items)
            p_info = imp.partner_map.get(p_name, {})
            partner_rows.append({
                "Партньор": p_name,
                "Номер в ERP": p_info.get('num', '—'),
                "Брой артикули": len(items),
                "Обща сума (€)": f"{p_val:.2f} €"
            })

        st.dataframe(pd.DataFrame(partner_rows), use_container_width=True)

        st.markdown("---")
        c_c1, c_c2 = st.columns([2, 1])
        with c_c1:
            confirm_deliv = st.checkbox("Потвърждавам записа на доставките в MONTEVIDEO.FDB и генерирането на репликация.", key="confirm_deliv_commit")
        with c_c2:
            if st.button("🚀 Изпълни запис на доставките", type="primary", disabled=not confirm_deliv, key="btn_deliv_commit"):
                with st.spinner("Записване на доставките в Монтевидео..."):
                    try:
                        live_imp = MontevideoDeliveryImporter(temp_path, dry_run=False, doc_date=doc_date)
                        live_imp.load_file()
                        live_imp.connect_db()
                        live_imp.validate_partners_and_materials()
                        live_imp.group_deliveries()
                        live_imp.execute_commit()
                        st.success("✅ Всички доставки бяха успешно записани в MONTEVIDEO.FDB с лог за репликация!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Грешка при запис на доставки: {e}")
