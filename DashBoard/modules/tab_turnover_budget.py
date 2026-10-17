# -*- coding: utf-8 -*-
"""
================================================================================
Модул: Засичане на обороти, Бюджет и Бонус Оборот (БО)
================================================================================
Интерфейс за Единния Dashboard на верига „Бакалия“.
Спазва формулата за Колона G (Сума от 11-те магазина) и търговските правила по ДТУ.
"""

import os
import sys
import glob
import pandas as pd
import streamlit as st

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
root_dir = os.path.dirname(parent_dir)

STORES_11 = [
    "АРСЕНАЛСКИ", "БАНКЯ", "ВЕЛЧО", "ЛОЗЕНЕЦ", "ЛЮЛИН 7",
    "МАНАСТИРСКИ ЛИВАДИ", "НАДЕЖДА", "ОБЕЛЯ2", "РЕДУТА", "СВОБОДА", "ХАЛИ ОКР"
]


def render_turnover_budget():
    """Главен рендер за Таб: Засичане на обороти & Бюджет."""
    st.subheader("📊 Засичане на обороти по обекти, Бюджет и Бонус Оборот (БО)")
    st.caption("Формиране на Колона G (11 обекта с ДДС), съпоставка с справки 123 и изчисляване на ДТУ")

    st.info("📐 **Златно правило за Колона G:** Оборотът се изчислява строго като чиста сума от 11-те търговски магазина без складови трансфери.")

    # Търсене на налични файлове за засичане
    budget_files = glob.glob(os.path.join(root_dir, "БЮДЖЕТ*.xls*"))
    report_123_files = glob.glob(os.path.join(root_dir, "123-*.xls*"))
    turnover_audits = glob.glob(os.path.join(root_dir, "ЗАСИЧАНЕ_ОБОРОТИ_*.xlsx"))

    c_b, c_123 = st.columns(2)
    with c_b:
        st.write("#### 📕 Бюджет (Master Read-Only):")
        if budget_files:
            b_selected = st.selectbox("Изберете файл с Бюджет:", [os.path.basename(f) for f in budget_files], key="sel_budget_file")
            st.caption(f"Път: `{b_selected}`")
        else:
            st.warning("Не са открити файлове `БЮДЖЕТ*.xls`.")

    with c_123:
        st.write("#### 📑 Справки 123 (Складови доставки):")
        if report_123_files:
            r_selected = st.selectbox("Изберете справка 123:", [os.path.basename(f) for f in report_123_files], key="sel_123_file")
            st.caption(f"Път: `{r_selected}`")
        else:
            st.warning("Не са открити файлове `123-*.xlsx`.")

    st.markdown("---")
    st.write("### 🧮 Интерактивен калкулатор за Бонус Оборот (БО) и такси по ДТУ:")

    col_c1, col_c2, col_c3, col_c4 = st.columns(4)
    with col_c1:
        partner_calc = st.text_input("Партньор:", value="Виста АВТ", key="tb_calc_partner")
    with col_c2:
        turnover_val = st.number_input("Реализиран оборот (€ с ДДС):", value=12450.00, step=100.0, key="tb_calc_turnover")
    with col_c3:
        bonus_pct = st.number_input("Договорен % БО:", value=5.0, step=0.5, key="tb_calc_pct")
    with col_c4:
        brochure_count = st.number_input("Брой брошури (110 лв / 56.24 €):", value=1, step=1, key="tb_calc_broch")

    # Сметки
    turnover_no_vat = turnover_val / 1.20
    bonus_amount = turnover_no_vat * (bonus_pct / 100.0)
    brochure_fee = brochure_count * 56.24  # 110 лв = 56.24 €
    total_to_invoice = bonus_amount + brochure_fee

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Оборот без ДДС", f"{turnover_no_vat:.2f} €")
    k2.metric("Бонус Оборот (БО)", f"{bonus_amount:.2f} €")
    k3.metric("Такси Брошури", f"{brochure_fee:.2f} €")
    k4.metric("Общо за фактуриране", f"{total_to_invoice:.2f} €")

    st.markdown("---")
    st.write("### 📁 Налични сравнителни анализи на оборотите:")
    if turnover_audits:
        for audit_f in turnover_audits:
            f_name = os.path.basename(audit_f)
            col_fn, col_dl = st.columns([3, 1])
            with col_fn:
                st.write(f"📄 **{f_name}** ({os.path.getsize(audit_f)/1024:.1f} KB)")
            with col_dl:
                with open(audit_f, "rb") as af:
                    st.download_button(
                        f"📥 Свали",
                        data=af.read(),
                        file_name=f_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_audit_{f_name}"
                    )
    else:
        st.info("Все още няма генерирани сравнителни таблици за обороти в папката.")
