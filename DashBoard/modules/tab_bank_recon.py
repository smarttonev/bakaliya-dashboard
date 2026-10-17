# -*- coding: utf-8 -*-
"""
================================================================================
Модул: Банково засичане и отразяване в Мистрал ERP (Монтевидео)
================================================================================
Интерфейс за Единния Dashboard на верига „Бакалия“.
База данни: 192.168.5.254:E:/db/MONTEVIDEO.FDB
Каса: ДСК Бакалия (ID = 179607, NAME ID = 52) | Оператор: Тонев (ID = 33)
"""

import os
import sys
import glob
import datetime
from decimal import Decimal
import pandas as pd
import streamlit as st

# Добавяне на пътищата за импорт на оригиналните модули
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
root_dir = os.path.dirname(parent_dir)
for p in [parent_dir, root_dir]:
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    from bank_reconciliation import MontevideoBankReconciliationEngine, DB_CONFIG
    HAS_ENGINE = True
except ImportError:
    HAS_ENGINE = False

try:
    import firebirdsql
    HAS_FIREBIRD = True
except ImportError:
    HAS_FIREBIRD = False


def find_candidate_bank_files():
    """Открива потенциални банкови файлове в папката на проекта и в Downloads."""
    search_dirs = [
        root_dir,
        os.path.join(root_dir, "DashBoard"),
        os.path.join(os.path.expanduser('~'), 'Downloads')
    ]
    patterns = ['report*.xml', 'report*.hml', 'report*.XML', 'report*.HML']
    candidates = []
    seen = set()

    for d in search_dirs:
        if not os.path.exists(d):
            continue
        for p in patterns:
            for f in glob.glob(os.path.join(d, p)):
                norm_f = os.path.abspath(f)
                if norm_f not in seen:
                    seen.add(norm_f)
                    try:
                        mtime = os.path.getmtime(norm_f)
                        candidates.append((norm_f, mtime))
                    except Exception:
                        pass

    candidates.sort(key=lambda x: x[1], reverse=True)
    return [c[0] for c in candidates]


def render_bank_recon():
    """Главен рендер за Таб: Банково засичане ДСК."""
    st.subheader("🏦 Банково засичане на извлечения ДСК в Мистрал ERP (Монтевидео)")
    st.caption("Автоматично разпознаване на фактури, насрещни компенсации, наеми със ЗДДФЛ и банкови такси")

    if not HAS_ENGINE:
        st.error("❌ Не може да се зареди основният модул `bank_reconciliation.py`. Моля, проверете файла.")
        return

    # Информационен ред за касата и базата
    col_kasa, col_db = st.columns(2)
    with col_kasa:
        st.info("💼 **Каса:** ДСК Бакалия (`179607`) | **Оператор:** Тонев (`ID 33`) | **Режим:** Anti-FIFO защита")
    with col_db:
        db_ok = False
        if HAS_FIREBIRD:
            try:
                conn = firebirdsql.connect(**DB_CONFIG)
                conn.close()
                db_ok = True
            except Exception:
                db_ok = False
        if db_ok:
            st.success("🟢 Свързан към Монтевидео: `192.168.5.254:E:/db/MONTEVIDEO.FDB`")
        else:
            st.warning("🟡 Връзката към MONTEVIDEO.FDB ще се изпълни при анализ или локален режим.")

    st.markdown("---")

    # Избор на източник на XML файл
    source_type = st.radio(
        "Изберете начин за зареждане на банково извлечение:",
        ["📁 Избери от откритите извлечения (в проекта / Downloads)", "⬆️ Качи ново XML извлечение (.xml)"],
        horizontal=True,
        key="bank_source_type"
    )

    xml_target_path = None

    if "Избери от откритите" in source_type:
        candidates = find_candidate_bank_files()
        if not candidates:
            st.warning("Не са намерени автоматично файлове `report*.xml` в папката или Downloads.")
        else:
            options = {f"{os.path.basename(f)} ({'Downloads' if 'Downloads' in f else 'Проект'}, {datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%d.%m.%Y %H:%M')})": f for f in candidates[:10]}
            selected_label = st.selectbox("Изберете налично извлечение:", list(options.keys()), key="select_bank_cand")
            xml_target_path = options[selected_label]
    else:
        uploaded_xml = st.file_uploader("Качете XML извлечение от Банка ДСК:", type=["xml", "hml"], key="bank_xml_uploader")
        if uploaded_xml is not None:
            temp_path = os.path.join(root_dir, f"temp_{uploaded_xml.name}")
            with open(temp_path, "wb") as f:
                f.write(uploaded_xml.getbuffer())
            xml_target_path = temp_path

    if not xml_target_path or not os.path.exists(xml_target_path):
        st.info("👆 Моля, изберете или качете файл с банково извлечение, за да стартирате засичането.")
        return

    st.success(f"📄 Зареден файл за анализ: **{os.path.basename(xml_target_path)}**")

    # Бутон за стартиране на симулацията
    if st.button("🔍 Стартирай предварителен Dry-Run анализ", type="primary", key="btn_run_bank_analysis"):
        with st.spinner("Извличане на движения и съпоставка със салдата в Монтевидео..."):
            engine = MontevideoBankReconciliationEngine(xml_target_path, dry_run=True)
            try:
                count = engine.parse_xml()
            except Exception as e:
                st.error(f"Грешка при парсване на XML: {e}")
                return

            db_conn = None
            if HAS_FIREBIRD:
                try:
                    db_conn = firebirdsql.connect(**DB_CONFIG)
                except Exception as e:
                    st.warning(f"Няма директна връзка с базата Монтевидео ({e}). Анализът работи по вградени правила.")

            try:
                engine.classify_movements(db_conn)
            finally:
                if db_conn:
                    try: db_conn.close()
                    except Exception: pass

            st.session_state["bank_engine"] = engine
            st.session_state["bank_xml_path"] = xml_target_path

    # Показване на резултатите от анализа, ако вече са в сесията
    if "bank_engine" in st.session_state and st.session_state.get("bank_xml_path") == xml_target_path:
        engine = st.session_state["bank_engine"]
        stats = engine.stats

        st.markdown("### 📊 Обобщени резултати от банковото засичане:")
        
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Общо движения", stats.get("total_rows", 0))
        m2.metric("🟢 За закриване", stats.get("exact_match_ready_count", 0), f"{stats.get('exact_match_ready_sum', 0):.2f} €")
        m3.metric("🔵 Банкови такси", stats.get("bank_fees_count", 0), f"{stats.get('bank_fees_sum', 0):.2f} €")
        m4.metric("🟡 Вече закрити", stats.get("already_closed_count", 0), f"{stats.get('already_closed_sum', 0):.2f} €")
        m5.metric("🟣 Вътрешни", stats.get("internal_transfers_count", 0), f"{stats.get('internal_transfers_sum', 0):.2f} €")
        m6.metric("⚠️ Неразпознати", stats.get("missing_accrual_count", 0))

        # Експорт на официален Excel доклад
        file_base = os.path.splitext(os.path.basename(xml_target_path))[0].replace(' ', '')
        output_excel_name = f"ОДИТ_МОНТЕВИДЕО_ДСК_{file_base}.xlsx"
        output_excel_path = os.path.join(root_dir, output_excel_name)
        
        try:
            generated_excel = engine.export_excel_report(output_excel_path)
            with open(generated_excel, "rb") as ef:
                excel_bytes = ef.read()
            st.download_button(
                label=f"📥 Свали Одитен Доклад: {output_excel_name}",
                data=excel_bytes,
                file_name=output_excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_bank_audit_excel"
            )
        except Exception as e:
            st.warning(f"Не можа да се генерира Excel файл за сваляне: {e}")

        st.markdown("---")
        st.markdown("#### 📋 Детайлен списък с класифицирани банкови транзакции:")

        # Филтриране
        filter_status = st.selectbox(
            "Филтър по статус на движението:",
            ["Всички транзакции", "🟢 Готово за закриване", "🟡 Вече закрито", "🔵 Банкови такси", "🟣 Вътрешни / Задържани", "⚠️ Липсва начисление / Ръчни"],
            key="bank_filter_status"
        )

        rows = []
        for r in engine.classified_rows:
            st_text = r.get("status_badge", "")
            if filter_status == "🟢 Готово за закриване" and "ГОТОВО" not in st_text:
                continue
            if filter_status == "🟡 Вече закрито" and "ВЕЧЕ ЗАКРИТО" not in st_text:
                continue
            if filter_status == "🔵 Банкови такси" and "ТАКСИ" not in st_text:
                continue
            if filter_status == "🟣 Вътрешни / Задържани" and "ВЪТРЕШЕН" not in st_text and "ЗАДЪРЖАН" not in st_text:
                continue
            if filter_status == "⚠️ Липсва начисление / Ръчни" and "ЛИПСВА" not in st_text and "НЕЯСЕН" not in st_text:
                continue

            amt = r.get("amount", Decimal('0.00'))
            m_type = r.get("type", "")
            sign_amt = f"-{amt:.2f} €" if "Debit" in m_type else f"+{amt:.2f} €"

            rows.append({
                "Дата": r.get("date", ""),
                "Контрагент": r.get("name", ""),
                "Тип": m_type,
                "Сума (€)": sign_amt,
                "Статус": st_text,
                "Основание": r.get("reason", ""),
                "Партньор в ERP": r.get("matched_partner_name", "—"),
                "Действие": r.get("action_note", "")
            })

        st.dataframe(pd.DataFrame(rows), use_container_width=True)

        st.markdown("---")
        st.markdown("### ⚡ Реално отразяване в каса „ДСК Бакалия“ (Commit):")
        
        c_commit1, c_commit2 = st.columns([2, 1])
        with c_commit1:
            commit_confirmed = st.checkbox(
                "Потвърждавам, че съм прегледал одитния доклад и разрешавам запис на готовите движения в MONTEVIDEO.FDB.",
                key="confirm_bank_commit"
            )
        with c_commit2:
            if st.button("🚀 Изпълни запис в Монтевидео", type="primary", disabled=not commit_confirmed, key="btn_bank_commit"):
                if not HAS_FIREBIRD:
                    st.error("Липсва firebirdsql драйвер за директен запис.")
                else:
                    with st.spinner("Записване на транзакциите в каса ДСК Бакалия..."):
                        try:
                            live_conn = firebirdsql.connect(**DB_CONFIG)
                            success_msg = engine.execute_real_commit(live_conn)
                            live_conn.close()
                            st.success(f"✅ Успешно приключване: {success_msg}")
                            st.balloons()
                        except Exception as e:
                            st.error(f"❌ Грешка при реален запис: {e}")
