# -*- coding: utf-8 -*-
"""
================================================================================
Верига супермаркети „Бакалия“ — Централен хъб за промоционални брошури 2027
================================================================================
Интерфейс: Streamlit Web Hub
Модули:
  - brochure_mapping.py: Трислоен мапинг мениджър (доставчик -> Мистрал -> предпечат)
  - brochure_rules.py: 4-те търговски правила на „Бакалия“
"""

import os
import sys
import io
import datetime
import pandas as pd
import openpyxl
import streamlit as st

# Локални модули за напасване и одит
from brochure_mapping import BrochureMappingManager, PRODUCT_GROUPS
from brochure_rules import audit_brochure_item, detect_segment
from brochure_email_bot import BrochureEmailBot

st.set_page_config(
    page_title="Бакалия — Брошурен Мениджър 2027",
    page_icon="🛒",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1B365D 0%, #2563EB 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
    }
    .kpi-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
    }
    .kpi-val {
        font-size: 24px;
        font-weight: bold;
        color: #1B365D;
    }
    .kpi-lbl {
        font-size: 13px;
        color: #64748B;
    }
    .rule-card {
        background-color: #FFFFFF;
        border-left: 4px solid #2563EB;
        padding: 12px;
        border-radius: 4px;
        margin-bottom: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .group-pill {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: bold;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h2 style='margin:0; color:white;'>🛒 ВЕРИГА СУПЕРМАРКЕТИ „БАКАЛИЯ“</h2>
    <p style='margin:5px 0 0 0; opacity:0.9;'>Интелигентен хъб за промоционални брошури 2027 — Трислоен мапинг, Търговски одит и Автоматизиран предпечат</p>
</div>
""", unsafe_allow_html=True)

# Инициализиране на мапинг мениджъра в сесията
if "mapping_mgr" not in st.session_state:
    st.session_state["mapping_mgr"] = BrochureMappingManager()

mgr = st.session_state["mapping_mgr"]

# ------------------------------------------------------------------------------
# График за 2027 г.
# ------------------------------------------------------------------------------
schedule_data = [
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

brochure_options = [f"Брошура № {num:02d} ({s} до {e})" for num, s, e, d in schedule_data]

tabs = st.tabs([
    "🛡️ 1. Търговски одит & Правила",
    "🔍 2. Интелигентно напасване (Matcher)",
    "📥 3. Обработка на бланка от доставчик",
    "📅 4. График & Календар 2027",
    "📚 5. Речник на съответствията (Памет)",
    "📧 6. Имейл Център (Поща за брошури)"
])

# ------------------------------------------------------------------------------
# ТАБ 1: Търговски одит & 4-те правила на Бакалия
# ------------------------------------------------------------------------------
with tabs[0]:
    st.subheader("🛡️ Автоматичен контрол по 4-те задължителни търговски правила")
    st.write("Системата прилага стриктен алгоритмичен филтър, за да защити маржа на веригата и да предотврати канибализация на рафта.")

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.markdown("""
        <div class="rule-card" style="border-left-color: #2563EB;">
            <b>🚫 Правило 1: Забрана за идентични артикули в 1 брошура</b><br>
            <span style="font-size:12px; color:#475569;">
            Не се допуска дублиране на артикули от една и съща тясна подгрупа (напр. 2 бр. кисело мляко 2%, 2 бр. насипно краве сирене, 2 бр. тоалетна хартия 8бр, 2 бр. олио 1л). При наличие на второ предложение, системата препоръчва преместване в следващото издание.
            </span>
        </div>
        <div class="rule-card" style="border-left-color: #DC2626;">
            <b>⛔ Правило 2: Забрана за продажба на загуба (Zero Loss)</b><br>
            <span style="font-size:12px; color:#475569;">
            Промоционалната продажна цена без ДДС (ПЦ / 1.20) не може да бъде по-ниска или равна на Промо доставната цена (ЗКП). Артикули с марж &le; 0% се <b>блокират автоматично</b> от включване в предпечата!
            </span>
        </div>
        """, unsafe_allow_html=True)
    with col_r2:
        st.markdown("""
        <div class="rule-card" style="border-left-color: #D97706;">
            <b>⚠️ Правило 3: Защита „Промо ПЦ vs Базова ДЦ“</b><br>
            <span style="font-size:12px; color:#475569;">
            Промоционалната продажна цена с ДДС не трябва да пада под <b>редовната базова доставна цена с ДДС</b>. Така веригата се предпазва от загуба при продажба на остатъчни складови наличности след края на брошурата.
            </span>
        </div>
        <div class="rule-card" style="border-left-color: #7C3AED;">
            <b>🆕 Правило 4: Забрана за дебют на нови артикули</b><br>
            <span style="font-size:12px; color:#475569;">
            В брошура се допускат <b>само утвърдени артикули с история на продажбите</b> и доставки в обектите на Бакалия в последните 90 дни. Нови неизпитани артикули не стартират директно в промоционална брошура.
            </span>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.write("### 🧪 Симулатор на търговския одит (Тествайте на живо):")
    
    t_cols = st.columns([2, 1, 1, 1, 1, 1])
    with t_cols[0]:
        sim_name = st.text_input("Наименование на артикул:", value="Здравец кисело мляко 2% 500гр")
    with t_cols[1]:
        sim_reg_dc = st.number_input("Редовна ДЦ (€ с ДДС):", value=0.72, step=0.05, format="%.2f")
    with t_cols[2]:
        sim_promo_dc = st.number_input("Промо ДЦ (€ с ДДС):", value=0.58, step=0.05, format="%.2f")
    with t_cols[3]:
        sim_promo_pc = st.number_input("Промо ПЦ (€ с ДДС):", value=0.79, step=0.05, format="%.2f")
    with t_cols[4]:
        sim_ed = st.selectbox("Издание:", ["Брошура 1", "Брошура 2"])
    with t_cols[5]:
        sim_is_new = st.checkbox("Нов артикул?", value=False)

    if st.button("🔍 Изпълни одит по правилата", type="primary"):
        test_item = {
            "code": 99999 if sim_is_new else 5147,
            "name": sim_name,
            "reg_dc": sim_reg_dc,
            "promo_dc": sim_promo_dc,
            "promo_pc": sim_promo_pc,
            "edition": sim_ed
        }
        known_codes = {5147, 3930, 15044, 12118, 9159, 2140}
        current_in_b1 = [
            {"code": 5147, "name": "Верея Кисело мляко 2% 400гр", "edition": "Брошура 1"}
        ]
        
        audit_res = audit_brochure_item(test_item, current_in_b1, known_codes)
        
        st.write("#### Резултат от одита:")
        m_pct = audit_res["margin_pct"]
        if audit_res["status"] == "BLOCKED":
            st.error(f"⛔ БЛОКИРАН ЗА БРОШУРАТА! Марж: {m_pct:.1f}%")
        elif audit_res["status"] == "WARNING":
            st.warning(f"⚠️ ДОПУСНАТ С ПРЕДУПРЕЖДЕНИЕ! Марж: {m_pct:.1f}%")
        else:
            st.success(f"✅ НАПЪЛНО ОДОБРЕН! Марж: {m_pct:.1f}%")
            
        for f in audit_res["flags"]:
            if f["level"] == "BLOCKED":
                st.error(f["msg"])
            else:
                st.warning(f["msg"])

# ------------------------------------------------------------------------------
# ТАБ 2: Интелигентно напасване (3-Tier Matcher)
# ------------------------------------------------------------------------------
with tabs[1]:
    st.subheader("🔍 Трислоен Мапинг & Разпознаване на артикули")
    st.write("Свързва свободен текст от доставчика с вътрешните кодове в Мистрал (CS.FDB) и маркетинговото име за печат.")

    m_col1, m_col2, m_col3 = st.columns([1.5, 2.5, 1])
    with m_col1:
        supplier_input = st.selectbox(
            "Изберете доставчик:",
            ["Аксон", "Белла", "Амперел", "Оптимо", "Сириус", "Астра и Ко", "Фикосота", "ОМК", "Тандем", "Друг..."]
        )
        if supplier_input == "Друг...":
            supplier_input = st.text_input("Въведете име на фирмата:", value="Аксон")

    with m_col2:
        query_text = st.text_input(
            "Наименование на артикула (по доставчик):",
            value="Барила макарони видове 500гр"
        )

    with m_col3:
        query_barcode = st.text_input("Баркод (по желание):", value="")

    btn_match = st.button("🚀 Резолвирай артикул", type="primary")

    if btn_match or query_text:
        with st.spinner("Резолвиране през трислойния мапинг..."):
            res = mgr.resolve_supplier_item(supplier_input, query_text, barcode=query_barcode)

            status = res.get("status")
            conf = res.get("confidence", 0.0)
            gid = res.get("group_id", 6)
            g_info = PRODUCT_GROUPS.get(gid, PRODUCT_GROUPS[6])
            mkt_name = res.get("marketing_name", query_text)
            unit_val = res.get("unit", "бр.")
            codes = res.get("mistral_codes", [])
            names = res.get("mistral_names", [])

            if status == "EXACT_MAPPING":
                st.success(f"✅ НАПЪЛНО ПОТВЪРДЕНО СЪОТВЕТСТВИЕ ОТ РЕЧНИКА! (100% сигурност)")
            elif status in ["BARCODE_MAPPING", "BARCODE_MISTRAL"]:
                st.success(f"📦 РАЗПОЗНАТ ПО БАРКОД В МИСТРАЛ! (Код: {codes})")
            elif status == "FUZZY_PROPOSED":
                st.info(f"🔍 ПРЕДЛОЖЕНО СЪОТВЕТСТВИЕ (Сходство: {conf*100:.1f}%)")
            else:
                st.warning("⚠️ Не е открито съответствие. Моля изберете кодове ръчно.")

            # Показване на трите слоя
            st.markdown("### 🧩 Резултат по трите слоя:")
            c_lay1, c_lay2, c_lay3 = st.columns(3)

            with c_lay1:
                st.markdown(f"""
                <div style='background:#F1F5F9; padding:12px; border-radius:8px;'>
                    <b>Слой А: Подадено от доставчика</b><br>
                    • Доставчик: <b>{supplier_input}</b><br>
                    • Име: <b>{query_text}</b><br>
                    • Баркод: <b>{query_barcode or '—'}</b>
                </div>
                """, unsafe_allow_html=True)

            with c_lay2:
                codes_str = ", ".join(str(c) for c in codes) if codes else "Няма код"
                st.markdown(f"""
                <div style='background:#EFF6FF; padding:12px; border-radius:8px;'>
                    <b>Слой Б: Връзка към Мистрал (CS.FDB)</b><br>
                    • Кодове в ERP ({len(codes)} бр.): <b>{codes_str}</b><br>
                    • Редовна ПЦ: <b>{res.get('reg_pc', 0.0):.2f} €</b><br>
                    • Базова ДЦ: <b>{res.get('reg_dc', 0.0):.2f} €</b>
                </div>
                """, unsafe_allow_html=True)

            with c_lay3:
                st.markdown(f"""
                <div style='background:#F8FAFC; border:1px solid #CBD5E1; padding:12px; border-radius:8px;'>
                    <b>Слой В: Маркетингов предпечат</b><br>
                    • Маркетингово име: <b>{mkt_name}</b><br>
                    • Мярка / Разфасовка: <b>{unit_val}</b><br>
                    • Група: <span style='background-color:{g_info["color"]}; color:black; padding:2px 8px; border-radius:4px; font-weight:bold;'>{gid} - {g_info["name"]}</span>
                </div>
                """, unsafe_allow_html=True)

            # Бутон за потвърждение с 1 клик
            st.markdown("---")
            col_save1, col_save2 = st.columns([2, 1])
            with col_save1:
                st.write("Желаете ли да запазите това съответствие в паметта за всички бъдещи брошури?")
            with col_save2:
                if st.button("💾 Потвърди и запиши в речника", type="primary"):
                    mgr.confirm_mapping(
                        supplier=supplier_input,
                        raw_name=query_text,
                        mistral_codes=codes,
                        mistral_names=names,
                        marketing_name=mkt_name,
                        unit=unit_val,
                        group_id=gid,
                        barcode=query_barcode,
                        reg_pc=res.get('reg_pc'),
                        reg_dc=res.get('reg_dc')
                    )
                    st.toast("✅ Съответствието е успешно записано в brochure_item_mapping.json!")

            if res.get("all_candidates") and len(res["all_candidates"]) > 1:
                with st.expander("Виж всички близки кандидати от Мистрал"):
                    st.dataframe(pd.DataFrame(res["all_candidates"])[["code", "mistral_name", "score", "reg_dc", "reg_pc"]], width='stretch')

# ------------------------------------------------------------------------------
# ТАБ 3: Обработка на бланка от доставчик с вграден ОДИТ и ЕКСПОРТ
# ------------------------------------------------------------------------------
with tabs[2]:
    st.subheader("📥 Списък с набраните артикули, Търговски Одит и Експорт")
    
    tab3_mode = st.radio(
        "Изберете режим на работа:",
        ["🎯 Сборен брошурен пул (ВСИЧКИ набрани 145 артикула от пощата)", "📄 Индивидуална бланка на доставчик"],
        horizontal=True
    )

    # ==========================================================================
    # РЕЖИМ 1: СБОРЕН БРОШУРЕН ПУЛ (КОЛЕКТОР НА ВСИЧКИ НАБРАНИ АРТИКУЛИ)
    # ==========================================================================
    if tab3_mode.startswith("🎯"):
        st.markdown("### 📚 Централен колектор на набраните оферти за брошурата")
        st.caption("Обединява всички получени бланки от доставчици в единен сборен списък за следене на запълването (цел: 201 позиции) и одит за канибализация между партньори.")

        pool_file = os.path.join(os.path.dirname(__file__), "brochure_master_pool.json")
        master_items = []
        if os.path.exists(pool_file):
            try:
                with open(pool_file, "r", encoding="utf-8") as f_pool:
                    master_items = json.load(f_pool)
            except Exception:
                master_items = []

        if not master_items:
            st.warning("Все още няма набрани артикули в сборния пул. Използвайте бутона по-долу за автоматично събиране на всички свалени бланки от пощата.")
            if st.button("🔄 Сканирай и събери всички входящи оферти сега"):
                inc_dir = os.path.join(os.path.dirname(__file__), "incoming_brochures")
                files = [os.path.join(inc_dir, f) for f in os.listdir(inc_dir) if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('temp_')]
                master_items = []
                for f in sorted(files):
                    parsed, err = mgr.parse_any_supplier_blank(f)
                    if parsed and parsed.get("items"):
                        fn = os.path.basename(f)
                        for it in parsed["items"]:
                            it["source_file"] = fn
                            it["supplier"] = parsed.get("supplier", "Доставчик")
                            master_items.append(it)
                with open(pool_file, "w", encoding="utf-8") as f_pool:
                    json.dump(master_items, f_pool, ensure_ascii=False, indent=2)
                st.rerun()
        else:
            # Извличане на активни кодове за правило 4 през бързия кеш
            all_known_mats = mgr._load_all_materials()
            active_codes = set(m["code"] for m in all_known_mats)

            # Филтри
            f_col1, f_col2, f_col3 = st.columns(3)
            with f_col1:
                all_supps = sorted(list(set(it.get("supplier", "Доставчик") for it in master_items)))
                sel_supp = st.selectbox("Филтър по доставчик:", ["Всички доставчици"] + all_supps)
            with f_col2:
                all_grps = sorted(list(set(f"{it.get('group_id', 6)} - {it.get('group_name', 'масови')}" for it in master_items)))
                sel_grp = st.selectbox("Филтър по стокова категория:", ["Всички категории"] + all_grps)
            with f_col3:
                sel_status = st.selectbox("Филтър по одит статус:", ["Всички статуси", "✅ Одобрени", "⚠️ Внимание / Предупреждение", "⛔ Блокирани"])

            # Прилагане на филтри
            display_items = []
            for itm in master_items:
                if sel_supp != "Всички доставчици" and itm.get("supplier") != sel_supp:
                    continue
                if sel_grp != "Всички категории" and f"{itm.get('group_id', 6)} - {itm.get('group_name', 'масови')}" != sel_grp:
                    continue
                display_items.append(itm)

            # Одит по правилата за всички елементи
            audit_table = []
            approved_pool = []
            for itm in display_items:
                audit_in = {
                    "code": itm["mistral_codes"][0] if itm.get("mistral_codes") else 0,
                    "name": itm["marketing_name"],
                    "reg_dc": itm.get("reg_dc", 0.0),
                    "promo_dc": itm.get("promo_dc", 0.0),
                    "promo_pc": itm.get("promo_pc", 0.0),
                    "edition": "Брошура 21"
                }
                a_res = audit_brochure_item(audit_in, master_items, active_codes)
                itm["audit_status"] = a_res["status"]
                itm["margin_pct"] = a_res["margin_pct"]
                
                if a_res["status"] != "BLOCKED":
                    approved_pool.append(itm)

                status_label = "⛔ БЛОКИРАН" if a_res["status"] == "BLOCKED" else "⚠️ ВНИМАНИЕ" if a_res["status"] == "WARNING" else "✅ ОДОБРЕН"
                if sel_status != "Всички статуси" and status_label != sel_status:
                    continue

                flag_msgs = [f["msg"] for f in a_res["flags"]]
                audit_table.append({
                    "Група": f"{itm.get('group_id', 6)} - {itm.get('group_name', 'масови')}",
                    "Доставчик": itm.get("supplier", "Доставчик"),
                    "Артикул (Маркетинг)": itm.get("marketing_name", itm.get("raw_name")),
                    "Код в Мистрал": ", ".join(str(c) for c in itm.get("mistral_codes", [])) if itm.get("mistral_codes") else "Липсва",
                    "Мярка": itm.get("unit", "бр."),
                    "Промо ДЦ (€)": f"{itm.get('promo_dc', 0.0):.2f}",
                    "Промо ПЦ (€)": f"{itm.get('promo_pc', 0.0):.2f}",
                    "Марж %": f"{a_res['margin_pct']:.1f}%",
                    "Одит статус": status_label,
                    "Източник (Файл)": itm.get("source_file", ""),
                    "Бележки": " | ".join(flag_msgs) if flag_msgs else "ОК"
                })

            # Метрики за цялата брошура
            st.markdown("#### 📊 Статус на запълване на Брошурата (Лимит: 201 позиции):")
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                total_cnt = len(master_items)
                fill_pct = (total_cnt / 201) * 100
                st.metric("Запълнени позиции", f"{total_cnt} / 201", f"{fill_pct:.1f}% готовност")
            with m2:
                supp_cnt = len(set(it.get("supplier") for it in master_items))
                st.metric("Доставчици с оферти", f"{supp_cnt} фирми")
            with m3:
                appr_cnt = len(approved_pool)
                st.metric("Одобрени за печат", f"{appr_cnt} бр.")
            with m4:
                blocked_cnt = total_cnt - appr_cnt
                st.metric("Блокирани по правилата", f"{blocked_cnt} бр.", delta=-blocked_cnt if blocked_cnt else 0, delta_color="inverse")

            # Разпределение по категории
            st.markdown("##### 📦 Разпределение по 8-те цветни продуктови групи:")
            grp_counts = {}
            for itm in master_items:
                g_title = itm.get("group_name", "масови")
                grp_counts[g_title] = grp_counts.get(g_title, 0) + 1
            
            grp_badges = " | ".join([f"**{g.capitalize()}**: {c} бр." for g, c in sorted(grp_counts.items(), key=lambda x: -x[1])])
            st.info(f"📋 {grp_badges}")

            # Таблица
            st.write(f"#### 📋 Списък с набраните позиции ({len(audit_table)} бр.):")
            st.dataframe(pd.DataFrame(audit_table), width='stretch')

            # Бутони за СБОРЕН ЕКСПОРТ на цялата брошура
            st.markdown("---")
            st.subheader("📦 Генериране на финални експорти за ЦЯЛАТА БРОШУРА")
            st.write("Сваля се сборен файл за всички одобрени артикули от всички доставчици, номерирани от 1 до 201.")

            e_col1, e_col2, e_col3 = st.columns(3)
            with e_col1:
                master_prep_path = os.path.join(os.path.dirname(__file__), "Предпечат_Сборен_Брошура.xlsx")
                mgr.export_prepress_excel(approved_pool, master_prep_path, brochure_title="Сборна Брошура 21")
                with open(master_prep_path, "rb") as f_mp:
                    st.download_button(
                        label="📄 Изтегли СБОРЕН ПРЕДПЕЧАТ (2 листа .xlsx)",
                        data=f_mp.read(),
                        file_name="Предпечат_Сборен_Брошура.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                st.caption("Sheet2 с 8-те цветни категории + Sheet1 решетка (1..201) за всички партньори")

            with e_col2:
                today = datetime.date.today()
                pc_s = today.strftime("%Y-%m-%d 00:00:00")
                pc_e = (today + datetime.timedelta(days=9)).strftime("%Y-%m-%d 23:59:59")
                zkp_s = (today - datetime.timedelta(days=7)).strftime("%Y-%m-%d 00:00:00")
                zkp_e = pc_e
                
                out_dir = os.path.join(os.path.dirname(__file__), "exports")
                pc_f, zkp_f = mgr.export_mistral_import_files(approved_pool, pc_s, pc_e, zkp_s, zkp_e, out_dir)

                with open(pc_f, "rb") as f_pc:
                    st.download_button(
                        label="⚙️ Изтегли СБОРЕН Мистрал ПЦ (.xlsx)",
                        data=f_pc.read(),
                        file_name="ПЦ_Сборен_Импорт_Мистрал.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.caption("Промо продажни цени с ДДС за всички одобрени артикули")

            with e_col3:
                with open(zkp_f, "rb") as f_zkp:
                    st.download_button(
                        label="⚙️ Изтегли СБОРЕН Мистрал ЗКП (.xlsx)",
                        data=f_zkp.read(),
                        file_name="ЗКП_Сборен_Импорт_Мистрал.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.caption("Промо ДЦ без ДДС за AGREEDDELIVERYPRICE (X-7 дни)")

    # ==========================================================================
    # РЕЖИМ 2: ИНДИВИДУАЛНА БЛАНКА НА ДОСТАВЧИК
    # ==========================================================================
    else:
        st.markdown("### 📄 Преглед и одит на единична бланка от доставчик")
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            uploaded_file = st.file_uploader("📤 Качете Excel бланка от компютъра (.xlsx, .xls)", type=["xlsx", "xls"])
        with col_up2:
            inc_dir = os.path.join(os.path.dirname(__file__), "incoming_brochures")
            inc_files = [f for f in os.listdir(inc_dir) if f.lower().endswith(('.xls', '.xlsx'))] if os.path.exists(inc_dir) else []
            selected_inc = st.selectbox("📬 Или изберете свалена оферта от пощата:", ["-- Изберете файл от пощата --"] + inc_files)

        target_file_path = None
        if uploaded_file is not None:
            temp_path = os.path.join(inc_dir, f"temp_upload_{uploaded_file.name}")
            with open(temp_path, "wb") as f_tmp:
                f_tmp.write(uploaded_file.getbuffer())
            target_file_path = temp_path
        elif selected_inc and selected_inc != "-- Изберете файл от пощата --":
            target_file_path = os.path.join(inc_dir, selected_inc)

        if target_file_path and os.path.exists(target_file_path):
            parsed_res, p_err = mgr.parse_any_supplier_blank(target_file_path)
            if p_err:
                st.error(f"❌ Грешка при разчитане на бланката: {p_err}")
            elif parsed_res:
                fmt_badge = "🟢 Нов шаблон 2027" if parsed_res.get("format") == "2027_TEMPLATE" else "🟡 Класическа (Legacy) бланка"
                st.markdown(f"**Разпознат формат:** `{fmt_badge}` | **Файл:** `{os.path.basename(target_file_path)}`")
                
                c_meta1, c_meta2 = st.columns(2)
                with c_meta1:
                    b_name = st.text_input("Брошура / Издание", value=parsed_res.get("brochure") or "Брошура 21")
                with c_meta2:
                    supplier_val = st.text_input("Доставчик / Контрагент", value=parsed_res.get("supplier") or "Доставчик")

                raw_items = parsed_res.get("items", [])
                if raw_items:
                    all_known_mats = mgr._load_all_materials()
                    active_codes = set(m["code"] for m in all_known_mats)

                    parsed_items = []
                    for it in raw_items:
                        p_dc = float(it.get("promo_dc", 0.0) or 0.0)
                        r_dc = float(it.get("reg_dc", 0.0) or 0.0)
                        rec_p = float(it.get("promo_pc", 0.0) or 0.0)
                        r_name = str(it.get("raw_name", "")).strip()
                        bc = str(it.get("barcode", "")).strip()

                        res_m = mgr.resolve_supplier_item(
                            supplier_val, r_name,
                            barcode=bc, unit=it.get("unit"),
                            promo_dc=p_dc, reg_dc=r_dc, rec_promo_pc=rec_p
                        )

                        calc_promo_pc = rec_p or (res_m.get("reg_pc", 0.0) * 0.85 if res_m.get("reg_pc") else (p_dc * 1.25))

                        parsed_items.append({
                            "raw_name": r_name,
                            "marketing_name": res_m.get("marketing_name") or r_name,
                            "supplier": supplier_val,
                            "barcode": bc,
                            "unit": res_m.get("unit") or it.get("unit") or "бр.",
                            "reg_dc": r_dc or res_m.get("reg_dc", 0.0),
                            "promo_dc": p_dc,
                            "promo_pc": round(calc_promo_pc, 2),
                            "reg_pc": res_m.get("reg_pc", 0.0),
                            "mistral_codes": res_m.get("mistral_codes", []),
                            "group_id": res_m.get("group_id", 6),
                            "group_name": res_m.get("group_name", "масови"),
                            "mapping_status": res_m.get("status"),
                            "confidence": res_m.get("confidence", 0.0),
                            "note": it.get("note", "")
                        })

                    audit_table = []
                    approved_for_export = []
                    unconfirmed_items = []

                    for itm in parsed_items:
                        audit_in = {
                            "code": itm["mistral_codes"][0] if itm["mistral_codes"] else 0,
                            "name": itm["marketing_name"],
                            "reg_dc": itm["reg_dc"],
                            "promo_dc": itm["promo_dc"],
                            "promo_pc": itm["promo_pc"],
                            "edition": b_name
                        }
                        a_res = audit_brochure_item(audit_in, parsed_items, active_codes)
                        itm["audit_status"] = a_res["status"]
                        itm["margin_pct"] = a_res["margin_pct"]
                        itm["audit_flags"] = a_res["flags"]

                        if a_res["status"] != "BLOCKED":
                            approved_for_export.append(itm)

                        if itm["mapping_status"] != "EXACT_MAPPING" and itm["mistral_codes"]:
                            unconfirmed_items.append(itm)

                        flag_msgs = [f["msg"] for f in a_res["flags"]]
                        audit_table.append({
                            "Група": f"{itm['group_id']} - {itm['group_name']}",
                            "Артикул (Маркетинг)": itm["marketing_name"],
                            "Код в Мистрал": ", ".join(str(c) for c in itm["mistral_codes"]) if itm["mistral_codes"] else "Липсва",
                            "Мярка": itm["unit"],
                            "Промо ДЦ (€)": f"{itm['promo_dc']:.2f}",
                            "Промо ПЦ (€)": f"{itm['promo_pc']:.2f}",
                            "Марж %": f"{a_res['margin_pct']:.1f}%",
                            "Одит правила": "⛔ БЛОКИРАН" if a_res["status"] == "BLOCKED" else "⚠️ ВНИМАНИЕ" if a_res["status"] == "WARNING" else "✅ ОДОБРЕН",
                            "Мапинг статус": "✅ Запомнен" if itm["mapping_status"] == "EXACT_MAPPING" else f"🔍 Предложен ({int(itm['confidence']*100)}%)",
                            "Бележки": " | ".join(flag_msgs) if flag_msgs else "ОК"
                        })

                    st.markdown("### 📊 Обобщение за тази бланка:")
                    k_c1, k_c2, k_c3, k_c4 = st.columns(4)
                    with k_c1:
                        st.metric("Общо артикули", len(parsed_items))
                    with k_c2:
                        confirmed_cnt = sum(1 for it in parsed_items if it["mapping_status"] == "EXACT_MAPPING")
                        st.metric("Потвърден мапинг", f"{confirmed_cnt} бр.")
                    with k_c3:
                        st.metric("Одобрени за експорт", f"{len(approved_for_export)} бр.")
                    with k_c4:
                        blocked_cnt = len(parsed_items) - len(approved_for_export)
                        st.metric("Блокирани по правилата", f"{blocked_cnt} бр.", delta=-blocked_cnt if blocked_cnt else 0, delta_color="inverse")

                    if unconfirmed_items:
                        st.info(f"💡 Открити са **{len(unconfirmed_items)}** нови съответствия с Мистрал.")
                        if st.button(f"💾 Потвърди и запомни всички {len(unconfirmed_items)} съответствия в Речника", type="primary"):
                            for u_it in unconfirmed_items:
                                mgr.confirm_mapping(
                                    supplier=supplier_val,
                                    raw_name=u_it["raw_name"],
                                    mistral_codes=u_it["mistral_codes"],
                                    marketing_name=u_it["marketing_name"],
                                    unit=u_it["unit"],
                                    group_id=u_it["group_id"],
                                    barcode=u_it.get("barcode"),
                                    reg_pc=u_it.get("reg_pc", 0.0),
                                    reg_dc=u_it.get("reg_dc", 0.0)
                                )
                            st.success("✅ Всички съответствия бяха успешно записани в `brochure_item_mapping.json`!")
                            st.rerun()

                    st.write("#### Подробен отчет по артикули:")
                    st.dataframe(pd.DataFrame(audit_table), width='stretch')

                    # Секция за експорти за този доставчик
                    st.markdown("---")
                    st.subheader("📦 Генериране на експорти за тази бланка")
                    
                    e_col1, e_col2, e_col3 = st.columns(3)
                    with e_col1:
                        prep_path = os.path.join(os.path.dirname(__file__), f"Предпечат_{b_name.replace(' ', '_')}_{supplier_val.replace(' ', '_')}.xlsx")
                        mgr.export_prepress_excel(approved_for_export, prep_path, brochure_title=b_name)
                        with open(prep_path, "rb") as f:
                            st.download_button(
                                label="📄 Изтегли за ПРЕДПЕЧАТ (2 листа .xlsx)",
                                data=f.read(),
                                file_name=f"Предпечат_{b_name.replace(' ', '_')}_{supplier_val.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                        st.caption("Sheet2 с 8-те цветни категории + Sheet1 решетка")

                    with e_col2:
                        today = datetime.date.today()
                        pc_s = today.strftime("%Y-%m-%d 00:00:00")
                        pc_e = (today + datetime.timedelta(days=9)).strftime("%Y-%m-%d 23:59:59")
                        zkp_s = (today - datetime.timedelta(days=7)).strftime("%Y-%m-%d 00:00:00")
                        zkp_e = pc_e
                        
                        out_dir = os.path.join(os.path.dirname(__file__), "exports")
                        pc_f, zkp_f = mgr.export_mistral_import_files(approved_for_export, pc_s, pc_e, zkp_s, zkp_e, out_dir)

                        with open(pc_f, "rb") as f_pc:
                            st.download_button(
                                label="⚙️ Изтегли Мистрал ПЦ (.xlsx)",
                                data=f_pc.read(),
                                file_name=f"ПЦ_Импорт_{supplier_val.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        st.caption("Промо продажни цени с ДДС")

                    with e_col3:
                        with open(zkp_f, "rb") as f_zkp:
                            st.download_button(
                                label="⚙️ Изтегли Мистрал ЗКП (.xlsx)",
                                data=f_zkp.read(),
                                file_name=f"ЗКП_Импорт_{supplier_val.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        st.caption("Промо ДЦ без ДДС за AGREEDDELIVERYPRICE")

# ------------------------------------------------------------------------------
# ТАБ 4: График 2027
# ------------------------------------------------------------------------------
with tabs[3]:
    st.subheader("📅 Официален график за 25-те промоционални брошури през 2027 г.")
    
    selected_b = st.selectbox("Изберете брошура за преглед на срокове:", brochure_options)
    b_idx = brochure_options.index(selected_b)
    num, s_str, e_str, days = schedule_data[b_idx]
    
    s_dt = datetime.date.fromisoformat(s_str)
    e_dt = datetime.date.fromisoformat(e_str)
    zkp_start = s_dt - datetime.timedelta(days=7)
    prep_deadline = s_dt - datetime.timedelta(days=10)
    reminder_date = s_dt - datetime.timedelta(days=12)

    k_cols = st.columns(4)
    with k_cols[0]:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-lbl">ПЕРИОД В МАГАЗИНИТЕ</div>
            <div class="kpi-val">{s_dt.strftime('%d.%m')} - {e_dt.strftime('%d.%m')}</div>
            <div style='font-size:11px; color:#64748B;'>{days} дни промоция</div>
        </div>
        """, unsafe_allow_html=True)
    with k_cols[1]:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-lbl">СТАРТ НАМАЛЕНА ДЦ (ЗКП)</div>
            <div class="kpi-val" style="color:#D97706;">{zkp_start.strftime('%d.%m.%Y')}</div>
            <div style='font-size:11px; color:#D97706;'>Точно 7 дни преди старта!</div>
        </div>
        """, unsafe_allow_html=True)
    with k_cols[2]:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-lbl">ПРЕДАВАНЕ ПРЕДПЕЧАТ</div>
            <div class="kpi-val" style="color:#DC2626;">{prep_deadline.strftime('%d.%m.%Y')}</div>
            <div style='font-size:11px; color:#64748B;'>Краен срок за финалния ексел</div>
        </div>
        """, unsafe_allow_html=True)
    with k_cols[3]:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-lbl">НАПОМНЯНЕ / ПРОЗВЪНЯВАНЕ</div>
            <div class="kpi-val" style="color:#2563EB;">{reminder_date.strftime('%d.%m.%Y')}</div>
            <div style='font-size:11px; color:#64748B;'>2 дни преди предпечата</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    sched_df = []
    for n, s, e, d in schedule_data:
        s_d = datetime.date.fromisoformat(s)
        e_d = datetime.date.fromisoformat(e)
        sched_df.append({
            "Брошура №": f"№ {n:02d}",
            "Период магазини": f"{s_d.strftime('%d.%m')} – {e_d.strftime('%d.%m.%Y')}",
            "Дни": d,
            "Старт доставна цена (X-7)": (s_d - datetime.timedelta(days=7)).strftime('%d.%m.%Y'),
            "Краен срок предпечат (X-10)": (s_d - datetime.timedelta(days=10)).strftime('%d.%m.%Y'),
            "Срок за предложения (X-12)": (s_d - datetime.timedelta(days=12)).strftime('%d.%m.%Y'),
            "Покана/Бланка (X-21)": (s_d - datetime.timedelta(days=21)).strftime('%d.%m.%Y'),
        })
    st.dataframe(pd.DataFrame(sched_df), width='stretch', height=350)

# ------------------------------------------------------------------------------
# ТАБ 5: Речник на съответствията (Памет на трислойния мапинг)
# ------------------------------------------------------------------------------
with tabs[4]:
    st.subheader("📚 Запаметени съответствия в локалния речник (brochure_item_mapping.json)")
    st.write("Списък на всички потвърдени съответствия между суровите имена от доставчиците, вътрешните кодове в Мистрал и предпечатните категории.")

    mappings_dict = mgr.data.get("mappings", {})
    st.metric("Общо запаметени артикули в базата", f"{len(mappings_dict)} бр.")

    if mappings_dict:
        m_list = []
        for k, v in mappings_dict.items():
            codes_val = v.get("mistral_codes", [])
            m_list.append({
                "Доставчик": v.get("supplier", ""),
                "Име по доставчик": v.get("raw_name", ""),
                "Баркод": v.get("barcode", ""),
                "Кодове в Мистрал": ", ".join(str(c) for c in codes_val) if codes_val else "—",
                "Маркетингово име (Предпечат)": v.get("marketing_name", ""),
                "Мярка": v.get("unit", ""),
                "Група": f"{v.get('group_id', 6)} - {v.get('group_name', 'масови')}",
                "Редовна ПЦ (€)": f"{v.get('reg_pc', 0.0):.2f}",
                "Обновено": str(v.get("updated_at", ""))[:19]
            })

        df_map = pd.DataFrame(m_list)

        filter_col1, filter_col2 = st.columns(2)
        with filter_col1:
            search_q = st.text_input("🔍 Търсене по име или доставчик:", value="")
        with filter_col2:
            group_filter = st.selectbox("Филтър по продуктова група:", ["Всички"] + [f"{gid} - {PRODUCT_GROUPS[gid]['name']}" for gid in sorted(PRODUCT_GROUPS.keys())])

        filtered_df = df_map
        if search_q:
            mask = filtered_df["Име по доставчик"].str.contains(search_q, case=False, na=False) | filtered_df["Доставчик"].str.contains(search_q, case=False, na=False) | filtered_df["Маркетингово име (Предпечат)"].str.contains(search_q, case=False, na=False)
            filtered_df = filtered_df[mask]

        if group_filter != "Всички":
            g_prefix = group_filter.split(" - ")[0]
            filtered_df = filtered_df[filtered_df["Група"].str.startswith(g_prefix)]

        st.dataframe(filtered_df, width='stretch', height=450)
    else:
        st.info("Речникът на съответствията все още е празен.")

# ------------------------------------------------------------------------------
# ТАБ 6: Имейл Център (Специализирана поща за промоционални брошури)
# ------------------------------------------------------------------------------
with tabs[5]:
    st.subheader("📧 Специализиран Имейл Център за промоционални брошури")
    st.write("Автоматизирано разпращане на покани, прихващане на попълнени оферти и контрол на закъснелите доставчици.")

    if "email_bot" not in st.session_state:
        st.session_state["email_bot"] = BrochureEmailBot()
    bot = st.session_state["email_bot"]

    m_subtabs = st.tabs([
        "📤 1. Разпращане на покани (X-21 дни)",
        "📥 2. Входящи оферти (Harvester)",
        "📞 3. Списък за прозвъняване (X-12 дни)",
        "👥 4. Доставчици & Контакти",
        "⚙️ 5. Настройки на пощата"
    ])

    # --------------------------------------------------------------------------
    # SUBTAB 1: Разпращане на покани
    # --------------------------------------------------------------------------
    with m_subtabs[0]:
        st.markdown("### 📤 Разпращане на покани към доставчиците")
        st.warning("🔒 **Автоматичното изпращане на реални имейли е СПРЯНО по инструкция.** Системата работи в защитен симулационен режим. Реални писма НЯМА да бъдат изпращани без ваше изрично потвърждение.")
        st.write("Изберете кампания за симулация на разпращането на поканите с прикрепена бланка.")

        col_inv1, col_inv2 = st.columns([2, 1])
        with col_inv1:
            sel_brochure = st.selectbox("Изберете брошура за покани:", brochure_options, key="inv_b_sel")
            b_i = brochure_options.index(sel_brochure)
            b_num, s_s, e_s, b_d = schedule_data[b_i]
            s_d_obj = datetime.date.fromisoformat(s_s)
            e_d_obj = datetime.date.fromisoformat(e_s)
            zkp_s_obj = s_d_obj - datetime.timedelta(days=7)
            sub_d_obj = s_d_obj - datetime.timedelta(days=12)

            st.info(f"**Брошура № {b_num:02d}** | Период в обектите: **{s_d_obj.strftime('%d.%m')} - {e_d_obj.strftime('%d.%m.%Y')}** | ЗКП старт: **{zkp_s_obj.strftime('%d.%m.%Y')}** | Краен срок оферти: **{sub_d_obj.strftime('%d.%m.%Y')}**")

        with col_inv2:
            test_mode = True
            st.checkbox("🔒 Защитен режим (Само симулация)", value=True, disabled=True, help="Реалното изпращане е заключено по инструкция на потребителя.")
            custom_test_email = st.text_input("Изпрати тест до имейл (опция):", value="", placeholder="test@bakaliya.bg", disabled=True)

        contacts_df = bot.load_contacts()
        active_contacts = contacts_df[contacts_df["Статус"] == "Активен"] if not contacts_df.empty and "Статус" in contacts_df.columns else contacts_df

        st.write(f"Готови за изпращане: **{len(active_contacts)} бр. активни доставчици** от регистъра.")

        if st.button("🚀 Стартирай разпращането на поканите", type="primary"):
            with st.spinner("Генериране и изпращане на покани..."):
                success_cnt = 0
                fail_cnt = 0
                log_reports = []

                p_bar = st.progress(0)
                total_c = len(active_contacts)

                for idx, (_, row) in enumerate(active_contacts.iterrows()):
                    f_name = row["Фирма / Партньор"]
                    t_mail = custom_test_email if custom_test_email else row.get("Имейл за оферти", "")
                    
                    ok, msg_out = bot.send_invitation_email(
                        to_email=t_mail,
                        supplier_name=f_name,
                        brochure_code=f"Брошура № {b_num:02d}",
                        store_period=f"{s_d_obj.strftime('%d.%m')} – {e_d_obj.strftime('%d.%m.%Y')}",
                        zkp_period=f"{zkp_s_obj.strftime('%d.%m')} – {e_d_obj.strftime('%d.%m.%Y')}",
                        deadline_date=sub_d_obj.strftime('%d.%m.%Y'),
                        test_mode=test_mode
                    )
                    if ok:
                        success_cnt += 1
                    else:
                        fail_cnt += 1
                    log_reports.append({"Доставчик": f_name, "Имейл": t_mail, "Резултат": msg_out})
                    p_bar.progress((idx + 1) / total_c)

                if test_mode:
                    st.success(f"✅ Симулацията завърши успешно! Подготвени за изпращане: {success_cnt} покани.")
                else:
                    st.success(f"✅ Успешно изпратени: {success_cnt} покани! Грешки: {fail_cnt}.")

                st.dataframe(pd.DataFrame(log_reports), width='stretch')

    # --------------------------------------------------------------------------
    # SUBTAB 2: Входящи оферти (Harvester)
    # --------------------------------------------------------------------------
    with m_subtabs[1]:
        st.markdown("### 📥 Автоматично прихващане на върнати бланки")
        st.write("Роботът проверява новата пощенска кутия, сваля прикачените файлове и ги анализира.")

        col_h1, col_h2 = st.columns([2, 1])
        with col_h1:
            if st.button("🔄 Провери пощата за нови оферти сега", type="primary"):
                with st.spinner("Свързване с пощенската кутия и търсене на входящи оферти..."):
                    downloaded = bot.fetch_incoming_brochures(test_mode=False if bot.config.get("app_password") else True)
                    if downloaded:
                        st.success(f"Намерени и свалени {len(downloaded)} файла с оферти!")
                        st.dataframe(pd.DataFrame(downloaded), width='stretch')
                    else:
                        st.info("Няма нови нечетени оферти в пощенската кутия.")

        with col_h2:
            st.caption("Папка за входящи файлове: `incoming_brochures/`")

        # Преглед на свалени файлове в incoming_brochures
        local_files = bot._scan_local_incoming_folder()
        if local_files:
            st.markdown("#### 📁 Свалени входящи оферти за обработка:")
            for f_info in local_files:
                f_path = f_info["file_path"]
                f_name = f_info["file_name"]
                with st.expander(f"📄 {f_name} (Получен: {f_info['received_at']})"):
                    parsed, err = bot.process_supplier_blank_file(f_path)
                    if parsed:
                        st.write(f"**Брошура:** {parsed['brochure']} | **Доставчик:** {parsed['supplier']} | **Артикули:** {parsed['items_count']} бр.")
                        st.dataframe(pd.DataFrame(parsed["items"])[["marketing_name", "unit", "reg_dc", "promo_dc", "promo_pc", "group_name", "mapping_status"]], width='stretch')
                    else:
                        st.error(err)

    # --------------------------------------------------------------------------
    # SUBTAB 3: Списък за прозвъняване (X-12 дни)
    # --------------------------------------------------------------------------
    with m_subtabs[2]:
        st.markdown("### 📞 Списък за прозвъняване на закъснели доставчици")
        st.write("Генерира се 2 дни преди предаване за предпечат за доставчиците, които все още нямат върната оферта.")

        contacts_df = bot.load_contacts()
        if not contacts_df.empty:
            # Намиране на върналите оферти
            submitted_suppliers = set()
            for f_info in bot._scan_local_incoming_folder():
                p_data, _ = bot.process_supplier_blank_file(f_info["file_path"])
                if p_data and p_data.get("supplier"):
                    submitted_suppliers.add(p_data["supplier"])

            # Закъснели
            late_df = contacts_df[~contacts_df["Фирма / Партньор"].isin(submitted_suppliers)]

            k1, k2, k3 = st.columns(3)
            with k1:
                st.metric("Всички поканени", f"{len(contacts_df)} бр.")
            with k2:
                st.metric("Подали оферта", f"{len(submitted_suppliers)} бр.")
            with k3:
                st.metric("Закъснели (за прозвъняване)", f"{len(late_df)} бр.", delta=-len(late_df), delta_color="inverse")

            st.write("#### 📋 Списък за спешно прозвъняване:")
            st.dataframe(late_df[["Фирма / Партньор", "Лице за контакт", "Телефон за връзка", "Имейл за оферти", "Основни брандове / Артикули"]], width='stretch')

            # Сваляне в Excel
            call_out = io.BytesIO()
            late_df.to_excel(call_out, index=False)
            call_out.seek(0)
            st.download_button(
                label="📥 Изтегли списък за прозвъняване (.xlsx)",
                data=call_out,
                file_name="Списък_Прозвъняване_Брошура.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    # --------------------------------------------------------------------------
    # SUBTAB 4: Доставчици & Контакти
    # --------------------------------------------------------------------------
    with m_subtabs[3]:
        st.markdown("### 👥 Регистър с контакти на 133-те доставчици за брошури")
        st.write("Извлечени от реалните участия на веригата. Можете да редактирате имейлите и телефоните директно оттук или през файла `suppliers_brochure_contacts.xlsx`.")

        c_df = bot.load_contacts()
        if not c_df.empty:
            st.dataframe(c_df, width='stretch', height=450)
            st.caption("Файл на регистъра: `suppliers_brochure_contacts.xlsx`")

    # --------------------------------------------------------------------------
    # SUBTAB 5: Настройки на новата пощенска кутия
    # --------------------------------------------------------------------------
    with m_subtabs[4]:
        st.markdown("### ⚙️ Конфигурация на новата поща само за брошури")
        st.write("Въведете данните на новосъздадената пощенска кутия. Всички настройки се запазват защитено в `brochure_email_config.json`.")

        c_cfg = bot.config
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            new_email = st.text_input("Имейл адрес на новата поща:", value=c_cfg.get("email_address", ""), placeholder="broshuri.bakaliya@gmail.com")
            new_pwd = st.text_input("Парола за приложения (App Password):", value=c_cfg.get("app_password", ""), type="password", help="За Gmail: генерира се от Google Account -> Security -> App passwords (16 символа).")
            sender_n = st.text_input("Име на подателя:", value=c_cfg.get("sender_name", "Търговски отдел — Супермаркети „Бакалия“"))

        with col_c2:
            new_smtp = st.text_input("SMTP Сървър (изходящи):", value=c_cfg.get("smtp_server", "smtp.gmail.com"))
            new_smtp_p = st.number_input("SMTP Порт:", value=int(c_cfg.get("smtp_port", 587)))
            new_imap = st.text_input("IMAP Сървър (входящи):", value=c_cfg.get("imap_server", "imap.gmail.com"))
            new_imap_p = st.number_input("IMAP Порт:", value=int(c_cfg.get("imap_port", 993)))

        if st.button("💾 Запази настройките на пощата", type="primary"):
            bot.save_config({
                "email_address": new_email.strip(),
                "app_password": new_pwd.strip(),
                "sender_name": sender_n.strip(),
                "smtp_server": new_smtp.strip(),
                "smtp_port": int(new_smtp_p),
                "imap_server": new_imap.strip(),
                "imap_port": int(new_imap_p)
            })
            st.toast("✅ Настройките на пощата са успешно запазени!")
            st.success("Конфигурацията е запазена в `brochure_email_config.json`.")

