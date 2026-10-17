# -*- coding: utf-8 -*-
"""
================================================================================
Верига „Бакалия“ — Модул за търговски одит и контрол на правилата за брошури
================================================================================
Правила:
1. Забрана за идентични/дублирани артикули в една брошура (канибализация)
2. Забрана за продажба на загуба (Марж <= 0)
3. Защита: Промоционалната цена без ДДС да не е под редовната базова доставна цена
4. Забрана за дебют на нови артикули директно в брошура (само активни с история)
"""

import re

# Речник за автоматично категоризиране на критични сегменти (за предотвратяване на дублиране)
CRITICAL_SEGMENTS = {
    "сирене_краве_насипно": [r"сирене.*краве.*(?:кг|насип|вак.*кг)", r"краве.*сирене.*(?:кг|насип)"],
    "кисело_мляко_2": [r"кисело.*мляко.*2%", r"к.*мл.*2%"],
    "кисело_мляко_3_6": [r"кисело.*мляко.*(?:3\.6%|4\.5%|3%)", r"к.*мл.*3\.6%"],
    "прясно_мляко": [r"прясно.*мляко", r"пр.*мл"],
    "тоалетна_хартия_8": [r"тоалетна.*хартия.*8.*бр", r"тх.*8.*бр", r"тоал.*хартия.*8"],
    "тоалетна_хартия_24_32": [r"тоалетна.*хартия.*(?:24|32|40).*бр", r"тх.*(?:24|32)"],
    "олио_1л": [r"олио.*1.*л", r"слънчогледово.*олио.*1.*л"],
    "захар_1кг": [r"захар.*1.*кг", r"бяла.*захар.*1.*кг"],
    "яйца": [r"яйца.*(?:10|6|30).*бр", r"яйца.*размер"],
    "кашкавал_краве_вак": [r"кашкавал.*краве.*(?:400|300|200|вак)", r"краве.*кашкавал.*вак"],
    "бира_кен": [r"бира.*кен.*500", r"кен.*500.*мл"],
    "кайма_свинска": [r"кайма.*свинска", r"мляно.*месо.*100%.*свинско"],
}

def detect_segment(item_name):
    clean = item_name.lower()
    for seg_id, patterns in CRITICAL_SEGMENTS.items():
        for pat in patterns:
            if re.search(pat, clean):
                return seg_id
    return None

def audit_brochure_item(item, current_brochure_items, known_active_codes):
    """
    Извършва пълен одит по 4-те задължителни правила на „Бакалия“.
    Връща списък от предупреждения и статус: 'APPROVED', 'WARNING', 'BLOCKED'.
    """
    flags = []
    status = "APPROVED"
    
    code = item.get("code")
    name = item.get("name", "")
    reg_dc = float(item.get("reg_dc", 0.0) or 0.0)
    promo_dc = float(item.get("promo_dc", 0.0) or 0.0)
    reg_pc = float(item.get("reg_pc", 0.0) or 0.0)
    promo_pc = float(item.get("promo_pc", 0.0) or 0.0)
    edition = item.get("edition", "Брошура 1")
    
    # Изчисляване на марж на база цени с ДДС
    margin_pct = ((promo_pc - promo_dc) / promo_pc * 100) if promo_pc > 0 else 0.0
    
    # ПРАВИЛО 4: Нов артикул директно в брошура
    if code and known_active_codes and code not in known_active_codes:
        flags.append({
            "rule": 4,
            "level": "BLOCKED",
            "msg": f"⛔ Нов артикул (Код {code} няма история на доставки в Бакалия). Забранено активиране директно в брошура!"
        })
        status = "BLOCKED"
    elif not code:
        flags.append({
            "rule": 4,
            "level": "WARNING",
            "msg": "⚠️ Липсва код в Мистрал — възможен нов неизпитан артикул!"
        })
        if status != "BLOCKED":
            status = "WARNING"

    # ПРАВИЛО 2: Продажба на загуба (Марж <= 0)
    if promo_pc > 0 and promo_dc > 0:
        if promo_pc <= promo_dc:
            loss = promo_dc - promo_pc
            flags.append({
                "rule": 2,
                "level": "BLOCKED",
                "msg": f"⛔ ПРОДАЖБА НА ЗАГУБА! Промо ПЦ с ДДС ({promo_pc:.2f} €) е под Промо ДЦ с ДДС ({promo_dc:.2f} €). Загуба: -{loss:.2f} €/бр.!"
            })
            status = "BLOCKED"
            
    # ПРАВИЛО 3: Промо ПЦ с ДДС е по-ниска от базовата редовна доставна цена с ДДС
    if promo_pc > 0 and reg_dc > 0:
        if promo_pc < reg_dc:
            diff = reg_dc - promo_pc
            flags.append({
                "rule": 3,
                "level": "WARNING",
                "msg": f"⚠️ Промо ПЦ с ДДС ({promo_pc:.2f} €) е под Базовата ДЦ с ДДС ({reg_dc:.2f} €). Риск при остатъчни наличности (-{diff:.2f} €)!"
            })
            if status != "BLOCKED":
                status = "WARNING"

    # ПРАВИЛО 1: Идентични / дублирани артикули в същото издание
    item_seg = detect_segment(name)
    if item_seg:
        for other in current_brochure_items:
            # Сравняваме само ако са в същото издание (напр. и двата са в Брошура 1)
            other_ed = other.get("edition", "Брошура 1")
            if other_ed == edition and other.get("code") != code:
                other_seg = detect_segment(other.get("name", ""))
                if other_seg == item_seg:
                    flags.append({
                        "rule": 1,
                        "level": "WARNING",
                        "msg": f"⚠️ ДУБЛИРАН АРТИКУЛ в {edition}! Вече има артикул от сегмент '{item_seg}': \"{other.get('name')}\". Препоръка: преместете в Брошура 2."
                    })
                    if status != "BLOCKED":
                        status = "WARNING"
                    break

    return {
        "status": status,
        "flags": flags,
        "margin_pct": margin_pct
    }
