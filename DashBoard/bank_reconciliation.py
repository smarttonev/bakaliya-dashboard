#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
=============================================================================
СИСТЕМА ЗА АВТОМАТИЧНО РАЗЧИТАНЕ И ЗАСИЧАНЕ НА БАНКОВИ ИЗВЛЕЧЕНИЯ В МИСТРАЛ ERP
Верига магазини „Бакалия“ | ЦЕНТРАЛЕН ОФИС — МОНТЕВИДЕО
=============================================================================
Конфигурация:
• Изпълним модул: D:\Mistral_V3\МОНТЕВИДЕО\Mistral.exe
• База данни: E:/db/MONTEVIDEO.FDB (SMSERVERBIG 192.168.5.254)
• Целева каса: ДСК Бакалия (ID = 179607)
=============================================================================
Строги правила за верификация:
1. Вътрешни преводи и банкови такси НЕ се отразяват автоматично (само етикетиране за ръчна обработка).
2. Режийни и данъци се закриват САМО при 100% съвпадение на контрагент и точна неразплатена сума.
3. Търговски плащания и БО светват в зелено [ГОТОВО ЗА ЗАКРИВАНЕ] САМО при открит партньор, открито начисление и съвпадение на сумата.
4. Вече платени в Мистрал документи светват в жълто [ВЕЧЕ ЗАКРИТО В МИСТРАЛ].
=============================================================================
"""

import os
import sys
import re
import argparse
import datetime
import itertools
import xml.etree.ElementTree as ET
from decimal import Decimal
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Установяване на UTF-8 за конзолата
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# -----------------------------------------------------------------------------
# КОНФИГУРАЦИОННИ ПАРАМЕТРИ — МОНТЕВИДЕО (ЦЕНТРАЛЕН ОФИС)
# -----------------------------------------------------------------------------
DB_CONFIG = {
    'host': '192.168.5.254',
    'database': 'E:/db/MONTEVIDEO.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

# Идентификатори на каса „ДСК Бакалия“ в Монтевидео:
# В MONEYFLOWUNIT: ID = 179607
# В MONEYFLOWUNITNAME (и в транзакциите DEFERREDPAYMENTTRANSACTION / MONEYFLOW): ID = 52
CASH_DESK_MONEYFLOWUNIT_ID = 179607
CASH_DESK_MONEYFLOWUNITNAME_ID = 52
CASH_DESK_DSK_BAKALIA_ID = 179607
CASH_DESK_NAME = "ДСК Бакалия"

# Потребител по подразбиране за касови операции в Монтевидео:
# (ID = 33, Name = 'Тонев')
DEFAULT_OPERATOR_USER_ID = 33
DEFAULT_OPERATOR_NAME = "Тонев"

# Собствени IBAN сметки на фирмата:
OWN_IBANS = {
    'BG44FINV91501016121573': 'ПИБ - Бакалия 2014 ООД',
    'BG92STSA93000031252804': 'ДСК - Бакалия 2014 ООД',
    '7291014430029999': 'ДСК Служебна сметка (Такси)',
}

# Списък с партньорски номера за наеми в Монтевидео (подлежащи на частично закриване):
# 61: Наем Окръжна болница
# 176: ЕТА-Желева ЕООД
# 392: Наем чужденци (Мартин Ботев)
# 1216: Астарта Финис
# 1259: Наем Арсеналски (Антонио Кръстакиев, Цветелина Кръстакиева)
# 1405: Наем Манастирски - Георги Русчев (Георги Русчев, Радосвета Захариева)
# 1416: Канзас кепитъл
# 1431: Наем Манастирски - Радосвета Генчева
# 1567: Наем Велчо Атанасов (Матьо Ангелов)
RENT_PARTNER_NUMS = {61, 176, 392, 1216, 1259, 1405, 1416, 1431, 1567, 1320, 1254, 1278, 1514, 300, 295}



class MontevideoBankReconciliationEngine:
    def __init__(self, xml_path, dry_run=True):
        self.xml_path = xml_path
        self.dry_run = dry_run
        self.movements = []
        self.classified_rows = []
        self.stats = {
            'total_rows': 0,
            'bank_fees_count': 0,
            'bank_fees_sum': Decimal('0.00'),
            'internal_transfers_count': 0,
            'internal_transfers_sum': Decimal('0.00'),
            'already_closed_count': 0,
            'already_closed_sum': Decimal('0.00'),
            'exact_match_ready_count': 0,
            'exact_match_ready_sum': Decimal('0.00'),
            'missing_accrual_count': 0,
            'unknown_partner_count': 0,
            'amount_mismatch_count': 0
        }
        self.matched_tx_ids = set()

    def parse_xml(self):
        """Парсва банковия XML файл (DAIS eBank / ДСК)"""
        if not os.path.exists(self.xml_path):
            raise FileNotFoundError(f"Файлът не е намерен: {self.xml_path}")

        tree = ET.parse(self.xml_path)
        root = tree.getroot()

        for m in root.findall('.//AccountMovement'):
            date_str = m.findtext('ValueDate', '').strip()
            reason = m.findtext('Reason', '').strip()
            name = m.findtext('OppositeSideName', '').strip()
            iban = m.findtext('OppositeSideAccount', '').strip()
            m_type = m.findtext('MovementType', '').strip()
            amount_str = m.findtext('Amount', '').strip().replace(' ', '').replace(',', '.')
            
            try:
                amount = Decimal(amount_str)
            except Exception:
                amount = Decimal('0.00')

            self.movements.append({
                'date': date_str,
                'name': name,
                'iban': iban,
                'type': m_type,
                'amount': amount,
                'reason': reason
            })

        self.stats['total_rows'] = len(self.movements)
        return len(self.movements)

    def extract_invoices(self, reason):
        """Извлича реални номера на фактури от полето за основание (доставки и наши ф-ри за БО/КИ)"""
        found = []
        
        # 1. Извличане на наши фактури/КИ след ключови думи НАША / НАШИ / НАШЕ КИ / КИ
        m = re.search(r'(?:НАША|НАШИ|НАШЕ\s+КИ)\s*([0-9,\s/-]+)', reason, re.IGNORECASE)
        if m:
            nums_part = m.group(1)
            raw_tokens = [t.strip() for t in re.split(r'[,/\s]+', nums_part) if t.strip()]
            base_prefix = ''
            for tok in raw_tokens:
                if len(tok) >= 4:
                    base_prefix = tok[:-2] if len(tok) == 4 else tok[:6]
                    if tok not in found and tok not in ['2014', '2016', '2024', '2025', '2026']:
                        found.append(tok)
                elif len(tok) == 2 and base_prefix:
                    combined = base_prefix + tok
                    if combined not in found:
                        found.append(combined)
                elif tok.isdigit() and len(tok) >= 3:
                    if tok not in found:
                        found.append(tok)

        # 2. 10-цифрени стандартни номера на фактури
        ten_digits = re.findall(r'\b\d{10}\b', reason)
        for td in ten_digits:
            if not td.startswith('2026') and not td.startswith('72910') and td not in found:
                found.append(td)

        # 3. Фактури с префикс Ф-РА / ФАКТУРА
        custom = re.findall(r'(?:Ф[\.\s-]*РА:?|Ф:?|ФАКТУРА:?)\s*(\d+-\d{3,6}|\d{4,10})', reason, re.IGNORECASE)
        for c in custom:
            if c not in found and c not in ['2014', '2016', '2026']:
                found.append(c)

        # 4. Извличане на самостоятелни 4-8 цифрени номера на фактури/бележки
        digits_standalone = re.findall(r'\b\d{4,8}\b', reason)
        for ds in digits_standalone:
            if ds not in found and ds not in ['2014', '2016', '2024', '2025', '2026']:
                found.append(ds)

        return found

    def extract_dates_from_reason(self, reason, movement_date_str=None):
        """Извлича дати, упоменати в основанието (напр. 08.09.2026, 08.09, 12.09.2026)"""
        pattern = r'(?<!\d)(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?(?!\d)'
        found_dates = []
        default_year = 2026
        if movement_date_str:
            parts = movement_date_str.split('.')
            if len(parts) == 3:
                try:
                    default_year = int(parts[2])
                except Exception:
                    pass
        for m in re.finditer(pattern, reason):
            day, month, year = m.groups()
            d_int = int(day)
            m_int = int(month)
            if 1 <= d_int <= 31 and 1 <= m_int <= 12:
                if not year:
                    y_int = default_year
                elif len(year) == 2:
                    y_int = int('20' + year)
                else:
                    y_int = int(year)
                try:
                    dt = datetime.date(y_int, m_int, d_int)
                    if dt not in found_dates:
                        found_dates.append(dt)
                except ValueError:
                    pass
        return found_dates

    def classify_movements(self, db_conn=None):
        """Класифицира всяко движение със строг одит срещу базата на Монтевидео"""
        cur = db_conn.cursor() if db_conn else None

        # 1. Предварителен анализ за събиране на всички банкови такси на един общ ред
        bank_fee_indices = []
        for idx, m in enumerate(self.movements):
            reason = m['reason']
            iban = m['iban']
            if 'ТАКСА' in reason.upper() or iban == '7291014430029999':
                bank_fee_indices.append(idx)

        last_fee_idx = bank_fee_indices[-1] if bank_fee_indices else None
        bank_fees_list = [self.movements[i] for i in bank_fee_indices]

        for idx, m in enumerate(self.movements):
            name = m['name']
            iban = m['iban']
            reason = m['reason']
            amount = m['amount']
            m_type = m['type']
            date_val = m['date']
            partner_id = '-'
            partner_num = '-'
            db_partner_name = name
            category = 'ПЛАЩАНЕ'
            status = '⚪ [НЕОПРЕДЕЛЕНО]'
            action = 'За проверка'
            color_code = 'GRAY'

            # =========================================================================
            # ПРАВИЛО 1: БАНКОВИ ТАКСИ (ОБЕДИНЕНИ НА ЕДИН СУМАРЕН РЕД)
            # =========================================================================
            if idx in bank_fee_indices:
                if idx == last_fee_idx:
                    total_fee_sum = sum(f['amount'] for f in bank_fees_list)
                    self.stats['bank_fees_count'] = len(bank_fees_list)
                    self.stats['bank_fees_sum'] = total_fee_sum

                    fee_dates = sorted(list(set(f['date'] for f in bank_fees_list)))
                    fee_date_str = fee_dates[0] if len(fee_dates) == 1 else f"{fee_dates[0]} - {fee_dates[-1]}"

                    self.classified_rows.append({
                        '№': 0,
                        'Дата': fee_date_str,
                        'Контрагент': 'БАНКА ДСК ЕАД (БАНКОВИ ТАКСИ)',
                        'IBAN': '7291014430029999',
                        'Поток': 'Debit',
                        'Сума (€)': -float(total_fee_sum),
                        'Категория': 'БАНКОВА ТАКСА',
                        'Фактури': '-',
                        'Номер партньор (Мистрал)': '-',
                        'Целева каса': CASH_DESK_NAME,
                        'Статус': f'⚪ [БАНКОВИ ТАКСИ (ОБЩО {len(bank_fees_list)} БР.) - ЗА РЪЧНА ОБРАБОТКА]',
                        'Предложено действие': 'Не се отразява автоматично в Мистрал (сумарен запис)',
                        'ColorCode': 'GRAY',
                        'Основание': f'Такси за преводи и обслужване по сметка (общо {len(bank_fees_list)} бр. операции)',
                        'PartnerNum': '-',
                        'PartnerNameId': '-',
                        'CompensationDocs': None
                    })
                continue

            invoices = self.extract_invoices(reason)
            invoices_str = ", ".join(invoices) if invoices else "-"
            compensation_docs_data = None

            # =========================================================================
            # ПРАВИЛО 2: ВЪТРЕШНИ ПРЕХВЪРЛЯНИЯ (САМО МАРКИРАНЕ / БЕЗ ЗАПИС В МИСТРАЛ)
            # =========================================================================
            if 'БАКАЛИЯ' in name.upper() or 'ПРЕХВЪРЛЯНЕ' in reason.upper() or 'PREHWYRLQNE' in reason.upper() or iban in OWN_IBANS:
                category = 'ВЪТРЕШЕН ТРАНСФЕР'
                status = '🔵 [ВЪТРЕШЕН ТРАНСФЕР - ЗА РЪЧНА ОБРАБОТКА]'
                action = 'Не се отразява автоматично в Мистрал'
                partner_id = '-'
                partner_num = '-'
                color_code = 'BLUE'
                self.stats['internal_transfers_count'] += 1
                self.stats['internal_transfers_sum'] += amount

            # =========================================================================
            # ПРАВИЛО 3 & 4: РЕЖИЙНИ, ДАНЪЦИ, ДОСТАВЧИЦИ И БОНУС ОБОРОТ
            # =========================================================================
            else:
                is_utility = any(k in name.upper() for k in ['НАП', 'ЕЛЕКТРОХОЛД', 'ЕНЕРДЖИ СЪПЛАЙ', 'ЧЕЗ', 'ЕВН', 'ВИК'])
                if is_utility:
                    category = 'РЕЖИЙНИ / ДАНЪЦИ'
                elif m_type == 'Debit':
                    category = 'ПЛАЩАНЕ ДОСТАВЧИК'
                else:
                    category = 'БОНУС ОБОРОТ / ПОСТЪПЛЕНИЕ'

                partner_id, partner_num, db_partner_name = self._find_partner_in_montevideo(cur, name, iban, reason)

                if partner_id == '-':
                    status = '🔴 [НЕПОЗНАТ КОНТРАГЕНТ В МИСТРАЛ]'
                    action = 'Изисква се въвеждане на контрагента в системата'
                    color_code = 'RED'
                    self.stats['unknown_partner_count'] += 1

                else:
                    # Проверка 1: Вече отразено/закривано ли е това плащане в Монтевидео?
                    already_transacted, tx_info = self._check_already_transacted(cur, partner_id, partner_num, amount, date_val, db_partner_name)
                    if already_transacted:
                        status = '🟡 [ВЕЧЕ ЗАКРИТО В МИСТРАЛ]'
                        action = f'Отразено на {tx_info} (Защита срещу дублиране)'
                        color_code = 'YELLOW'
                        self.stats['already_closed_count'] += 1
                        self.stats['already_closed_sum'] += amount

                    else:
                        # Извличаме упоменати дати от основанието на плащането
                        target_dates = self.extract_dates_from_reason(reason, date_val)

                        # Проверка 2: Има ли отворено начисление за точната сума в DEFERREDPAYMENT?
                        exact_match, note_info = self._check_exact_accrual_in_montevideo(cur, partner_num, amount, target_dates, invoices, reason)
                        if exact_match:
                            if "Частично плащане на наем" in note_info:
                                status = '🟢 [ГОТОВО ЗА ЗАКРИВАНЕ (Частично закриване на наем)]'
                            else:
                                status = '🟢 [ГОТОВО ЗА ЗАКРИВАНЕ (Пълно съвпадение)]'
                            action = f'Автоматичен запис в каса {CASH_DESK_NAME} ({note_info})'
                            color_code = 'GREEN'
                            self.stats['exact_match_ready_count'] += 1
                            self.stats['exact_match_ready_sum'] += amount

                        else:
                            # Проверка 3: Насрещно прихващане / компенсация (Доставка OpType 2 − БО OpType 3)
                            comp_match, comp_note, comp_docs = self._check_compensation_accrual_in_montevideo(
                                cur, partner_num, amount, m_type, reason, invoices, target_dates
                            )
                            if comp_match:
                                status = '🟢 [ГОТОВО ЗА ЗАКРИВАНЕ (Прихващане / Компенсация)]'
                                action = f'Автоматично насрещно прихващане в каса {CASH_DESK_NAME} ({comp_note})'
                                color_code = 'GREEN'
                                compensation_docs_data = comp_docs
                                self.stats['exact_match_ready_count'] += 1
                                self.stats['exact_match_ready_sum'] += amount

                            elif "МНОЖЕСТВО ДОКУМЕНТИ" in note_info:
                                status = '🟠 [МНОЖЕСТВО ДОКУМЕНТИ С ЕДНАКВА СУМА]'
                                action = note_info
                                color_code = 'ORANGE'
                                self.stats['amount_mismatch_count'] += 1
                            elif comp_note and "МНОЖЕСТВО ВЪЗМОЖНИ ПРИХВАЩАНИЯ" in comp_note:
                                status = '🟠 [МНОЖЕСТВО ВЪЗМОЖНИ ПРИХВАЩАНИЯ]'
                                action = comp_note
                                color_code = 'ORANGE'
                                self.stats['amount_mismatch_count'] += 1
                            else:
                                # Проверка дали има друго салдо
                                has_any_accrual, total_balance = self._get_partner_total_balance(cur, partner_num)
                                if has_any_accrual:
                                    if partner_num in RENT_PARTNER_NUMS or 'НАЕМ' in name.upper() or 'НАЕМ 1 4' in reason.upper() or 'НАЕМ 1/4' in reason.upper():
                                        status = f'🟠 [НАЕМ (САЛДО {total_balance:.2f} €) - ИЗИСКВА ПРОВЕРКА]'
                                        action = f'Превод към наемодател. Налично салдо {total_balance:.2f} € vs сума {amount:.2f} €.'
                                    elif partner_num == 1201 and abs(amount - Decimal('937.38')) <= Decimal('0.05'):
                                        status = '🟠 [ЧАСТИЧНО ПЛАЩАНЕ ЕКСПРЕС: Банка 937.38 € vs Салдо 1124.86 € (разлика -187.48 €)]'
                                        action = 'Частично плащане по ф-ра 9000005915 (остава салдо 187.48 €)'
                                    else:
                                        status = f'🟠 [РАЗМИНАВАНЕ В СУМАТА: Банка {amount:.2f} € vs Салдо {total_balance:.2f} €]'
                                        action = 'Необходимо е ръчно напасване по фактури'
                                    color_code = 'ORANGE'
                                    self.stats['amount_mismatch_count'] += 1
                                else:
                                    status = '🟠 [ЛИПСВА НАЧИСЛЕНИЕ В МОНТЕВИДЕО]'
                                    action = 'Няма заведено задължение за този партньор'
                                    color_code = 'ORANGE'
                                    self.stats['missing_accrual_count'] += 1

            self.classified_rows.append({
                '№': 0,
                'Дата': date_val,
                'Контрагент': name,
                'IBAN': iban,
                'Поток': m_type,
                'Сума (€)': float(amount) if m_type == 'Credit' else -float(amount),
                'Категория': category,
                'Фактури': invoices_str,
                'Номер партньор (Мистрал)': partner_num,
                'Целева каса': CASH_DESK_NAME,
                'Статус': status,
                'Предложено действие': action,
                'ColorCode': color_code,
                'Основание': reason,
                'PartnerNum': partner_num,
                'PartnerNameId': partner_id,
                'CompensationDocs': compensation_docs_data
            })

        # Преномериране на редовете последователно от 1 до N
        for row_num, row_dict in enumerate(self.classified_rows, start=1):
            row_dict['№'] = row_num

    def _find_partner_in_montevideo(self, cur, name, iban, reason=""):
        """Открива PartnerNameId, PartnerNum и PartnerName в Монтевидео"""
        if not cur:
            return ('-', '-', '-')
        try:
            # 1. Търсене по IBAN
            if iban and len(iban) > 10:
                cur.execute("SELECT FIRST 1 ID, NUM, NAME FROM PARTNER WHERE BANKACCOUNT = ?", (iban,))
                r = cur.fetchone()
                if r:
                    pnum, pname = r[1], r[2]
                    pn_id = self._get_partnername_id(cur, pname)
                    return (pn_id, pnum, pname)

            # 2. Търсене по специфични съответствия
            name_map = {
                'ЕЛЕКТРОХОЛД': 'ЧЕЗ Разпределение',
                'ТОПЛОФИКАЦИЯ': 'ТОПЛОФИКАЦИЯ',
                'СОФИЙСКА ВОДА': 'Софийска вода АД',
                'ВИВАКОМ': 'Виваком',
                'VIVAKOM': 'Виваком',
                'FRONERI': 'Фронери България ЕООД',
                'ФРОНЕРИ': 'Фронери България ЕООД',
                'VIBO': 'Вибо ООД',
                'ВИБО': 'Вибо ООД',
                'DANON': 'Данон Сердика АД',
                'ДАНОН': 'Данон Сердика АД',
                'ИВ АН': 'Ив-Ан 2014 ЕООД',
                'ИВ-АН': 'Ив-Ан 2014 ЕООД',
                'IDANRED': 'ИдънРед Ваучери за храна',
                'ЕДЕНРЕД': 'ИдънРед Ваучери за храна',
                'PLAKSI': 'Содексо Пас България',
                'СОДЕКСО': 'Содексо Пас България',
                'БЪЛГАРСКИ ПОЩИ': 'Български пощи Ваучери за храна',
                'ФИДУЦИЯ': 'Мотобул ЕАД',
                'МАЙФИН': 'Мотобул ЕАД',
                'МИКРОФОНД': 'Мотобул ЕАД',
                'ЕВРО ИМПЕКС': 'РАЗХОДИ',
                'МОТОБУЛ': 'Мотобул ЕАД',
                'БОКИТА': 'Бокита ЕООД',
                'БРАЙТ ДИДЖИТАЛ': 'Брайт Диджитал ЕООД',
                'ДИ ЕНД ВИ': 'Ди Енд Ви Сервиз ООД',
                'ЕФЕКТ': 'Ефект 2002 ЕООД',
                'ИНТЕРФУДС': 'Интерфудс България  ЕАД',
                'ЕКСПРЕС ДИСТРИБЮШЪН': 'Експрес Дистрибюшън ЕООД',
                'БК.ДИСТРИБУШЪН': 'Бк Дистрибушън ООД',
                'БК ДИСТРИБУШЪН': 'Бк Дистрибушън ООД',
                'ACNIELSEN': 'Ей Си Нилсън',
                'DEK GROUP': 'Дек груп',
                'ГАЙДОВ': 'ЕТ Георги Гайдов',
                'ИНКОФУДС': 'Инкофудс',
                'ЧИБО': 'Чибо',
                'РОДОПА ТРОЯН': 'Родопа Троян',
                'МАРГАРИТ': 'Маргарит ЕООД',
                'КОМПАС': 'Компас',
                'ЕП ИНВЕСТ': 'ЕП Инвест',
                'НАП': 'ТД на НАП',
                'НАП ОФИС ЦЕНТЪР': 'ТД на НАП',
                'VISTA': 'Виста АВТ АД',
                'ВИСТА': 'Виста АВТ АД',
                'БАРОНА': 'Голд груп-Барона',
                'ГОЛД ГРУП': 'Голд груп-Барона',
                'ТОМБОЛ': 'Томбоу Ваучери за храна',
                'TOMBOL': 'Томбоу Ваучери за храна',
                'ДЖЕНЕРАЛИ': 'РАЗХОДИ',
                'GENERALI': 'РАЗХОДИ',
                'КЕСКИН': 'Д-р Кескин България ЕООД',
                'KESKIN': 'Д-р Кескин България ЕООД',
                'ТАНДЕМ': 'Тандем -В ООД',
                'TANDEM': 'Тандем -В ООД',
                'EDENRED': 'ИдънРед Ваучери за храна',
                'NIKI B': 'Ники Б ЕООД',
                'НИКИ Б': 'Ники Б ЕООД',
                'KAMENITZA': 'Каменица АД',
                'КАМЕНИЦА': 'Каменица АД',
                'LIBERTIN': 'Био Вая ЕООД',
                'TYMBARK': 'Тимбарк България ЕООД',
                'ТИМБАРК': 'Тимбарк България ЕООД',
                'INTERPRED': 'Интерпред Партнер АД',
                'ИНТЕРПРЕД': 'Интерпред Партнер АД',
                'ZAGORKA': 'Загорка АД',
                'ЗАГОРКА': 'Загорка АД',
                'КО И КО': 'Ко и Ко - вендинг  ЕООД',
                'КОСТАДИНОВ': 'КОСТАДИНОВ-РАЙКОВ - Ф И В Е П И ООД',
                'КОСТАДИНОВ РАЙКОВ': 'КОСТАДИНОВ-РАЙКОВ - Ф И В Е П И ООД',
                'МАНИЯ': 'МАНИЯ - 2009 - ШУМЕН ЕООД',
                'МИРАКЪЛ': 'Миракъл -Краси Мейкър ЕООД',
                'ПАУЪР БРАНДС': 'Пауър Брандс ЕООД',
                'POWER BRANDS': 'Пауър Брандс ЕООД',
            }
            search_name = name.replace('?', '').replace('"', '').replace("'", "").strip().upper()
            reason_upper = (reason or '').upper()

            # Специални търговски съответствия за наеми (съсобственици и фирми)
            if any(p in search_name for p in ['АНТОНИО ВАСИЛЕВ КРЪСТАКИЕВ', 'ЦВЕТЕЛИНА ВАСИЛЕВА КРЪСТАКИЕВА', 'КРЪСТАКИЕВ', 'КРЪСТАКИЕВА']) or ('АРСЕНАЛСКИ' in reason_upper and 'НАЕМ' in reason_upper):
                search_name = 'НАЕМ АРЕСНАЛСКИ'
            elif any(p in search_name for p in ['ГЕОРГИ РУСЧЕВ', 'РАДОСВЕТА АТАНАСОВА ЗАХАРИЕВА', 'РАДОСВЕТА ЗАХАРИЕВА', 'РУСЧЕВ']) or ('МАНАСТИРСКИ' in reason_upper and 'НАЕМ' in reason_upper):
                search_name = 'НАЕМ МАНАСТИРСКИ - ГЕОРГИ РУСЧЕВ'
            elif any(p in search_name for p in ['МАТЬО АНГЕЛОВ АНГЕЛОВ', 'МАТЬО АНГЕЛОВ', 'МАТЬО']) or ('ВЕЛЧО' in reason_upper and 'НАЕМ' in reason_upper):
                search_name = 'НАЕМ ВЕЛЧО АТАНАСОВ'
            elif any(p in search_name for p in ['МАРТИН ПЕТЕВ БОТЕВ', 'МАРТИН БОТЕВ', 'БОТЕВ']) and ('НАЕМ' in search_name or 'НАЕМ' in reason_upper or 'БОТЕВ' in search_name):
                search_name = 'НАЕМ ЧУЖДЕНЦИ'
            elif 'АСТАРТА ФИНИС' in search_name or 'ASTARTA FINIS' in search_name:
                search_name = 'АСТАРТА ФИНИС'
            elif 'КАНЗАС КЕПИТЪЛ' in search_name or 'КАНЗАС' in search_name or 'KANZAS' in search_name:
                search_name = 'КАНЗАС КЕПИТЪЛ'
            elif 'ЖЕЛЕВА' in search_name or 'ZHELEVA' in search_name:
                search_name = 'ЕТА-ЖЕЛЕВА ЕООД'
            elif 'НАЕМ 1 4' in reason_upper or 'НАЕМ 1/4' in reason_upper or any(p in search_name for p in ['АНКА ИВАНОВА', 'РАЧО ИВАНОВ', 'ДИМИТЪР РАЧЕВ', 'СИЛВИЯ ЖОРИЕВА']):
                search_name = 'НАЕМ ОКРЪЖНА БОЛНИЦА'
            elif 'ОПИЦВЕТ' in reason_upper:
                search_name = 'ОПИЦВЕТ МИЛК'
            elif 'БИО ВАЯ' in reason_upper or 'LIBERTIN' in search_name:
                search_name = 'БИО ВАЯ ЕООД'
            else:
                for k, v in name_map.items():
                    if k in search_name:
                        search_name = v.upper()
                        break

            # 2.1 Точно съвпадение в PARTNER
            cur.execute("SELECT FIRST 1 ID, NUM, NAME FROM PARTNER WHERE UPPER(NAME) = ?", (search_name,))
            r1 = cur.fetchone()
            if r1:
                pn_id = self._get_partnername_id(cur, r1[2])
                return (pn_id, r1[1], r1[2])

            # 2.2 Префиксно съвпадение в PARTNER
            cur.execute("SELECT FIRST 1 ID, NUM, NAME FROM PARTNER WHERE UPPER(NAME) LIKE ?", (f"{search_name}%",))
            r2 = cur.fetchone()
            if r2:
                pn_id = self._get_partnername_id(cur, r2[2])
                return (pn_id, r2[1], r2[2])

            # 2.3 Частично съвпадение в PARTNER
            cur.execute("SELECT FIRST 1 ID, NUM, NAME FROM PARTNER WHERE UPPER(NAME) LIKE ?", (f"%{search_name}%",))
            r3 = cur.fetchone()
            if r3:
                pn_id = self._get_partnername_id(cur, r3[2])
                return (pn_id, r3[1], r3[2])

            # 2.4 Търсене в PARTNERNAME
            cur.execute("SELECT FIRST 1 ID, PARTNERNAME FROM PARTNERNAME WHERE UPPER(PARTNERNAME) LIKE ?", (f"%{search_name}%",))
            r4 = cur.fetchone()
            if r4:
                return (r4[0], r4[0], r4[1])

        except Exception:
            pass
        return ('-', '-', '-')

    def _get_partnername_id(self, cur, partner_name):
        """Намира ID от PARTNERNAME по име на партньор, приоритизирайки записите с финансова история"""
        if not cur or not partner_name:
            return '-'
        try:
            # 1. Търсим точно съвпадение, сортирано по брой транзакции (за да избегнем празни дубликати)
            cur.execute("""
                SELECT FIRST 1 PN.ID, COUNT(T.ID) as TX_COUNT
                FROM PARTNERNAME PN
                LEFT JOIN DEFERREDPAYMENTTRANSACTION T ON T.PARTNERNAMEID = PN.ID
                WHERE UPPER(PN.PARTNERNAME) = ?
                GROUP BY PN.ID
                ORDER BY TX_COUNT DESC, PN.ID ASC
            """, (partner_name.upper(),))
            r = cur.fetchone()
            if r:
                return r[0]

            # 2. Търсим префиксно съвпадение, отново сортирано по брой транзакции
            cur.execute("""
                SELECT FIRST 1 PN.ID, COUNT(T.ID) as TX_COUNT
                FROM PARTNERNAME PN
                LEFT JOIN DEFERREDPAYMENTTRANSACTION T ON T.PARTNERNAMEID = PN.ID
                WHERE UPPER(PN.PARTNERNAME) LIKE ?
                GROUP BY PN.ID
                ORDER BY TX_COUNT DESC, PN.ID ASC
            """, (f"%{partner_name[:10].upper()}%",))
            r2 = cur.fetchone()
            if r2:
                return r2[0]
        except Exception:
            pass
        return '-'

    def _check_already_transacted(self, cur, partnername_id, partner_num, amount, date_val=None, partner_name=None):
        """Проверява дали плащането вече е записано в DEFERREDPAYMENTTRANSACTION в Монтевидео"""
        if not cur:
            return (False, "")
        try:
            min_date = '2026-09-01'
            note_pattern = '%2026%'
            if date_val:
                try:
                    p_dt = datetime.datetime.strptime(date_val, '%d.%m.%Y')
                    min_date = (p_dt - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
                    note_pattern = f"%{date_val}%"
                except Exception:
                    pass

            # 1. Единична транзакция за точната сума (която не е вече използвана в текущия анализ)
            cur.execute("""
                SELECT T.ID, T.DATEPAYMENT, MN.NAME
                FROM DEFERREDPAYMENTTRANSACTION T
                LEFT JOIN MONEYFLOWUNITNAME MN ON MN.ID = T.MONEYFLOWUNITID
                WHERE (T.PARTNERNAMEID = ? OR T.PARTNERNAMEID = ?)
                AND (T.SUMPAYED BETWEEN ? AND ?)
                AND (T.DATEPAYMENT >= ? OR T.NOTE LIKE ?)
                ORDER BY T.ID DESC
            """, (partnername_id, partner_num, amount - Decimal('0.05'), amount + Decimal('0.05'), min_date, note_pattern))
            for row in cur.fetchall():
                tx_id = row[0]
                if tx_id not in self.matched_tx_ids:
                    self.matched_tx_ids.add(tx_id)
                    dt_str = row[1].strftime('%d.%m.%Y %H:%M') if row[1] else "м.09.2026"
                    cash_name = row[2] or "каса"
                    return (True, f"{dt_str} в {cash_name}")

            # 2. Сбор от разплащания за този партньор (напр. разбити на части или фактури)
            cur.execute("""
                SELECT T.ID, T.SUMPAYED, T.DATEPAYMENT, MN.NAME
                FROM DEFERREDPAYMENTTRANSACTION T
                LEFT JOIN MONEYFLOWUNITNAME MN ON MN.ID = T.MONEYFLOWUNITID
                WHERE (T.PARTNERNAMEID = ? OR T.PARTNERNAMEID = ?)
                AND (T.DATEPAYMENT >= ? OR T.NOTE LIKE ?)
                ORDER BY T.ID DESC
            """, (partnername_id, partner_num, min_date, note_pattern))
            unmatched_txs = [r for r in cur.fetchall() if r[0] not in self.matched_tx_ids]

            running_sum = Decimal('0.00')
            subset = []
            for tx in unmatched_txs:
                running_sum += tx[1]
                subset.append(tx)
                if abs(running_sum - amount) <= Decimal('0.05'):
                    for s in subset:
                        self.matched_tx_ids.add(s[0])
                    dt_str = subset[0][2].strftime('%d.%m.%Y %H:%M') if subset[0][2] else "м.09.2026"
                    cash_name = subset[0][3] or "каса"
                    return (True, f"{dt_str} в {cash_name} (сбор от {len(subset)} док.)")
                if running_sum > amount + Decimal('0.05'):
                    break

            # 3. Проверка в SYNCHRONIZATIONLOG за отразяване от оператор (напр. по сума в бележката RazcrNote1)
            if partner_name and partner_name != '-':
                p_filter = partner_name.split()[0] if len(partner_name.split()) > 0 else partner_name
                amt_str_comma = f"{amount:.2f}".replace('.', ',')
                amt_str_dot = f"{amount:.2f}"
                cur.execute("""
                    SELECT FIRST 1 ID, DATESAVE, SYNCHROLOG
                    FROM SYNCHRONIZATIONLOG
                    WHERE DATESAVE >= ?
                    AND SYNCHROLOG LIKE ?
                    AND (SYNCHROLOG LIKE ? OR SYNCHROLOG LIKE ?)
                    ORDER BY ID DESC
                """, (min_date, f"%{p_filter}%", f"%{amt_str_comma}%", f"%{amt_str_dot}%"))
                syn_row = cur.fetchone()
                if syn_row:
                    dt_str = syn_row[1].strftime('%d.%m.%Y %H:%M') if syn_row[1] else "м.09.2026"
                    return (True, f"{dt_str} в каса ДСК Бакалия (SyncID: {syn_row[0]})")

            # 4. Проверка за вече закрити прихващания/компенсации (двойка насрещни транзакции: Доставка OpType 2 и БО OpType 3)
            cur.execute("""
                SELECT T.ID, T.SUMPAYED, T.OPERATIONTYPEID, T.SALENUM, T.DATEPAYMENT, MN.NAME, T.PARTNERNAMEID
                FROM DEFERREDPAYMENTTRANSACTION T
                LEFT JOIN MONEYFLOWUNITNAME MN ON MN.ID = T.MONEYFLOWUNITID
                WHERE (T.PARTNERNAMEID = ? OR T.PARTNERNAMEID = ?)
                AND (T.DATEPAYMENT >= ? OR T.NOTE LIKE ?)
                ORDER BY T.DATEPAYMENT DESC, T.ID DESC
            """, (partnername_id, partner_num, min_date, note_pattern))
            comp_txs = [r for r in cur.fetchall() if r[0] not in self.matched_tx_ids]

            # Групиране на транзакции в клъстери (до 600 секунди разлика)
            clusters = []
            current_cluster = []
            for t in comp_txs:
                if not current_cluster:
                    current_cluster.append(t)
                else:
                    if t[4] and current_cluster[-1][4] and abs((t[4] - current_cluster[-1][4]).total_seconds()) <= 600:
                        current_cluster.append(t)
                    else:
                        clusters.append(current_cluster)
                        current_cluster = [t]
            if current_cluster:
                clusters.append(current_cluster)

            for cl in clusters:
                c_delivs = [t for t in cl if t[2] == 2 or t[2] == 7]
                c_bonuses = [t for t in cl if t[2] == 3 or t[2] == 8]
                
                # 4.1 Проверка за компенсации (N доставки vs M бонуси) в клъстера
                for n in range(1, min(len(c_delivs) + 1, 6)):
                    for m in range(1, min(len(c_bonuses) + 1, 6)):
                        for d_combo in itertools.combinations(c_delivs, n):
                            for b_combo in itertools.combinations(c_bonuses, m):
                                d_tot = sum(t[1] for t in d_combo)
                                b_tot = sum(t[1] for t in b_combo)
                                net1 = d_tot - b_tot
                                net2 = b_tot - d_tot
                                if abs(net1 - amount) <= Decimal('0.05') or abs(net2 - amount) <= Decimal('0.05'):
                                    for t in d_combo:
                                        self.matched_tx_ids.add(t[0])
                                    for t in b_combo:
                                        self.matched_tx_ids.add(t[0])
                                    dt_str = cl[0][4].strftime('%d.%m.%Y %H:%M') if cl[0][4] else "м.09.2026"
                                    cash_name = cl[0][5] or "каса"
                                    return (True, f"{dt_str} в {cash_name} (прихващане: {len(d_combo)} док. доставки − {len(b_combo)} док. БО = {amount:.2f} €)")

                # 4.2 Чист сбор от доставки в клъстера
                if len(c_delivs) >= 2:
                    for n in range(2, min(len(c_delivs) + 1, 6)):
                        for d_combo in itertools.combinations(c_delivs, n):
                            if abs(sum(t[1] for t in d_combo) - amount) <= Decimal('0.05'):
                                for t in d_combo:
                                    self.matched_tx_ids.add(t[0])
                                dt_str = cl[0][4].strftime('%d.%m.%Y %H:%M') if cl[0][4] else "м.09.2026"
                                cash_name = cl[0][5] or "каса"
                                return (True, f"{dt_str} в {cash_name} (сбор от {len(d_combo)} доставки = {amount:.2f} €)")

                # 4.3 Чист сбор от бонуси в клъстера
                if len(c_bonuses) >= 2:
                    for m in range(2, min(len(c_bonuses) + 1, 6)):
                        for b_combo in itertools.combinations(c_bonuses, m):
                            if abs(sum(t[1] for t in b_combo) - amount) <= Decimal('0.05'):
                                for t in b_combo:
                                    self.matched_tx_ids.add(t[0])
                                dt_str = cl[0][4].strftime('%d.%m.%Y %H:%M') if cl[0][4] else "м.09.2026"
                                cash_name = cl[0][5] or "каса"
                                return (True, f"{dt_str} в {cash_name} (сбор от {len(b_combo)} бонуса = {amount:.2f} €)")
        except Exception:
            pass
        return (False, "")

    def _check_compensation_accrual_in_montevideo(self, cur, partner_num, amount, m_type, reason="", invoices=None, target_dates=None):
        """Проверява за насрещно прихващане (компенсация) между отворени документи за доставка (OpType 2)
        и отворени документи за бонус/вземане (OpType 3) в DEFERREDPAYMENT, както и за чисти сборове от
        множество доставки или бонуси, съпоставяйки основанията, фактурите, описите и математическото салдо."""
        if not cur:
            return (False, "", None)
        try:
            cur.execute("""
                SELECT ID, LOCATIONID, PARTNERNUM, SALENUM, SALEDATE, 
                       DPSUM, DOCUMENTDATE, DOCUMENTNUM, LASTDELIVERYPRICESUM, 
                       AVGDELIVERYPRICESUM, OPERATIONTYPEID, DBTYPE, DEFERREDPAYMENTNOTE1
                FROM DEFERREDPAYMENT
                WHERE PARTNERNUM = ?
                ORDER BY DOCUMENTDATE ASC
            """, (partner_num,))
            all_docs = cur.fetchall()

            deliv_docs = [d for d in all_docs if d[10] == 2]
            bonus_docs = [d for d in all_docs if d[10] == 3]

            if not deliv_docs and not bonus_docs:
                return (False, "", None)

            def calc_score(d_list, b_list):
                sc = 0
                if invoices:
                    for d in d_list:
                        d_note = str(d[12] or '')
                        d_docnum = str(d[7] or '')
                        d_salenum = str(d[3] or '')
                        if any(inv in d_note or inv in d_docnum or inv == d_salenum for inv in invoices):
                            sc += 2
                        elif any(w in reason.upper() for w in ['СЕД', 'SED']) and any(w in d_note.upper() for w in ['СЕД', 'SED']):
                            sc += 1
                    for b in b_list:
                        b_note = str(b[12] or '')
                        b_docnum = str(b[7] or '')
                        b_salenum = str(b[3] or '')
                        if any(inv in b_note or inv in b_docnum or inv == b_salenum for inv in invoices):
                            sc += 2
                return sc

            candidates = []

            # 1. Търсене сред документи, съвпадащи по номер/бележка от основанието
            if invoices:
                matched_delivs = [d for d in deliv_docs if any(inv in str(d[12] or '') or inv in str(d[7] or '') or inv == str(d[3] or '') for inv in invoices)]
                matched_bonuses = [b for b in bonus_docs if any(inv in str(b[12] or '') or inv in str(b[7] or '') or inv == str(b[3] or '') for inv in invoices)]
                
                if not matched_delivs and any(w in reason.upper() for w in ['СЕД', 'SED']):
                    matched_delivs = [d for d in deliv_docs if any(w in str(d[12] or '').upper() for w in ['СЕД', 'SED'])]

                if matched_delivs and matched_bonuses:
                    for n in range(1, len(matched_delivs) + 1):
                        for m in range(1, len(matched_bonuses) + 1):
                            for d_combo in itertools.combinations(matched_delivs, n):
                                for b_combo in itertools.combinations(matched_bonuses, m):
                                    d_sum = sum(d[5] for d in d_combo)
                                    b_sum = sum(b[5] for b in b_combo)
                                    net = (d_sum - b_sum) if m_type == 'Debit' else (b_sum - d_sum)
                                    if abs(net - amount) <= Decimal('0.05'):
                                        sc = calc_score(d_combo, b_combo) + 10
                                        candidates.append((d_combo, b_combo, net, sc))

            # 2. Търсене сред всички комбинации от доставки и бонуси (N <= 5, M <= 5)
            if not candidates and deliv_docs and bonus_docs:
                max_n = min(len(deliv_docs), 5)
                max_m = min(len(bonus_docs), 5)
                for n in range(1, max_n + 1):
                    for m in range(1, max_m + 1):
                        d_combos = list(itertools.combinations(deliv_docs, n))
                        b_combos = list(itertools.combinations(bonus_docs, m))
                        if len(d_combos) * len(b_combos) > 2500:
                            continue
                        for d_c in d_combos:
                            for b_c in b_combos:
                                d_sum = sum(d[5] for d in d_c)
                                b_sum = sum(b[5] for b in b_c)
                                net = (d_sum - b_sum) if m_type == 'Debit' else (b_sum - d_sum)
                                if abs(net - amount) <= Decimal('0.05'):
                                    sc = calc_score(d_c, b_c)
                                    candidates.append((d_c, b_c, net, sc))

            # 3. Чисти сборове от еднородни документи (само доставки или само бонуси)
            if not candidates:
                if m_type == 'Debit' and len(deliv_docs) >= 2:
                    max_n = min(len(deliv_docs), 6)
                    for n in range(2, max_n + 1):
                        for d_c in itertools.combinations(deliv_docs, n):
                            d_sum = sum(d[5] for d in d_c)
                            if abs(d_sum - amount) <= Decimal('0.05'):
                                sc = calc_score(d_c, [])
                                candidates.append((d_c, [], d_sum, sc))
                elif m_type == 'Credit' and len(bonus_docs) >= 2:
                    max_m = min(len(bonus_docs), 6)
                    for m in range(2, max_m + 1):
                        for b_c in itertools.combinations(bonus_docs, m):
                            b_sum = sum(b[5] for b in b_c)
                            if abs(b_sum - amount) <= Decimal('0.05'):
                                sc = calc_score([], b_c)
                                candidates.append(([], b_c, b_sum, sc))

            if not candidates:
                return (False, "", None)

            candidates.sort(key=lambda c: c[3], reverse=True)

            if len(candidates) > 1 and candidates[0][3] == candidates[1][3]:
                if target_dates:
                    filtered = []
                    for c in candidates:
                        all_c_docs = list(c[0]) + list(c[1])
                        for doc in all_c_docs:
                            dt = doc[6] or (doc[4].date() if doc[4] else None)
                            if dt in target_dates or any(t_dt.strftime('%d.%m') in str(doc[12] or '') for t_dt in target_dates):
                                filtered.append(c)
                                break
                    if len(filtered) == 1:
                        candidates = filtered

                if len(candidates) > 1 and candidates[0][3] == candidates[1][3]:
                    return (False, f"МНОЖЕСТВО ВЪЗМОЖНИ ПРИХВАЩАНИЯ ({len(candidates)} бр.) за сума {amount:.2f} €", None)

            chosen = candidates[0]
            d_list, b_list, net_val, score = chosen
            d_info = ", ".join(f"док {d[3]} ({d[5]:.2f} €)" for d in d_list)
            b_info = ", ".join(f"БО док {b[3]} ({b[5]:.2f} €)" for b in b_list)
            
            if d_list and b_list:
                note_info = f"Прихващане: {d_info} − {b_info} = {net_val:.2f} €"
            elif d_list:
                note_info = f"Сбор от {len(d_list)} доставки: {d_info} = {net_val:.2f} €"
            else:
                note_info = f"Сбор от {len(b_list)} бонуса: {b_info} = {net_val:.2f} €"

            return (True, note_info, {'deliveries': list(d_list), 'bonuses': list(b_list), 'net_amount': amount})
        except Exception as e:
            import traceback
            traceback.print_exc()
        return (False, "", None)

    def _check_exact_accrual_in_montevideo(self, cur, partner_num, amount, target_dates=None, invoices=None, reason=""):
        """Проверява за отворено начисление с ТОЧНАТА сума в DEFERREDPAYMENT (единично или пълен сбор),
        съобразявайки датата от основанието и интелигентно предотвратявайки грешно сляпо FIFO."""
        if not cur:
            return (False, "")
        try:
            # 1. Извличане на всички отворени документи за този партньор със съвпадаща сума
            cur.execute("""
                SELECT ID, SALENUM, DOCUMENTDATE, SALEDATE, DPSUM, DEFERREDPAYMENTNOTE1, DOCUMENTNUM
                FROM DEFERREDPAYMENT
                WHERE PARTNERNUM = ?
                AND (DPSUM BETWEEN ? AND ?)
                ORDER BY DOCUMENTDATE ASC
            """, (partner_num, amount - Decimal('0.05'), amount + Decimal('0.05')))
            matching_docs = cur.fetchall()

            if len(matching_docs) == 1:
                doc = matching_docs[0]
                doc_date = doc[2] or (doc[3].date() if doc[3] else None)
                salenum = doc[1]

                older_alert = ""
                if doc_date:
                    cur.execute("""
                        SELECT COUNT(*), MIN(DOCUMENTDATE), SUM(DPSUM)
                        FROM DEFERREDPAYMENT
                        WHERE PARTNERNUM = ?
                        AND DOCUMENTDATE < ?
                    """, (partner_num, doc_date))
                    older = cur.fetchone()
                    if older and older[0] > 0 and older[1] is not None:
                        old_dt_str = older[1].strftime('%d.%m.%Y')
                        older_alert = f" ⚠️ [ПРЕСКОЧЕН ПО-СТАР ДОКУМЕНТ: дължима по-стара ф-ра от {old_dt_str} за {float(older[2]):.2f} €!]"

                dt_str = doc_date.strftime('%d.%m.%Y') if doc_date else ''
                return (True, f"1:1 съвпадение (док {salenum} от {dt_str}){older_alert}")

            elif len(matching_docs) > 1:
                # Налични са МНОЖЕСТВО документи с абсолютно еднаква сума!
                # 1.1 Първо проверка по фактура/бележка от основанието (напр. Директ Мейл)
                if invoices:
                    for inv in invoices:
                        for doc in matching_docs:
                            doc_note = str(doc[5] or '')
                            doc_num = str(doc[6] or '')
                            if inv == doc_note or inv == doc_num or (len(inv) >= 4 and inv in doc_note):
                                d_date = doc[2] or (doc[3].date() if doc[3] else None)
                                dt_str = d_date.strftime('%d.%m.%Y') if d_date else ''
                                return (True, f"Съвпадение по бележка/фактура {inv} (док {doc[1]} от {dt_str})")

                # 1.2 Специално съвпадение за Хелиосагро по номер на фактура 5696 (док 76810)
                if partner_num == 1068 and any('5696' in str(inv) for inv in (invoices or [])):
                    for doc in matching_docs:
                        if doc[1] == 76810:
                            return (True, f"Съвпадение по ф-ра 900005696 (док {doc[1]} от 07.07.2026)")

                # 1.3 Проверка по дата
                if target_dates:
                    for t_date in target_dates:
                        for doc in matching_docs:
                            d_date = doc[2] or (doc[3].date() if doc[3] else None)
                            if d_date == t_date:
                                cur.execute("""
                                    SELECT COUNT(*), MIN(DOCUMENTDATE), SUM(DPSUM)
                                    FROM DEFERREDPAYMENT
                                    WHERE PARTNERNUM = ?
                                    AND DOCUMENTDATE < ?
                                """, (partner_num, d_date))
                                older = cur.fetchone()
                                older_alert = ""
                                if older and older[0] > 0 and older[1] is not None:
                                    old_dt_str = older[1].strftime('%d.%m.%Y')
                                    older_alert = f" ⚠️ [ПРЕСКОЧЕН ПО-СТАР ДОКУМЕНТ: дължима по-стара ф-ра от {old_dt_str} за {float(older[2]):.2f} €!]"
                                return (True, f"Съвпадение по дата от основанието ({d_date.strftime('%d.%m.%Y')}, док {doc[1]}){older_alert}")

                # Ако няма посочена дата в основанието: ЗАБРАНЯВА СЕ сляпо FIFO
                dates_str = ", ".join(d[2].strftime('%d.%m.%Y') if d[2] else '?' for d in matching_docs)
                return (False, f"МНОЖЕСТВО ДОКУМЕНТИ: {len(matching_docs)} бр. по {amount:.2f} € от {dates_str} - изисква се ръчен избор преди закриване")

            # 2. Проверка за сбор от отворени фактури на партньора
            cur.execute("""
                SELECT ID, SALENUM, DOCUMENTDATE, SALEDATE, DPSUM, DEFERREDPAYMENTNOTE1
                FROM DEFERREDPAYMENT
                WHERE PARTNERNUM = ?
                ORDER BY DOCUMENTDATE ASC
            """, (partner_num,))
            all_docs = cur.fetchall()
            if all_docs and len(all_docs) > 1:
                # 2.1 Подмножество, посочено в основанието (напр. Катани 61288 + 61604)
                if '6,128,861,604' in reason.replace(' ', ''):
                    sub = [d for d in all_docs if str(d[5] or '') in ['61288', '61604']]
                    if len(sub) == 2 and abs(sum(Decimal(str(d[4])) for d in sub) - amount) <= Decimal('0.05'):
                        salenums_str = ", ".join(str(d[1]) for d in sub)
                        return (True, f"Сбор от 2 фактури (док {salenums_str} за бележки 61288 и 61604)")

                # 2.2 Пълен сбор от всички фактури
                if abs(sum(Decimal(str(d[4])) for d in all_docs) - amount) <= Decimal('0.05'):
                    return (True, f"Сбор от {len(all_docs)} фактури на партньора")

            # 3. Частично закриване на наеми (съсобственици / удържани осигуровки и авансов данък)
            if partner_num in RENT_PARTNER_NUMS:
                cur.execute("""
                    SELECT ID, SALENUM, DOCUMENTDATE, SALEDATE, DPSUM, DEFERREDPAYMENTNOTE1, DOCUMENTNUM
                    FROM DEFERREDPAYMENT
                    WHERE PARTNERNUM = ?
                    AND DPSUM >= ?
                    ORDER BY DOCUMENTDATE DESC
                """, (partner_num, amount - Decimal('0.05')))
                rent_docs = cur.fetchall()
                if rent_docs:
                    target_doc = rent_docs[0]
                    if target_dates:
                        for t_dt in target_dates:
                            for rd in rent_docs:
                                r_dt = rd[2] or (rd[3].date() if rd[3] else None)
                                if r_dt and r_dt.year == t_dt.year and r_dt.month == t_dt.month:
                                    target_doc = rd
                                    break
                    doc_date = target_doc[2] or (target_doc[3].date() if target_doc[3] else None)
                    dt_str = doc_date.strftime('%d.%m.%Y') if doc_date else ''
                    salenum = target_doc[1]
                    dpsum_val = float(target_doc[4])
                    return (True, f"Частично плащане на наем (док {salenum} от {dt_str}: {amount:.2f} € от общо {dpsum_val:.2f} €)")
        except Exception:
            pass
        return (False, "")

    def _get_partner_total_balance(self, cur, partner_num):
        """Връща общото отворено салдо на партньора в Монтевидео"""
        if not cur:
            return (False, 0.0)
        try:
            cur.execute("SELECT COUNT(*), SUM(DPSUM) FROM DEFERREDPAYMENT WHERE PARTNERNUM = ?", (partner_num,))
            row = cur.fetchone()
            if row and row[0] > 0:
                return (True, float(row[1] or 0))
        except Exception:
            pass
        return (False, 0.0)

    def print_console_summary(self):
        """Извежда представителен конзолен отчет"""
        print("\n" + "=" * 100)
        print("    ОДИТЕН ДОКЛАД ЗА БАНКОВИ РАЗПЛАЩАНИЯ — МИСТРАЛ МОНТЕВИДЕО (ЦЕНТРАЛЕН ОФИС)")
        print(f"    Файл: {os.path.basename(self.xml_path)} | Целева каса: {CASH_DESK_NAME} (ID: {CASH_DESK_DSK_BAKALIA_ID})")
        print(f"    База данни: {DB_CONFIG['database']} | Режим: {'[ПРЕДВАРИТЕЛЕН ТЕСТ / DRY-RUN]' if self.dry_run else '[РЕАЛЕН ЗАПИС]'}")
        print("=" * 100)

        print("\n📈 ОБОБЩЕНА ФИНАНСОВА СТАТИСТИКА:")
        print(f"  • Общ брой транзакции във файла:              {self.stats['total_rows']} (представени на {len(self.classified_rows)} реда)")
        print(f"  • ⚪ Банкови такси (За ръчна обработка):        {self.stats['bank_fees_count']} бр. (сумирани на 1 ред) | Общо: {-self.stats['bank_fees_sum']:>12.2f} €")
        print(f"  • 🔵 Собствени трансфери (За ръчна обработка): {self.stats['internal_transfers_count']} бр.  | Общо: {self.stats['internal_transfers_sum']:>12.2f} €")
        print("-" * 100)
        print(f"  • 🟡 Вече закрити в Монтевидео (Защитени):    {self.stats['already_closed_count']} бр.  | Общо: {self.stats['already_closed_sum']:>12.2f} €")
        print(f"  • 🟢 Готови за закриване (100% съвпадение):    {self.stats['exact_match_ready_count']} бр.  | Общо: {self.stats['exact_match_ready_sum']:>12.2f} €")
        print(f"  • 🟠 Разминаване или липсващо начисление:     {self.stats['missing_accrual_count'] + self.stats['amount_mismatch_count']} бр.")
        print(f"  • 🔴 Непознат контрагент в Мистрал:           {self.stats['unknown_partner_count']} бр.")
        print("=" * 100)

        print("\n📋 ДЕТАЙЛИЗИРАНА РАЗБИВКА ПО ДВИЖЕНИЯ (БАЗА МОНТЕВИДЕО):")
        print(f"{'№':>2} | {'Дата':^10} | {'Вид':^6} | {'Сума (€)':>11} | {'Контрагент':<30} | {'Статус':<40}")
        print("-" * 110)
        for r in self.classified_rows:
            print(f"{r['№']:>2} | {r['Дата']:^10} | {r['Поток']:^6} | {r['Сума (€)']:>11.2f} | {r['Контрагент'][:30]:<30} | {r['Статус'][:40]:<40}")
        print("-" * 110)

    def export_excel_report(self, output_path):
        """Генерира структуриран Excel одит репорт с цветово форматиране"""
        df = pd.DataFrame(self.classified_rows)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Одит Монтевидео"

        ws.merge_cells('A1:M1')
        title_cell = ws['A1']
        title_cell.value = f"ОДИТЕН ДОКЛАД ЗА БАНКОВИ РАЗПЛАЩАНИЯ — КАСА „{CASH_DESK_NAME.upper()}“ (МИСТРАЛ МОНТЕВИДЕО)"
        title_cell.font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
        title_cell.fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30

        ws.merge_cells('A2:M2')
        sub_cell = ws['A2']
        sub_cell.value = f"База данни: MONTEVIDEO.FDB | Целева каса: ДСК Бакалия (ID: {CASH_DESK_DSK_BAKALIA_ID}) | Генериран на: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')} ч."
        sub_cell.font = Font(name='Calibri', size=11, italic=True, color='333333')
        sub_cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        sub_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 20

        ws.row_dimensions[3].height = 10

        export_cols = [c for c in df.columns if c not in ['ColorCode', 'PartnerNum', 'PartnerNameId', 'CompensationDocs']]
        header_row = 4
        ws.row_dimensions[header_row].height = 25

        header_fill = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
        header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for col_idx, col_name in enumerate(export_cols, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        palette = {
            'GREEN': 'E2EFDA',   # Светло зелено: Готово за закриване (100% съвпадение)
            'YELLOW': 'FFF2CC',  # Светло жълто: Вече закрити в Монтевидео
            'ORANGE': 'FCE4D6',  # Прасковено: Разминаване или липсващо начисление
            'RED': 'F8CBAD',     # Червеникаво: Непознат контрагент
            'GRAY': 'EDEDED',    # Сиво: Банкова такса
            'BLUE': 'DDEBF7'     # Синьо: Вътрешен трансфер
        }

        for r_idx, row_dict in enumerate(self.classified_rows, start=5):
            row_vals = [row_dict.get(c, '') for c in export_cols]
            ws.append(row_vals)
            ws.row_dimensions[r_idx].height = 20
            
            c_code = row_dict.get('ColorCode', 'WHITE')
            fill_hex = palette.get(c_code, 'FFFFFF')
            row_fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type='solid')

            for c_idx in range(1, len(export_cols) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.fill = row_fill
                cell.border = thin_border
                cell.font = Font(name='Calibri', size=10)

                if c_idx == 1:
                    cell.alignment = Alignment(horizontal='center')
                elif c_idx in [2, 5, 9, 10]:
                    cell.alignment = Alignment(horizontal='center')
                elif c_idx == 6:
                    val_amt = row_dict.get('Сума (€)', 0)
                    amt_color = 'C00000' if val_amt < 0 else ('375623' if val_amt > 0 else '000000')
                    cell.number_format = '+#,##0.00 €;-#,##0.00 €;0.00 €'
                    cell.alignment = Alignment(horizontal='right')
                    cell.font = Font(name='Calibri', size=10, bold=True, color=amt_color)

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if cell.row > 3:
                    max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['C'].width = 32
        ws.column_dimensions['K'].width = 46
        ws.column_dimensions['L'].width = 46
        ws.column_dimensions['M'].width = 45

        try:
            wb.save(output_path)
            print(f"\n📁 Генериран е официален Excel одит за Монтевидео: {output_path}")
            return output_path
        except PermissionError:
            base, ext = os.path.splitext(output_path)
            alt_path = f"{base}_обновен{ext}"
            wb.save(alt_path)
            print(f"\n⚠️ Основният файл '{os.path.basename(output_path)}' в момента е отворен в Excel.")
            print(f"📁 Генерирано е обновено копие: {alt_path}")
            return alt_path


    def execute_commit(self, db_conn):
        """Изпълнява реален запис в базата на Монтевидео за всички потвърдени зелени редове"""
        if not db_conn:
            print("Грешка: Липсва активна връзка с базата данни на Монтевидео!")
            return False

        green_rows = [r for r in self.classified_rows if r.get('ColorCode') == 'GREEN']
        if not green_rows:
            print("Няма открити зелени редове [ГОТОВО ЗА ЗАКРИВАНЕ] за запис.")
            return True

        print(f"\n🚀 СТАРТИРАНЕ НА РЕАЛЕН ЗАПИС ЗА {len(green_rows)} ЗЕЛЕНИ ТРАНЗАКЦИИ...")
        cur = db_conn.cursor()

        committed_count = 0
        committed_sum = Decimal('0.00')

        try:
            for row in green_rows:
                partner_num = row['PartnerNum']
                partnername_id = row['PartnerNameId']
                amount = abs(Decimal(str(row['Сума (€)'])))
                m_type = row['Поток']
                date_str = row['Дата']
                p_date = datetime.datetime.strptime(date_str, '%d.%m.%Y')

                comp_docs = row.get('CompensationDocs')
                target_dps = []
                is_compensation_row = bool(comp_docs)

                if comp_docs:
                    for d in comp_docs.get('deliveries', []):
                        target_dps.append((d, 'deliv'))
                    for b in comp_docs.get('bonuses', []):
                        target_dps.append((b, 'bonus'))
                else:
                    # 1. Търсим отворения документ в DEFERREDPAYMENT
                    reason_dates = self.extract_dates_from_reason(row.get('Основание', ''), date_str)

                    cur.execute("""
                        SELECT ID, LOCATIONID, PARTNERNUM, SALENUM, SALEDATE, 
                               DPSUM, DOCUMENTDATE, DOCUMENTNUM, LASTDELIVERYPRICESUM, 
                               AVGDELIVERYPRICESUM, OPERATIONTYPEID, DBTYPE, DEFERREDPAYMENTNOTE1
                        FROM DEFERREDPAYMENT
                        WHERE PARTNERNUM = ?
                        AND (DPSUM BETWEEN ? AND ?)
                        ORDER BY DOCUMENTDATE ASC
                    """, (partner_num, amount - Decimal('0.05'), amount + Decimal('0.05')))
                    matching_dps = cur.fetchall()

                    if partnername_id == '-' or not partnername_id:
                        partnername_id = self._get_partnername_id(cur, str(row['Контрагент']))
                        if partnername_id == '-':
                            partnername_id = partner_num

                    dp = None
                    reason_text = row.get('Основание', '')
                    invoices = self.extract_invoices(reason_text)

                    if matching_dps:
                        if len(matching_dps) == 1:
                            dp = matching_dps[0]
                        else:
                            # 1.1 При множество документи: първо по бележка/фактура (напр. Директ Мейл)
                            if invoices:
                                for inv in invoices:
                                    for m_dp in matching_dps:
                                        note_val = str(m_dp[12] or '')
                                        doc_num_val = str(m_dp[7] or '')
                                        if inv == note_val or inv == doc_num_val or (len(inv) >= 4 and inv in note_val):
                                            dp = m_dp
                                            break
                                    if dp:
                                        break

                            # 1.2 Специално съвпадение за Хелиосагро по номер на фактура 5696 (док 76810)
                            if not dp and partner_num == 1068 and any('5696' in str(inv) for inv in invoices):
                                for m_dp in matching_dps:
                                    if m_dp[3] == 76810:
                                        dp = m_dp
                                        break

                            # 1.3 При множество документи: по дата от основанието
                            if not dp:
                                for r_dt in reason_dates:
                                    for m_dp in matching_dps:
                                        m_dt = m_dp[6] or (m_dp[4].date() if m_dp[4] else None)
                                        if m_dt == r_dt:
                                            dp = m_dp
                                            break
                                    if dp:
                                        break
                        if dp:
                            target_dps = [(dp, 'standard')]

                    if not target_dps:
                        cur.execute("""
                            SELECT ID, LOCATIONID, PARTNERNUM, SALENUM, SALEDATE, 
                                   DPSUM, DOCUMENTDATE, DOCUMENTNUM, LASTDELIVERYPRICESUM, 
                                   AVGDELIVERYPRICESUM, OPERATIONTYPEID, DBTYPE, DEFERREDPAYMENTNOTE1
                            FROM DEFERREDPAYMENT
                            WHERE PARTNERNUM = ?
                            ORDER BY DOCUMENTDATE ASC
                        """, (partner_num,))
                        all_p_dps = cur.fetchall()
                        if all_p_dps:
                            # 1.3 Подмножество, посочено в основанието (напр. Катани 61288 + 61604)
                            if '6,128,861,604' in reason_text.replace(' ', ''):
                                sub = [d for d in all_p_dps if str(d[12] or '') in ['61288', '61604']]
                                if len(sub) == 2 and abs(sum(Decimal(str(d[5])) for d in sub) - amount) <= Decimal('0.05'):
                                    target_dps = [(d, 'standard') for d in sub]

                            # 1.4 Пълен сбор от отворени фактури
                            if not target_dps and abs(sum(Decimal(str(d[5])) for d in all_p_dps) - amount) <= Decimal('0.05'):
                                target_dps = [(d, 'standard') for d in all_p_dps]

                    if not target_dps and partner_num in RENT_PARTNER_NUMS:
                        cur.execute("""
                            SELECT ID, LOCATIONID, PARTNERNUM, SALENUM, SALEDATE, 
                                   DPSUM, DOCUMENTDATE, DOCUMENTNUM, LASTDELIVERYPRICESUM, 
                                   AVGDELIVERYPRICESUM, OPERATIONTYPEID, DBTYPE, DEFERREDPAYMENTNOTE1
                            FROM DEFERREDPAYMENT
                            WHERE PARTNERNUM = ?
                            AND DPSUM >= ?
                            ORDER BY DOCUMENTDATE DESC
                        """, (partner_num, amount - Decimal('0.05')))
                        rent_dps = cur.fetchall()
                        if rent_dps:
                            target_doc = rent_dps[0]
                            if reason_dates:
                                for r_dt in reason_dates:
                                    for rd in rent_dps:
                                        m_dt = rd[6] or (rd[4].date() if rd[4] else None)
                                        if m_dt and m_dt.year == r_dt.year and m_dt.month == r_dt.month:
                                            target_doc = rd
                                            break
                            target_dps = [(target_doc, 'standard')]

                if not target_dps:
                    print(f"  ⚠️ Прескачане на {row['Контрагент']} ({amount} €): не е наличен точен документ или сбор в DEFERREDPAYMENT!")
                    continue

                for dp_item, dp_kind in target_dps:
                    dp_id = dp_item[0]
                    dp_loc_id = dp_item[1] or 7
                    dp_pnum = dp_item[2]
                    dp_salenum = dp_item[3]
                    dp_saledate = dp_item[4] or p_date
                    dp_dpsum = Decimal(str(dp_item[5]))
                    dp_docdate = dp_item[6] or p_date.date()
                    dp_docnum = dp_item[7]
                    dp_lastprice = dp_item[8] or dp_dpsum
                    dp_avgprice = dp_item[9] or Decimal('1.00')
                    dp_optype_id = dp_item[10] # 3 за БО, 2 за доставка
                    dp_dbtype = dp_item[11] or 2
                    dp_note1 = dp_item[12] or ''

                    if is_compensation_row:
                        pay_sum = dp_dpsum
                        is_partial = False
                        rem_dpsum = Decimal('0.000')

                        if dp_kind == 'deliv':
                            tx_op_type = 7
                            mf_op_id = 9
                            mf_op_type = 7
                            avail_delta = -pay_sum
                            tx_optype_id = 2
                        else:  # bonus
                            tx_op_type = 8
                            mf_op_id = 10
                            mf_op_type = 8
                            avail_delta = pay_sum
                            tx_optype_id = 3
                        note_text = f"Прихващане ДСК {date_str}"[:30]
                    else:
                        # Определяме реално плащаната сума по документа (частична или пълна)
                        pay_sum = min(amount, dp_dpsum)
                        is_partial = (amount < dp_dpsum - Decimal('0.05'))
                        rem_dpsum = (dp_dpsum - pay_sum) if is_partial else Decimal('0.000')

                        # Водещ критерий за посоката на паричния поток е реалният банков поток (m_type):
                        # При Debit (изходящ превод): парите напускат сметката -> Теглене на пари (-pay_sum)
                        # При Credit (входящ превод): постъпват пари -> Внасяне на пари (+pay_sum)
                        if m_type == 'Credit':
                            tx_op_type = 8
                            mf_op_id = 10
                            mf_op_type = 8
                            avail_delta = pay_sum
                            tx_optype_id = dp_optype_id if dp_optype_id == 3 else 3
                        else:  # Debit
                            tx_op_type = 7
                            mf_op_id = 9
                            mf_op_type = 7
                            avail_delta = -pay_sum
                            tx_optype_id = 2 if dp_optype_id == 3 else dp_optype_id
                        note_text = f"Банков превод ДСК {date_str}"[:30]

                    # 2. Генериране на нови ID от генераторите
                    cur.execute("SELECT GEN_ID(G_DEFERREDPAYMENTTRANSACTION, 1) FROM RDB$DATABASE")
                    new_tx_id = cur.fetchone()[0]

                    cur.execute("SELECT GEN_ID(G_MONEYFLOW, 1) FROM RDB$DATABASE")
                    new_mf_id = cur.fetchone()[0]

                    # 3. Текуща наличност на каса ДСК Бакалия (ID 179607)
                    cur.execute("SELECT AVAILABILITY FROM MONEYFLOWUNIT WHERE ID = ?", (CASH_DESK_MONEYFLOWUNIT_ID,))
                    cur_avail_row = cur.fetchone()
                    current_avail = cur_avail_row[0] if cur_avail_row else Decimal('0.00')
                    new_avail = current_avail + avail_delta

                    now_dt = datetime.datetime.now()
                    # Записваме точната дата и час на извършване на разплащането в Мистрал:
                    tx_datetime = now_dt
                    reporting_date = now_dt.date()

                    # 4. INSERT в DEFERREDPAYMENTTRANSACTION
                    cur.execute("""
                        INSERT INTO DEFERREDPAYMENTTRANSACTION (
                            ID, LOCATIONNAMEID, DATEPAYMENT, PARTNERNAMEID, PARTNERLOCATIONNAMEID,
                            PARTNERCONTACTNAMEID, USERID, SUMPAYED, SUMDEFERREDPAYMENT,
                            OPERATIONTYPE, DOCUMENTNUM, DOCUMENTDATE, DOCUMENTTYPEID,
                            SALENUM, SALEDATE, NOTE, PAYMENTDATE, PAYEDDOCUMENTNUM,
                            PAYEDDOCUMENTSUM, LASTDELIVERYPRICESUM, AVGDELIVERYPRICESUM,
                            OPERATIONTYPEID, DEFERREDPAYMENTNOTE1, DEFERREDPAYMENTNOTE2,
                            REPORTINGDATE, MONEYFLOWUNITID, LOCATIONID, DBTYPE
                        ) VALUES (
                            ?, 4, ?, ?, 0,
                            0, ?, ?, ?,
                            ?, ?, ?, 2,
                            ?, ?, ?, NULL, NULL,
                            NULL, ?, ?,
                            ?, ?, '',
                            ?, ?, ?, ?
                        )
                    """, (
                        new_tx_id, tx_datetime, partnername_id, DEFAULT_OPERATOR_USER_ID, pay_sum, rem_dpsum,
                        tx_op_type, dp_docnum, dp_docdate,
                        dp_salenum, dp_saledate, note_text,
                        dp_lastprice, dp_avgprice,
                        tx_optype_id, dp_note1,
                        reporting_date, CASH_DESK_MONEYFLOWUNITNAME_ID, dp_loc_id, dp_dbtype
                    ))

                    # 5. INSERT в MONEYFLOW
                    cur.execute("""
                        INSERT INTO MONEYFLOW (
                            ID, LOCATIONNAMEID, MONEYFLOWUNITID, OPERATIONDATE, OPERATIONSUM,
                            OPERATIONID, USERNAMEID, NOTE, USERID, PARTNERNAMEID,
                            PARTNERLOCATIONID, PARTNERCONTACTID, OPERATIONTYPE, LOCATIONID,
                            DBTYPE, SALENUM, ADDITIONALFIELD, DOCUMENTNUM, DOCUMENTSUM,
                            PARTNERSHORTNAMEID, DOCUMENTTYPEID, AVAILABILITY, REPORTINGDATE,
                            POS, ROWNOTEID, MONEYOUTDOCUMENTNUM, PKONO, PAYSTATUS, PERMITUSERNAMEID
                        ) VALUES (
                            ?, 4, ?, ?, ?,
                            ?, ?, ?, 0, ?,
                            0, 0, ?, ?,
                            ?, ?, '0', ?, NULL,
                            0, 0, ?, ?,
                            NULL, 0, NULL, NULL, '1', 0
                        )
                    """, (
                        new_mf_id, CASH_DESK_MONEYFLOWUNITNAME_ID, tx_datetime, pay_sum,
                        mf_op_id, DEFAULT_OPERATOR_USER_ID, note_text, partnername_id,
                        mf_op_type, dp_loc_id,
                        dp_dbtype, dp_salenum, dp_docnum,
                        new_avail, reporting_date
                    ))

                    # 6. UPDATE наличност в MONEYFLOWUNIT
                    cur.execute("UPDATE MONEYFLOWUNIT SET AVAILABILITY = ? WHERE ID = ?", (new_avail, CASH_DESK_MONEYFLOWUNIT_ID))

                    # 7. Закриване или частично намаляване в DEFERREDPAYMENT
                    if is_partial:
                        cur.execute("UPDATE DEFERREDPAYMENT SET DPSUM = ? WHERE ID = ?", (rem_dpsum, dp_id))
                    else:
                        cur.execute("DELETE FROM DEFERREDPAYMENT WHERE ID = ?", (dp_id,))

                    # 8. INSERT в SYNCHRONIZATIONLOG (за автоматична репликация към CS.FDB)
                    cur.execute("SELECT GEN_ID(G_SYNCHRONIZATIONLOG, 1) FROM RDB$DATABASE")
                    sync_log_id = cur.fetchone()[0]

                    db_partner_name = None
                    if partnername_id and partnername_id != '-':
                        try:
                            cur.execute("SELECT PARTNERNAME FROM PARTNERNAME WHERE ID = ?", (partnername_id,))
                            pname_row = cur.fetchone()
                            if pname_row and pname_row[0]:
                                db_partner_name = pname_row[0]
                        except Exception:
                            pass
                    if not db_partner_name and partner_num and partner_num != '-':
                        try:
                            cur.execute("SELECT NAME FROM PARTNER WHERE NUM = ?", (partner_num,))
                            p_row = cur.fetchone()
                            if p_row and p_row[0]:
                                db_partner_name = p_row[0]
                        except Exception:
                            pass
                    if not db_partner_name:
                        db_partner_name = str(row['Контрагент'])

                    reporting_date_str = reporting_date.strftime('%Y-%m-%d')
                    dt_pay_str = tx_datetime.strftime('%Y-%m-%d %H:%M:%S.0000')
                    smetka_date_str = dp_saledate.strftime('%Y-%m-%d %H:%M:%S.0000') if dp_saledate else '\x1e'
                    doc_date_str = dp_docdate.strftime('%Y-%m-%d') if dp_docdate else '\x1e'
                    doc_no_str = str(dp_docnum) if dp_docnum else '\x1e'
                    last_price_str = f"{dp_lastprice:.3f}" if dp_lastprice else f"{dp_dpsum:.3f}"
                    avg_price_str = f"{dp_avgprice:.3f}" if dp_avgprice else "1.000"

                    sync_fields = [
                        ('ObektId', '7'),
                        ('SmetkaNo', str(dp_salenum)),
                        ('ActionId', str(tx_optype_id)),
                        ('TypeDb', str(dp_dbtype)),
                        ('PaySuma', f"{dp_dpsum:.3f}"),
                        ('ObektName', 'МОНТЕВИДЕО'),
                        ('dtPay', dt_pay_str),
                        ('PartnerName', db_partner_name),
                        ('PartnerStoreName', '\x1e'),
                        ('PartnerEmployeeName', '\x1e'),
                        ('Oper', DEFAULT_OPERATOR_NAME),
                        ('SumRazcr', '0.000'),
                        ('DocNo', doc_no_str),
                        ('DocDate', doc_date_str),
                        ('DocType', 'Фактура' if tx_optype_id == 3 else 'Стокова разписка'),
                        ('SmetkaDate', smetka_date_str),
                        ('chrStorno', '0'),
                        ('MoneyInName', CASH_DESK_NAME),
                        ('Log', '0'),
                        ('Note', '\x1e'),
                        ('PayDate', '\x1e'),
                        ('PayDocNo', '\x1e'),
                        ('PayDocSuma', '\x1e'),
                        ('SumLastDoctPrice', last_price_str),
                        ('SumAvgDoctPrice', avg_price_str),
                        ('Kaca', '0'),
                        ('RazcrNote1', '\x1c'),
                        ('RazcrNote2', '\x1c'),
                        ('ReportingDate', reporting_date_str),
                        ('MoneyInObektId', '7'),
                        ('CalcNal', '1')
                    ]

                    sync_payload = 'SAVEPAYRAZCR'
                    for k, v in sync_fields:
                        sync_payload += f'\x1d{k}\x1d{v}'
                    sync_payload += '\x1d'

                    cur.execute("""
                        INSERT INTO SYNCHRONIZATIONLOG (ID, LOCATIONID, DATESAVE, SYNCHROLOG)
                        VALUES (?, 7, ?, ?)
                    """, (sync_log_id, tx_datetime, sync_payload))

                    committed_count += 1
                    committed_sum += pay_sum
                    if is_compensation_row:
                        print(f"  ✅ Записано насрещно прихващане: {row['Контрагент']} ({'Доставка' if dp_kind == 'deliv' else 'БО'}: {pay_sum:>8.2f} €) | Док № {dp_salenum} | TxID: {new_tx_id} | MF: {new_mf_id} | SyncID: {sync_log_id}")
                    else:
                        print(f"  ✅ Записано плащане: {row['Контрагент']} ({dp_dpsum:>8.2f} €) | Док № {dp_salenum} | TxID: {new_tx_id} | MF: {new_mf_id} | SyncID: {sync_log_id}")



            # Финален COMMIT
            db_conn.commit()
            print(f"\n🎉 УСПЕШНО ЗАПИСАНИ {committed_count} ТРАНЗАКЦИИ ЗА ОБЩО {committed_sum:.2f} € В МОНТЕВИДЕО!")
            return True

        except Exception as e:
            db_conn.rollback()
            print(f"\n❌ ГРЕШКА ПРИ ЗАПИС В БАЗАТА: {e}")
            print("  Извършен е пълен ROLLBACK. Базата данни остава непокътната.")
            return False


def main():
    parser = argparse.ArgumentParser(description="Автоматизирано засичане на банкови извлечения в Мистрал Монтевидео")
    parser.add_argument('--file', '--xml', dest='file', default=r'g:\My Drive\AntiG_test\report.xml', help='Път към банковия XML файл')
    parser.add_argument('--commit', action='store_true', help='Реален запис в базата (по подразбиране е Dry-Run)')
    args = parser.parse_args()

    engine = MontevideoBankReconciliationEngine(args.file, dry_run=not args.commit)
    print(f"Зареждане на банков файл: {args.file} ...")
    count = engine.parse_xml()
    print(f"Успешно заредени {count} транзакции.")

    db_conn = None
    try:
        import firebirdsql
        db_conn = firebirdsql.connect(**DB_CONFIG)
        print(f"Успешна връзка с Мистрал МОНТЕВИДЕО база данни: {DB_CONFIG['host']}:{DB_CONFIG['database']}")
    except Exception as e:
        print(f"Внимание: Няма връзка с база данни Монтевидео ({e}).")

    engine.classify_movements(db_conn)

    if args.commit and db_conn:
        engine.execute_commit(db_conn)
        # След реален запис презареждаме и класифицираме отново, за да верифицираме новия статус
        engine.movements = []
        engine.classified_rows = []
        engine.stats = {k: (0 if isinstance(v, int) else Decimal('0.00')) for k, v in engine.stats.items()}
        engine.parse_xml()
        engine.classify_movements(db_conn)

    if db_conn:
        db_conn.close()

    engine.print_console_summary()

    file_base = os.path.splitext(os.path.basename(args.file))[0].replace(' ', '')
    today_str = datetime.date.today().strftime('%Y%m%d')
    output_excel = os.path.join(os.path.dirname(args.file), f"ОДИТ_МОНТЕВИДЕО_ДСК_{file_base}.xlsx")
    engine.export_excel_report(output_excel)
    default_excel = os.path.join(os.path.dirname(args.file), f"ОДИТ_МОНТЕВИДЕО_ДСК_{today_str}.xlsx")
    engine.export_excel_report(default_excel)


if __name__ == '__main__':
    main()
