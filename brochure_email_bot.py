# -*- coding: utf-8 -*-
r"""
================================================================================
Верига супермаркети „Бакалия“ — Имейл Робот за Промоционални Брошури 2027
================================================================================
Функционалности:
  1. Разпращане на покани с прикрепена бланка (X-21 дни преди брошурата)
  2. Автоматично прихващане и сваляне на попълнени оферти (.xlsx)
  3. Търговски одит и интеграция с трислойния мапинг мениджър
  4. Генериране на списък за прозвъняване на закъснели доставчици (X-12 дни)
"""

import os
import sys
import re
import json
import imaplib
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
import pandas as pd
import openpyxl

from brochure_mapping import BrochureMappingManager
from brochure_rules import audit_brochure_item

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "brochure_email_config.json")
CONTACTS_FILE = os.path.join(os.path.dirname(__file__), "suppliers_brochure_contacts.xlsx")
INCOMING_DIR = os.path.join(os.path.dirname(__file__), "incoming_brochures")
BLANK_TEMPLATE = os.path.join(os.path.dirname(__file__), "БЛАНКА_ДОСТАВЧИЦИ_2027.xlsx")


def get_default_config():
    return {
        "email_address": "",
        "app_password": "",
        "imap_server": "imap.gmail.com",
        "imap_port": 993,
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 587,
        "sender_name": "Търговски отдел — Супермаркети „Бакалия“",
        "auto_mark_read": True
    }


class BrochureEmailBot:
    def __init__(self, config_file=None, contacts_file=None):
        self.config_file = config_file or CONFIG_FILE
        self.contacts_file = contacts_file or CONTACTS_FILE
        self.config = self.load_config()
        self.mapping_mgr = BrochureMappingManager()
        os.makedirs(INCOMING_DIR, exist_ok=True)

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    base = get_default_config()
                    base.update(cfg)
                    return base
            except Exception:
                pass
        return get_default_config()

    def save_config(self, new_cfg):
        self.config.update(new_cfg)
        with open(self.config_file, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def load_contacts(self):
        """Зарежда регистъра с контакти на доставчиците."""
        if not os.path.exists(self.contacts_file):
            return pd.DataFrame()
        try:
            return pd.read_excel(self.contacts_file)
        except Exception:
            return pd.DataFrame()

    def save_contacts(self, df):
        """Записва обновения списък с контакти."""
        df.to_excel(self.contacts_file, index=False)

    # --------------------------------------------------------------------------
    # 1. РАЗПРАЩАНЕ НА ПОКАНИ ЗА БРОШУРА
    # --------------------------------------------------------------------------
    def generate_invitation_html(self, supplier_name, brochure_code, store_period, zkp_period, deadline_date):
        """Генерира елегантен HTML шаблон на поканата към доставчика."""
        return f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #1E293B; line-height: 1.6; }}
                .container {{ max-width: 650px; margin: 0 auto; border: 1px solid #E2E8F0; border-radius: 8px; overflow: hidden; }}
                .header {{ background-color: #1B365D; color: #FFFFFF; padding: 20px; text-align: center; }}
                .content {{ padding: 25px; }}
                .highlight-box {{ background-color: #F8FAFC; border-left: 4px solid #2563EB; padding: 15px; margin: 20px 0; border-radius: 4px; }}
                .alert-box {{ background-color: #FEF3C7; border: 1px solid #F59E0B; padding: 12px; margin: 15px 0; border-radius: 6px; font-weight: bold; color: #92400E; }}
                .footer {{ background-color: #F1F5F9; padding: 15px; font-size: 12px; color: #64748B; text-align: center; }}
                .btn {{ background-color: #2563EB; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; font-weight: bold; margin-top: 10px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h2 style="margin:0;">🛒 ВЕРИГА СУПЕРМАРКЕТИ „БАКАЛИЯ“</h2>
                    <p style="margin:5px 0 0 0; opacity:0.9;">Покана за промоционална брошура — {brochure_code}</p>
                </div>
                <div class="content">
                    <p>Уважаеми партньори от <b>{supplier_name}</b>,</p>
                    <p>Стартираме кампанията по набиране на артикули за предстоящата промоционална брошура на верига „Бакалия“ (11 търговски обекта + Централен склад).</p>
                    
                    <div class="highlight-box">
                        <h4 style="margin:0 0 10px 0; color:#1B365D;">📅 ПАРАМЕТРИ НА КАМПАНИЯТА:</h4>
                        • <b>Издание:</b> {brochure_code}<br>
                        • <b>Период в магазините:</b> {store_period}<br>
                        • <b>Период намалена доставна цена (ЗКП):</b> <span style="color:#D97706; font-weight:bold;">{zkp_period}</span> (влиза в сила 7 дни преди старта)<br>
                        • <b>Краен срок за връщане на оферти:</b> <span style="color:#DC2626; font-weight:bold;">{deadline_date} г. до 17:00 ч.</span>
                    </div>

                    <div class="alert-box">
                        ⚠️ ВАЖНО: Всички цени (доставни и препоръчителни продажни) в приложената бланка се попълват С ВКЛЮЧЕН ДДС в ЕВРО (€)!
                    </div>

                    <p>Моля попълнете прикачената стандартизирана Excel бланка и ни я върнете в отговор на този имейл в указания срок.</p>
                    <p>Благодарим за успешното търговско партньорство!</p>
                    
                    <p style="margin-top:25px;">
                        С уважение,<br>
                        <b>Търговски отдел</b><br>
                        Верига супермаркети „Бакалия“<br>
                        Имейл: {self.config.get('email_address') or 'broshuri@bakaliya.bg'}
                    </p>
                </div>
                <div class="footer">
                    Това съобщение е автоматично генерирано от системата за управление на брошури на „Бакалия“.
                </div>
            </div>
        </body>
        </html>
        """

    def send_invitation_email(self, to_email, supplier_name, brochure_code, store_period, zkp_period, deadline_date, test_mode=False):
        """Изпраща имейл с покана и прикрепена празна бланка."""
        if not to_email or not isinstance(to_email, str) or '@' not in to_email:
            return False, "Невалиден имейл адрес"

        # Стриктна блокировка: ако allow_sending не е True, не се изпращат реални имейли!
        if not self.config.get("allow_sending", False) or test_mode or not self.config.get("email_address") or not self.config.get("app_password"):
            return True, f"[🔒 ЗАЩИТЕН ТЕСТОВ РЕЖИМ] Поканата към {supplier_name} ({to_email}) е подготвена (реалното изпращане е заключено)."

        try:
            msg = MIMEMultipart()
            msg['From'] = f"{self.config.get('sender_name', 'Бакалия')} <{self.config['email_address']}>"
            msg['To'] = to_email
            msg['Subject'] = f"Покана за Промоционална брошура {brochure_code} — Супермаркети „Бакалия“"

            html_body = self.generate_invitation_html(supplier_name, brochure_code, store_period, zkp_period, deadline_date)
            msg.attach(MIMEText(html_body, 'html', 'utf-8'))

            # Прикачане на файла БЛАНКА_ДОСТАВЧИЦИ_2027.xlsx
            if os.path.exists(BLANK_TEMPLATE):
                with open(BLANK_TEMPLATE, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="Бланка_Бакалия_{brochure_code.replace(" ", "_")}.xlsx"')
                msg.attach(part)

            # SMTP изпращане
            server = smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port'])
            server.starttls()
            server.login(self.config['email_address'], self.config['app_password'])
            server.sendmail(self.config['email_address'], [to_email], msg.as_string())
            server.quit()
            return True, f"Успешно изпратен към {to_email}"
        except Exception as e:
            return False, f"Грешка при изпращане: {str(e)}"

    # --------------------------------------------------------------------------
    # 2. ПРИХВАЩАНЕ И СВАЛЯНЕ НА ВЪРНАТИ БЛАНКИ (IMAP HARVESTER)
    # --------------------------------------------------------------------------
    def fetch_incoming_brochures(self, test_mode=False):
        """
        Свързва се с новата поща през IMAP, намира нечетените имейли с прикачени
        файлове .xlsx, сваля ги и ги регистрира за обработка.
        """
        if test_mode or not self.config.get("email_address") or not self.config.get("app_password"):
            # Симулиран преглед на вече свалени файлове в папката incoming_brochures
            return self._scan_local_incoming_folder()

        downloaded_files = []
        try:
            mail = imaplib.IMAP4_SSL(self.config['imap_server'], self.config['imap_port'])
            mail.login(self.config['email_address'], self.config['app_password'])
            mail.select("INBOX")

            # Търсим нечетени имейли
            status, messages = mail.search(None, 'UNSEEN')
            if status != "OK" or not messages[0]:
                mail.logout()
                return []

            email_ids = messages[0].split()
            for e_id in email_ids:
                res, msg_data = mail.fetch(e_id, '(RFC822)')
                if res != "OK":
                    continue

                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        raw_msg = email.message_from_bytes(response_part[1])
                        subject = self._decode_header(raw_msg["Subject"])
                        sender = self._decode_header(raw_msg["From"])
                        date_str = raw_msg["Date"]

                        # Проверка за прикачени файлове
                        for part in raw_msg.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            if part.get('Content-Disposition') is None:
                                continue

                            filename = part.get_filename()
                            if filename:
                                clean_name = self._decode_header(filename)
                                if clean_name.lower().endswith(('.xlsx', '.xls')):
                                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                                    save_path = os.path.join(INCOMING_DIR, f"{timestamp}_{clean_name}")
                                
                                with open(save_path, "wb") as f_out:
                                    f_out.write(part.get_payload(decode=True))

                                downloaded_files.append({
                                    "file_path": save_path,
                                    "file_name": clean_name,
                                    "sender": sender,
                                    "subject": subject,
                                    "received_at": date_str
                                })

            mail.logout()
            return downloaded_files
        except Exception as e:
            return [{"error": str(e)}]

    def _decode_header(self, text):
        if not text:
            return ""
        decoded_list = email.header.decode_header(text)
        res = ""
        for bytes_str, charset in decoded_list:
            if isinstance(bytes_str, bytes):
                res += bytes_str.decode(charset or 'utf-8', errors='replace')
            else:
                res += str(bytes_str)
        return res

    def _scan_local_incoming_folder(self):
        """Сканира локалната папка с входящи оферти."""
        files = []
        if os.path.exists(INCOMING_DIR):
            for fname in os.listdir(INCOMING_DIR):
                if fname.lower().endswith(('.xlsx', '.xls')):
                    files.append({
                        "file_path": os.path.join(INCOMING_DIR, fname),
                        "file_name": fname,
                        "sender": "Локален импорт / Тест",
                        "subject": "Входяща оферта",
                        "received_at": datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
                    })
        return files

    # --------------------------------------------------------------------------
    # 3. АВТОМАТИЧНО РАЗЧИТАНЕ И ОДИТ НА СВАЛЕНА БЛАНКА
    # --------------------------------------------------------------------------
    def process_supplier_blank_file(self, file_path):
        """
        Прочита сведена бланка (универсален парсер за 2027 и Legacy формати),
        напасва артикулите през трислойния мапинг и връща структуриран резултат.
        """
        return self.mapping_mgr.parse_any_supplier_blank(file_path)
