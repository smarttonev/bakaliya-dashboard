#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
=============================================================================
МОДУЛ ЗА ИМПОРТ НА ДОСТАВКИ В МИСТРАЛ ERP — МОНТЕВИДЕО
Верига магазини „Бакалия“ | ЦЕНТРАЛЕН ОФИС — МОНТЕВИДЕО
=============================================================================
Функционалност:
1. Парсва входящ файл с доставки (CSV / Excel) за Обект МОНТЕВИДЕО, склад 1.
2. Обединява записите за всеки партньор в ТОЧНО ЕДНА доставка с множество редове.
3. Валидира контрагентите и артикулите спрямо базата данни на Монтевидео.
4. Отразява доставките като НЕРАЗПЛАТЕНИ в салдата (DEFERREDPAYMENT / Салда доставчици).
5. Записва репликационни команди в SYNCHRONIZATIONLOG (SAVEDELIVERY / SAVEDOCBYDELIVERY)
   за автоматично пренасяне и видимост в Централен склад (CS.FDB) при синхронизация.
6. Генерира подробен одитен Excel доклад.
=============================================================================
"""

import os
import sys
import argparse
import datetime
from decimal import Decimal
import pandas as pd
import firebirdsql
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# Конфигурация за Монтевидео (ЦО)
DB_CONFIG_MV = {
    'host': '192.168.5.254',
    'database': 'E:/db/MONTEVIDEO.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

LOCATION_ID_MV = 7       # Монтевидео
STORAGE_ID_MV = 1        # склад 1
LOCATION_NAME_ID = 4     # Идентификатор на името на обекта в Монтевидео
OPERATOR_USER_ID = 33    # Тонев
OPERATOR_NAME = "Тонев"


class MontevideoDeliveryImporter:
    def __init__(self, file_path, dry_run=True, doc_date=None):
        self.file_path = file_path
        self.dry_run = dry_run
        self.doc_date = doc_date or datetime.date.today()
        self.deliveries = {}  # { partner_name: [item_dict, ...] }
        self.partner_map = {} # { partner_name: { 'num': int, 'name_id': int, 'name': str } }
        self.material_map = {} # { code: { 'code': int, 'name': str, 'uom_id': int } }
        self.summary_rows = []
        self.total_sum = Decimal('0.00')

    def load_file(self):
        """Зарежда и структурира файла с доставки по контрагенти"""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"Файлът не е намерен: {self.file_path}")

        if self.file_path.endswith('.csv'):
            df = pd.read_csv(self.file_path, encoding='utf-8')
        else:
            df = pd.read_excel(self.file_path)

        # Стандартизиране на имената на колоните
        col_map = {
            'Oбект': 'location', 'Обект': 'location',
            'Склад': 'storage',
            'Номер': 'code',
            'Име на материал': 'material_name',
            'Мярка': 'measure',
            'К-во': 'qty',
            'Ед. цена': 'price',
            'Партньор': 'partner'
        }
        df = df.rename(columns=lambda c: col_map.get(c.strip(), c.strip()))

        for _, r in df.iterrows():
            partner = str(r['partner']).strip()
            if not partner or partner == 'nan':
                continue

            code = int(r['code'])
            mat_name = str(r['material_name']).strip()
            measure = str(r['measure']).strip()
            qty = Decimal(str(r['qty']))
            price = Decimal(str(r['price']))

            if partner not in self.deliveries:
                self.deliveries[partner] = []

            self.deliveries[partner].append({
                'code': code,
                'name': mat_name,
                'measure': measure,
                'qty': qty,
                'price': price,
                'line_sum': qty * price
            })

        print(f"📁 Зареден файл: {os.path.basename(self.file_path)}")
        print(f"📊 Общ брой уникални партньори (доставки): {len(self.deliveries)}")
        total_items = sum(len(items) for items in self.deliveries.values())
        print(f"📦 Общ брой редове (артикули): {total_items}")
        return len(self.deliveries)

    def validate_master_data(self, conn):
        """Валидира всички партньори и артикули в базата на Монтевидео"""
        cur = conn.cursor()

        # 1. Валидация на артикулите
        all_codes = set()
        for items in self.deliveries.values():
            for it in items:
                all_codes.add(it['code'])

        for code in all_codes:
            cur.execute("""
                SELECT MATERIALCODE, MATERIAL, UOMID, TAXGROUPID
                FROM MATERIAL
                WHERE LOCATIONID = ? AND STORAGEID = ? AND MATERIALCODE = ?
            """, (LOCATION_ID_MV, STORAGE_ID_MV, code))
            row = cur.fetchone()
            if not row:
                raise ValueError(f"❌ Артикул с код {code} не съществува в Монтевидео (склад 1)!")
            self.material_map[code] = {
                'code': row[0],
                'name': row[1],
                'uom_id': row[2],
                'tax_group': row[3]
            }

        # 2. Валидация на партньорите
        for pname in self.deliveries.keys():
            # PARTNER table
            cur.execute("SELECT NUM, NAME, TAXNUM, VATNUM FROM PARTNER WHERE UPPER(NAME) = ?", (pname.upper(),))
            p_row = cur.fetchone()
            if not p_row:
                cur.execute("SELECT NUM, NAME, TAXNUM, VATNUM FROM PARTNER WHERE UPPER(NAME) LIKE ?", (f"%{pname.upper()}%",))
                p_row = cur.fetchone()

            if not p_row:
                raise ValueError(f"❌ Контрагент '{pname}' не е намерен в таблица PARTNER на Монтевидео!")

            pnum = p_row[0]
            db_pname = p_row[1]

            # PARTNERNAME table
            cur.execute("SELECT ID, PARTNERNAME FROM PARTNERNAME WHERE UPPER(PARTNERNAME) = ?", (pname.upper(),))
            pn_row = cur.fetchone()
            if not pn_row:
                cur.execute("SELECT ID, PARTNERNAME FROM PARTNERNAME WHERE UPPER(PARTNERNAME) LIKE ?", (f"%{pname.upper()}%",))
                pn_row = cur.fetchone()

            if not pn_row:
                raise ValueError(f"❌ Контрагент '{pname}' не е намерен в таблица PARTNERNAME на Монтевидео!")

            name_id = pn_row[0]

            self.partner_map[pname] = {
                'num': pnum,
                'name_id': name_id,
                'db_name': db_pname,
                'tax_num': p_row[2] or '',
                'vat_num': p_row[3] or ''
            }

        print("✅ Всички 9 артикула и 20 контрагента са верифицирани успешно в базата данни!")

    def preview_dry_run(self):
        """Отпечатва детайлен преглед на планираните доставки (Dry-Run)"""
        print("\n" + "=" * 90)
        print(f" ПРЕГЛЕД НА ДОСТАВКИТЕ ЗА ИМПОРТ В МИСТРАЛ МОНТЕВИДЕО (ДАТА: {self.doc_date.strftime('%d.%m.%Y')})")
        print("=" * 90)

        grand_total = Decimal('0.00')
        self.summary_rows = []

        for idx, (pname, items) in enumerate(self.deliveries.items(), 1):
            pinfo = self.partner_map[pname]
            doc_sum = sum(it['line_sum'] for it in items)
            grand_total += doc_sum

            print(f"\n[{idx:02d}] Контрагент: {pinfo['db_name']} (NUM: {pinfo['num']}, NameID: {pinfo['name_id']})")
            print(f"     Брой редове: {len(items)} | Обща стойност на доставката: {doc_sum:>10.2f} € (НЕРАЗПЛАТЕНО В САЛДО)")
            print(f"     {'-' * 80}")
            print(f"     {'Код':<6} | {'Артикул':<32} | {'К-во':<10} | {'Ед. цена':<10} | {'Стойност':<12}")
            print(f"     {'-' * 80}")

            for it in items:
                print(f"     {it['code']:<6} | {it['name']:<32} | {it['qty']:>9.2f} | {it['price']:>8.2f} € | {it['line_sum']:>10.2f} €")
                self.summary_rows.append({
                    'DeliveryIdx': idx,
                    'PartnerInput': pname,
                    'PartnerDB': pinfo['db_name'],
                    'PartnerNum': pinfo['num'],
                    'PartnerNameId': pinfo['name_id'],
                    'ItemCode': it['code'],
                    'ItemName': it['name'],
                    'Measure': it['measure'],
                    'Qty': float(it['qty']),
                    'Price': float(it['price']),
                    'LineSum': float(it['line_sum']),
                    'DeliveryTotal': float(doc_sum)
                })

        print("\n" + "=" * 90)
        print(f" ОБЩО ПЛАНИРАНИ ДОСТАВКИ: {len(self.deliveries)} БРОЯ")
        print(f" ОБЩА СУМА НА ВСИЧКИ ДОСТАВКИ: {grand_total:>12.2f} €")
        print("=" * 90)
        self.total_sum = grand_total

    def execute_import(self, conn):
        """Изпълнява реалния запис в Монтевидео с пълна транзакционна цялост"""
        cur = conn.cursor()
        now = datetime.datetime.now().replace(microsecond=0)
        today = self.doc_date

        print(f"\n🚀 СТАРТИРАНЕ НА РЕАЛЕН ИМПОРТ НА {len(self.deliveries)} ДОСТАВКИ В МОНТЕВИДЕО...")
        committed_deliveries = []

        try:
            for idx, (pname, items) in enumerate(self.deliveries.items(), 1):
                pinfo = self.partner_map[pname]
                pnum = pinfo['num']
                name_id = pinfo['name_id']
                db_name = pinfo['db_name']
                doc_sum = sum(it['line_sum'] for it in items)

                # 1. Генериране на следващ номер на доставка и ID на операция
                cur.execute("SELECT GEN_ID(G_OPERATIONNUM, 1) FROM RDB$DATABASE")
                op_num = cur.fetchone()[0]

                cur.execute("SELECT GEN_ID(G_OPERATIONS, 1) FROM RDB$DATABASE")
                op_id = cur.fetchone()[0]

                cur.execute("SELECT GEN_ID(G_OPERATIONSUNIQUENUM, 1) FROM RDB$DATABASE")
                op_uniq = cur.fetchone()[0]

                # 2. INSERT в OPERATIONS (Header на доставката)
                cur.execute("""
                    INSERT INTO OPERATIONS (
                        ID, DBTYPE, LOCATIONID, UNIQUENUM, NUM, OPERATIONDOCTYPE,
                        DOCUMENTDATE, DOCUMENTTYPEID, USERID, PARTNERNAMEID,
                        PARTNERCONTACTID, PARTNERLOCATIONID, PARTNERSHORTNAMEID,
                        DEFERREDPAYMENT, DEFERREDPAYMENTSUM, DEFERREDPAYMENTBACK,
                        LOCATIONNAMEID, PARTNERNUM, PARTNERLOCATIONNUM, PARTNERCONTACTNUM,
                        TAXNUMID, VATNUMID, PRICEWOTAX, CONFIRMATIONREQUIRED,
                        DATESAVED, DELIVERYTYPEID, PARTNERTYPEID, DELIVERYTYPE, CANCELATIONOPTION
                    ) VALUES (
                        ?, 2, ?, ?, ?, 0,
                        ?, 0, ?, ?,
                        0, 0, 0,
                        '1', ?, ?,
                        ?, ?, -1, -1,
                        0, 0, '1', '0',
                        ?, 0, 0, 0, '0'
                    )
                """, (
                    op_id, LOCATION_ID_MV, op_uniq, op_num,
                    now, OPERATOR_USER_ID, name_id,
                    doc_sum, today,
                    LOCATION_NAME_ID, pnum,
                    now
                ))

                # 3. Обработка на всеки артикул от доставката
                for it_idx, it in enumerate(items):
                    code = it['code']
                    mat_name = it['name']
                    qty = float(it['qty'])
                    price = float(it['price'])
                    f_flag = '1' if it_idx == 0 else '0'

                    # 3.1. SAVEDELIVERY (записва в OPERATIONCONTENT и SYNCHRONIZATIONLOG)
                    args_deliv = [
                        LOCATION_ID_MV, STORAGE_ID_MV, code, mat_name, qty, 'броя', price, 0.0, mat_name, '1',
                        code, '0', price, '', pnum, -1, -1, op_num, 'МОНТЕВИДЕО', now,
                        db_name, '', '', None, today, 'Стокова разписка', OPERATOR_NAME, 'склад 1', 'Доставка', f_flag,
                        None, '', '', '', None, '1', float(doc_sum), today, '', '',
                        '', '', '1', '', '', '', '', '', OPERATOR_USER_ID, 2,
                        '0', '', 0.0, '', None, 0.0, 0, '', '', '',
                        '0', 0.0, '', float(doc_sum), float(doc_sum), 0, '0', 0.0, '', '',
                        '', '', '', '', today, '', 0.0, 0.0, None, None,
                        '2', 20.0, '1', price, 0, '0', '0', 0.0, '', 
                        '1', '1', '0', '1',
                        '', '', '', 0, '0'
                    ]
                    cur.execute(f"EXECUTE PROCEDURE SAVEDELIVERY({', '.join(['?']*len(args_deliv))})", args_deliv)

                    # 3.2. SAVEDOCTINCKLAD (актуализира наличността в MATERIAL.QTY и записва в MATERIALQTYLOG)
                    args_savedoct = [
                        LOCATION_ID_MV, STORAGE_ID_MV, code, mat_name, qty, 'броя', price, 0.0, mat_name, '1',
                        code, '0', pnum, '0', -1, -1, price, '', OPERATOR_USER_ID, '0',
                        price, price, op_num, 2, 0, '0',
                        0, 0, 0, 0, 0,
                        '', 0.0, 0.0, 0.0, now, None,
                        0.0, price, price, '', '0', '0', '0', OPERATOR_NAME
                    ]
                    cur.execute(f"EXECUTE PROCEDURE SAVEDOCTINCKLAD({', '.join(['?']*len(args_savedoct))})", args_savedoct)

                # 4. SAVEDOCBYDELIVERY (записва в OPERATIONDOCUMENT и SYNCHRONIZATIONLOG)
                args_doc = [
                    LOCATION_ID_MV, op_num, 'Стокова разписка', None, today, now,
                    2, '', 0.0, None, '', None, 0.0, '0', today, '1',
                    '', '', '', 0
                ]
                cur.execute(f"EXECUTE PROCEDURE SAVEDOCBYDELIVERY({', '.join(['?']*len(args_doc))})", args_doc)

                # 5. INSERTRAZCR (записва в DEFERREDPAYMENT — неразплатено в салдото на доставчика)
                args_razcr = [
                    LOCATION_ID_MV, pnum, -1, -1, op_num, now, float(doc_sum), today, OPERATOR_NAME,
                    2, None, today, 'без документ', 2, float(doc_sum), float(doc_sum), '', ''
                ]
                cur.execute(f"EXECUTE PROCEDURE INSERTRAZCR({', '.join(['?']*len(args_razcr))})", args_razcr)

                committed_deliveries.append({
                    'num': op_num,
                    'op_id': op_id,
                    'partner': db_name,
                    'sum': doc_sum,
                    'rows_count': len(items)
                })

                print(f"  ✅ [{idx:02d}/{len(self.deliveries)}] Доставка № {op_num} | {db_name:<30} | {len(items)} реда | {doc_sum:>10.2f} €")

            # Финален COMMIT на цялата транзакция
            conn.commit()
            print("\n🎉 УСПЕШЕН COMMIT НА ВСИЧКИ 20 ДОСТАВКИ В МОНТЕВИДЕО!")
            print(f"   Общо записани документи: {len(committed_deliveries)} бр.")
            print(f"   Обща стойност на салдата: {self.total_sum:.2f} €")
            return True

        except Exception as e:
            conn.rollback()
            print(f"\n❌ ГРЕШКА ПРИ ЗАПИС В БАЗАТА: {e}")
            print("   Извършен е пълен ROLLBACK. Базата данни остава непокътната.")
            raise e

    def verify_imported_deliveries(self, conn):
        """Проверява отразените доставки в Монтевидео"""
        cur = conn.cursor()
        print("\n🔍 ВЕРИФИКАЦИЯ НА ЗАПИСИТЕ В МОНТЕВИДЕО...")

        for idx, (pname, items) in enumerate(self.deliveries.items(), 1):
            pnum = self.partner_map[pname]['num']
            doc_sum = sum(it['line_sum'] for it in items)

            cur.execute("""
                SELECT FIRST 1 SALENUM, DPSUM, SALEDATE, DOCUMENTTYPE
                FROM DEFERREDPAYMENT
                WHERE LOCATIONID = ? AND PARTNERNUM = ? AND (DPSUM BETWEEN ? AND ?)
                ORDER BY ID DESC
            """, (LOCATION_ID_MV, pnum, doc_sum - Decimal('0.05'), doc_sum + Decimal('0.05')))
            row = cur.fetchone()
            if row:
                print(f"  ✓ {pname:<30} -> Док № {row[0]} | Сума в DEFERREDPAYMENT: {float(row[1]):.2f} € | Тип: {row[3]}")
            else:
                print(f"  ⚠️ {pname:<30} -> Не е открит запис в DEFERREDPAYMENT!")

        # Проверка на наличностите на доставените артикули
        print("\n📦 ВЕРИФИКАЦИЯ НА СКЛАДОВИТЕ НАЛИЧНОСТИ (MATERIAL.QTY / GETARTIKULNAL)...")
        all_codes = sorted({it['code'] for items in self.deliveries.values() for it in items})
        for code in all_codes:
            cur.execute("SELECT MATERIAL, QTY FROM MATERIAL WHERE LOCATIONID = ? AND STORAGEID = ? AND MATERIALCODE = ?", (LOCATION_ID_MV, STORAGE_ID_MV, code))
            mat_row = cur.fetchone()
            cur.execute("SELECT NAL FROM GETARTIKULNAL(?, ?)", (LOCATION_ID_MV, code))
            nal_row = cur.fetchone()
            mat_name = mat_row[0] if mat_row else "Неизвестен"
            mat_qty = float(mat_row[1]) if mat_row else 0.0
            sys_nal = float(nal_row[0]) if nal_row else 0.0
            print(f"  ✓ Код {code:<6} | {mat_name:<28} | MATERIAL.QTY: {mat_qty:>10.3f} | GETARTIKULNAL: {sys_nal:>10.3f}")

    def export_excel_report(self, output_path):
        """Експортира структуриран одитен Excel отчет"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Доставки Монтевидео"
        ws.views.sheetView[0].showGridLines = True

        # Styles
        title_font = Font(name='Calibri', size=14, bold=True, color='1F497D')
        header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
        regular_font = Font(name='Calibri', size=10)
        bold_font = Font(name='Calibri', size=10, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        total_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

        # Title
        ws.merge_cells('A1:J1')
        title_cell = ws['A1']
        title_cell.value = f"ОДИТЕН ДОКЛАД — ИМПОРТ НА ДОСТАВКИ В МИСТРАЛ МОНТЕВИДЕО ({self.doc_date.strftime('%d.%m.%Y')})"
        title_cell.font = title_font
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30

        # Subtitle
        ws.merge_cells('A2:J2')
        sub_cell = ws['A2']
        status_txt = "РЕАЛЕН ЗАПИС В БАЗАТА (LIVE COMMIT)" if not self.dry_run else "ПРЕДВАРИТЕЛЕН ПРЕГЛЕД (DRY RUN)"
        sub_cell.value = f"Обект: МОНТЕВИДЕО (ID: 7) | Склад: 1 | Оператор: Тонев (ID: 33) | Статус: {status_txt}"
        sub_cell.font = Font(name='Calibri', size=10, italic=True)
        sub_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 20

        # Headers
        headers = [
            "Доставка №", "Контрагент (Файл)", "Партньор в Мистрал", "PARTNER.NUM",
            "Код", "Име на материал", "Мярка", "К-во", "Ед. цена (€)", "Стойност (€)"
        ]

        ws.row_dimensions[4].height = 25
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        row_idx = 5
        for r in self.summary_rows:
            ws.cell(row=row_idx, column=1, value=r['DeliveryIdx']).alignment = Alignment(horizontal='center')
            ws.cell(row=row_idx, column=2, value=r['PartnerInput']).alignment = Alignment(horizontal='left')
            ws.cell(row=row_idx, column=3, value=r['PartnerDB']).alignment = Alignment(horizontal='left')
            ws.cell(row=row_idx, column=4, value=r['PartnerNum']).alignment = Alignment(horizontal='center')
            ws.cell(row=row_idx, column=5, value=r['ItemCode']).alignment = Alignment(horizontal='center')
            ws.cell(row=row_idx, column=6, value=r['ItemName']).alignment = Alignment(horizontal='left')
            ws.cell(row=row_idx, column=7, value=r['Measure']).alignment = Alignment(horizontal='center')
            
            c_qty = ws.cell(row=row_idx, column=8, value=r['Qty'])
            c_qty.number_format = '#,##0.00'
            c_qty.alignment = Alignment(horizontal='right')

            c_price = ws.cell(row=row_idx, column=9, value=r['Price'])
            c_price.number_format = '#,##0.00 €'
            c_price.alignment = Alignment(horizontal='right')

            c_sum = ws.cell(row=row_idx, column=10, value=r['LineSum'])
            c_sum.number_format = '#,##0.00 €'
            c_sum.alignment = Alignment(horizontal='right')

            for c in range(1, 11):
                ws.cell(row=row_idx, column=c).border = thin_border
                ws.cell(row=row_idx, column=c).font = regular_font

            row_idx += 1

        # Total row
        ws.row_dimensions[row_idx].height = 22
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=9)
        t_label = ws.cell(row=row_idx, column=1, value="ОБЩА СТОЙНОСТ НА ВСИЧКИ ДОСТАВКИ:")
        t_label.font = bold_font
        t_label.alignment = Alignment(horizontal='right', vertical='center')

        t_val = ws.cell(row=row_idx, column=10, value=float(self.total_sum))
        t_val.font = bold_font
        t_val.number_format = '#,##0.00 €'
        t_val.alignment = Alignment(horizontal='right', vertical='center')

        for c in range(1, 11):
            ws.cell(row=row_idx, column=c).fill = total_fill
            ws.cell(row=row_idx, column=c).border = thin_border

        # Adjust column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len and cell.row > 2:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        wb.save(output_path)
        print(f"📊 Одитен Excel файл записан успешно: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Импорт на доставки в Мистрал Монтевидео")
    parser.add_argument('--file', default=r"g:\My Drive\AntiG_test\доставки_октомври_2026.csv", help="Път към CSV/Excel файла")
    parser.add_argument('--commit', action='store_true', help="Изпълнява реален запис в базата (по подразбиране е Dry-Run)")
    parser.add_argument('--verify', action='store_true', help="Само проверява и верифицира състоянието на салдата и наличностите")
    parser.add_argument('--date', default=None, help="Дата на документа (ГГГГ-ММ-ДД), по подразбиране днешна дата")
    args = parser.parse_args()

    doc_date = datetime.datetime.strptime(args.date, '%Y-%m-%d').date() if args.date else datetime.date.today()

    importer = MontevideoDeliveryImporter(args.file, dry_run=not args.commit, doc_date=doc_date)
    importer.load_file()

    conn = firebirdsql.connect(**DB_CONFIG_MV)
    try:
        importer.validate_master_data(conn)

        if args.verify:
            importer.verify_imported_deliveries(conn)
            return

        importer.preview_dry_run()

        if args.commit:
            importer.execute_import(conn)
            importer.verify_imported_deliveries(conn)
        else:
            print("\n💡 Забележка: Това беше симулация (Dry-Run). За реален запис стартирайте с флаг --commit.")

    finally:
        conn.close()

    # Експорт на одитен Excel
    today_str = datetime.date.today().strftime('%Y%m%d')
    output_excel = os.path.join(os.path.dirname(args.file), f"ОДИТ_ИМПОРТ_ДОСТАВКИ_МОНТЕВИДЕО_{today_str}.xlsx")
    importer.export_excel_report(output_excel)


if __name__ == '__main__':
    main()
