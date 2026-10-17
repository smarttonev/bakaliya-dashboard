#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
===============================================================================
ИНТЕРАКТИВЕН СТАРТЕР ЗА БАНКОВО ЗАСИЧАНЕ — ВЕРИГА „БАКАЛИЯ“ (МОНТЕВИДЕО)
===============================================================================
Функции:
1. Автоматично открива най-новия свален report*.xml / report*.hml от Банка ДСК
   (както в работната папка, така и в папка Downloads).
2. Поддържа Drag & Drop на XML файл директно върху иконата.
3. Изпълнява предварителен Dry-Run анализ без да променя базата данни.
4. Генерира официалния оформен файл ОДИТ_МОНТЕВИДЕО_ДСК_...xlsx и го отваря
   АВТОМАТИЧНО в Microsoft Excel за визуален преглед.
5. При потвърждение от потребителя (Y) извършва реален запис в MONTEVIDEO.FDB
   със 100% спазване на всички правила (Anti-FIFO, ActionId репликация,
   Наем Окръжна, обединени такси).
===============================================================================
"""

import os
import sys
import glob
import datetime
from decimal import Decimal

# Осигуряване на UTF-8 в Windows конзолата
if sys.platform == 'win32':
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Импорт на основния двигател
from bank_reconciliation import MontevideoBankReconciliationEngine, DB_CONFIG

try:
    import firebirdsql
    HAS_FIREBIRD = True
except ImportError:
    HAS_FIREBIRD = False


def find_candidate_files():
    """Открива потенциални банкови файлове в папката на проекта и в Downloads"""
    search_dirs = [
        os.path.dirname(os.path.abspath(__file__)),
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

    # Сортиране по дата на промяна (най-новите най-отгоре)
    candidates.sort(key=lambda x: x[1], reverse=True)
    return [c[0] for c in candidates]


def select_bank_file():
    """Определя кой файл ще се обработва"""
    # 1. Ако е подаден като аргумент (например чрез Drag & Drop върху .bat)
    if len(sys.argv) > 1:
        # Първо опитваме с директния sys.argv[1]
        arg_file = sys.argv[1].strip().strip('"')
        if os.path.exists(arg_file):
            return os.path.abspath(arg_file)
        # Ако пътят съдържа интервали и не е бил в кавички, обединяваме всички аргументи:
        joined_arg = " ".join(sys.argv[1:]).strip().strip('"')
        if os.path.exists(joined_arg):
            return os.path.abspath(joined_arg)

    # 2. Търсене на налични файлове
    candidates = find_candidate_files()
    if not candidates:
        print("\n❌ Не бяха открити банкови файлове (report*.xml / report*.hml)!")
        print("   Моля, изтеглете извлечението от ДСК в Downloads или го поставете тук.")
        return None

    print("\n" + "=" * 78)
    print("  🏦 ОТКРИТИ БАНКОВИ ИЗВЛЕЧЕНИЯ ДСК (СОРТИРАНИ ПО ДАТА):")
    print("=" * 78)
    for idx, f in enumerate(candidates[:5], 1):
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f))
        size_kb = os.path.getsize(f) / 1024.0
        folder_tag = "[Downloads]" if "Downloads" in f else "[Работна папка]"
        print(f"  [{idx}] {folder_tag} {os.path.basename(f):<25} ({mtime.strftime('%d.%m.%Y %H:%M')}, {size_kb:.1f} KB)")

    print("-" * 78)
    default_file = candidates[0]
    prompt_text = f"  Изберете номер [1-{min(len(candidates), 5)}] или натиснете [Enter] за най-новия ({os.path.basename(default_file)}): "
    try:
        choice = input(prompt_text).strip()
    except (EOFError, KeyboardInterrupt):
        choice = ""

    if not choice:
        return default_file

    try:
        choice_idx = int(choice)
        if 1 <= choice_idx <= len(candidates):
            return candidates[choice_idx - 1]
    except ValueError:
        pass

    if os.path.exists(choice):
        return os.path.abspath(choice)

    return default_file


def main():
    print("=" * 78)
    print("  🚀 БАКАЛИЯ — АВТОМАТИЗИРАНО ЗАСИЧАНЕ НА БАНКОВИ ИЗВЛЕЧЕНИЯ ДСК")
    print("  Централен офис — Монтевидео | База данни: SMSERVERBIG (192.168.5.254)")
    print("=" * 78)

    xml_file = select_bank_file()
    if not xml_file:
        input("\nНатиснете [Enter] за изход...")
        return

    print(f"\n📂 Избран файл за обработка: {xml_file}")

    # 1. Зареждане на банковия файл
    engine = MontevideoBankReconciliationEngine(xml_file, dry_run=True)
    try:
        count = engine.parse_xml()
        print(f"  -> Успешно заредени {count} транзакции от банковото извлечение.")
    except Exception as e:
        print(f"\n❌ Грешка при четене на XML файла: {e}")
        input("\nНатиснете [Enter] за изход...")
        return

    # 2. Връзка с базата данни
    db_conn = None
    if HAS_FIREBIRD:
        try:
            db_conn = firebirdsql.connect(**DB_CONFIG)
            print(f"  -> Успешна връзка с Мистрал МОНТЕВИДЕО ({DB_CONFIG['host']}:{DB_CONFIG['database']})")
        except Exception as e:
            print(f"\n⚠️ ВНИМАНИЕ: Няма връзка с база данни Монтевидео: {e}")
            print("   Анализът ще се извърши без справка с базата.")
    else:
        print("\n⚠️ Модулът firebirdsql не е инсталиран. Свързването с базата е невъзможно.")

    # 3. Предварителен Dry-Run анализ
    print("\n🔍 Анализиране на движенията и съпоставяне със задълженията в Мистрал...")
    engine.classify_movements(db_conn)
    engine.print_console_summary()

    # 4. Генериране на официалния Excel файл
    file_base = os.path.splitext(os.path.basename(xml_file))[0].replace(' ', '')
    today_str = datetime.date.today().strftime('%Y%m%d')
    output_excel = os.path.join(os.path.dirname(xml_file), f"ОДИТ_МОНТЕВИДЕО_ДСК_{file_base}.xlsx")
    generated_path = engine.export_excel_report(output_excel)

    # Записваме и работно копие в директорията на проекта, ако файлът е бил в Downloads
    project_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.abspath(os.path.dirname(xml_file)) != os.path.abspath(project_dir):
        alt_project_path = os.path.join(project_dir, f"ОДИТ_МОНТЕВИДЕО_ДСК_{file_base}.xlsx")
        try:
            engine.export_excel_report(alt_project_path)
        except Exception:
            pass

    # 5. Автоматично отваряне на Excel файла за одит
    if generated_path and os.path.exists(generated_path):
        print(f"\n🖥️ Отваряне на одитния файл в Excel: {os.path.basename(generated_path)} ...")
        try:
            os.startfile(generated_path)
        except Exception as e:
            print(f"  (Не успяхме да стартираме Excel автоматично: {e})")

    # 6. Проверка за готови за запис редове
    green_count = engine.stats.get('exact_match_ready_count', 0)
    green_sum = engine.stats.get('exact_match_ready_sum', Decimal('0.00'))

    print("\n" + "=" * 78)
    print("  РЕЗУЛТАТ ОТ ПРЕДВАРИТЕЛНАТА ПРОВЕРКА:")
    print("=" * 78)
    if green_count > 0:
        print(f"  🟢 ГОТОВИ ЗА ЗАКРИВАНЕ: {green_count} транзакции за общо {green_sum:,.2f} €")
        print("  🟡 ВЕЧЕ ЗАКРИТИ:        Всички предходни плащания са разпознати и защитени.")
        print("  ⚪ БАНКОВИ ТАКСИ:       Обединени на 1 ред за ръчно осчетоводяване.")
        print("-" * 78)
        print("  ПРЕГЛЕДАЙТЕ ОТВОРЕНИЯ ЕКСЕЛСКИ ФАЙЛ. Ако всичко е наред:")
        print("  Натиснете [ Y ] за ДА (Реален запис в Мистрал Монтевидео)")
        print("  Натиснете [ N ] за НЕ (Отказ - тестов режим)")
        print("-" * 78)
        
        try:
            user_choice = input("  Желаете ли реален запис в Мистрал? [Y / N]: ").strip().upper()
        except (EOFError, KeyboardInterrupt):
            user_choice = 'N'

        if user_choice in ['Y', 'YES', 'Д', 'ДА']:
            if not db_conn:
                print("\n❌ Грешка: Няма активна връзка с базата данни! Записът е невъзможен.")
            else:
                success = engine.execute_commit(db_conn)
                if success:
                    print("\n" + "=" * 78)
                    print("  🎉 УСПЕШНО ЗАПИСАНИ ПЛАЩАНИЯ В МОНТЕВИДЕО!")
                    print("=" * 78)
                    print("  ✅ Каса 'ДСК Бакалия' и салдата на доставчиците са актуализирани.")
                    print("  ✅ Генерирани са пакети в SYNCHRONIZATIONLOG с динамичен ActionId.")
                    print("  ℹ️ При следващата синхронизация на SyncManager сметките автоматично")
                    print("     ще бъдат закрити и в Централен склад (CS.FDB).")
                    print("=" * 78)

                    # Обновяваме одитния Excel файл със статуси [ВЕЧЕ ЗАКРИТО]
                    engine.movements = []
                    engine.classified_rows = []
                    engine.stats = {k: (0 if isinstance(v, int) else Decimal('0.00')) for k, v in engine.stats.items()}
                    engine.parse_xml()
                    engine.classify_movements(db_conn)
                    engine.export_excel_report(output_excel)
        else:
            print("\n" + "=" * 78)
            print("  🛑 ОПЕРАЦИЯТА Е ОТКАЗАНА.")
            print("  Базата данни НЕ е променяна. Одитният Excel файл остава за ваша справка.")
            print("=" * 78)
    else:
        print("  ℹ️ Няма нови зелени транзакции за автоматично закриване.")
        print("     Всички плащания са или вече закрити, или изискват ръчен одит.")
        print("  Базата данни не изисква промени.")
        print("=" * 78)

    if db_conn:
        try:
            db_conn.close()
        except Exception:
            pass

    print()
    try:
        input("Натиснете [Enter] за изход...")
    except (EOFError, KeyboardInterrupt):
        pass


if __name__ == '__main__':
    main()
