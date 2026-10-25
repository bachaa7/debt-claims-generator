# -*- coding: utf-8 -*-
"""
Главный скрипт: превращает выгрузку отчёта по ДЗ из 1С в готовые претензии
(и, при необходимости, исковые заявления) по каждому договору с просрочкой.

Использование:
    python3 main.py pretenzii  --report Otchet.xls --contragents data/contragents.csv --min-days 5
    python3 main.py iski       --report Otchet.xls --contragents data/contragents.csv --min-days 5

Перед запуском:
    1. Экспортируйте отчёт "Отчет по ДЗ" из 1С (кнопка "Сформировать" на скрине) и сохраните файл.
    2. Один раз заполните data/contragents.csv реквизитами покупателей (адрес, УНП, банк) —
       возьмите их из папки с договорами. Это нужно, чтобы претензия/иск имели юридическую силу.
    3. При необходимости поправьте ставку пени в колонке "Ставка_пени_%_override", если в конкретном
       договоре она отличается от типовой (0,1%/день — биржевой, 0,5%/день — прямой).
"""
import argparse
import csv
import datetime as dt
import os
from collections import defaultdict

from parser import parse_debt_report, classify_contract_type, DebtLine
from docgen import generate_pretenzia, generate_isk, BuyerInfo
from company import DEFAULT_PENALTY_RATE
import journal

STATUTE_OF_LIMITATIONS_DAYS = 1095  # ~3 года, общий срок исковой давности по ГК РБ (ст.197)


def load_contragents(path: str) -> dict:
    data = {}
    if not os.path.exists(path):
        return data
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            data[row["Контрагент"]] = row
    return data


def build_buyer(cp_row: dict, name: str) -> BuyerInfo:
    return BuyerInfo(
        name=name,
        name_full=(cp_row.get("Полное_наименование", "") if cp_row else "") or name,
        address=cp_row.get("Юр_адрес", "") if cp_row else "",
        unp=cp_row.get("УНП", "") if cp_row else "",
        bank_account=cp_row.get("Расчетный_счет", "") if cp_row else "",
        bank_name=cp_row.get("Банк", "") if cp_row else "",
        bank_code=cp_row.get("Код_банка", "") if cp_row else "",
        phone=cp_row.get("Телефон", "") if cp_row else "",
        director_fio=cp_row.get("Директор_ФИО", "") if cp_row else "",
    )


def group_by_contract(lines: list[DebtLine]):
    groups = defaultdict(list)
    for l in lines:
        groups[(l.counterparty, l.contract_number, l.contract_date)].append(l)
    return groups


def run(mode: str, report: str, contragents_path: str, out_dir: str, min_days: int,
        as_of: dt.date, contragent: str = None, contract: str = None, only_keys=None,
        progress=None):
    """Генерирует документы.

    contragent/contract — необязательный отбор из командной строки: достаточно
    части названия или номера, регистр не важен.
    only_keys — набор ключей (контрагент, номер договора, дата договора); им
    пользуется окно программы, когда галочками отмечены конкретные строки.
    progress — необязательная функция обратного вызова f(сделано, всего, текст).
    """
    lines = parse_debt_report(report)
    # новые покупатели из отчёта попадают в справочник сами; заполненное не трогается
    from contracts import ensure_counterparties
    added = ensure_counterparties(lines, contragents_path)
    if added:
        print(f"В справочник добавлено новых покупателей: {len(added)} — {', '.join(added)}")
    contragents = load_contragents(contragents_path)
    groups = group_by_contract(lines)

    os.makedirs(out_dir, exist_ok=True)
    summary_rows = []
    journal_entries = []
    skipped_by_filter = 0

    for (cp, contract_number, contract_date), grp_lines in sorted(groups.items()):
        qualifying = [l for l in grp_lines if l.overdue_days >= min_days]
        if not qualifying:
            continue

        if only_keys is not None and (cp, contract_number, contract_date) not in only_keys:
            skipped_by_filter += 1
            continue
        if contragent and contragent.lower() not in cp.lower():
            skipped_by_filter += 1
            continue
        if contract and contract.lower() not in contract_number.lower():
            skipped_by_filter += 1
            continue

        max_days = max(l.overdue_days for l in qualifying)
        contract_type = classify_contract_type(contract_number)
        cp_row = contragents.get(cp)
        rate_override = None
        if cp_row and cp_row.get("Ставка_пени_%_override"):
            try:
                rate_override = float(cp_row["Ставка_пени_%_override"])
            except ValueError:
                rate_override = None

        buyer = build_buyer(cp_row, cp)
        statute_risk = max_days > STATUTE_OF_LIMITATIONS_DAYS

        # Раскладка: папка на покупателя, внутри — претензии и иски.
        #   output/СтальТрейд/pretenzii/pretenzia__010126/1.docx
        #   output/СтальТрейд/iski/isk__010126/1.docx
        # Повторная генерация перезаписывает файл — актуальной остаётся одна версия.
        safe_cp = "".join(c for c in cp if c.isalnum() or c in " _-").strip().replace(" ", "_")
        safe_contract = contract_number.replace("/", "-")
        case_dir = os.path.join(out_dir, safe_cp, mode)
        os.makedirs(case_dir, exist_ok=True)

        result = {}
        out_path = ""
        if mode == "pretenzii":
            out_path = os.path.join(case_dir, f"pretenzia__{safe_contract}.docx")
            result = generate_pretenzia(
                buyer=buyer,
                contract_number=contract_number,
                contract_date=contract_date,
                contract_type=contract_type,
                lines=qualifying,
                out_path=out_path,
                as_of=as_of,
                penalty_rate=rate_override,
            )
        elif mode == "iski":
            out_path = os.path.join(case_dir, f"isk__{safe_contract}.docx")
            result = generate_isk(
                buyer=buyer,
                contract_number=contract_number,
                contract_date=contract_date,
                contract_type=contract_type,
                lines=qualifying,
                out_path=out_path,
                as_of=as_of,
                penalty_rate=rate_override,
            )

        if progress:
            progress(len(summary_rows) + 1, cp, contract_number)

        journal_entries.append(dict(
            Контрагент=cp,
            Договор=contract_number,
            Дата_договора=contract_date,
            Документ=("претензия" if mode == "pretenzii" else "иск"),
            Дата_формирования=as_of.strftime("%d.%m.%Y"),
            Долг_на_момент=result.get("total_debt", ""),
            Итого_на_момент=result.get("total_due", ""),
            Файл=out_path,
        ))

        summary_rows.append(
            dict(
                Контрагент=cp,
                Договор=contract_number,
                Тип=contract_type,
                Дней_просрочки_max=max_days,
                Осн_долг=result.get("total_debt", ""),
                Пеня=result.get("total_penalty", ""),
                Проценты_ст366=result.get("total_interest", ""),
                Итого=result.get("total_due", ""),
                Госпошлина=result.get("duty", ""),
                Госпошлина_подтверждена=("да" if result.get("duty_verified") else "ориентировочно — сверьте data/tariffs.json"),
                Срок_ответа=result.get("deadline", ""),
                Риск_срока_давности=("ДА — проверьте срок исковой давности (>3 лет)!" if statute_risk else ""),
                Требуются_реквизиты=("ДА — заполните contragents.csv" if not cp_row or not cp_row.get("УНП") else ""),
                Файл=out_path,
            )
        )

    if journal_entries:
        journal.record(journal_entries)

    if mode == "iski" and summary_rows:
        # число листов иска в приложении — настоящее (считает Word, если он есть)
        from docgen import fix_isk_sheets
        fixed = fix_isk_sheets([r["Файл"] for r in summary_rows])
        if fixed:
            print(f"Число листов иска в приложении проставлено: {fixed} шт.")

    from rates import rate_freshness
    from company import CHARGE_INTEREST_366
    if CHARGE_INTEREST_366:
        fresh = rate_freshness(as_of)
        if fresh["устарела"]:
            print("ВНИМАНИЕ: " + fresh["текст"])

    summary_path = os.path.join(out_dir, f"_summary_{mode}.csv")
    if summary_rows:
        with open(summary_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(summary_rows[0].keys()))
            writer.writeheader()
            writer.writerows(summary_rows)

    if not summary_rows and (contragent or contract):
        print("По заданному отбору ничего не найдено.")
        print(f'  контрагент содержит: {contragent or "— (не задан)"}')
        print(f'  договор содержит:    {contract or "— (не задан)"}')
        print(f"Пропущено по отбору договоров: {skipped_by_filter}.")
        print('Посмотрите точные названия командой:  python main.py list --report <файл>')
        return summary_rows

    print(f"Сгенерировано документов: {len(summary_rows)}")
    for r in summary_rows[:20]:
        print(f'  {r["Контрагент"]} — дог. {r["Договор"]} — итого {r["Итого"]} руб. — {r["Файл"]}')
    if len(summary_rows) > 20:
        print(f"  ... и ещё {len(summary_rows) - 20}")
    print(f"Сводка: {summary_path}")
    return summary_rows


def show_list(report: str, min_days: int):
    """Печатает, по каким контрагентам и договорам есть просрочка."""
    lines = parse_debt_report(report)
    groups = group_by_contract(lines)
    rows = []
    for (cp, contract_number, _), grp in sorted(groups.items()):
        qualifying = [l for l in grp if l.overdue_days >= min_days]
        if not qualifying:
            continue
        rows.append((cp, contract_number, max(l.overdue_days for l in qualifying),
                     round(sum(l.overdue_amount for l in qualifying), 2), len(qualifying)))

    print(f"Договоров с просрочкой от {min_days} дн.: {len(rows)}")
    print()
    print(f'{"Контрагент":<45} {"Договор":<18} {"Дней":>5} {"Накл.":>6} {"Сумма долга":>14}')
    print("-" * 92)
    for cp, num, days, total, cnt in rows:
        print(f"{cp[:44]:<45} {num[:17]:<18} {days:>5} {cnt:>6} {total:>14,.2f}".replace(",", " "))
    print()
    print("Чтобы сделать документ по одному из них:")
    print('  python main.py pretenzii --report <файл> --contragent "СтальТрейд" --contract "020126/2"')
    return rows


def current_debts(report: str, min_days: int) -> dict:
    """{(контрагент, договор, дата_договора): сумма_просроченного_долга} из свежей выгрузки."""
    result = {}
    for (cp, number, date), grp in group_by_contract(parse_debt_report(report)).items():
        qualifying = [l for l in grp if l.overdue_days >= min_days]
        if qualifying:
            result[(cp, number, date)] = round(sum(l.overdue_amount for l in qualifying), 2)
    return result


def show_journal(report: str, min_days: int, as_of: dt.date):
    """Что по каким договорам уже сделано, отправлено, оплачено."""
    debts = current_debts(report, min_days)
    state = journal.build_state()
    keys = sorted(set(debts) | set(state))

    rows = []
    for key in keys:
        st = journal.status_for(key, state, debts.get(key, 0.0), as_of)
        rows.append((key, debts.get(key, 0.0), st))

    print(f"Дата: {as_of.strftime('%d.%m.%Y')}. Записей: {len(rows)}")
    print()
    print(f'{"Контрагент":<38} {"Договор":<16} {"Долг сейчас":>13}  {"Статус":<26} Подсказка')
    print("-" * 130)
    for (cp, num, _), debt, st in rows:
        money = f"{debt:,.2f}".replace(",", " ")
        print(f'{cp[:37]:<38} {num[:15]:<16} {money:>13}  {st["статус"]:<26} {st["подсказка"]}')

    ready = [r for r in rows if r[2]["можно_иск"] and r[2]["статус"] != journal.SUIT_MADE]
    print()
    if ready:
        print(f"Готовы к иску (срок претензии истёк, оплаты нет): {len(ready)}")
        for (cp, num, _), debt, _st in ready:
            print(f'  {cp} — договор {num} — {f"{debt:,.2f}".replace(",", " ")} руб.')
        print()
        print('Сформировать иски по ним:  python main.py iski --report <файл> --ready-for-suit')
    else:
        print("Готовых к иску пока нет.")
    return rows


def do_mark_sent(report: str, min_days: int, contragent: str, contract: str, sent: dt.date):
    """Отметить претензии как отправленные."""
    debts = current_debts(report, min_days)
    keys = [k for k in debts
            if (not contragent or contragent.lower() in k[0].lower())
            and (not contract or contract.lower() in k[1].lower())]
    if not keys:
        print("Под этот отбор ничего не нашлось.")
        return 0
    count = journal.mark_sent(keys, sent)
    print(f'Отмечено отправленными: {count} претензий, дата {sent.strftime("%d.%m.%Y")}.')
    print(f'Срок ответа истекает {(sent + dt.timedelta(days=journal.PRETENZIA_DEADLINE_DAYS)).strftime("%d.%m.%Y")}.')
    return count


def ready_for_suit_keys(report: str, min_days: int, as_of: dt.date) -> set:
    """Договоры, по которым претензия отправлена, срок вышел, оплаты нет."""
    debts = current_debts(report, min_days)
    state = journal.build_state()
    return {k for k in debts
            if journal.status_for(k, state, debts[k], as_of)["можно_иск"]}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("mode", choices=["pretenzii", "iski", "list", "journal", "mark-sent",
                                     "update-rate"],
                    help="pretenzii | iski | list (список должников) | "
                         "journal (что уже отправлено и что готово к иску) | "
                         "mark-sent (отметить претензии отправленными) | "
                         "update-rate (скачать свежую ставку рефинансирования)")
    ap.add_argument("--report", default=None,
                    help="файл отчёта (не нужен только для update-rate)")
    ap.add_argument("--contragents", default="data/contragents.csv")
    ap.add_argument("--out", default=None)
    ap.add_argument("--min-days", type=int, default=5)
    ap.add_argument("--as-of", default=None, help="дд.мм.гггг, по умолчанию — сегодня")
    ap.add_argument("--contragent", default=None,
                    help='сделать документ только по одному покупателю (часть названия, напр. "СтальТрейд")')
    ap.add_argument("--contract", default=None,
                    help='только по одному договору (часть номера, напр. "020126/2")')
    ap.add_argument("--ready-for-suit", action="store_true",
                    help="только те, у кого претензия отправлена, срок вышел, оплаты нет")
    ap.add_argument("--sent-date", default=None,
                    help="для mark-sent: дата отправки дд.мм.гггг (по умолчанию — сегодня)")
    args = ap.parse_args()

    as_of = dt.datetime.strptime(args.as_of, "%d.%m.%Y").date() if args.as_of else dt.date.today()

    if args.mode == "update-rate":
        from rates import update_refi_file, mark_checked
        res = update_refi_file()
        print(res["текст"])
        if res["ok"]:
            mark_checked()
        raise SystemExit(0 if res["ok"] else 1)

    if not args.report:
        ap.error("нужен --report")

    if args.mode == "list":
        show_list(args.report, args.min_days)
    elif args.mode == "journal":
        show_journal(args.report, args.min_days, as_of)
    elif args.mode == "mark-sent":
        sent = dt.datetime.strptime(args.sent_date, "%d.%m.%Y").date() if args.sent_date else dt.date.today()
        do_mark_sent(args.report, args.min_days, args.contragent, args.contract, sent)
    else:
        out_dir = args.out or "output"
        only = ready_for_suit_keys(args.report, args.min_days, as_of) if args.ready_for_suit else None
        if args.ready_for_suit and not only:
            print("Готовых к иску нет: либо претензии ещё не отправлены, либо срок не вышел.")
        else:
            run(args.mode, args.report, args.contragents, out_dir, args.min_days, as_of,
                contragent=args.contragent, contract=args.contract, only_keys=only)
