# -*- coding: utf-8 -*-
"""Контрольный набор: 7 договоров с заранее посчитанными ответами.

Запуск:  python tests/control_set.py

Ничего не читает с диска, кроме data/refinancing_rate.csv, и ничего не пишет —
реестр и справочник не трогаются. Гоняйте после любой правки в расчёте.

Долг и пеня зафиксированы навсегда: они не зависят ни от каких настроек.
Проценты по ст. 366 зависят от ставки рефинансирования в
data/refinancing_rate.csv — если вы дописали туда новую строку с датой раньше
19.08.2026, расхождение по колонке %366 законно, пересчитайте ожидания.
"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import docgen
import payments as pay
from company import DEFAULT_PENALTY_RATE
from parser import DebtLine, classify_contract_type

AS_OF = dt.date(2026, 8, 19)

# контрагент, договор, дата дог., накладная, дата накл., срок оплаты, остаток, платежи
CASES = [
    ("СтальТрейд", "010726/1", "01.07.26", "5000001", "03.07.26",
     dt.date(2026, 7, 10), 10000.00, []),
    ("СтальТрейд", "020726/2", "02.07.26", "5000002", "05.07.26",
     dt.date(2026, 7, 15), 6000.00,
     [(dt.date(2026, 8, 1), 3000.00, "9301"), (dt.date(2026, 8, 10), 1000.00, "9315")]),
    ("АвтоЗавод", "01/13900-101", "10.07.26", "6100001", "12.07.26",
     dt.date(2026, 7, 20), 50000.00, []),
    ("АвтоЗавод", "01/13900-102", "14.07.26", "6100002", "16.07.26",
     dt.date(2026, 7, 25), 20000.00, [(dt.date(2026, 8, 5), 15000.00, "9420")]),
    ("АвтоЗавод", "01/13900-102", "14.07.26", "6100003", "18.07.26",
     dt.date(2026, 7, 28), 8000.00, []),
    ("АгроМаш", "01/13901-110", "20.07.26", "6100004", "22.07.26",
     dt.date(2026, 8, 1), 15000.00, []),
    ("ПродТорг", "050826/1", "05.08.26", "5000003", "06.08.26",
     dt.date(2026, 8, 5), 45.00, []),
    ("Кузнецов Кирилл Кириллович", "100122/1", "10.01.22", "4000001", "12.01.22",
     dt.date(2022, 1, 15), 3000.00, []),   # просрочка 1677 дн. — пеня режется до 1095
]

# договор -> (долг, пеня, проценты ст.366, итого)
EXPECTED = {
    "010726/1":     (10000.00,  2000.00, 102.78, 12102.78),
    "020726/2":     ( 6000.00,  1415.00,  72.71,  7487.71),
    "01/13900-101": (50000.00,  1500.00, 385.42, 51885.42),
    "01/13900-102": (28000.00,   826.00, 212.23, 29038.23),
    "01/13901-110": (15000.00,   270.00,  69.38, 15339.38),
    "050826/1":     (   45.00,     3.15,   0.16,    48.31),
    "100122/1":     ( 3000.00, 16425.00, 872.31, 20297.31),
}

# разбивка по частям там, где были оплаты: договор -> [(сумма, дней, пеня), ...]
EXPECTED_PARTS = {
    "020726/2":     [(3000.00, 16, 240.00), (1000.00, 25, 125.00), (6000.00, 35, 1050.00)],
    "01/13900-102": [(15000.00, 10, 150.00), (20000.00, 25, 500.00), (8000.00, 22, 176.00)],
}


def build():
    index, groups = {}, {}
    for cp, dog, dd, ttn, td, due, rest, pays in CASES:
        if pays:
            index[(cp, dog, dd, ttn)] = [
                dict(дата=d, сумма=s, пп=pp, назначение="ЗА МЕТАЛЛОПРОКАТ")
                for d, s, pp in pays]
        groups.setdefault((cp, dog, dd), []).append(
            DebtLine(counterparty=cp, contract_number=dog, contract_date=dd,
                     due_date=due, doc_type="Руб.накл", doc_number=ttn, doc_date=td,
                     total=rest, overdue_amount=rest, overdue_days=(AS_OF - due).days))
    return index, groups


def main():
    index, groups = build()
    pay.build_index = lambda *a, **k: index      # подменяем справочник оплат

    money = lambda v: f"{v:,.2f}".replace(",", " ")
    print(f'{"договор":<16}{"тип":<10}{"долг":>12}{"пеня":>11}{"%366":>9}{"итого":>13}   ')
    print("-" * 74)

    failures, rate_notes = [], []
    for (cp, dog, dd), lines in groups.items():
        kind = classify_contract_type(dog)
        rate = DEFAULT_PENALTY_RATE[kind]
        rows, debt, penalty = docgen._penalty(lines, rate, AS_OF, index, cp, dog, dd)
        _, interest, _ = docgen._interest(lines, AS_OF, index, cp, dog, dd)
        total = round(debt + penalty + interest, 2)

        want = EXPECTED[dog]
        ok_hard = abs(debt - want[0]) < 0.005 and abs(penalty - want[1]) < 0.005
        ok_rate = abs(interest - want[2]) < 0.005
        if not ok_hard:
            failures.append(f'{dog}: долг/пеня {debt:.2f}/{penalty:.2f}, ждали {want[0]:.2f}/{want[1]:.2f}')
        if not ok_rate:
            rate_notes.append(f'{dog}: %366 {interest:.2f}, ждали {want[2]:.2f}')

        mark = "OK " if ok_hard and ok_rate else ("~  " if ok_hard else "БАГ")
        label = "биржевой" if kind == "exchange" else "прямой"
        print(f'{dog:<16}{label:<10}{money(debt):>12}{money(penalty):>11}'
              f'{money(interest):>9}{money(total):>13}   {mark}')

        for want_part, got in zip(EXPECTED_PARTS.get(dog, []),
                                  [p for r in rows for p in r["parts"]]):
            amount, days, fine = want_part
            if not (abs(got["сумма"] - amount) < 0.005 and got["дней"] == days
                    and abs(got["пеня"] - fine) < 0.005):
                failures.append(
                    f'{dog}: часть {got["сумма"]:.2f}×{got["дней"]}дн={got["пеня"]:.2f}, '
                    f'ждали {amount:.2f}×{days}дн={fine:.2f}')

    print("-" * 74)
    if failures:
        print("РАСХОЖДЕНИЯ В РАСЧЁТЕ — это ошибка:")
        for f in failures:
            print("  ", f)
    if rate_notes:
        print("Проценты ст.366 не сошлись. Это законно, если менялась ставка в "
              "data/refinancing_rate.csv:")
        for f in rate_notes:
            print("  ", f)
    if not failures and not rate_notes:
        print("Всё сходится.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
