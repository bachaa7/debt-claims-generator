# -*- coding: utf-8 -*-
"""Что в рабочих файлах настоящее, а что осталось от тестов. Ничего не удаляет.

Запуск:  python tests/audit_data.py [путь_к_свежему_отчёту.xls]

Тестовым считается всё, чьего номера договора нет в переданном отчёте.
Пригодится перед передачей проекта и перед выходом «в бой».
"""
import csv
import glob
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
sys.path.insert(0, PROJECT)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from parser import parse_debt_report

REPORT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    PROJECT, "data", "uploaded", "Отчет.xls")
if not os.path.exists(REPORT):
    sys.exit("Нет отчёта для сверки. Передайте путь: python tests/audit_data.py D:\\путь\\Отчет.xls")

lines = parse_debt_report(REPORT)
real_contracts = {l.contract_number for l in lines}
real_cps = {l.counterparty for l in lines}
print(f'Сверяю с {os.path.basename(REPORT)}: {len(real_contracts)} договоров, '
      f'{len(real_cps)} покупателей\n')


def rows_of(name):
    path = os.path.join(PROJECT, "data", name)
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


jr = rows_of("journal.csv")
alien_j = [r for r in jr if r["Договор"] not in real_contracts]
print(f'journal.csv — строк {len(jr)}, посторонних {len(alien_j)}')
for r in alien_j:
    print(f'   {r["Контрагент"]:<30}{r["Договор"]:<16}{r["Документ"]}')

cg = rows_of("contragents.csv")
alien_c = [r for r in cg if r["Контрагент"] not in real_cps]
filled = [r["Контрагент"] for r in cg if (r.get("УНП") or "").strip()]
print(f'\ncontragents.csv — покупателей {len(cg)}, с реквизитами {len(filled)}')
print(f'   {", ".join(filled) if filled else "реквизиты не заполнены ни у кого"}')
print(f'   нет в отчёте: {len(alien_c)}',
      [r["Контрагент"] for r in alien_c] if alien_c else "— все настоящие")

pm = rows_of("payments.csv")
alien_p = [r for r in pm if r["Договор"] not in real_contracts]
print(f'\npayments.csv — записей {len(pm)}, посторонних {len(alien_p)}')
for r in alien_p:
    print(f'   {r["Контрагент"]:<20}{r["Договор"]:<16}накл {r["Накладная"]:<10}'
          f'{r["Дата_оплаты"]}  {r["Сумма"]}')

docs = glob.glob(os.path.join(PROJECT, "output", "**", "*.docx"), recursive=True)
safe = {c.replace("/", "-") for c in real_contracts}
alien_d = [f for f in docs if os.path.basename(f).split("__")[-1][:-5] not in safe]
print(f'\noutput — документов {len(docs)}, посторонних {len(alien_d)}')
for f in sorted(alien_d):
    print("   ", os.path.relpath(f, os.path.join(PROJECT, "output")))

for folder in ("uploaded", "contracts"):
    files = glob.glob(os.path.join(PROJECT, "data", folder, "*"))
    print(f'\ndata/{folder} — файлов {len(files)}')
    for f in files:
        print("   ", os.path.basename(f))

total = len(alien_j) + len(alien_p) + len(alien_d)
print(f'\nИТОГ: посторонних записей и файлов — {total}.',
      "Чисто." if not total else "Удалять только вручную, посмотрев список выше.")
