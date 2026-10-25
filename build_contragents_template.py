# -*- coding: utf-8 -*-
"""Строит справочник реквизитов покупателей (contragents.csv) на основе отчёта по ДЗ.

Справочник заполняется ОДИН РАЗ (данными из папки договоров) и затем переиспользуется
при каждой генерации претензий/исков. Уже известные по присланным договорам компании
(СтальТрейд, АвтоЗавод) заполнены полностью — остальные нужно дозаполнить.
"""
import csv
from collections import defaultdict, Counter

from parser import parse_debt_report, classify_contract_type
from company import DEFAULT_PENALTY_RATE

KNOWN = {
    'СтальТрейд': dict(
        unp='100000002',
        address='220000, г. Минск, ул. Заводская, 2',
        bank_account='BY00TEST00000000000000000002',
        bank_name='ОАО "Банк-Пример", г. Минск, ул. Банковская, 2',
        bank_code='TESTBY2X',
        phone='+375 29 000-00-11',
        director_fio='Сергеев С.С.',
        director_basis='Устава',
    ),
    'АвтоЗавод': dict(
        unp='100000003',
        address='220000, г. Минск, ул. Автозаводская, 3',
        bank_account='BY00TEST00000000000000000003',
        bank_name='ОАО "Банк-Пример", г. Минск, ул. Банковская, 3',
        bank_code='TESTBY2X',
        phone='+375 29 000-00-12',
        director_fio='(уточнить по доверенности/уставу)',
        director_basis='',
    ),
}


def main(report_path: str, out_csv: str):
    """Дополняет справочник покупателями из отчёта. Существующие строки не трогает.

    Раньше этот скрипт переписывал файл с нуля и стирал заполненные реквизиты.
    Теперь он только добавляет недостающих — запускать можно сколько угодно раз.
    """
    from contracts import ensure_counterparties, _read_directory, _write_directory

    lines = parse_debt_report(report_path)
    rows, fieldnames = _read_directory(out_csv)

    # реквизиты, известные из присланных договоров, — только в пустые ячейки
    for row in rows:
        known = KNOWN.get(row.get("Контрагент", ""))
        if not known:
            continue
        for src, dst in (("unp", "УНП"), ("address", "Юр_адрес"),
                         ("bank_account", "Расчетный_счет"), ("bank_name", "Банк"),
                         ("bank_code", "Код_банка"), ("phone", "Телефон"),
                         ("director_fio", "Директор_ФИО"),
                         ("director_basis", "Директор_основание")):
            if not (row.get(dst) or "").strip():
                row[dst] = known.get(src, "")
    _write_directory(rows, fieldnames, out_csv)

    added = ensure_counterparties(lines, out_csv)
    total = len(_read_directory(out_csv)[0])
    print(f"Справочник {out_csv}: всего {total}, добавлено новых {len(added)}"
          + (f" — {', '.join(added)}" if added else ""))


if __name__ == '__main__':
    import sys
    import os
    here = os.path.dirname(os.path.abspath(__file__))
    report = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, 'Отчет.xls')
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(here, 'data', 'contragents.csv')
    main(report, out)
