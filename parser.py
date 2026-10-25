# -*- coding: utf-8 -*-
"""
Разбор отчёта по дебиторской задолженности 1С ("Отчет по ДЗ 60 дней")
в структурированный список долгов.

Ожидаемый формат (после экспорта из 1С в Excel и открытия/сохранения как .xls/.xlsx):
№ п/п | Контрагент (в этой же ячейке второй строкой - "Договор №... от ...") |
Срок оплаты (ддммгг без разделителей, напр. 210726 = 21.07.26) |
Документ (напр. "Руб.накл 0424502 (08.06.26)") | ВСЕГО | Просроченная задолженность | Дней просрочено
"""
import re
import datetime as dt
from dataclasses import dataclass, field
from typing import Optional
import pandas as pd


# Между словом "Договор" и номером в отчёте 1С встречается что угодно:
# "Договор № 01/13711-155", "Договор возмездного оказания услуг №030521",
# "Договор возм.оказ.услуг №04/01/19", "Договор N01/13803-152" (латинская N).
CONTRACT_RE = re.compile(
    r"Договор.*?[№N]\s*([^\s]+)\s*от\s*([\d.]+)",
    re.IGNORECASE,
)
DOC_RE = re.compile(r"([^\d(]+)\s*([^\s(]+)\s*\(([\d.]+)\)")


@dataclass
class DebtLine:
    counterparty: str
    contract_number: str
    contract_date: str  # dd.mm.yy как в отчёте
    due_date: dt.date
    doc_type: str
    doc_number: str
    doc_date: str  # dd.mm.yy
    total: float
    overdue_amount: float
    overdue_days: int
    contract_type: str = ""  # 'exchange' | 'direct' — заполняется классификатором


def _parse_ddmmyy_compact(value) -> dt.date:
    """Преобразует 210726 / '210726' / '20826' в дату (2026-07-21 / 2026-08-02)."""
    s = str(int(value)) if not isinstance(value, str) else value.strip()
    s = s.zfill(6)
    day = int(s[0:2])
    month = int(s[2:4])
    year = 2000 + int(s[4:6])
    return dt.date(year, month, day)


def _parse_ddmmyy_dotted(s: str) -> str:
    return s.strip()


def parse_debt_report(path: str, sheet_name=0) -> list[DebtLine]:
    """Читает выгрузку отчёта 1С (xls/xlsx/csv) и возвращает список DebtLine.

    Строка "ВСЕГО:" и пустые строки пропускаются.
    """
    if path.lower().endswith(".csv"):
        df = pd.read_csv(path, header=None)
    elif path.lower().endswith(".xls"):
        # Старый бинарный .xls из 1С не содержит записи CODEPAGE, поэтому xlrd по
        # умолчанию читает кириллицу как iso-8859-1 ("Àìêîäîð-ÓÍÈÊÀÁ").
        # Явно задаём cp1251 — это работает без LibreOffice, на любой машине.
        import xlrd

        book = xlrd.open_workbook(path, encoding_override="cp1251")
        sheet = book.sheet_by_index(sheet_name if isinstance(sheet_name, int) else 0)
        df = pd.DataFrame([[c.value for c in sheet.row(r)] for r in range(sheet.nrows)])
    else:
        df = pd.read_excel(path, header=None)

    lines: list[DebtLine] = []
    for _, row in df.iterrows():
        vals = list(row.values)
        # находим ячейку с "Договор №"
        cp_cell = None
        for v in vals:
            if isinstance(v, str) and "Договор" in v:
                cp_cell = v
                break
        if cp_cell is None:
            continue

        parts = cp_cell.split("\n")
        counterparty = parts[0].strip()
        contract_line = parts[1] if len(parts) > 1 else ""
        m = CONTRACT_RE.search(contract_line)
        if not m:
            continue
        contract_number, contract_date = m.group(1), m.group(2)

        # остальные значения — по позиции относительно найденной ячейки
        idx = vals.index(cp_cell)
        due_raw = vals[idx + 1]
        doc_raw = vals[idx + 2]
        total_raw = vals[idx + 3]
        overdue_raw = vals[idx + 4]
        days_raw = vals[idx + 5]

        try:
            due_date = _parse_ddmmyy_compact(due_raw)
        except Exception:
            continue

        doc_str = str(doc_raw).strip()
        dm = DOC_RE.search(doc_str)
        if dm:
            doc_type, doc_number, doc_date = dm.group(1).strip(), dm.group(2).strip(), dm.group(3).strip()
        else:
            doc_type, doc_number, doc_date = doc_str, "", ""

        def to_float(x):
            try:
                return round(float(x), 2)
            except Exception:
                return 0.0

        total = to_float(total_raw)
        overdue_amount = to_float(overdue_raw)
        try:
            overdue_days = int(float(days_raw))
        except Exception:
            overdue_days = 0

        if overdue_amount <= 0 or overdue_days <= 0:
            continue

        lines.append(
            DebtLine(
                counterparty=counterparty,
                contract_number=contract_number,
                contract_date=contract_date,
                due_date=due_date,
                doc_type=doc_type,
                doc_number=doc_number,
                doc_date=doc_date,
                total=total,
                overdue_amount=overdue_amount,
                overdue_days=overdue_days,
            )
        )
    return lines


def classify_contract_type(contract_number: str) -> str:
    """Биржевые договора БУТБ регистрируются в формате '01/NNNNN-NNN'.
    Прямые договора обычно 'ддммгг/N' или произвольная внутренняя нумерация.
    """
    if re.match(r"^0?1/\d{4,6}-\d{2,4}$", contract_number.strip()):
        return "exchange"
    return "direct"
