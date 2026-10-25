# -*- coding: utf-8 -*-
"""Частичные оплаты по накладным.

В отчёте по ДЗ виден только остаток долга — сколько и когда должник заплатил,
там не написано. А пеня считается по частям: каждая погашенная часть «живёт»
до своей даты оплаты. Поэтому платежи хранятся отдельно, в data/payments.csv.

Правило расчёта (как в судебной практике по частичной оплате):

    долг по ТТН 10 000,00, просрочка с 28.01
    06.02 п/п №101 оплачено 3 000,00  ->  28.01–05.02 = 9 дн.
    09.02 п/п №102 оплачено 2 000,00  ->  28.01–08.02 = 12 дн.
    остаток          5 000,00          ->  28.01–18.02 = 22 дн.

То есть день оплаты в просрочку не входит, а остаток считается до даты документа.
"""
import csv
import datetime as dt
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PAYMENTS_PATH = os.path.join(BASE_DIR, "data", "payments.csv")

FIELDS = ["Контрагент", "Договор", "Дата_договора", "Накладная",
          "Дата_оплаты", "Сумма", "Номер_ПП", "Назначение"]


def _date(value):
    value = (value or "").strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def _amount(value) -> float:
    try:
        return round(float(str(value).replace(" ", "").replace(" ", "").replace(",", ".")), 2)
    except (TypeError, ValueError):
        return 0.0


def load_payments(path: str = PAYMENTS_PATH) -> list:
    """Список платежей как есть, строками из файла."""
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig") as f:
        return [dict(r) for r in csv.DictReader(f)]


def save_payments(rows: list, path: str = PAYMENTS_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in FIELDS})


def _key(row: dict, field: str) -> str:
    return (row.get(field) or "").strip()


def add_payment(row: dict, path: str = PAYMENTS_PATH) -> dict:
    """Добавляет платёж. Возвращает dict(ok=, текст=)."""
    when = _date(row.get("Дата_оплаты"))
    amount = _amount(row.get("Сумма"))
    if not row.get("Контрагент") or not row.get("Договор"):
        return dict(ok=False, текст="Не указан покупатель или договор.")
    if not row.get("Накладная"):
        return dict(ok=False, текст="Не указана накладная — пеня считается по каждой отдельно.")
    if when is None:
        return dict(ok=False, текст="Дата оплаты должна быть в виде ДД.ММ.ГГГГ.")
    if amount <= 0:
        return dict(ok=False, текст="Сумма оплаты должна быть больше нуля.")

    rows = load_payments(path)
    entry = {k: (row.get(k) or "").strip() for k in FIELDS}
    entry["Дата_оплаты"] = when.strftime("%d.%m.%Y")
    entry["Сумма"] = f"{amount:.2f}"

    # Тот же платёж дважды — почти наверняка описка. Пеня от этого занижается,
    # а сумма поставки в иске завышается, и в расчёте это выглядит как две
    # одинаковых строки подряд — глаз их не ловит. Два реальных платежа одного
    # размера в один день различаются номером п/п, так что их пропускаем.
    same = [r for r in rows
            if all(_key(r, k) == _key(entry, k) for k in
                   ("Контрагент", "Договор", "Дата_договора", "Накладная",
                    "Дата_оплаты", "Сумма", "Номер_ПП"))]
    if same:
        пп = f' п/п №{entry["Номер_ПП"]}' if entry["Номер_ПП"] else ""
        return dict(ok=False,
                    текст=f'Такая оплата уже внесена: {amount:.2f} руб. '
                          f'от {entry["Дата_оплаты"]}{пп} по накладной '
                          f'{entry["Накладная"]}. Если это второй платёж — '
                          f'укажите его номер платёжного поручения.')

    rows.append(entry)
    save_payments(rows, path)
    return dict(ok=True, текст=f'Оплата {amount:.2f} руб. от {entry["Дата_оплаты"]} записана.')


def delete_payment(index: int, path: str = PAYMENTS_PATH) -> dict:
    rows = load_payments(path)
    if not 0 <= index < len(rows):
        return dict(ok=False, текст="Такой записи нет.")
    removed = rows.pop(index)
    save_payments(rows, path)
    return dict(ok=True, текст=f'Удалена оплата {removed.get("Сумма")} руб. '
                              f'от {removed.get("Дата_оплаты")}.')


def build_index(path: str = PAYMENTS_PATH) -> dict:
    """{(контрагент, договор, дата_договора, накладная): [(дата, сумма, № п/п, назначение)]}

    Платежи внутри каждой накладной отсортированы по дате.
    """
    index = {}
    for r in load_payments(path):
        when = _date(r.get("Дата_оплаты"))
        amount = _amount(r.get("Сумма"))
        if when is None or amount <= 0:
            continue
        key = (r.get("Контрагент", "").strip(), r.get("Договор", "").strip(),
               r.get("Дата_договора", "").strip(), r.get("Накладная", "").strip())
        index.setdefault(key, []).append(
            dict(дата=when, сумма=amount, пп=r.get("Номер_ПП", "").strip(),
                 назначение=r.get("Назначение", "").strip()))
    for key in index:
        index[key].sort(key=lambda p: p["дата"])
    return index


def for_line(index: dict, counterparty: str, contract: str, contract_date: str,
             doc_number: str, as_of=None) -> list:
    """Платежи по конкретной накладной. Дата договора в ключе необязательна.

    as_of — дата документа: платежи, прошедшие позже неё, не возвращаются.
    Иначе на старой выгрузке (долг ещё полный) в документ попадали оплаты,
    которых на его дату ещё не было.
    """
    if not index:
        return []
    rows = index.get((counterparty, contract, contract_date, doc_number))
    if rows is None:
        # запись могли внести без даты договора
        rows = index.get((counterparty, contract, "", doc_number), [])
    if as_of is None:
        return rows
    return [p for p in rows if p["дата"] <= as_of]
