# -*- coding: utf-8 -*-
"""Реестр (журнал) документов: что по какому договору уже сформировано,
отправлено ли, истёк ли срок ответа на претензию, оплачено ли.

Журнал — обычный CSV (data/journal.csv), его можно открыть в Excel и править
руками. Программа сама дописывает строки при формировании документов и
пересчитывает статус при каждой новой выгрузке из 1С.
"""
import csv
import datetime as dt
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JOURNAL_PATH = os.path.join(BASE_DIR, "data", "journal.csv")

# Срок, который даётся на ответ по претензии (совпадает с текстом претензии)
PRETENZIA_DEADLINE_DAYS = 10

FIELDS = [
    "Контрагент", "Договор", "Дата_договора", "Документ",
    "Дата_формирования", "Долг_на_момент", "Итого_на_момент", "Файл",
    "Дата_отправки", "Дата_оплаты", "Примечание",
]

# Статусы, которые программа вычисляет сама (в файле не хранятся)
NOT_MADE = "не формировалась"
MADE = "сформирована"
WAITING = "ждём ответа"
EXPIRED = "срок истёк — можно в суд"
PAID = "оплачено"
PART_PAID = "оплачено частично"
SUIT_MADE = "иск сформирован"


def _d(value: str):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def _f(value) -> float:
    try:
        return float(str(value).replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


def load_journal(path: str = JOURNAL_PATH) -> list:
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig") as f:
        return [dict(r) for r in csv.DictReader(f)]


def save_journal(rows: list, path: str = JOURNAL_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in FIELDS})


def _key(row: dict) -> tuple:
    return (row.get("Контрагент", ""), row.get("Договор", ""),
            row.get("Дата_договора", ""), row.get("Документ", ""))


def record(entries: list, path: str = JOURNAL_PATH):
    """Дописывает/обновляет строки журнала.

    Если по этому договору документ такого типа уже формировался, строка
    обновляется (дата, суммы, файл), но отметки об отправке и оплате
    сохраняются — их проставляет человек.
    """
    rows = load_journal(path)
    index = {_key(r): r for r in rows}
    for e in entries:
        existing = index.get(_key(e))
        if existing:
            for field in ("Дата_формирования", "Долг_на_момент", "Итого_на_момент", "Файл"):
                existing[field] = e.get(field, existing.get(field, ""))
        else:
            rows.append(e)
            index[_key(e)] = e
    save_journal(rows, path)
    return rows


def mark_sent(keys: list, sent_date: dt.date, path: str = JOURNAL_PATH) -> int:
    """Проставляет дату отправки претензий. keys — [(контрагент, договор, дата_договора)]."""
    rows = load_journal(path)
    wanted = {tuple(k) for k in keys}
    count = 0
    for r in rows:
        if r.get("Документ") != "претензия":
            continue
        if (r.get("Контрагент"), r.get("Договор"), r.get("Дата_договора")) in wanted:
            r["Дата_отправки"] = sent_date.strftime("%d.%m.%Y")
            count += 1
    save_journal(rows, path)
    return count


def build_state(path: str = JOURNAL_PATH) -> dict:
    """Журнал в виде {(контрагент, договор, дата_договора): {'претензия': row, 'иск': row}}."""
    state = {}
    for r in load_journal(path):
        key = (r.get("Контрагент", ""), r.get("Договор", ""), r.get("Дата_договора", ""))
        state.setdefault(key, {})[r.get("Документ", "")] = r
    return state


def status_for(key: tuple, state: dict, current_debt: float, today: dt.date = None) -> dict:
    """Что сейчас с претензией по этому договору.

    current_debt — сумма долга в свежей выгрузке из 1С (0, если долга больше нет).
    Возвращает dict(статус=, срок_ответа=, можно_иск=, подсказка=).
    """
    today = today or dt.date.today()
    entry = state.get(key, {})
    pret = entry.get("претензия")
    isk = entry.get("иск")

    if not pret:
        return dict(статус=NOT_MADE, срок_ответа=None, можно_иск=False,
                    подсказка="Сначала претензия.")

    was = _f(pret.get("Долг_на_момент"))
    paid_date = _d(pret.get("Дата_оплаты"))
    sent = _d(pret.get("Дата_отправки"))

    if paid_date or current_debt <= 0:
        return dict(статус=PAID, срок_ответа=None, можно_иск=False,
                    подсказка="Долг погашен — иск не нужен.")

    if was and current_debt < was - 0.01:
        deadline = sent + dt.timedelta(days=PRETENZIA_DEADLINE_DAYS) if sent else None
        return dict(статус=PART_PAID, срок_ответа=deadline,
                    можно_иск=bool(deadline and today > deadline),
                    подсказка=f"Было {was:.2f}, стало {current_debt:.2f} — платят частями.")

    if not sent:
        return dict(статус=MADE, срок_ответа=None, можно_иск=False,
                    подсказка=f'Сформирована {pret.get("Дата_формирования", "")}, '
                              f'отметьте дату отправки.')

    deadline = sent + dt.timedelta(days=PRETENZIA_DEADLINE_DAYS)
    if today <= deadline:
        return dict(статус=WAITING, срок_ответа=deadline, можно_иск=False,
                    подсказка=f'Ответ ждём до {deadline.strftime("%d.%m.%Y")}.')

    if isk:
        return dict(статус=SUIT_MADE, срок_ответа=deadline, можно_иск=True,
                    подсказка=f'Иск сформирован {isk.get("Дата_формирования", "")}.')

    return dict(статус=EXPIRED, срок_ответа=deadline, можно_иск=True,
                подсказка=f'Срок истёк {deadline.strftime("%d.%m.%Y")}, оплаты нет.')
