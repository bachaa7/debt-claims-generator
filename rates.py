# -*- coding: utf-8 -*-
"""Ставка рефинансирования НБ РБ, проценты по ст. 366 ГК РБ и госпошлина.

Все меняющиеся законодательством числа лежат в data/ (refinancing_rate.csv и
tariffs.json) — их правят вручную, без изменения кода.
"""
import csv
import datetime as dt
import json
import os

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
REFI_PATH = os.path.join(DATA_DIR, "refinancing_rate.csv")
TARIFFS_PATH = os.path.join(DATA_DIR, "tariffs.json")

# По сложившейся в РБ практике проценты по ст. 366 ГК считают исходя из 360 дней
# в году. Если ваш юрист считает по фактическому числу дней — поставьте 365.
DAYS_IN_YEAR = 360


def load_refi_rates(path: str = REFI_PATH) -> list:
    """Возвращает [(дата_начала_действия, ставка_годовых_%)], отсортировано по дате."""
    rows = []
    if not os.path.exists(path):
        return rows
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            date_s = (r.get("Действует_с") or "").strip()
            rate_s = (r.get("Ставка_годовых_%") or "").strip().replace(",", ".")
            if not date_s or not rate_s:
                continue
            rows.append((dt.datetime.strptime(date_s, "%d.%m.%Y").date(), float(rate_s)))
    rows.sort()
    return rows


# Через сколько дней после последнего изменения ставки считать данные
# устаревшими и предупреждать. НБ РБ меняет ставку примерно раз в год-полтора.
RATE_STALE_DAYS = 180


def rate_freshness(as_of: dt.date, rates: list = None) -> dict:
    """Насколько свежа таблица ставок.

    Программа не знает, объявил ли Нацбанк новую ставку — она продолжит считать
    по последней записи в файле. Поэтому если запись давняя, надо предупредить.
    """
    rates = load_refi_rates() if rates is None else rates
    if not rates:
        return dict(устарела=True, последняя=None, ставка=None,
                    текст="Файл со ставкой рефинансирования пуст — проценты по ст. 366 не посчитать.")
    last_date, last_rate = rates[-1]
    days = (as_of - last_date).days
    if days > RATE_STALE_DAYS:
        return dict(
            устарела=True, последняя=last_date, ставка=last_rate,
            текст=(f'Последняя запись о ставке рефинансирования — {last_date.strftime("%d.%m.%Y")} '
                   f'({str(last_rate).replace(".", ",")}%), это было {days} дн. назад. '
                   f'Проверьте на сайте Нацбанка, не менялась ли ставка, и допишите строку '
                   f'в data/refinancing_rate.csv — иначе проценты считаются по устаревшей ставке.'),
        )
    return dict(устарела=False, последняя=last_date, ставка=last_rate, текст="")


# Открытый адрес Национального банка, откуда берётся история ставки.
# Никаких ваших данных туда не уходит — это обычный запрос на чтение.
NBRB_API = "https://api.nbrb.by/refinancingrate"
LAST_CHECK_PATH = os.path.join(DATA_DIR, "rate_last_check.txt")


def fetch_refi_rates(timeout: int = 15) -> list:
    """Скачивает историю ставки рефинансирования с сайта НБ РБ.

    Возвращает [(дата, ставка)]. Бросает исключение, если нет связи —
    вызывающий код обязан это обработать.
    """
    import json as _json
    import urllib.request

    req = urllib.request.Request(NBRB_API, headers={"User-Agent": "ClaimsGenerator/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = _json.load(resp)

    rows = []
    for item in data:
        raw_date = (item.get("Date") or "")[:10]
        value = item.get("Value")
        if not raw_date or value is None:
            continue
        rows.append((dt.date.fromisoformat(raw_date), float(value)))
    rows.sort()

    # оставляем только моменты, когда ставка реально менялась
    changes, prev = [], None
    for day, rate in rows:
        if rate != prev:
            changes.append((day, rate))
            prev = rate
    return changes


def update_refi_file(path: str = REFI_PATH, timeout: int = 15) -> dict:
    """Дописывает в файл ставки, которых там ещё нет.

    Существующие строки не трогает — ваши пометки в колонке «Источник»
    остаются на месте. Возвращает dict(ok=, добавлено=, текст=).
    """
    try:
        fetched = fetch_refi_rates(timeout)
    except Exception as e:
        return dict(ok=False, добавлено=0, последняя=None, текст=(
            f"Не удалось связаться с сайтом Национального банка ({type(e).__name__}). "
            f"Проверьте интернет или впишите ставку вручную в data/refinancing_rate.csv."))

    if not fetched:
        return dict(ok=False, добавлено=0, последняя=None,
                    текст="Национальный банк вернул пустой список — ничего не изменено.")

    existing_rows = []
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            existing_rows = [r for r in csv.DictReader(f) if (r.get("Действует_с") or "").strip()]
    have = {(r["Действует_с"] or "").strip() for r in existing_rows}

    added = []
    for day, rate in fetched:
        key = day.strftime("%d.%m.%Y")
        if key in have:
            continue
        added.append({"Действует_с": key,
                      "Ставка_годовых_%": f"{rate:g}",
                      "Источник": "api.nbrb.by, загружено автоматически"})

    if not added:
        last = fetched[-1]
        return dict(ok=True, добавлено=0, последняя=last[0], текст=(
            f'Ставка актуальна: последнее изменение {last[0].strftime("%d.%m.%Y")} — '
            f'{str(last[1]).replace(".", ",")}%. Новых записей нет.'))

    rows = existing_rows + added
    rows.sort(key=lambda r: dt.datetime.strptime(r["Действует_с"], "%d.%m.%Y").date())
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["Действует_с", "Ставка_годовых_%", "Источник"])
        w.writeheader()
        w.writerows(rows)

    last = added[-1]
    return dict(ok=True, добавлено=len(added),
                последняя=dt.datetime.strptime(last["Действует_с"], "%d.%m.%Y").date(),
                текст=(f'Добавлено записей: {len(added)}. Последняя — с {last["Действует_с"]}, '
                       f'ставка {last["Ставка_годовых_%"].replace(".", ",")}%.'))


def checked_today(path: str = LAST_CHECK_PATH) -> bool:
    """Проверяли ли ставку сегодня (чтобы не дёргать сайт при каждом запуске)."""
    try:
        with open(path, encoding="utf-8") as f:
            return f.read().strip() == dt.date.today().isoformat()
    except OSError:
        return False


def mark_checked(path: str = LAST_CHECK_PATH):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(dt.date.today().isoformat())
    except OSError:
        pass


def rate_on(day: dt.date, rates: list):
    """Ставка, действовавшая на указанную дату, или None, если данных нет."""
    result = None
    for start, rate in rates:
        if start <= day:
            result = rate
        else:
            break
    return result


def interest_366(amount: float, first_day: dt.date, last_day: dt.date,
                 rates: list = None, days_in_year: int = DAYS_IN_YEAR):
    """Проценты за пользование чужими денежными средствами (ст. 366 ГК РБ).

    Начисляются с first_day по last_day включительно. Если за это время ставка
    рефинансирования менялась, период разбивается на отрезки и по каждому
    считается своя ставка.

    Возвращает (сумма, [отрезки], дней_без_данных_о_ставке).
    Каждый отрезок: dict(с=, по=, дней=, ставка=, сумма=).
    """
    rates = load_refi_rates() if rates is None else rates
    if not rates or last_day < first_day:
        return 0.0, [], max(0, (last_day - first_day).days + 1) if last_day >= first_day else 0

    # границы отрезков — начало периода + все даты изменения ставки внутри него
    bounds = [first_day]
    for start, _ in rates:
        if first_day < start <= last_day:
            bounds.append(start)
    bounds.append(last_day + dt.timedelta(days=1))

    total = 0.0
    periods = []
    days_without_rate = 0
    for i in range(len(bounds) - 1):
        seg_start = bounds[i]
        seg_end = bounds[i + 1] - dt.timedelta(days=1)
        days = (seg_end - seg_start).days + 1
        if days <= 0:
            continue
        rate = rate_on(seg_start, rates)
        if rate is None:
            # ставка на этот период неизвестна — молча считать нельзя
            days_without_rate += days
            continue
        amount_seg = round(amount * rate / 100 / days_in_year * days, 2)
        total += amount_seg
        periods.append(dict(с=seg_start, по=seg_end, дней=days, ставка=rate, сумма=amount_seg))

    return round(total, 2), periods, days_without_rate


def _num(x) -> str:
    """Число в привычном виде: 5 вместо 5.0, 1 050,00 вместо 1050.0."""
    x = float(x)
    if x == int(x) and abs(x) < 1000:
        return str(int(x))
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


def load_tariffs(path: str = TARIFFS_PATH) -> dict:
    if not os.path.exists(path):
        return {"проверено": False}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def state_duty(claim_amount: float, kind: str = "исковое_экономсуд", tariffs: dict = None) -> dict:
    """Госпошлина (экономсуд) или арбитражный сбор (БУТБ) от цены иска.

    kind: 'исковое_экономсуд' | 'приказное_экономсуд' | 'арбитражный_сбор_БУТБ'.

    Возвращает dict(сумма=, проверено=, описание=). Пока в tariffs.json стоит
    "проверено": false, значение считается ориентировочным и в текст искового
    заявления не подставляется.
    """
    tariffs = load_tariffs() if tariffs is None else tariffs
    verified = bool(tariffs.get("проверено"))

    if kind == "арбитражный_сбор_БУТБ":
        cfg = tariffs.get("арбитражный_сбор_БУТБ", {})
        amount = claim_amount * float(cfg.get("процент_от_цены_иска", 0)) / 100
        amount = max(amount, float(cfg.get("минимум_BYN", 0)))
        description = (
            f'{_num(cfg.get("процент_от_цены_иска", 0))}% от цены иска '
            f'по Регламенту арбитражной комиссии БУТБ'
        )
    else:
        cfg = tariffs.get("госпошлина", {}).get(kind, {})
        base_value = float(tariffs.get("базовая_величина_BYN", 0))
        pct = float(cfg.get("процент_от_цены_иска", 0))
        min_bv = float(cfg.get("минимум_базовых_величин", 0))
        amount = max(claim_amount * pct / 100, min_bv * base_value)
        description = (
            f'{_num(pct)}% от цены иска, но не менее {_num(min_bv)} базовых величин '
            f'({_num(round(min_bv * base_value, 2))} руб. при базовой величине {_num(base_value)} руб.)'
        )

    return dict(сумма=round(amount, 2), проверено=verified, описание=description)
