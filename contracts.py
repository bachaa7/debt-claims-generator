# -*- coding: utf-8 -*-
"""Извлечение реквизитов покупателя из файла договора.

Понимает .docx, .doc, .rtf (через установленный Word) и .pdf с текстовым слоем.
Скан-картинку без текста разобрать нельзя — про это программа честно сообщает.

Работает так: в договоре две стороны, наша и покупателя. Всё, что совпадает с
реквизитами ООО «Ромашка», отбрасывается — остальное и есть покупатель.
"""
import os
import re
import tempfile

from company import SELLER

OUR_UNP = SELLER["unp"]
OUR_IBAN = SELLER["bank_account"]

SUPPORTED = (".docx", ".doc", ".rtf", ".pdf")

# В документах реквизиты набирают вперемешку: в одном из исков счёт записан как
# «BY00АВВВ300000000...», где АКВВ — кириллица. Для человека неотличимо,
# для программы — другие символы. Приводим двойники к латинице.
HOMOGLYPHS = str.maketrans({
    "А": "A", "В": "B", "С": "C", "Е": "E", "К": "K", "М": "M", "Н": "H",
    "О": "O", "Р": "P", "Т": "T", "У": "Y", "Х": "X", "І": "I",
})


def _latin(text: str) -> str:
    """«BY00АВВВ3000…» -> «BY00ABBB3000…»"""
    return (text or "").translate(HOMOGLYPHS)

# --- регулярные выражения по реквизитам --------------------------------------
RE_UNP = re.compile(r"УНП[\s:№]*([0-9]{9})")
RE_IBAN = re.compile(r"\b(BY\s*\d{2}(?:\s*[A-ZАВСЕКМНОРТУХІ]){4}(?:\s*[0-9A-ZАВСЕКМНОРТУХІ]){20})", re.I)
RE_ACCOUNT_LABEL = re.compile(r"(?:Сч\.?\s*№|р/с|расч[её]тный\s+счет|IBAN)", re.I)
RE_SWIFT = re.compile(r"(?:SWIFT|БИК|код(?:\s*банка)?)[\s:]*([A-ZАВСЕКМНОРТУХІ]{6}[0-9A-ZАВСЕКМНОРТУХІ]{2,5})", re.I)
RE_SWIFT_ANY = re.compile(r"\b([A-Z]{4}BY2[0-9A-Z])\b")
RE_ADDRESS = re.compile(
    r"(?:Юридический\s+адрес|Место\s+нахождения|Адрес)\s*:?\s*([^|\n]{10,160})", re.I)
RE_PHONE = re.compile(r"(?:тел(?:ефон)?\.?/?ф?\.?)\s*:?\s*([+()\d][\d\s\-()]{6,22})", re.I)
# Адрес не всегда подписан словом «адрес»: в иске он идёт сразу после
# наименования и часто разорван переносом строки —
#   220000, Минская область, г.Минск,
#   ул.Примерная, д.5
# поэтому переносы внутри допускаем, а заканчиваем на следующем реквизите.
RE_ADDRESS_INDEX = re.compile(
    r"\b(\d{6}\s*,\s*[^|]{5,140}?)"
    r"(?=\s*(?:тел\.?|Тел\.?|[Рр]/с|Сч\.?\s*№|IBAN|УНП|ОКПО|ОКЮЛП|e-?mail|"
    r"код\s*банка|$))", re.I)
# В иске сторона называется «Ответчик», а не «Покупатель». Блок тянется
# до цены иска, номера дела или заголовка документа.
RE_DEFENDANT = re.compile(
    r"Ответчик\s*:?\s*(.{20,1200}?)"
    r"(?=\n\s*(?:Цена\s+иска|Дело\s*№|Исковое\s+заявление|Заявление|"
    r"Третье\s+лицо|ПРОШУ)|$)", re.I | re.S)
RE_SIGNER = re.compile(r"в\s+лице\s+([^,]{5,90}?)\s*,\s*действующ", re.I)
# «… именуемое в дальнейшем Покупатель, в лице X, действующего на основании Y»
RE_BUYER_SIGNER = re.compile(
    r"Покупател[ьяю][^.]{0,200}?в\s+лице\s+([^,]{5,90}?)\s*,\s*действующ[^,]*?"
    r"на\s+основании\s+([^,;]{4,80}?)(?=\s*г\.|\s*,|\s*;)", re.I | re.S)
RE_BASIS = re.compile(r"на\s+основании\s+([^,;]{4,80}?)(?=\s*г\.|\s*,|\s*;|\s*$)", re.I)
RE_ORG = re.compile(
    r"((?:Открытое\s+акционерное\s+общество|Закрытое\s+акционерное\s+общество|"
    r"Общество\s+с\s+ограниченной\s+ответственностью|Унитарное\s+предприятие|"
    r"Частное\s+(?:торговое\s+|производственное\s+)?унитарное\s+предприятие|"
    r"Совместное\s+общество|Индивидуальный\s+предприниматель|ОАО|ЗАО|ООО|ЧУП|УП|СП|ИП)"
    r"\s*[«\"']?[^«\"'»\n|]{2,120}[»\"']?)")


# ---------------------------------------------------------------------------
# Чтение файла
# ---------------------------------------------------------------------------

def _read_docx(path):
    from docx import Document
    d = Document(path)
    paragraphs = [" ".join(p.text.split()) for p in d.paragraphs if p.text.strip()]
    tables = []
    for t in d.tables:
        tables.append([[" ".join(c.text.split()) for c in row.cells] for row in t.rows])
    return paragraphs, tables


def _via_word(path):
    """Конвертируем .doc/.rtf в .docx установленным Word и читаем уже его.

    Сервер обрабатывает каждый запрос в своём потоке, а COM (механизм, через
    который Python управляет Word) нужно включать отдельно в каждом потоке —
    иначе со второго раза Word не запускается.
    """
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    app = None
    try:
        tmp = os.path.join(tempfile.mkdtemp(), "converted.docx")
        app = win32.Dispatch("Word.Application")
        app.Visible = False
        app.DisplayAlerts = False
        doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True)
        doc.SaveAs2(tmp, FileFormat=12)
        doc.Close(False)
        return _read_docx(tmp)
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _read_pdf(path):
    import pdfplumber
    paragraphs, tables = [], []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                line = " ".join(line.split())
                if line:
                    paragraphs.append(line)
            for t in page.extract_tables() or []:
                tables.append([[" ".join((c or "").split()) for c in row] for row in t])
    return paragraphs, tables


def read_document(path):
    """(абзацы, таблицы). Бросает ValueError с понятным текстом при неудаче."""
    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED:
        raise ValueError(f"Формат {ext} не поддерживается. Нужен .docx, .doc, .rtf или .pdf.")
    if ext == ".docx":
        paragraphs, tables = _read_docx(path)
    elif ext in (".doc", ".rtf"):
        try:
            paragraphs, tables = _via_word(path)
        except Exception as e:
            raise ValueError(
                f"Не удалось открыть {ext} — для этого формата нужен установленный "
                f"Microsoft Word ({type(e).__name__}). Пересохраните договор как .docx.")
    else:
        paragraphs, tables = _read_pdf(path)

    if not any(len(p) > 20 for p in paragraphs) and not tables:
        raise ValueError(
            "В файле нет текста — похоже, это скан (картинка). Программа не умеет "
            "распознавать изображения. Возьмите договор в электронном виде "
            "(.docx/.doc/.rtf) или наберите реквизиты вручную.")
    return paragraphs, tables


# ---------------------------------------------------------------------------
# Разбор
# ---------------------------------------------------------------------------

def _is_ours(text: str) -> bool:
    """Наш ли это кусок документа.

    Счёт у компании в разных договорах разный (бывает в разных банках),
    поэтому опираемся на УНП, название и метку «Поставщик».
    """
    if OUR_UNP in text or OUR_IBAN in text or "омашка" in text:
        return True
    head = text.strip()[:40].lower()
    return head.startswith("поставщик") or head.startswith("продавец")


def _buyer_text(paragraphs, tables) -> str:
    """Собирает куски документа, относящиеся к покупателю, а не к нам."""
    # Исковое заявление: блок «Ответчик:» — самый надёжный источник,
    # там реквизиты уже собраны в одном месте.
    m = RE_DEFENDANT.search("\n".join(paragraphs))
    if m:
        block = m.group(1)
        if [u for u in RE_UNP.findall(block) if u != OUR_UNP]:
            return block

    parts = []

    # 1. Двухколоночные таблицы реквизитов: наша колонка отбрасывается
    for table in tables:
        widths = {len(row) for row in table}
        if widths == {2}:
            col0 = " ".join(row[0] for row in table)
            col1 = " ".join(row[1] for row in table)
            if _is_ours(col0) and not _is_ours(col1):
                parts.append(col1)
            elif _is_ours(col1) and not _is_ours(col0):
                parts.append(col0)
            else:
                # метка «Покупатель» в первой ячейке — берём вторую
                for row in table:
                    if re.fullmatch(r"Покупатель", row[0].strip(), re.I):
                        parts.append(row[1])
        else:
            for row in table:
                joined = " ".join(row)
                if not _is_ours(joined):
                    parts.append(joined)

    # 2. Текст после слова «Покупатель» — типовой блок реквизитов сторон
    whole = "\n".join(paragraphs)
    m = re.search(r"(Покупатель|ПОКУПАТЕЛЬ)\s*:?\s*(.{0,1500})", whole, re.S)
    if m:
        tail = m.group(2)
        parts.extend(chunk for chunk in tail.split("\n") if not _is_ours(chunk))

    # 3. Если ничего не нашлось — весь документ без наших строк
    if not parts:
        parts = [p for p in paragraphs if not _is_ours(p)]

    return "\n".join(parts)


def _first(pattern, text, group=1):
    m = pattern.search(text)
    return m.group(group).strip(" .,;:") if m else ""


def _clean_basis(text: str) -> str:
    """«доверенности №73 от 23.12.24г с одной стороны» -> «доверенности №73 от 23.12.24»."""
    text = " ".join((text or "").split())
    text = re.split(r"\s+с\s+одной\s+стороны|\s+с\s+другой\s+стороны", text, flags=re.I)[0]
    text = re.sub(r"\s*г\.?\s*$", "", text)
    return text.strip(" .,;:")


def _clean_org(name: str) -> str:
    name = " ".join(name.split()).strip(" .,;:")
    name = re.sub(r'"([^"]+)"', r"«\1»", name)
    return name


def _signer_from_tables(tables):
    """Бланк БУТБ помечает поля подписью снизу: «(фамилия, имя, отчество)» и
    «(устава/доверенности № от …)». Берём те, что идут после строки «Покупатель»."""
    signer = basis = ""
    seen_buyer = False
    for table in tables:
        flat = [c for row in table for c in row]
        if any(re.fullmatch(r"Покупатель", c.strip(), re.I) for c in flat):
            seen_buyer = True
            continue
        if not seen_buyer:
            continue
        hint = " ".join(flat).lower()
        value = flat[0].strip() if flat else ""
        if not value or value.startswith("("):
            continue
        if "фамилия" in hint and not signer:
            signer = value
        elif ("устава" in hint or "доверенности" in hint) and not basis:
            basis = value
    return signer, basis


def extract_requisites(paragraphs, tables) -> dict:
    """Возвращает найденные реквизиты покупателя (пустая строка — не найдено)."""
    text = _buyer_text(paragraphs, tables)
    whole = "\n".join(paragraphs) + "\n" + "\n".join(
        " ".join(row) for table in tables for row in table)

    # полное наименование: сначала явная метка «Покупатель», затем поиск по тексту
    name = ""
    for table in tables:
        for row in table:
            if len(row) >= 2 and re.fullmatch(r"Покупатель", row[0].strip(), re.I):
                name = row[1]
                break
        if name:
            break
    if not name:
        # в иске наименование часто разорвано переносом строки
        flat = " ".join(text.split())
        for candidate in RE_ORG.findall(flat) or RE_ORG.findall(text):
            if "омашка" not in candidate and "товарная биржа" not in candidate.lower():
                name = candidate
                break

    unp = ""
    for found in RE_UNP.findall(text) or RE_UNP.findall(whole):
        if found != OUR_UNP:
            unp = found
            break

    iban = ""
    for found in RE_IBAN.findall(text) or RE_IBAN.findall(whole):
        candidate = _latin(re.sub(r"\s+", "", found)).upper()
        if candidate != OUR_IBAN and len(candidate) == 28:
            iban = candidate
            break

    swift = _latin(_first(RE_SWIFT, text) or _first(RE_SWIFT_ANY, text)).upper()
    if swift and swift == SELLER["bank_code"]:
        swift = ""

    address = _first(RE_ADDRESS, text) or _first(RE_ADDRESS_INDEX, text)
    if address:
        address = re.sub(r"^РЕСПУБЛИКА\s+БЕЛАРУСЬ,\s*", "", address, flags=re.I)
        # адрес заканчивается там, где начинается следующий реквизит
        address = re.split(r"\s*(?:УНП|Расчетный\s+счет|р/с|IBAN|SWIFT|БИК|код\s*банка|"
                           r"тел|Тел|e-?mail|Email)\b", address)[0]
        address = " ".join(address.split()).strip(" .,;:|")

    # банк — между IBAN и SWIFT в строке «Расчетный счет: IBAN ..., ЗАО «...», ...»
    bank = ""
    if iban:
        # счёт в тексте может быть записан с пробелами — ищем по «гибкому» шаблону
        spaced = r"\s*".join(re.escape(ch) for ch in iban)
        m = re.search(spaced + r"(.{0,160}?)(?:код|SWIFT|БИК)", text, re.I | re.S)
        window = m.group(1) if m else ""
        if not window:
            m = re.search(spaced + r"\s*,?\s*([^,|]{4,90})", text)
            window = m.group(1) if m else ""
        for candidate in RE_ORG.findall(window):
            if "омашка" not in candidate:
                bank = candidate
                break
        if not bank and window:
            bank = window.strip(" ,.;в")
    if not bank:
        for candidate in RE_ORG.findall(text):
            if re.search(r"банк", candidate, re.I) and "омашка" not in candidate:
                bank = candidate
                break

    m = RE_BUYER_SIGNER.search(whole)
    signer = m.group(1).strip(" .,;:") if m else _first(RE_SIGNER, text)
    basis = _clean_basis(m.group(2)) if m else _clean_basis(_first(RE_BASIS, text))
    if not signer or not basis:
        t_signer, t_basis = _signer_from_tables(tables)
        signer = signer or t_signer
        basis = basis or _clean_basis(t_basis)

    phone = _first(RE_PHONE, text)

    return {
        "Полное_наименование": _clean_org(name),
        "УНП": unp,
        "Юр_адрес": address,
        "Расчетный_счет": iban,
        "Банк": _clean_org(bank),
        "Код_банка": swift,
        "Телефон": " ".join(phone.split()) if phone else "",
        "Директор_ФИО": signer,
        "Директор_основание": basis,
    }


def parse_contract(path: str) -> dict:
    """Полный разбор файла договора. Бросает ValueError с понятным сообщением."""
    paragraphs, tables = read_document(path)
    data = extract_requisites(paragraphs, tables)
    data["_найдено"] = sum(1 for k, v in data.items() if not k.startswith("_") and v)
    return data


# ---------------------------------------------------------------------------
# Сопоставление с контрагентом из справочника
# ---------------------------------------------------------------------------

def _norm(name: str) -> str:
    name = (name or "").lower()
    name = re.sub(r"[«»\"'`]", " ", name)
    name = re.sub(r"\b(открытое|закрытое|акционерное|общество|с|ограниченной|"
                  r"ответственностью|унитарное|предприятие|частное|торговое|"
                  r"производственное|совместное|индивидуальный|предприниматель|"
                  r"оао|зао|ооо|чуп|уп|сп|ип|филиал)\b", " ", name)
    return re.sub(r"[^a-zа-яё0-9]+", "", name)


def match_counterparty(full_name: str, known: list) -> list:
    """Возвращает список кандидатов из справочника, отсортированный по похожести."""
    target = _norm(full_name)
    if not target:
        return []
    scored = []
    for name in known:
        n = _norm(name)
        if not n:
            continue
        if n == target:
            score = 1.0
        elif n in target or target in n:
            score = 0.8
        else:
            common = len(set(n) & set(target))
            score = common / max(len(set(n) | set(target)), 1) * 0.5
        scored.append((score, name))
    scored.sort(reverse=True)
    # Слабое совпадение хуже, чем никакого: «Минский завод литья и нормалей»
    # и «Минский лифтовый завод» похожи по буквам, но это разные заводы.
    # Если уверенности нет — не подсказываем, человек выберет из полного списка.
    return [name for score, name in scored if score >= 0.6][:5]


# ---------------------------------------------------------------------------
# Запись в справочник
# ---------------------------------------------------------------------------

FIELDS = ["Полное_наименование", "УНП", "Юр_адрес", "Расчетный_счет",
          "Банк", "Код_банка", "Телефон", "Директор_ФИО", "Директор_основание"]


DIRECTORY_COLUMNS = [
    "Контрагент", "Полное_наименование", "Тип_договора_default", "УНП", "Юр_адрес",
    "Расчетный_счет", "Банк", "Код_банка", "Телефон", "Email",
    "Директор_ФИО", "Директор_основание", "Ставка_пени_%_override",
    "Максимум_дней_просрочки", "Сумма_просрочки_всего_BYN",
    "Риск_срока_давности_3_года", "Комментарий",
]


def _read_directory(path):
    """(строки, имена колонок). Если файла нет — пустой справочник."""
    import csv
    if not os.path.exists(path):
        return [], list(DIRECTORY_COLUMNS)
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fieldnames = list(reader.fieldnames or [])
    for col in DIRECTORY_COLUMNS:
        if col not in fieldnames:
            fieldnames.append(col)
    return rows, fieldnames


def _write_directory(rows, fieldnames, path):
    import csv
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fieldnames})


def ensure_counterparties(lines, path) -> list:
    """Дописывает в справочник покупателей из отчёта, которых там ещё нет.

    Существующие строки не трогает — ни реквизиты, ни ваши правки. Возвращает
    список добавленных названий. Так справочник пополняется сам при каждой
    загрузке отчёта, а заполненное никогда не пропадает.
    """
    from collections import defaultdict
    from parser import classify_contract_type

    rows, fieldnames = _read_directory(path)
    have = {(r.get("Контрагент") or "").strip() for r in rows}

    by_cp = defaultdict(list)
    for l in lines:
        by_cp[l.counterparty].append(l)

    added = []
    for cp in sorted(by_cp):
        if cp in have:
            continue
        ls = by_cp[cp]
        types = [classify_contract_type(l.contract_number) for l in ls]
        max_days = max(l.overdue_days for l in ls)
        rows.append({
            "Контрагент": cp,
            "Тип_договора_default": max(set(types), key=types.count),
            "Максимум_дней_просрочки": max_days,
            "Сумма_просрочки_всего_BYN": round(sum(l.overdue_amount for l in ls), 2),
            "Риск_срока_давности_3_года": "ДА - проверить срок давности!" if max_days > 1095 else "",
            "Комментарий": "ЗАПОЛНИТЬ реквизиты из папки договоров перед отправкой претензии",
        })
        added.append(cp)

    if added:
        _write_directory(rows, fieldnames, path)
    return added


def save_requisites(counterparty: str, values: dict, path: str, create: bool = False) -> dict:
    """Вписывает реквизиты в data/contragents.csv по имени контрагента.

    Пустые значения не затирают то, что уже заполнено вручную.
    create=True — завести новую строку, если такого контрагента ещё нет
    (покупатель, которого пока нет в отчёте по ДЗ).
    """
    rows, fieldnames = _read_directory(path)

    target = None
    for row in rows:
        if row.get("Контрагент", "").strip() == counterparty.strip():
            target = row
            break
    if target is None and create:
        target = {"Контрагент": counterparty.strip(),
                  "Комментарий": "заведён вручную, в отчёте по ДЗ пока не встречался"}
        rows.append(target)
    if target is None:
        return dict(ok=False, изменено=[],
                    текст=f"Контрагента «{counterparty}» нет в справочнике.")

    changed = []
    for field in FIELDS:
        new = (values.get(field) or "").strip()
        if new and new != (target.get(field) or "").strip():
            target[field] = new
            changed.append(field)

    if changed and (target.get("Комментарий") or "").startswith("ЗАПОЛНИТЬ"):
        target["Комментарий"] = "реквизиты загружены из договора"

    _write_directory(rows, fieldnames, path)

    return dict(ok=True, изменено=changed,
                текст=(f"Записано полей: {len(changed)}." if changed
                       else "Всё уже было заполнено — изменений нет."))
