# -*- coding: utf-8 -*-
"""Веб-интерфейс программы: открывается в браузере, работает только на этом
компьютере (127.0.0.1), никуда ничего не отправляет.

Запуск: двойной щелчок по «Претензии.bat» либо  python web.py
Ничего доустанавливать не нужно — сервер собран на стандартной библиотеке Python.
"""
import base64
import datetime as dt
import json
import os
import socket
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from main import group_by_contract, load_contragents, run, STATUTE_OF_LIMITATIONS_DAYS
from parser import parse_debt_report
from company import MIN_CLAIM_AMOUNT, CHARGE_INTEREST_366
from rates import rate_freshness, update_refi_file, checked_today, mark_checked
import journal
import contracts
import payments as pay

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, "templates", "index.html")
CONTRAGENTS = os.path.join(BASE_DIR, "data", "contragents.csv")
UPLOAD_DIR = os.path.join(BASE_DIR, "data", "uploaded")
CONTRACTS_DIR = os.path.join(BASE_DIR, "data", "contracts")

def collect_rows(report: str, min_days: int, as_of: dt.date = None) -> list:
    """Список договоров с просрочкой — то, что показывается в таблице."""
    as_of = as_of or dt.date.today()
    lines = parse_debt_report(report)
    contracts.ensure_counterparties(lines, CONTRAGENTS)
    contragents = load_contragents(CONTRAGENTS)
    state = journal.build_state()
    pay_index = pay.build_index()
    rows = []
    for (cp, number, date), grp in sorted(group_by_contract(lines).items()):
        qualifying = [l for l in grp if l.overdue_days >= min_days]
        if not qualifying:
            continue
        max_days = max(l.overdue_days for l in qualifying)
        cp_row = contragents.get(cp)
        debt = round(sum(l.overdue_amount for l in qualifying), 2)
        st = journal.status_for((cp, number, date), state, debt, as_of)
        paid = sum(p["сумма"] for l in qualifying
                   for p in pay.for_line(pay_index, cp, number, date, l.doc_number, as_of))
        rows.append(dict(
            invoices=[dict(number=l.doc_number, date=l.doc_date,
                           amount=round(l.overdue_amount, 2),
                           due=l.due_date.strftime("%d.%m.%Y")) for l in qualifying],
            paid=round(paid, 2),
            cp=cp,
            contract=number,
            contract_date=date,
            days=max_days,
            docs=len(qualifying),
            debt=debt,
            type=("Биржевой" if number.strip().startswith(("01/", "1/")) else "Прямой"),
            risk=max_days > STATUTE_OF_LIMITATIONS_DAYS,
            nodata=not cp_row or not cp_row.get("УНП"),
            tiny=debt < MIN_CLAIM_AMOUNT,
            status=st["статус"],
            hint=st["подсказка"],
            ready=st["можно_иск"],
            deadline=st["срок_ответа"].strftime("%d.%m.%Y") if st["срок_ответа"] else "",
        ))
    return rows


class Handler(BaseHTTPRequestHandler):
    server_version = "ClaimsGenerator/1.0"

    # ------------------------------------------------------------- служебное
    def log_message(self, fmt, *args):
        pass  # не засоряем консоль

    def _send(self, code: int, body: bytes, ctype: str):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code: int = 200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    # ------------------------------------------------------------------ GET
    def do_GET(self):
        if self.path in ("/", "/index.html"):
            with open(TEMPLATE, "rb") as f:
                self._send(200, f.read(), "text/html; charset=utf-8")
        elif self.path == "/api/start":
            # Путь к отчёту намеренно не подставляем: файл каждый раз выбирает
            # человек, чтобы случайно не сформировать документы по старой выгрузке.
            self._json(dict(report="", today=dt.date.today().strftime("%d.%m.%Y")))
        else:
            self._send(404, b"not found", "text/plain; charset=utf-8")

    # ----------------------------------------------------------------- POST
    def do_POST(self):
        try:
            if self.path == "/api/load":
                self.api_load()
            elif self.path == "/api/upload":
                self.api_upload()
            elif self.path == "/api/generate":
                self.api_generate()
            elif self.path == "/api/payments":
                self.api_payments()
            elif self.path == "/api/payment-add":
                self.api_payment_add()
            elif self.path == "/api/payment-delete":
                self.api_payment_delete()
            elif self.path == "/api/contract-parse":
                self.api_contract_parse()
            elif self.path == "/api/contract-save":
                self.api_contract_save()
            elif self.path == "/api/update-rate":
                self.api_update_rate()
            elif self.path == "/api/mark-sent":
                self.api_mark_sent()
            elif self.path == "/api/open":
                self.api_open()
            else:
                self._send(404, b"not found", "text/plain; charset=utf-8")
        except PermissionError as e:
            name = os.path.basename(getattr(e, "filename", "") or "") or "файл"
            self._json(dict(error=f"Не могу записать «{name}» — он открыт в другой программе "
                                  f"(Excel или Word?). Закройте его и повторите."), 500)
        except Exception as e:
            self._json(dict(error=f"{type(e).__name__}: {e}"), 500)

    def api_load(self):
        data = self._read_json()
        report = (data.get("report") or "").strip()
        min_days = int(data.get("min_days") or 5)
        if not report or not os.path.exists(report):
            return self._json(dict(error="Файл отчёта не найден. Проверьте путь или выберите файл."), 400)
        try:
            as_of = dt.datetime.strptime((data.get("as_of") or "").strip(), "%d.%m.%Y").date()
        except ValueError:
            return self._json(dict(error="Дата расчёта должна быть в виде ДД.ММ.ГГГГ."), 400)
        rows = collect_rows(report, min_days, as_of)
        fresh = rate_freshness(as_of) if CHARGE_INTEREST_366 else dict(устарела=False, текст="")
        self._json(dict(
            rows=rows,
            total_debt=round(sum(r["debt"] for r in rows), 2),
            report=report,
            rate_warning=fresh["текст"] if fresh["устарела"] else "",
        ))

    def api_upload(self):
        data = self._read_json()
        name = os.path.basename(data.get("name") or "report.xls")
        raw = base64.b64decode(data.get("data") or "")
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        path = os.path.join(UPLOAD_DIR, name)
        with open(path, "wb") as f:
            f.write(raw)
        self._json(dict(report=path))

    def api_generate(self):
        data = self._read_json()
        mode = data.get("mode")
        if mode not in ("pretenzii", "iski"):
            return self._json(dict(error="Неизвестный режим."), 400)
        report = (data.get("report") or "").strip()
        if not os.path.exists(report):
            return self._json(dict(error="Файл отчёта не найден."), 400)
        keys = {(k[0], k[1], k[2]) for k in data.get("keys") or []}
        if not keys:
            return self._json(dict(error="Не отмечено ни одного договора."), 400)
        min_days = int(data.get("min_days") or 5)
        try:
            as_of = dt.datetime.strptime((data.get("as_of") or "").strip(), "%d.%m.%Y").date()
        except ValueError:
            return self._json(dict(error="Дата расчёта должна быть в виде ДД.ММ.ГГГГ."), 400)

        out_dir = os.path.join(BASE_DIR, "output")
        rows = run(mode, report, CONTRAGENTS, out_dir, min_days, as_of, only_keys=keys)

        files = [dict(
            cp=r["Контрагент"], contract=r["Договор"],
            debt=r["Осн_долг"], penalty=r["Пеня"],
            interest=r.get("Проценты_ст366", ""), total=r["Итого"],
            file=os.path.relpath(r["Файл"], out_dir),
        ) for r in rows]
        self._json(dict(
            count=len(rows),
            total=round(sum(float(r["Итого"]) for r in rows), 2) if rows else 0.0,
            dir=out_dir,
            files=files,
        ))

    def _counterparties(self):
        with open(CONTRAGENTS, encoding="utf-8-sig") as f:
            return [r for r in __import__("csv").DictReader(f)]

    def api_payments(self):
        """Все внесённые оплаты, с порядковым номером для удаления."""
        rows = pay.load_payments()
        self._json(dict(payments=[dict(i=i, **r) for i, r in enumerate(rows)]))

    def api_payment_add(self):
        res = pay.add_payment(self._read_json() or {})
        if not res["ok"]:
            return self._json(dict(error=res["текст"]), 400)
        self._json(dict(ok=True, text=res["текст"]))

    def api_payment_delete(self):
        data = self._read_json() or {}
        try:
            index = int(data.get("i"))
        except (TypeError, ValueError):
            return self._json(dict(error="Не указана запись."), 400)
        res = pay.delete_payment(index)
        if not res["ok"]:
            return self._json(dict(error=res["текст"]), 400)
        self._json(dict(ok=True, text=res["текст"]))

    def api_contract_parse(self):
        """Принимает файл договора, возвращает найденные реквизиты покупателя."""
        data = self._read_json()
        name = os.path.basename(data.get("name") or "contract")
        raw = base64.b64decode(data.get("data") or "")
        os.makedirs(CONTRACTS_DIR, exist_ok=True)
        path = os.path.join(CONTRACTS_DIR, name)
        with open(path, "wb") as f:
            f.write(raw)

        try:
            found = contracts.parse_contract(path)
        except ValueError as e:
            return self._json(dict(error=str(e)), 400)

        rows = self._counterparties()
        known = [r["Контрагент"] for r in rows]
        candidates = contracts.match_counterparty(found.get("Полное_наименование", ""), known)
        # у кого реквизиты ещё не заполнены — таких показываем первыми
        empty = [r["Контрагент"] for r in rows if not (r.get("УНП") or "").strip()]

        self._json(dict(
            file=name,
            fields={k: v for k, v in found.items() if not k.startswith("_")},
            found=found["_найдено"],
            candidates=candidates,
            all=known,
            empty=empty,
        ))

    def api_contract_save(self):
        data = self._read_json()
        counterparty = (data.get("counterparty") or "").strip()
        values = data.get("fields") or {}
        create = bool(data.get("create"))
        if not counterparty:
            return self._json(dict(error="Не выбран контрагент."), 400)
        res = contracts.save_requisites(counterparty, values, CONTRAGENTS, create=create)
        if not res["ok"]:
            return self._json(dict(error=res["текст"]), 400)
        self._json(dict(ok=True, changed=res["изменено"], text=res["текст"]))

    def api_update_rate(self):
        res = update_refi_file()
        if res["ok"]:
            mark_checked()
        self._json(dict(ok=res["ok"], added=res["добавлено"], text=res["текст"]))

    def api_mark_sent(self):
        data = self._read_json()
        keys = [(k[0], k[1], k[2]) for k in data.get("keys") or []]
        if not keys:
            return self._json(dict(error="Не отмечено ни одного договора."), 400)
        try:
            sent = dt.datetime.strptime((data.get("date") or "").strip(), "%d.%m.%Y").date()
        except ValueError:
            return self._json(dict(error="Дата отправки должна быть в виде ДД.ММ.ГГГГ."), 400)
        count = journal.mark_sent(keys, sent)
        deadline = sent + dt.timedelta(days=journal.PRETENZIA_DEADLINE_DAYS)
        self._json(dict(count=count, deadline=deadline.strftime("%d.%m.%Y")))

    def api_open(self):
        data = self._read_json()
        path = (data.get("path") or os.path.join(BASE_DIR, "output")).strip()
        os.makedirs(path, exist_ok=True)
        try:
            os.startfile(path)  # Windows
            self._json(dict(ok=True))
        except AttributeError:
            self._json(dict(ok=False, path=path))


def free_port(preferred: int = 8731) -> int:
    for port in range(preferred, preferred + 20):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return 0


def main():
    port = free_port()
    if not port:
        raise SystemExit("Не удалось занять порт для интерфейса.")
    url = f"http://127.0.0.1:{port}/"
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print("Интерфейс запущен:", url)
    print("Окно программы откроется в браузере. Чтобы закрыть программу — закройте это чёрное окно.")
    def daily_rate_check():
        """Раз в сутки тихо проверяем ставку. Нет интернета — молча пропускаем."""
        if checked_today():
            return
        res = update_refi_file(timeout=8)
        if res["ok"]:
            mark_checked()
            if res["добавлено"]:
                print("Ставка рефинансирования обновлена:", res["текст"])

    threading.Thread(target=daily_rate_check, daemon=True).start()
    threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
