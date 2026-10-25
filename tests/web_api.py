# -*- coding: utf-8 -*-
"""Дымовая проверка окна программы: поднимает сервер на свободном порту
и дёргает все точки API так, как это делает страница.

Запуск:  python tests/web_api.py [путь_к_отчёту.xls]

Реестр не пишется (journal.record подменён). Документы формируются в
tests/output, настоящую папку output не трогает.
"""
import datetime as dt
import json
import os
import sys
import threading
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
sys.path.insert(0, PROJECT)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
os.chdir(PROJECT)

import journal
import web

journal.record = lambda entries, path=None: None

REPORT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    PROJECT, "data", "uploaded", "Отчет.xls")
AS_OF = "19.08.2026"

httpd = ThreadingHTTPServer(("127.0.0.1", 0), web.Handler)
threading.Thread(target=httpd.serve_forever, daemon=True).start()
BASE = f"http://127.0.0.1:{httpd.server_address[1]}"


def get(path):
    with urllib.request.urlopen(BASE + path, timeout=60) as r:
        return r.status, r.read()


def post(path, obj):
    req = urllib.request.Request(BASE + path, data=json.dumps(obj).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=600) as r:
            return r.status, json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read().decode("utf-8"))


def check(label, condition, detail=""):
    print(f'  {"OK " if condition else "БАГ"} {label}{"  " + detail if detail else ""}')
    return bool(condition)


print("Отчёт:", REPORT)
if not os.path.exists(REPORT):
    sys.exit("Файла нет. Передайте путь параметром: python tests/web_api.py D:\\путь\\Отчет.xls")

ok = True
print("\nСтраница и старт")
st, body = get("/")
ok &= check("GET /", st == 200 and len(body) > 10000, f'{len(body)} байт')
st, body = get("/api/start")
ok &= check("GET /api/start", st == 200, body.decode("utf-8"))

print("\nЧтение отчёта")
st, j = post("/api/load", dict(report=REPORT, min_days=5, as_of=AS_OF))
ok &= check("/api/load", st == 200 and j.get("rows"),
            f'{len(j.get("rows", []))} договоров, долг {j.get("total_debt")}')
st, j2 = post("/api/load", dict(report=REPORT, min_days=5, as_of="19/08/2026"))
ok &= check("/api/load отвергает кривую дату", st == 400, j2.get("error", ""))
st, j3 = post("/api/load", dict(report="нет такого файла.xls", min_days=5, as_of=AS_OF))
ok &= check("/api/load отвергает несуществующий файл", st == 400, j3.get("error", ""))

first = j["rows"][0]
key = [first["cp"], first["contract"], first["contract_date"]]

print("\nФормирование (в tests/output)")
real_out, web.BASE_DIR = web.BASE_DIR, HERE
try:
    for mode in ("pretenzii", "iski"):
        st, j4 = post("/api/generate", dict(mode=mode, report=REPORT, min_days=5,
                                            as_of=AS_OF, keys=[key]))
        ok &= check(f'/api/generate {mode}', st == 200 and j4.get("count") == 1,
                    j4.get("error") or f'итого {j4.get("total")}')
    st, j5 = post("/api/generate", dict(mode="iski", report=REPORT, min_days=5,
                                        as_of=AS_OF, keys=[]))
    ok &= check("/api/generate без отметок", st == 400, j5.get("error", ""))
    st, j6 = post("/api/generate", dict(mode="iski", report=REPORT, min_days=5,
                                        as_of="19/08/2026", keys=[key]))
    ok &= check("/api/generate отвергает кривую дату", st == 400, j6.get("error", ""))
finally:
    web.BASE_DIR = real_out

print("\nОплаты")
st, j7 = post("/api/payments", {})
ok &= check("/api/payments", st == 200, f'{len(j7.get("payments", []))} записей')
st, j8 = post("/api/payment-add", dict(Контрагент="X", Договор="Y", Дата_договора="01.01.26",
                                       Накладная="1", Дата_оплаты="32.13.2026", Сумма="abc"))
ok &= check("/api/payment-add отвергает мусор", st == 400, j8.get("error", ""))

print("\nСтавка и реестр")
st, j9 = post("/api/update-rate", {})
ok &= check("/api/update-rate", st == 200, j9.get("text", ""))
st, j10 = post("/api/mark-sent", dict(keys=[key], date="19/08/2026"))
ok &= check("/api/mark-sent отвергает кривую дату", st == 400, j10.get("error", ""))

print("\nРазбор договора")
contracts_dir = os.path.join(PROJECT, "data", "contracts")
samples = [f for f in os.listdir(contracts_dir)] if os.path.isdir(contracts_dir) else []
if samples:
    import base64
    with open(os.path.join(contracts_dir, samples[0]), "rb") as f:
        blob = base64.b64encode(f.read()).decode("ascii")
    st, j11 = post("/api/contract-parse", dict(name=samples[0], data=blob))
    ok &= check(f'/api/contract-parse ({samples[0]})', st == 200,
                j11.get("error") or f'найдено полей: {j11.get("found")}')
else:
    print("   — пропущено: в data/contracts нет файлов")

httpd.shutdown()
print("\nИТОГ:", "всё работает" if ok else "ЕСТЬ ПРОБЛЕМЫ (см. строки «БАГ»)")
sys.exit(0 if ok else 1)
