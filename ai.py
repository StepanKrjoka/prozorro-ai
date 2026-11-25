#!/usr/bin/env python3
"""
Закупівлі штучного інтелекту в Prozorro: хто купує, у кого і за скільки.

Конвеєр (кожен крок можна запускати окремо):
  probe    — перевірка текстового пошуку: скільки результатів дає кожен запит, які типи процедур індексуються
  search   — збір кандидатів через пошук prozorro.gov.ua за текстом (окремого CPV для ШІ немає)
  details  — повні дані тендерів з public-api.prozorro.gov.ua; у вибірку йдуть лише ті, де ШІ явно згадано
             в назві, описі, лоті чи позиції; для них — договори з модуля contracting
  build    — CSV-таблиці, рейтинги, docs/index.html
  all      — search + details + build;  resume — details + build

Лише стандартна бібліотека Python 3.10+.
"""
import argparse
import csv
import datetime as dt
import gzip
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")
RAW_T = os.path.join(DATA, "raw", "tenders")
RAW_C = os.path.join(DATA, "raw", "contracts")
DOCS = os.path.join(ROOT, "docs")

SEARCH_URL = "https://prozorro.gov.ua/api/search/tenders"
SKIP_STATUSES = ("unsuccessful", "cancelled")
SUMMARY_URL = "https://prozorro.gov.ua/api/tenders/{}/summary"
TENDER_URL = "https://public-api.prozorro.gov.ua/api/2.5/tenders/{}"
CONTRACT_URL = "https://public-api.prozorro.gov.ua/api/2.5/contracts/{}"
PUBLIC_TENDER = "https://prozorro.gov.ua/tender/{}"
SEARCH_CAP = 1000           # пошук віддає не більше 1000 результатів на запит
HEX_RE = re.compile(r"^[0-9a-f]{32}$")
FINAL_STATUSES = {"complete", "cancelled", "unsuccessful"}

HEADERS = {"User-Agent": "prozorro-ai-research/1.0 (journalism; github actions)",
           "Accept": "application/json"}
GAP = {"prozorro.gov.ua": 1.0, "public-api.prozorro.gov.ua": 0.15}   # пауза між запитами до хоста, с
_last = {}
T0 = time.time()


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


# ------------------------------------------------------------------ HTTP
def http(url, body=None, tries=6, timeout=60):
    host = urllib.parse.urlparse(url).hostname
    headers = dict(HEADERS)
    data = None
    if body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    for i in range(tries):
        wait = _last.get(host, 0) + GAP.get(host, 0.2) - time.time()
        if wait > 0:
            time.sleep(wait)
        _last[host] = time.time()
        try:
            with urllib.request.urlopen(urllib.request.Request(url, data, headers), timeout=timeout) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (404, 410):
                return None
            if e.code in (400, 401, 403) or i == tries - 1:
                raise RuntimeError(f"HTTP {e.code} від {host}: {url} :: {e.read()[:300]!r}")
            time.sleep((15 if e.code == 429 else 3) * (i + 1))
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            if i == tries - 1:
                raise RuntimeError(f"{host} недоступний: {e}")
            time.sleep(3 * (i + 1))


# ------------------------------------------------------------------ files
def gz_write(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with gzip.open(tmp, "wt", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, path)


def gz_read(path):
    try:
        with gzip.open(path, "rt", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, OSError, ValueError):
        return None


def json_read(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError):
        return default


def json_write(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def csv_write(path, rows, fields):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:   # BOM — щоб Excel відкрив кирилицю
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


# ------------------------------------------------------------------ config & classification
def load_config():
    cfg = json_read(os.path.join(ROOT, "config.json"), None)
    if not cfg:
        sys.exit("Немає config.json")
    for g in cfg["groups"]:
        g["_inc"] = [re.compile(p, re.I) for p in g.get("patterns", [])]
        g["_exc"] = [re.compile(p, re.I) for p in g.get("exclude", [])]
    for k in ("communal", "state"):
        cfg["_own_" + k] = [re.compile(p, re.I) for p in cfg["ownership"][k]]
    cfg["_ai"] = [re.compile(p, re.I) for p in cfg["ai_patterns"]]
    cfg["_ai_exc"] = [re.compile(p, re.I) for p in cfg.get("ai_exclude_context", [])]
    cfg["_genai"] = [re.compile(p, re.I) for p in cfg.get("genai_patterns", [])]
    return cfg


def read_override(name, value_col):
    out = {}
    path = os.path.join(ROOT, "overrides", name)
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                code = re.sub(r"\D", "", r.get("edrpou") or "")
                if code and (r.get(value_col) or "").strip():
                    out[code] = r[value_col].strip()
    return out


def norm(name):
    return " ".join((name or "").lower().replace("’", "'").replace("`", "'").split())


def classify(cfg, name, kind=None, edrpou=None, overrides=None):
    """Повертає (group_id, причина). Остання група в config — «інші», тож групу отримує кожен замовник;
    None — лише якщо замовника вручну виключено (overrides/buyers.csv, group=exclude)."""
    if overrides and edrpou and edrpou in overrides:
        g = overrides[edrpou]
        return (None, None) if g == "exclude" else (g, "ручне віднесення (overrides/buyers.csv)")
    n = norm(name)
    for g in cfg["groups"]:
        if kind and kind in g.get("kind", []):
            return g["id"], f"kind = {kind}"
        if g["_inc"] and any(p.search(n) for p in g["_inc"]) and not any(p.search(n) for p in g["_exc"]):
            return g["id"], "назва замовника"
    return cfg["groups"][-1]["id"], "не підпадає під жодне правило"


def ownership(cfg, name, kind=None):
    n = norm(name)
    if any(p.search(n) for p in cfg["_own_communal"]):
        return "комунальна"
    if kind in ("central", "defense") or any(p.search(n) for p in cfg["_own_state"]):
        return "державна"
    return "не визначено"



def ai_terms(cfg, text):
    """Знайдені згадки ШІ в тексті (нормалізовані збіги). Контекст на кшталт «клавіша Copilot» не рахується."""
    t = norm(text)
    for p in cfg["_ai_exc"]:
        t = p.sub(" ", t)
    out = []
    for p in cfg["_ai"]:
        m = p.search(t)
        if m:
            out.append(m.group(0))
    return out


def tender_ai(cfg, t):
    """Де саме в тендері згадано ШІ: назва/опис тендера, лоти, позиції. Повертає (місця, терміни, позиції з ШІ)."""
    where, terms, items = [], [], []
    for label, text in (("назва", t.get("title")), ("опис", t.get("description"))):
        f = ai_terms(cfg, text)
        if f:
            where.append(label); terms += f
    ai_lots = set()
    for l in t.get("lots") or []:
        f = ai_terms(cfg, (l.get("title") or "") + " " + (l.get("description") or ""))
        if f:
            ai_lots.add(l.get("id")); terms += f
    if ai_lots:
        where.append("лот")
    for i in t.get("items") or []:
        f = ai_terms(cfg, i.get("description"))
        if f or i.get("relatedLot") in ai_lots:
            items.append(i); terms += f
    if items and "лот" not in where:
        where.append("позиція")
    return where, sorted(set(terms)), items


def category(cfg, items):
    """Категорія за CPV позицій з ШІ (або всіх позицій тендера): перша категорія, до якої належить найбільше позицій."""
    cnt = defaultdict(int)
    for i in items:
        cpv = ((i.get("classification") or {}).get("id") or "")
        for c in cfg["categories"]:
            if any(cpv.startswith(p) for p in c["cpv"]):
                cnt[c["id"]] += 1
                break
    if not cnt:
        return cfg["categories"][-1]["id"]
    return max(cnt, key=lambda k: (cnt[k], -[c["id"] for c in cfg["categories"]].index(k)))


def is_genai(cfg, text):
    t = norm(text)
    return any(p.search(t) for p in cfg["_genai"])


def edr(org):
    return re.sub(r"\s", "", str(((org or {}).get("identifier") or {}).get("id") or ""))


def amount(v):
    try:
        return float((v or {}).get("amount"))
    except (TypeError, ValueError):
        return None


def _org(o):
    o = o or {}
    return {"name": o.get("name"), "identifier": {"id": edr(o), "legalName": (o.get("identifier") or {}).get("legalName")},
            "kind": o.get("kind"), "address": {k: (o.get("address") or {}).get(k) for k in ("region", "locality")}}


def _items(items):
    return [{"id": i.get("id"), "description": (i.get("description") or "")[:300],
             "classification": {"id": (i.get("classification") or {}).get("id")},
             "quantity": i.get("quantity"), "unit": i.get("unit"), "relatedLot": i.get("relatedLot")}
            for i in items or []]



def trim_tender(t):
    return {
        "id": t.get("id"), "tenderID": t.get("tenderID"), "title": (t.get("title") or "")[:500],
        "description": (t.get("description") or "")[:1500],
        "status": t.get("status"), "procurementMethodType": t.get("procurementMethodType"),
        "date": t.get("date"), "dateCreated": t.get("dateCreated"), "dateModified": t.get("dateModified"),
        "value": t.get("value"), "procuringEntity": _org(t.get("procuringEntity")),
        "items": _items(t.get("items")),
        "lots": [{"id": l.get("id"), "title": l.get("title"), "description": (l.get("description") or "")[:500],
                  "status": l.get("status"), "value": l.get("value")} for l in t.get("lots") or []],
        "awards": [{"id": a.get("id"), "lotID": a.get("lotID"), "status": a.get("status"), "value": a.get("value"),
                    "suppliers": [_org(s) for s in a.get("suppliers") or []]} for a in t.get("awards") or []],
        "contracts": [{"id": k.get("id"), "awardID": k.get("awardID"), "status": k.get("status"),
                       "value": k.get("value"), "dateSigned": k.get("dateSigned"), "period": k.get("period"),
                       "contractNumber": k.get("contractNumber"), "suppliers": [_org(s) for s in k.get("suppliers") or []],
                       "items": _items(k.get("items"))} for k in t.get("contracts") or []],
        "n_bids": len(t.get("bids") or []),
    }


def trim_contract(c):
    return {
        "id": c.get("id"), "status": c.get("status"), "value": c.get("value"), "amountPaid": c.get("amountPaid"),
        "dateSigned": c.get("dateSigned"), "period": c.get("period"), "contractNumber": c.get("contractNumber"),
        "suppliers": [_org(s) for s in c.get("suppliers") or []], "items": _items(c.get("items")),
        "changes": [{"date": (ch.get("date") or "")[:10], "status": ch.get("status"),
                     "rationaleTypes": ch.get("rationaleTypes")} for ch in c.get("changes") or []],
    }


class Shards:
    """Словник {ключ: запис}, розкладений по файлах data/raw/<kind>/YYYY-MM.json.gz."""
    def __init__(self, folder):
        self.folder, self.cache, self.dirty = folder, {}, set()

    def _load(self, month):
        if month not in self.cache:
            self.cache[month] = gz_read(os.path.join(self.folder, month + ".json.gz")) or {}
        return self.cache[month]

    def get(self, month, key):
        return self._load(month).get(key)

    def put(self, month, key, obj):
        self._load(month)[key] = obj
        self.dirty.add(month)

    def flush(self):
        for m in sorted(self.dirty):
            gz_write(os.path.join(self.folder, m + ".json.gz"), self.cache[m])
        self.dirty.clear()

    def months(self):
        if os.path.isdir(self.folder):
            for f in sorted(os.listdir(self.folder)):
                if f.endswith(".json.gz"):
                    self._load(f[:-8])
        return self.cache


def month_of(tid):
    m = re.match(r"UA-(\d{4})-(\d{2})", tid or "")
    return f"{m.group(1)}-{m.group(2)}" if m else "other"



def month_of(tid):
    m = re.match(r"UA-(\d{4})-(\d{2})", tid or "")
    return f"{m.group(1)}-{m.group(2)}" if m else "other"


# ------------------------------------------------------------------ search
def month_windows(start, end):
    d = dt.date.fromisoformat(start)
    end = dt.date.fromisoformat(end)
    while d <= end:
        nxt = (d.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
        yield d.isoformat(), min(nxt - dt.timedelta(days=1), end).isoformat()
        d = nxt


def search_page(base, start, end, page):
    body = dict(base)
    body["date"] = {"notice": {"start": start, "end": end}}
    return http(f"{SEARCH_URL}?page={page}", body) or {}


def search_window(base, start, end, sink, stats, label):
    """Збирає всі результати за вікно дат; якщо їх >= ліміту (1000) — ділить вікно навпіл."""
    first = search_page(base, start, end, 1)
    total = int(first.get("total") or 0)
    s, e = dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    if total >= SEARCH_CAP and s < e:
        mid = s + (e - s) // 2
        log(f"  {label} {start}..{end}: {total} ≥ ліміту, ділю вікно")
        search_window(base, start, mid.isoformat(), sink, stats, label)
        search_window(base, (mid + dt.timedelta(days=1)).isoformat(), end, sink, stats, label)
        return
    if total >= SEARCH_CAP:
        stats["truncated"].append({"query": label, "date": start, "total": total})
        log(f"  УВАГА: {label} {start}: {total} результатів за один день, отримаю лише перші {SEARCH_CAP}")
    per = int(first.get("per_page") or 20) or 20
    pages = min((total + per - 1) // per, SEARCH_CAP // per)
    got = 0
    for p in range(1, pages + 1):
        r = first if p == 1 else search_page(base, start, end, p)
        rows = r.get("data") or []
        if not rows:
            break
        for x in rows:
            sink(x, label)
        got += len(rows)
    stats["windows"].append({"query": label, "start": start, "end": end, "total": total, "fetched": got})
    log(f"  {label} {start}..{end}: {total} знайдено, отримано {got}")



def queries(cfg):
    for t in cfg.get("text_search", []):
        yield f"text:{t}", {"text": t}


def cmd_probe(cfg, args):
    """Діагностика текстового пошуку: формат відповіді, скільки дає кожен запит за місяць і чи збігаються назви з ШІ."""
    start = args.probe_from or cfg["date_from"]
    end = ((dt.date.fromisoformat(start).replace(day=1) + dt.timedelta(days=32)).replace(day=1)
           - dt.timedelta(days=1)).isoformat()
    r = search_page({"text": "штучний інтелект"}, start, end, 1)
    print("== Ключі відповіді пошуку:", sorted(r.keys()))
    print("== Перший запис:", json.dumps((r.get("data") or [None])[0], ensure_ascii=False)[:1500])
    print(f"\n== Результатів за {start}..{end}; у дужках — скільки з першої сторінки мають ШІ в назві:")
    for t in cfg.get("text_search", []):
        x = search_page({"text": t}, start, end, 1)
        rows = x.get("data") or []
        hit = sum(1 for row in rows if ai_terms(cfg, row.get("title")))
        print(f"  «{t}»: {x.get('total')} ({hit}/{len(rows)})")
        for row in rows[:3]:
            print("     ·", (row.get("title") or "")[:120])
    for p in cfg.get("probe_proc_types", []):
        print(f"  тип {p}: {search_page({'text': 'штучний інтелект', 'proc_type': [p]}, start, end, 1).get('total')}")
    print("  фраза в лапках «\"штучний інтелект\"»:", search_page({"text": '"штучний інтелект"'}, start, end, 1).get("total"))
    print("  лише «штучний»:", search_page({"text": "штучний"}, start, end, 1).get("total"))
    print("  лише «інтелект»:", search_page({"text": "інтелект"}, start, end, 1).get("total"))


def cmd_search(cfg, args):
    today = dt.date.today().isoformat()
    date_to = cfg.get("date_to") or today
    state = json_read(os.path.join(DATA, "state.json"), {})
    done = set(state.get("searched_months", []))
    refresh_from = (dt.date.today() - dt.timedelta(days=args.refresh_days)).isoformat()
    cand_path = os.path.join(DATA, "candidates.json")
    cands = json_read(cand_path, {})
    overrides = read_override("buyers.csv", "group")
    stats = {"windows": [], "truncated": [], "seen": 0, "kept": 0}

    def sink(x, label):
        stats["seen"] += 1
        tid = x.get("tenderID")
        if not tid:
            return
        pe = x.get("procuringEntity") or {}
        name = pe.get("name")
        group, _ = classify(cfg, name, None, None, overrides)
        if group is None:
            return
        stats["kept"] += 1
        c = cands.setdefault(tid, {"found_by": []})
        if HEX_RE.match(str(x.get("id") or "")):
            c["id"] = x["id"]
        c["status"] = x.get("status")
        if ai_terms(cfg, x.get("title")):
            c["title_ai"] = True
        if label not in c["found_by"]:
            c["found_by"].append(label)

    for start, end in month_windows(cfg["date_from"], date_to):
        month = start[:7]
        if month in done and end < refresh_from and not args.full:
            continue
        log(f"Місяць {month}")
        for label, base in queries(cfg):
            search_window(base, start, end, sink, stats, label)
        if end < refresh_from:
            done.add(month)
        json_write(cand_path, cands)
        state["searched_months"] = sorted(done)
        json_write(os.path.join(DATA, "state.json"), state)

    state["last_search"] = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    json_write(os.path.join(DATA, "state.json"), state)
    log_path = os.path.join(DATA, "search_log.json")
    hist = json_read(log_path, [])
    hist.append({"run": state["last_search"], **stats})
    json_write(log_path, hist[-10:])
    log(f"Пошук завершено: переглянуто {stats['seen']}, кандидатів {len(cands)}, обрізаних вікон {len(stats['truncated'])}")



# ------------------------------------------------------------------ details
def cmd_details(cfg, args):
    cand_path = os.path.join(DATA, "candidates.json")
    cands = json_read(cand_path, {})
    T, C = Shards(RAW_T), Shards(RAW_C)
    stale_end = (dt.date.today() - dt.timedelta(days=90)).isoformat()
    n_t = n_c = n_skip = 0
    incomplete = False
    # спершу ті, де ШІ видно вже в назві з пошуку, далі — найновіші
    order = sorted(cands, reverse=True)
    order = sorted(order, key=lambda k: not cands[k].get("title_ai"))   # стабільне сортування
    for i, tid in enumerate(order):
        if time.time() - T0 > args.budget_min * 60:
            incomplete = True
            log(f"Ліміт часу {args.budget_min} хв: оброблено {i}/{len(order)}, решту продовжить наступний запуск")
            break
        c, month = cands[tid], month_of(tid)
        t = T.get(month, tid)
        stale = t is not None and t.get("status") not in FINAL_STATUSES and \
            time.time() - float(t.get("_fetched") or 0) > args.tender_ttl * 86400
        if t is None or stale or args.refetch:
            h = c.get("id")
            if not h:
                s = http(SUMMARY_URL.format(tid))
                h = (s or {}).get("id")
                if not h:
                    n_skip += 1
                    continue
                c["id"] = h
            r = http(TENDER_URL.format(h))
            if not r:
                n_skip += 1
                continue
            t = trim_tender(r["data"])
            t["_fetched"] = time.time()
            T.put(month, tid, t)
            n_t += 1
        where, _, _ = tender_ai(cfg, t)
        if not where:
            continue                                     # ШІ не згадано — договори не тягнемо
        for k in t.get("contracts") or []:
            if k.get("status") not in ("active", "terminated"):
                continue
            cc = C.get(month, k["id"])
            if cc and not args.refetch:
                if cc.get("status") == "terminated":
                    continue
                p_end = ((cc.get("period") or {}).get("endDate") or "")[:10]
                if p_end and p_end < stale_end:
                    continue
                if time.time() - float(cc.get("_fetched") or 0) < args.contract_ttl * 86400:
                    continue
            r = http(CONTRACT_URL.format(k["id"]))
            cc = trim_contract(r["data"]) if r else {"_missing": True}
            cc["_fetched"] = time.time()
            C.put(month, k["id"], cc)
            n_c += 1
        if (i + 1) % 300 == 0:
            log(f"  оброблено {i + 1}/{len(order)}: нових тендерів {n_t}, договорів {n_c}")
            T.flush(); C.flush(); json_write(cand_path, cands)
    T.flush(); C.flush(); json_write(cand_path, cands)
    state = json_read(os.path.join(DATA, "state.json"), {})
    state["details_incomplete"] = incomplete
    json_write(os.path.join(DATA, "state.json"), state)
    log(f"Деталі: завантажено тендерів {n_t}, договорів {n_c}, пропущено {n_skip}; "
        f"{'НЕ ЗАВЕРШЕНО' if incomplete else 'завершено'}")
    return incomplete


# ------------------------------------------------------------------ build
STATUS_UA = {"complete": "завершено", "cancelled": "скасовано", "unsuccessful": "не відбувся",
             "active.tendering": "подання пропозицій", "active.enquiries": "уточнення",
             "active.qualification": "кваліфікація", "active.awarded": "визначено переможця",
             "active.auction": "аукціон", "active.pre-qualification": "прекваліфікація", "active": "активний"}


def cmd_build(cfg, args):
    cands = json_read(os.path.join(DATA, "candidates.json"), {})
    overrides = read_override("buyers.csv", "group")
    holdings = read_override("suppliers.csv", "holding")
    tender_ovr = read_override_tenders()
    date_from = cfg["date_from"]
    groups = {g["id"]: g["title"] for g in cfg["groups"]}
    cats = {c["id"]: c["title"] for c in cfg["categories"]}
    T, C = Shards(RAW_T), Shards(RAW_C)
    T.months(); C.months()
    tenders_rows, contract_rows = [], []
    rejected = defaultdict(int)
    for tid in sorted(cands):
        month = month_of(tid)
        t = T.get(month, tid)
        if not t:
            rejected["ще не завантажено"] += 1
            continue
        if tender_ovr.get(tid) == "exclude":
            rejected["виключено вручну (overrides/tenders.csv)"] += 1
            continue
        where, terms, ai_items = tender_ai(cfg, t)
        if not where:
            rejected["ШІ не згадано (лише збіг пошуку)"] += 1
            continue
        notice = (t.get("date") or t.get("dateCreated") or "")[:10]
        if notice and notice < date_from:
            rejected["оголошено до початку періоду"] += 1
            continue
        pe = t.get("procuringEntity") or {}
        b_code, b_name = edr(pe), pe.get("name") or ""
        group, why = classify(cfg, b_name, pe.get("kind"), b_code, overrides)
        if not group:
            rejected["замовника виключено вручну"] += 1
            continue
        own = ownership(cfg, b_name, pe.get("kind"))
        all_items = t.get("items") or []
        whole = bool({"назва", "опис"} & set(where)) or len(ai_items) == len(all_items)
        cat = category(cfg, ai_items or all_items)
        genai = is_genai(cfg, " ".join([t.get("title") or "", t.get("description") or ""] +
                                        [i.get("description") or "" for i in ai_items]))
        tval = amount(t.get("value"))
        awards = {a.get("id"): a for a in t.get("awards") or []}
        lots = {l.get("id"): l for l in t.get("lots") or []}
        ai_text = "; ".join(sorted({(i.get("description") or "")[:160] for i in ai_items}))[:600]
        base = {"tender_id": tid, "url": PUBLIC_TENDER.format(tid), "title": t.get("title"),
                "status": t.get("status"), "status_ua": STATUS_UA.get(t.get("status"), t.get("status")),
                "procedure": t.get("procurementMethodType"), "notice_date": notice, "year": notice[:4],
                "buyer": b_name, "buyer_edrpou": b_code, "buyer_kind": pe.get("kind"),
                "buyer_region": (pe.get("address") or {}).get("region"),
                "group": group, "group_title": groups.get(group, group), "group_reason": why, "ownership": own,
                "category": cat, "category_title": cats.get(cat, cat), "genai": genai,
                "ai_where": ", ".join(where), "ai_terms": ", ".join(terms), "ai_scope": "весь тендер" if whole else "окремі позиції",
                "ai_items": ai_text}
        tenders_rows.append({**base, "expected_value": tval, "currency": (t.get("value") or {}).get("currency"),
                             "bids": t.get("n_bids"),
                             "n_contracts": sum(1 for k in t.get("contracts") or [] if k.get("status") in ("active", "terminated")),
                             "large": bool(tval and tval >= cfg["large_tender_uah"])})
        ai_lot_ids = {i.get("relatedLot") for i in ai_items}
        for k in t.get("contracts") or []:
            if k.get("status") not in ("active", "terminated"):
                continue
            full = C.get(month, k["id"]) or {}
            if full.get("_missing"):
                full = {}
            award = awards.get(k.get("awardID")) or {}
            lot_id = award.get("lotID")
            if lot_id and not whole and ai_lot_ids and lot_id not in ai_lot_ids:
                continue                                 # договір за лотом без ШІ
            sup = ((full.get("suppliers") or k.get("suppliers") or award.get("suppliers") or [{}])[0]) or {}
            signed = (full.get("dateSigned") or k.get("dateSigned") or "")[:10]
            period = full.get("period") or k.get("period") or {}
            status = full.get("status") or k.get("status")
            v_init = amount(k.get("value"))
            v_cur = amount(full.get("value")) if full else None
            paid = amount(full.get("amountPaid"))
            v_rank = paid if (status == "terminated" and paid is not None) else (v_cur if v_cur is not None else v_init)
            changes = [ch for ch in full.get("changes") or [] if ch.get("status") in (None, "active")]
            s_code = edr(sup)
            contract_rows.append({**base,
                "contract_id": k["id"], "contract_number": full.get("contractNumber") or k.get("contractNumber"),
                "contract_status": status, "date_signed": signed, "signed_year": signed[:4],
                "period_start": (period.get("startDate") or "")[:10], "period_end": (period.get("endDate") or "")[:10],
                "supplier": sup.get("name") or "не розкрито", "supplier_edrpou": s_code,
                "holding": holdings.get(s_code, ""),
                "value_initial": v_init, "value_current": v_cur, "amount_paid": paid, "value_rank": v_rank,
                "currency": ((full.get("value") or k.get("value") or {}).get("currency")),
                "vat_included": ((full.get("value") or k.get("value") or {}).get("valueAddedTaxIncluded")),
                "n_changes": len(changes),
                "change_types": "; ".join(sorted({r for ch in changes for r in ch.get("rationaleTypes") or []})),
                "expected_value": amount((lots.get(lot_id) or {}).get("value")) if lot_id else tval,
                "bids": t.get("n_bids"),
                "contract_data": "contracting" if full else "лише з тендера",
                "large": bool(v_rank and v_rank >= cfg["large_contract_uah"]),
            })

    def rank(rows, key_fn, label_fn, code_fn, partner_fn):
        agg = {}
        for r in rows:
            if (r.get("currency") or "UAH") != "UAH" or r["value_rank"] is None:
                continue
            key = key_fn(r)
            a = agg.setdefault(key, {"name": label_fn(r), "edrpou": code_fn(r), "total": 0.0, "contracts": 0,
                                     "partners": set(), "cats": set(), "max_contract": 0.0, "genai": 0, "names": set()})
            a["total"] += r["value_rank"]; a["contracts"] += 1
            a["cats"].add(r["category_title"]); a["names"].add(label_fn(r))
            a["max_contract"] = max(a["max_contract"], r["value_rank"])
            a["genai"] += 1 if r["genai"] else 0
            a["partners"].add(partner_fn(r))
        grand = sum(a["total"] for a in agg.values()) or 1
        out = []
        for i, a in enumerate(sorted(agg.values(), key=lambda x: -x["total"]), 1):
            out.append({"rank": i, "name": a["name"], "edrpou": a["edrpou"], "total": round(a["total"], 2),
                        "share_pct": round(100 * a["total"] / grand, 2), "contracts": a["contracts"],
                        "partners": len(a["partners"]), "genai_contracts": a["genai"],
                        "categories": ", ".join(sorted(a["cats"])), "max_contract": round(a["max_contract"], 2),
                        "other_names": " | ".join(sorted(a["names"] - {a["name"]}))[:300]})
        return out

    sup_rank = rank(contract_rows, lambda r: r["supplier_edrpou"] or "? " + r["supplier"], lambda r: r["supplier"],
                    lambda r: r["supplier_edrpou"], lambda r: r["buyer_edrpou"] or r["buyer"])
    buy_rank = rank(contract_rows, lambda r: r["buyer_edrpou"] or "? " + r["buyer"], lambda r: r["buyer"],
                    lambda r: r["buyer_edrpou"], lambda r: r["supplier_edrpou"] or r["supplier"])
    r_fields = ["rank", "name", "edrpou", "total", "share_pct", "contracts", "partners", "genai_contracts",
                "categories", "max_contract", "other_names"]
    csv_write(os.path.join(DATA, "suppliers.csv"), sup_rank, r_fields)
    csv_write(os.path.join(DATA, "buyers_rank.csv"), buy_rank, r_fields)
    c_fields = list(contract_rows[0].keys()) if contract_rows else ["tender_id"]
    csv_write(os.path.join(DATA, "contracts.csv"), sorted(contract_rows, key=lambda r: -(r["value_rank"] or 0)), c_fields)
    t_fields = list(tenders_rows[0].keys()) if tenders_rows else ["tender_id"]
    csv_write(os.path.join(DATA, "tenders.csv"), sorted(tenders_rows, key=lambda r: r["notice_date"], reverse=True), t_fields)

    meta = {
        "built": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "date_from": date_from, "groups": groups, "categories": cats,
        "candidates": len(cands), "tenders": len(tenders_rows), "contracts": len(contract_rows),
        "rejected": dict(rejected),
        "details_incomplete": json_read(os.path.join(DATA, "state.json"), {}).get("details_incomplete", False),
        "search_log": {k: v for k, v in ((json_read(os.path.join(DATA, "search_log.json"), []) or [{}])[-1]).items()
                       if k != "windows"},
        "text_search": cfg.get("text_search", []),
    }
    json_write(os.path.join(DATA, "summary.json"), meta)
    keep_t = ["tender_id", "url", "title", "status_ua", "status", "notice_date", "year", "buyer", "buyer_edrpou", "group",
              "ownership", "category", "genai", "ai_where", "ai_terms", "ai_scope", "ai_items", "expected_value", "bids",
              "n_contracts", "procedure"]
    keep_c = ["tender_id", "url", "title", "buyer", "buyer_edrpou", "group", "ownership", "category", "genai", "ai_scope",
              "supplier", "supplier_edrpou", "value_rank", "value_initial", "date_signed", "signed_year", "contract_status",
              "n_changes", "bids", "notice_date", "year"]
    payload = {"meta": meta,
               "tenders": [{k: r[k] for k in keep_t} for r in tenders_rows],
               "contracts": [{k: r[k] for k in keep_c} for r in contract_rows
                             if (r.get("currency") or "UAH") == "UAH" and r["value_rank"] is not None]}
    with open(os.path.join(ROOT, "report_template.html"), encoding="utf-8") as f:
        html = f.read()
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    os.makedirs(DOCS, exist_ok=True)
    with open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8") as f:
        f.write(html.replace("/*__DATA__*/null", blob))
    log(f"Звіт: тендерів з ШІ {len(tenders_rows)}, договорів {len(contract_rows)}, "
        f"сторінка {len(blob) / 1e6:.1f} МБ; відсіяно {dict(rejected)}")


def read_override_tenders():
    out = {}
    path = os.path.join(ROOT, "overrides", "tenders.csv")
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                if (r.get("tender_id") or "").strip():
                    out[r["tender_id"].strip()] = (r.get("action") or "").strip()
    return out


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["probe", "search", "details", "build", "all", "resume"])
    ap.add_argument("--full", action="store_true", help="переглянути пошук за всі місяці заново")
    ap.add_argument("--refetch", action="store_true", help="перезавантажити всі тендери і договори")
    ap.add_argument("--tender-ttl", type=float, default=3, help="оновлювати незавершені тендери раз на N днів")
    ap.add_argument("--refresh-days", type=int, default=60, help="скільки останніх днів шукати повторно")
    ap.add_argument("--contract-ttl", type=int, default=30, help="оновлювати чинні договори раз на N днів")
    ap.add_argument("--budget-min", type=float, default=300, help="ліміт часу на details, хв")
    ap.add_argument("--probe-from", help="місяць для probe, напр. 2026-03-01")
    args = ap.parse_args()
    cfg = load_config()
    if args.command == "probe":
        cmd_probe(cfg, args)
    if args.command in ("search", "all"):
        cmd_search(cfg, args)
    if args.command in ("details", "all", "resume"):
        cmd_details(cfg, args)
    if args.command in ("build", "all", "resume"):
        cmd_build(cfg, args)


if __name__ == "__main__":
    main()
