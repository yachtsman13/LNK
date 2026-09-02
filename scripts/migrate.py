# -*- coding: utf-8 -*-
"""Разбор старого плоского журнала: объекты и перевод в широкий формат."""
import re, collections

METHODS = ["ВИК", "УЗД", "УЗТ", "АЭ", "ТВ", "ТК", "ВД", "МК"]

# Маркеры номеров, встречающиеся в наименованиях объектов старого журнала.
_MARKERS = [
    ("зав", re.compile(r"зав(?:одско[йг]о?)?\s*\.?\s*№", re.I)),
    ("рег", re.compile(r"рег(?:истрационн\w*)?\s*\.?\s*№", re.I)),
    ("уч",  re.compile(r"уч(?:[её]тн\w*)?\s*\.?\s*№", re.I)),
    ("инв", re.compile(r"инв(?:ентарн\w*)?\s*\.?\s*№", re.I)),
]


# Один и тот же человек в старом журнале записан по-разному: «Д.С.Бирюков»,
# «Бирюков Д.С.». Приводим к тому виду, который стоит в справочнике специалистов,
# иначе выпадающий список в журнале не совпадёт с мигрированными значениями.
FIO_CANON = {
    "д.с.бирюков": "Д.С. Бирюков",
    "бирюков д.с.": "Д.С. Бирюков",
    "д.с. бирюков": "Д.С. Бирюков",
    "а.в. матяш": "А.В. Матяш",
    "а.в.матяш": "А.В. Матяш",
    "матяш а.в.": "А.В. Матяш",
    "п.с. сиволап": "П.С. Сиволап",
    "п.с.сиволап": "П.С. Сиволап",
    "сиволап п.с.": "П.С. Сиволап",
}


def canon_fio(s):
    v = norm(s)
    return FIO_CANON.get(v.lower().replace(" ", " "), v)


def norm(s):
    """Схлопывает переносы и повторяющиеся пробелы."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\n", " ")).strip()


def parse_object(raw):
    """Наименование объекта → (имя, зав. №, рег. №, уч. №, инв. №).

    Номера в старом журнале вписаны прямо в наименование, в разном порядке и
    с разной пунктуацией. Ищем все маркеры, значение каждого — текст до
    следующего маркера или до конца строки.
    """
    s = norm(raw)
    if not s:
        return "", "", "", "", ""
    hits = []
    for key, rx in _MARKERS:
        for m in rx.finditer(s):
            hits.append((m.start(), m.end(), key))
    hits.sort()
    if not hits:
        return s.strip(" ,;"), "", "", "", ""
    nums = {}
    for i, (start, end, key) in enumerate(hits):
        stop = hits[i + 1][0] if i + 1 < len(hits) else len(s)
        val = s[end:stop].strip(" ,;.")
        if val and key not in nums:
            nums[key] = val
    name = s[: hits[0][0]].strip(" ,;")
    return name, nums.get("зав", ""), nums.get("рег", ""), nums.get("уч", ""), nums.get("инв", "")


def object_key(name, zav, reg, uch, inv):
    """Ключ тождества объекта: заводской номер важнее наименования."""
    return (name.lower(), zav.lower(), reg.lower(), uch.lower(), inv.lower())


def _clean_num(v):
    """Номер из старого журнала: пробельный мусор превращаем в отсутствие номера."""
    if isinstance(v, str):
        v = v.strip()
        if not v:
            return None
        return int(v) if v.isdigit() else v
    return v


def read_flat(ws, first_row=4):
    """Читает старый плоский лист «Журнал»."""
    out = []
    for r in range(first_row, ws.max_row + 1):
        method = ws.cell(r, 2).value
        if not method or method == "Метод контроля":
            continue
        out.append({
            "row": r,
            "method": norm(method),
            "date": ws.cell(r, 6).value,
            "cust": norm(ws.cell(r, 7).value),
            "obj_raw": ws.cell(r, 8).value,
            "exe": canon_fio(ws.cell(r, 9).value),
            "arch": ws.cell(r, 10).value,
            "out": ws.cell(r, 11).value,
            "sender": canon_fio(ws.cell(r, 12).value),
            "orig": _clean_num(ws.cell(r, 13).value),
        })
    return out


def to_wide(flat):
    """Плоские строки → широкие. Группировка: дата + объект + заказчик.

    Если внутри группы метод повторяется, значит в этот день контролировали
    два одинаковых объекта (в старом журнале они неразличимы) — такая группа
    разбивается на несколько широких строк по порядку следования.
    """
    groups = collections.OrderedDict()
    for f in flat:
        name, zav, reg, uch, inv = parse_object(f["obj_raw"])
        f["obj"] = (name, zav, reg, uch, inv)
        key = (f["date"], object_key(*f["obj"]), f["cust"])
        groups.setdefault(key, []).append(f)

    wide, warnings = [], []
    for key, items in groups.items():
        buckets = [{}]
        for f in items:
            placed = False
            for b in buckets:
                if f["method"] not in b:
                    b[f["method"]] = f
                    placed = True
                    break
            if not placed:
                buckets.append({f["method"]: f})
        if len(buckets) > 1:
            warnings.append(
                "%s, «%s»: %d одинаковых объекта в один день — разнесены по %d строкам"
                % (key[0].strftime("%d.%m.%Y") if key[0] else "без даты",
                   items[0]["obj"][0][:50], len(buckets), len(buckets)))
        for b in buckets:
            first = list(b.values())[0]
            wide.append({
                "date": first["date"],
                "cust": first["cust"],
                "obj": first["obj"],
                "methods": b,
                "arch": first["arch"],
                "out": first["out"],
                "sender": first["sender"],
                "origs": ", ".join("%s:%s" % (k, v["orig"]) for k, v in b.items() if v["orig"] is not None),
            })
    wide.sort(key=lambda w: (w["date"] or 0, w["obj"][0]))
    return wide, warnings


def build_objects(wide):
    """Справочник объектов из широких строк, в порядке первого появления."""
    objects = collections.OrderedDict()
    for w in wide:
        k = object_key(*w["obj"])
        if k not in objects:
            objects[k] = {"name": w["obj"][0], "zav": w["obj"][1], "reg": w["obj"][2],
                          "uch": w["obj"][3], "inv": w["obj"][4], "cust": w["cust"], "n": 0}
        objects[k]["n"] += 1
    for i, o in enumerate(objects.values(), start=1):
        o["id"] = "ОБ-%04d" % i
    return objects


# ── Номера актов из исходного широкого журнала ───────────────────────
#
# Акт — отдельный уровень учёта: он один на работу и охватывает
# несколько объектов, у каждого из которых свои номера протоколов по
# методам. В плоском журнале, из которого шла первая миграция, этого
# уровня не было, поэтому номера актов берутся из исходного файла.

МЕТОД_КОЛОНКИ = {9: "ВИК", 11: "УЗТ", 13: "УЗД", 15: "ТВ",
                 17: "АЭ", 19: "ВД", 21: "ТК", 23: "МК"}


def read_acts(path, sheet="2026"):
    """(дата, ключ объекта) → список номеров актов в порядке файла."""
    import datetime
    import openpyxl
    import collections

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    out = collections.defaultdict(list)
    for r in range(3, ws.max_row + 1):
        д = ws.cell(r, 3).value
        if not isinstance(д, (datetime.date, datetime.datetime)):
            continue
        if not any(ws.cell(r, c).value not in (None, "") for c in МЕТОД_КОЛОНКИ):
            continue
        акт = norm(ws.cell(r, 2).value)
        объект = parse_object(ws.cell(r, 5).value)
        ключ = (д.date() if hasattr(д, "date") else д, object_key(*объект))
        out[ключ].append(акт)
    return out


def attach_acts(wide, acts):
    """Проставляет номер акта каждой широкой строке. Возвращает счёт."""
    import collections
    остаток = {k: list(v) for k, v in acts.items()}
    нашлось = 0
    for w in wide:
        д = w["date"]
        ключ = (д.date() if hasattr(д, "date") else д, object_key(*w["obj"]))
        w["act"] = ""
        if остаток.get(ключ):
            w["act"] = остаток[ключ].pop(0)
            if w["act"]:
                нашлось += 1
    return нашлось
