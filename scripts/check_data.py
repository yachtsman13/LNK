# -*- coding: utf-8 -*-
"""Проверка собранной книги ЛНК_Данные.xlsx: сверка с исходным журналом и связность ссылок."""
import sys, os, re, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from migrate import METHODS, read_flat

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "Реестры", "ЛНК_Данные.xlsx")
SRC = os.path.join(ROOT, "Реестры", "ЛНК_Реестры_v2.xlsx")
FIRST = 4

errors, notes = [], []


def check(cond, msg):
    (notes if cond else errors).append(("OK  " if cond else "ОШИБКА ") + msg)


def main():
    wb = openpyxl.load_workbook(OUT)
    src = openpyxl.load_workbook(SRC, data_only=True)
    flat = read_flat(src["Журнал"])

    # ── 1. сверка количества записей по методам ──────────────────────
    ws = wb["Журнал"]
    hdr = {ws.cell(3, c).value: c for c in range(1, ws.max_column + 1)}
    cnt = collections.Counter()
    rows = 0
    for r in range(FIRST, ws.max_row + 1):
        if not ws.cell(r, hdr["Дата контроля"]).value:
            continue
        rows += 1
        for m in METHODS:
            if ws.cell(r, hdr["%s\n№" % m]).value not in (None, ""):
                cnt[m] += 1
    was = collections.Counter(f["method"] for f in flat)
    print("Сверка по методам (старый плоский журнал → новый широкий):")
    print("  %-6s %8s %8s %s" % ("метод", "было", "стало", "расхождение"))
    for m in METHODS:
        d = cnt[m] - was[m]
        print("  %-6s %8d %8d %s" % (m, was[m], cnt[m], "—" if d == 0 else "%+d" % d))
        check(d == 0, "метод %s: %d → %d" % (m, was[m], cnt[m]))
    check(sum(cnt.values()) == len(flat),
          "всего работ: %d → %d" % (len(flat), sum(cnt.values())))
    print("  %-6s %8d %8d\n" % ("ИТОГО", len(flat), sum(cnt.values())))
    check(rows == 315, "широких строк: %d" % rows)

    # ── 2. пропуски в нумерации по каждому методу ────────────────────
    unknown = sum(1 for r in range(FIRST, ws.max_row + 1) for m in METHODS
                  if ws.cell(r, hdr["%s\n№" % m]).value == "?")
    check(True, "записей без номера в старом журнале (стоит «?»): %d" % unknown)

    print("Нумерация по методам:")
    for m in METHODS:
        nums = sorted(v for r in range(FIRST, ws.max_row + 1)
                      for v in [ws.cell(r, hdr["%s\n№" % m]).value]
                      if isinstance(v, int))
        if not nums:
            continue
        gaps = sorted(set(range(min(nums), max(nums) + 1)) - set(nums))
        dups = [n for n, k in collections.Counter(nums).items() if k > 1]
        print("  %-4s %d..%d, записей %d, пропусков %d%s%s"
              % (m, min(nums), max(nums), len(nums), len(gaps),
                 (", пропущены " + ", ".join(map(str, gaps[:12]))) if gaps else "",
                 (", ДУБЛИ: " + ", ".join(map(str, dups))) if dups else ""))
        check(not dups, "метод %s: дублей номеров нет" % m)

    # ── 3. выбранные объекты есть в справочнике ──────────────────────
    # Раньше в журнале хранился ID объекта. Теперь там читаемое
    # обозначение, а ID считается формулой — поэтому сверяем обозначения:
    # промах здесь означает, что выпадающий список ничего не найдёт.
    def обозначение(ws, r):
        s = str(ws.cell(r, 2).value or "").strip()
        for c, префикс in ((3, ", зав. № "), (4, ", рег. № "), (5, ", уч. № ")):
            v = str(ws.cell(r, c).value or "").strip()
            if v:
                s += префикс + v
        return s

    об = wb["Объекты"]
    справочник = {обозначение(об, r) for r in range(FIRST, об.max_row + 1)
                  if об.cell(r, 1).value}
    выбрано = {str(ws.cell(r, hdr["Объект контроля\n(выбор из справочника)"]).value or "").strip()
               for r in range(FIRST, ws.max_row + 1)
               if ws.cell(r, hdr["Дата контроля"]).value}
    выбрано.discard("")
    check(not (выбрано - справочник),
          "выбранные в журнале объекты есть в справочнике (объектов %d); лишние: %s"
          % (len(справочник), list(выбрано - справочник)[:3]))

    # ── 4. исполнители совпадают со справочником специалистов ────────
    staff = {wb["Специалисты"].cell(r, 3).value
             for r in range(FIRST, wb["Специалисты"].max_row + 1)
             if wb["Специалисты"].cell(r, 3).value}
    used = set()
    for r in range(FIRST, ws.max_row + 1):
        for m in METHODS:
            v = ws.cell(r, hdr["%s\nисполнитель" % m]).value
            if v:
                used.add(v)
        v = ws.cell(r, hdr["Отправитель"]).value
        if v:
            used.add(v)
    check(not (used - staff), "ФИО в журнале совпадают со справочником; лишние: %s" % (used - staff))

    # ── 5. связка ППР ↔ СИ по номеру карточки ────────────────────────
    cards_si = {wb["СИ"].cell(r, 2).value for r in range(FIRST, wb["СИ"].max_row + 1)
                if wb["СИ"].cell(r, 2).value}
    cards_to = {wb["СИ_Обслуживание"].cell(r, 2).value
                for r in range(FIRST, wb["СИ_Обслуживание"].max_row + 1)
                if wb["СИ_Обслуживание"].cell(r, 2).value}
    check(not (cards_to - cards_si),
          "все карточки из обслуживания есть в реестре СИ (карточек %d)" % len(cards_si))

    # ── 6. формулы ссылаются на существующие листы ───────────────────
    sheets = set(wb.sheetnames)
    refs = set()
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for nm in re.findall(r"([А-Яа-яЁёA-Za-z_][А-Яа-яЁёA-Za-z0-9_]*)!", c.value):
                        refs.add(nm)
    check(not (refs - sheets), "формулы ссылаются только на существующие листы: %s" % (refs - sheets))

    # ── 7. запрещённые функции (нет в LibreOffice / старом Excel) ────
    banned = ("FILTER(", "XLOOKUP(", "UNIQUE(", "SORT(", "TEXTJOIN(", "MAXIFS(", "IFS(")
    hits = []
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for b in banned:
                        if b in c.value.upper():
                            hits.append("%s!%s %s" % (w.title, c.coordinate, b))
    check(not hits, "нет функций, несовместимых с LibreOffice: %s" % hits[:5])

    # ── 8. даты — настоящие даты, а не текст ─────────────────────────
    import datetime
    bad_dates = []
    for r in range(FIRST, 320):
        v = ws.cell(r, hdr["Дата контроля"]).value
        if v is not None and not isinstance(v, (datetime.date, datetime.datetime)):
            bad_dates.append(r)
    check(not bad_dates, "даты контроля — настоящие даты (проверено 316 строк)")

    print("\n" + "\n".join(notes))
    if errors:
        print("\n".join(errors))
        print("\nПРОВЕРКА НЕ ПРОЙДЕНА: %d ошибок" % len(errors))
        return 1
    print("\nПроверка пройдена: %d контролей." % len(notes))
    return 0


if __name__ == "__main__":
    sys.exit(main())
