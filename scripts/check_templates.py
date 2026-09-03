# -*- coding: utf-8 -*-
"""Сверка плейсхолдеров шаблонов Word с источниками подстановки.

Плейсхолдер, которому неоткуда взяться, останется в готовом документе видимым
«#Заключение» — заметят это в лучшем случае при вычитке, в худшем у заказчика.
Проверка находит такие места до печати.

Источниками считаются: колонки листов «Объекты» и «Данные_протоколов»
(их заголовки превращаются в имена полей), явные подстановки в коде и
таблицы точек замеров.
"""
import io, os, re, sys, glob, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "Реестры", "ЛНК_Данные.xlsx")
TPL = os.path.join(ROOT, "Шаблоны")
MAC = os.path.join(ROOT, "Макросы")

# Заполняется вручную по существу дела — не пробел, а осознанное решение.
РУЧНЫЕ = {"Место_для_схемы_или_вставленного_рисунка",
          "Место_для_графика_или_вставленного_рисунка"}

# Вне выбранного объёма автоподстановки (сейчас — УЗТ и ТВ).
ВНЕ_ОБЪЁМА = {}


def norm(s):
    return re.sub(r"\s+", " ", str(s or "").replace("\n", " ")).strip()


def field(s):
    return re.sub(r"[ -]", "_", norm(s))


def источники():
    import openpyxl
    out = set()
    wb = openpyxl.load_workbook(DATA, read_only=True)
    for лист, старт in (("Объекты", 2), ("Данные_протоколов", 4), ("Данные_АЭ", 3)):
        строка = next(wb[лист].iter_rows(min_row=3, max_row=3))
        for c in list(строка)[старт - 1:]:
            f = field(c.value)
            if f and not re.match(
                    r"^(ID|Кем_внесено|Изменено|Работ_в_журнале|Переопределения|Прежние)", f):
                out.add(f)
    код = "".join(io.open(p, encoding="utf-8").read()
                  for p in glob.glob(os.path.join(MAC, "*.bas")))
    for m in re.finditer(r'(?:Подставить|ЗаменитьВДокументе)\s+\w+,\s*"([^"]+)"', код):
        out.add(m.group(1))
    for k in (1, 2):
        for s in ("Наименование", "Тип", "Зав_номер", "Поверка_до"):
            out.add("СИ%d_%s" % (k, s))
    for t in range(1, 5):
        for c in range(1, 7):
            out.add("Т%dС%d" % (t, c))
        for s in ("HB", "Расч", "Норм", "ЗавНомер"):
            out.add("Т%d_%s" % (t, s))
    return out


def плейсхолдеры(путь):
    xml = zipfile.ZipFile(путь).read("word/document.xml").decode("utf-8")
    txt = re.sub(r"<[^>]+>", "", re.sub(r"</w:(p|tc)>", "\n", xml))
    return sorted(set(re.findall(r"#([А-Яа-яЁёA-Za-z0-9_№]+)", txt)))


def main():
    ист = источники()
    print("Источников подстановки: %d\n" % len(ист))
    всего = пробелов = 0
    ошибки = []
    for путь in sorted(glob.glob(os.path.join(TPL, "*.docx"))):
        имя = os.path.basename(путь)
        ph = плейсхолдеры(путь)
        нет = [p for p in ph if p not in ист and p not in РУЧНЫЕ]
        всего += len(ph)
        пробелов += len(нет)
        пометка = ""
        if нет:
            пометка = "— " + ВНЕ_ОБЪЁМА.get(имя, "ТРЕБУЕТ ВНИМАНИЯ")
            if имя not in ВНЕ_ОБЪЁМА:
                ошибки.append("%s: %s" % (имя, ", ".join(нет)))
        print("%-26s полей %3d, без источника %2d %s" % (имя, len(ph), len(нет), пометка))
        if нет:
            print("      " + ", ".join(нет))

    print("\nИтого: %d плейсхолдеров, без источника %d (%d%%)."
          % (всего, пробелов, 100 * пробелов // максимум(всего)))
    print("Рисунки и схемы вставляются вручную по существу — они не в счёт.")
    if ошибки:
        print("\nПРОВЕРКА НЕ ПРОЙДЕНА — пробелы вне ожидаемых:")
        for e in ошибки:
            print("  " + e)
        return 1
    print("Проверка пройдена: все пробелы объяснимы выбранным объёмом работ.")
    return 0


def максимум(n):
    return n if n else 1


if __name__ == "__main__":
    sys.exit(main())
