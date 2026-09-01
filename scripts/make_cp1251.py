# -*- coding: utf-8 -*-
"""Готовит копии модулей в cp1251 для импорта в редактор VBA.

Редактор VBA на русской Windows читает .bas в системной кодировке (cp1251).
Файл в UTF-8 он импортирует как набор нечитаемых символов, причём молча —
код при этом не компилируется, а причина неочевидна. Поэтому исходники
лежат в UTF-8 (их видно в истории изменений), а для импорта собирается
отдельная папка в cp1251.
"""
import io, os, glob, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "Макросы")
DST = os.path.join(SRC, "Для_импорта")


def main():
    if not os.path.isdir(DST):
        os.makedirs(DST)
    n, bad = 0, []
    for path in sorted(glob.glob(os.path.join(SRC, "*.bas")) +
                       glob.glob(os.path.join(SRC, "*.txt"))):
        text = io.open(path, encoding="utf-8").read()
        # Заменяем только то, чего в cp1251 нет, и только в оформлении.
        # Кавычки-ёлочки, тире, многоточие и особенно N — в cp1251 есть и
        # остаются как были: N входит в заголовки колонок, по которым код
        # ищет данные, и его подмена сломала бы весь поиск.
        for a, b in (("═", "="), ("─", "-"), ("▸", ">"),
                     ("→", "->"), ("≤", "<="), ("≥", ">="), ("•", "-")):
            text = text.replace(a, b)
        out = os.path.join(DST, os.path.basename(path))
        try:
            io.open(out, "w", encoding="cp1251", newline="\r\n").write(text)
        except UnicodeEncodeError as e:
            bad.append("%s: %s" % (os.path.basename(path), e))
            continue
        # Обратная сверка: файл должен читаться из cp1251 ровно тем же
        # текстом. Молча испорченная кириллица — худший вид поломки:
        # код не компилируется, а причина неочевидна.
        back = io.open(out, encoding="cp1251", newline="").read().replace("\r\n", "\n")
        if back != text:
            bad.append("%s: обратное чтение не совпало" % os.path.basename(path))
            continue
        if "№" in text and "№" not in back:
            bad.append("%s: потерян знак номера" % os.path.basename(path))
            continue
        n += 1
    print("Записано в cp1251: %d файлов → %s" % (n, DST))
    for b in bad:
        print("  НЕ УДАЛОСЬ  " + b)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
