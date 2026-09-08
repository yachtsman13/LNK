# -*- coding: utf-8 -*-
"""Сверка ленты с кодом: каждая кнопка должна что-то делать.

Ошибка в ленте не видна ничем: Excel просто не показывает вкладку (если
сломан XML) или показывает кнопку, которая молча ничего не делает (если
tag разошёлся с текстом команды). Ни то, ни другое не поймает ни один
контроль в книге, поэтому лента сверяется отдельно.
"""
import io, os, re, sys
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ЛЕНТА = os.path.join(ROOT, "Лента", "customUI14.xml")
MAC = os.path.join(ROOT, "Макросы")
NS = "{http://schemas.microsoft.com/office/2009/07/customui}"

errors, notes = [], []


def check(cond, msg):
    (notes if cond else errors).append(("OK  " if cond else "ОШИБКА ") + msg)


def main():
    xml_текст = io.open(ЛЕНТА, encoding="utf-8").read()
    try:
        корень = ET.fromstring(xml_текст)
    except ET.ParseError as e:
        print("ОШИБКА лента не разбирается как XML: %s" % e)
        return 1
    check(True, "лента разбирается как XML")

    check(корень.tag == NS + "customUI",
          "пространство имён customUI 2009/07 (иначе Excel молча не покажет вкладку)")

    кнопки = корень.iter(NS + "button")
    команды, ids, действия = [], [], set()
    for b in кнопки:
        ids.append(b.get("id"))
        действия.add(b.get("onAction"))
        команды.append((b.get("id"), b.get("tag")))

    check(len(ids) == len(set(ids)),
          "id кнопок не повторяются (%d кнопок)" % len(ids))

    без_tag = [i for i, t in команды if not t]
    check(not без_tag, "у каждой кнопки есть tag с командой%s"
          % ("" if not без_tag else ": " + ", ".join(без_tag)))

    # onAction должен быть Public-процедурой в модулях
    исходники = ""
    for имя in sorted(os.listdir(MAC)):
        if имя.endswith(".bas"):
            исходники += io.open(os.path.join(MAC, имя), encoding="utf-8").read() + "\n"
    нет_обработчика = [a for a in действия
                       if not re.search(r"^\s*Public\s+Sub\s+%s\b" % re.escape(a or ""),
                                        исходники, re.M)]
    check(not нет_обработчика,
          "onAction ведёт в существующую Public-процедуру%s"
          % ("" if not нет_обработчика else ": " + ", ".join(map(str, нет_обработчика))))

    # tag должен совпадать с веткой Case в ЗапуститьПунктМеню
    меню = io.open(os.path.join(MAC, "modMenu.bas"), encoding="utf-8").read()
    ветви = set(re.findall(r'Case\s+"([^"]+)"\s*:', меню))
    чужие = sorted({t for _, t in команды if t and t not in ветви})
    check(not чужие,
          "tag каждой кнопки разбирается обработчиком (кнопок %d, ветвей %d)%s"
          % (len(команды), len(ветви),
             "" if not чужие else "; нет ветви для: " + "; ".join(чужие)))

    print("\n".join(notes))
    if errors:
        print("\n".join(errors))
        print("\nПРОВЕРКА НЕ ПРОЙДЕНА: %d ошибок" % len(errors))
        return 1
    print("\nПроверка пройдена: %d контролей." % len(notes))
    return 0


if __name__ == "__main__":
    sys.exit(main())
