# -*- coding: utf-8 -*-
"""Вкладывает вкладку «ЛНК» (Лента/customUI14.xml) внутрь ЛНК_Программа.xlsm.

Зачем отдельный скрипт. Лента Excel описывается XML-файлом внутри книги, а
VBA такого делать не умеет: создать вкладку из макроса нельзя. Зато .xlsm —
это обычный ZIP-архив, и положить в него ещё одну часть можно снаружи, без
Excel. Двоичный блок с кодом (vbaProject.bin) при этом не трогается вовсе:
все части исходного архива переписываются дословно, добавляется одна.

Та же вкладка вкладывается и в книгу данных — ключом --данные. Вкладка
показывается, только пока активна книга, внутри которой она лежит, а работают
в книге данных; без второго вложения ленты в работе не видно вовсе. Книга
данных макросов не содержит, поэтому обработчик в её ленте зовётся по имени
книги программы — как макрос у кнопки на листе.

Запуск:
    python3 scripts/inject_ribbon.py ЛНК_Программа.xlsm [куда_сохранить.xlsm]
    python3 scripts/inject_ribbon.py --данные Реестры/ЛНК_Данные.xlsx

Повторный запуск на уже обработанной книге безопасен: скрипт заменит
прежнюю ленту новой, а не положит вторую.
"""
import io, os, re, shutil, sys, zipfile
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ЛЕНТА = os.path.join(ROOT, "Лента", "customUI14.xml")

# Для customUI14.xml тип связки именно 2007-й, хотя схема самой ленты —
# 2009/07. Это не опечатка, так устроено в Office.
ТИП_СВЯЗКИ = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
ЧАСТЬ = "customUI/customUI14.xml"
РЕЛСЫ = "_rels/.rels"
КНИГА_ПРОГРАММЫ = "ЛНК_Программа.xlsm"


def для_книги_данных(xml_текст):
    """Обработчики ленты книги данных зовутся с именем книги программы.

    Ленту Excel показывает у активной книги, а активна в работе книга
    данных. Своего кода в ней нет и быть не может — .xlsx макросов не
    хранит, — поэтому в onAction пишется полное имя: 'книга'!процедура,
    ровно так же, как у кнопок пульта, которые в этой книге давно живут
    и работают.
    """
    return re.sub(r'onAction="([^"\'!]+)"',
                  lambda m: 'onAction="\'%s\'!%s"' % (КНИГА_ПРОГРАММЫ, m.group(1)),
                  xml_текст)


def вложить(книга, куда, xml_текст):
    ET.fromstring(xml_текст)          # лента обязана быть валидным XML

    with zipfile.ZipFile(книга) as zin:
        части = [(i, zin.read(i.filename)) for i in zin.infolist()]

    имена = [i.filename for i, _ in части]
    if РЕЛСЫ not in имена:
        raise SystemExit("В файле нет %s — это не книга Office." % РЕЛСЫ)

    # [Content_Types].xml по стандарту OPC — первая часть архива. Excel
    # свои книги так и пишет, а openpyxl кладёт его последним: книгу
    # данных Excel и в таком виде читает, но ленту он ищет при открытии
    # файла, и проверять на живом Excel, простит ли он это и здесь, мне
    # нечем. Раз порядок ничего не стоит — ставим правильный.
    ТИПЫ = "[Content_Types].xml"
    новые = []
    for info, data in части:
        if info.filename == ЧАСТЬ:
            continue                  # прежняя лента выбрасывается
        if info.filename == РЕЛСЫ:
            data = поправить_связки(data.decode("utf-8")).encode("utf-8")
        новые.append((info, data))
    новые.sort(key=lambda пара: пара[0].filename != ТИПЫ)

    with zipfile.ZipFile(куда, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in новые:
            zout.writestr(info, data)
        zout.writestr(ЧАСТЬ, xml_текст.encode("utf-8"))


def поправить_связки(текст):
    # Убираем прежнюю связку на ленту, если книга уже обрабатывалась.
    текст = re.sub(r'<Relationship[^>]*ui/extensibility[^>]*/>', "", текст)
    занятые = set(re.findall(r'Id="(rId\d+)"', текст))
    n = 1
    while ("rId%d" % n) in занятые:
        n += 1
    связка = '<Relationship Id="rId%d" Type="%s" Target="%s"/>' % (n, ТИП_СВЯЗКИ, ЧАСТЬ)
    if "</Relationships>" not in текст:
        raise SystemExit("Неожиданный вид %s — вложение отменено." % РЕЛСЫ)
    return текст.replace("</Relationships>", связка + "</Relationships>")


def проверить(файл):
    """Сверка после записи: части на месте, оба XML разбираются."""
    with zipfile.ZipFile(файл) as z:
        имена = z.namelist()
        if ЧАСТЬ not in имена:
            raise SystemExit("Лента не попала в файл.")
        rels = z.read(РЕЛСЫ).decode("utf-8")
        if ТИП_СВЯЗКИ not in rels:
            raise SystemExit("Связка на ленту не записалась.")
        ET.fromstring(rels)
        ET.fromstring(z.read(ЧАСТЬ))
        if имена[0] != "[Content_Types].xml":
            raise SystemExit("[Content_Types].xml не первой частью — вложение отменено.")
        if 'Extension="xml"' not in z.read("[Content_Types].xml").decode("utf-8"):
            raise SystemExit("Для .xml нет типа по умолчанию — Excel счёл бы книгу битой.")
        return имена


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    аргументы = [a for a in sys.argv[1:] if a != "--данные"]
    книга_данных = "--данные" in sys.argv
    if not аргументы:
        print(__doc__)
        return 1
    книга = аргументы[0]
    куда = аргументы[1] if len(аргументы) > 1 else книга
    if not os.path.exists(книга):
        raise SystemExit("Не найден файл: %s" % книга)

    xml_текст = io.open(ЛЕНТА, encoding="utf-8").read()
    if книга_данных:
        xml_текст = для_книги_данных(xml_текст)
    было = zipfile.ZipFile(книга).namelist()
    имел_vba = "xl/vbaProject.bin" in было

    временный = куда + ".tmp"
    вложить(книга, временный, xml_текст)
    стало = проверить(временный)

    потеряно = [n for n in было if n not in стало and n != ЧАСТЬ]
    if потеряно:
        os.remove(временный)
        raise SystemExit("Части пропали бы: %s — вложение отменено." % потеряно[:5])
    if имел_vba and "xl/vbaProject.bin" not in стало:
        os.remove(временный)
        raise SystemExit("Пропал бы код книги — вложение отменено.")

    shutil.move(временный, куда)
    print("Лента вложена: %s" % куда)
    print("  частей было %d, стало %d" % (len(было), len(стало)))
    print("  код книги (vbaProject.bin): %s" %
          ("на месте" if "xl/vbaProject.bin" in стало else "в книге его нет"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
