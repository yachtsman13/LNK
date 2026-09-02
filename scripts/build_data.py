# -*- coding: utf-8 -*-
"""Сборка книги данных ЛНК_Данные.xlsx (этапы 1-2 + полный состав колонок реестров)."""
import sys, os, re, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment

from lnk_common import *
from migrate import METHODS, read_flat, to_wide, build_objects, parse_object, norm

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_V2 = os.path.join(ROOT, "Реестры", "ЛНК_Реестры_v2.xlsx")
SRC_PPR = os.path.join(ROOT, "Исходники", "План-график_ППР_2026-2030.xlsx")
OUT = os.path.join(ROOT, "Реестры", "ЛНК_Данные.xlsx")

HDR = 3          # строка заголовков таблиц
FIRST = HDR + 1  # первая строка данных
MAXR = 1200      # журнал и объекты: формулы и проверки протянуты до этой строки
MAXR_REG = 400   # остальные реестры — их строки прибывают на порядок медленнее

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
PPR_YEARS = [2026, 2027, 2028, 2029, 2030]
TO_KINDS = {  # категория → (название периодичности, интервал в днях)
    "ТО-1": ("Ежедневно (по рабочим дням)", 1),
    "ТО-2": ("Подекадно (раз в 10 дней)", 10),
    "ТО-3": ("Ежемесячно", 30),
    "ТО-4": ("Ежеквартально", 91),
}


def col(t, w):
    return {"t": t, "w": w}


def mark_input(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).font = F_INPUT


def mark_calc(ws, r1, r2, cols):
    for r in range(r1, r2 + 1):
        for c in cols:
            cell = ws.cell(r, c)
            cell.font, cell.fill = F_CALC, FILL_LOCK


def add_status_cf(ws, letter, r1, r2):
    rng = "%s%d:%s%d" % (letter, r1, letter, r2)
    for text, fill, font in STATUS_COLORS:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=['ISNUMBER(SEARCH("%s",%s%d))' % (text, letter, r1)],
            fill=PatternFill("solid", fgColor=fill),
            font=Font(color=font, bold=(text == "ПРОСРОЧЕНО"))))


def add_list_dv(ws, letter, source, r1=FIRST, r2=MAXR, allow_blank=True):
    dv = DataValidation(type="list", formula1=source, allow_blank=allow_blank, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add("%s%d:%s%d" % (letter, r1, letter, r2))
    return dv


def banded(ws, r1, r2, ncols):
    """Тонкая зебра — читать длинные строки без линейки."""
    fill = PatternFill("solid", fgColor=C_BAND)
    for r in range(r1, r2 + 1):
        if r % 2 == 0:
            for c in range(1, ncols + 1):
                if ws.cell(r, c).fill.fgColor.rgb in (None, "00000000"):
                    ws.cell(r, c).fill = fill


def set_date(cell, value):
    if isinstance(value, datetime.datetime):
        cell.value = value.date()
    elif isinstance(value, datetime.date):
        cell.value = value
    else:
        cell.value = value
        return
    cell.number_format = DATE_FMT


# ══════════════════════════════════════════════════════════════════════
# Настройки
# ══════════════════════════════════════════════════════════════════════
SETTINGS = [
    ("Рабочее место", None, None),
    ("Код рабочего места", "", "ПК1 … ПК6. Задаётся один раз при установке, входит в ID записей. В документах не печатается."),
    ("ФИО сотрудника", "", "Подставляется в поле «Кем внесено»."),
    ("Это мастер-ПК", "Нет", "Да / Нет. Мастер собирает свод. Роль переключается кнопкой, файл Свод\\мастер.txt в общей папке."),
    ("Пути", None, None),
    ("Папка программы", "D:\\ЛНК\\", "Где лежат ЛНК_Программа.xlsm и ЛНК_Данные.xlsx."),
    ("Папка шаблонов", "D:\\ЛНК\\Шаблоны\\", "Шаблоны Word с #Плейсхолдерами."),
    ("Папка готовых документов", "D:\\ЛНК\\Готовые\\", "Сформированные акты и протоколы, по годам и методам."),
    ("Папка локальных копий", "D:\\ЛНК\\Копии\\", "Копии данных при закрытии книги, 30 дней + помесячно."),
    ("Режим обмена", "Авто", "Авто — сначала сеть, если недоступна, штатная флешка. "
                             "Только флешка / Только сеть — если нужно жёстко."),
    ("Общая папка обмена", "", "Сетевой путь, например \\\\сервер\\ЛНК_Обмен\\. Пока сети нет, "
                               "оставьте пустым: программа сама найдёт штатную флешку по "
                               "серийному номеру тома, буква диска значения не имеет."),
    ("Носитель", None, None),
    ("Серийные номера разрешённых флешек", "", "Через запятую. Заполняется мастером, дублирует Версия\\носители.txt."),
    ("Хранение копий", None, None),
    ("Хранить ежедневных копий, дней", 30, "После этого срока ежедневные копии удаляются, помесячные остаются."),
    ("Хранить помесячных копий, мес.", 12, ""),
    ("Хранить состояний свода", 60, "История свода у мастера, перед каждой сборкой."),
    ("Учётные номера журналов", None, None),
    ("Учётный номер журнала ВИК", "", "Ставится в шапку печатного журнала по методу."),
    ("Учётный номер журнала УЗД", "", ""),
    ("Учётный номер журнала УЗТ", "", ""),
    ("Учётный номер журнала АЭ", "", ""),
    ("Учётный номер журнала ТВ", "", ""),
    ("Учётный номер журнала ТК", "", ""),
    ("Учётный номер журнала ВД", "", ""),
    ("Учётный номер журнала МК", "", ""),
    ("Учётный номер журнала электроинструмента", "", ""),
    ("Паспорт", None, None),
    ("Полное наименование ЛНК", "", "Подставляется в документы Word вместо #Наименование_ЛНК. "
                                     "То же значение — на листе «Паспорт_постоянная»."),
    ("Сшив и печать", None, None),
    ("Текст бирки", "", "Что печатается на бирке при сшиве. Подстановки: #ЧИСЛО, #ПРОПИСЬЮ, "
                        "#ЛИСТОВ, #ОБЪЕКТ. Пусто — берётся типовой текст."),
    ("Служебное", None, None),
    ("Версия программы", "1.0", "Заполняется при обновлении и публикации версии."),
    ("Версия структуры данных", "1.0", "Не менять вручную."),
    ("Счётчик записей рабочего места", 0, "Последний выданный порядковый номер в ID записи. Не менять вручную."),
]


def build_settings(wb):
    ws = wb.create_sheet("Настройки")
    sheet_head(ws, "Настройки рабочего места",
               "Заполняется один раз при установке. Синим — то, что вводится руками.", 3)
    table_head(ws, HDR, [col("Параметр", 42), col("Значение", 34), col("Пояснение", 86)])
    r = FIRST
    for name, value, hint in SETTINGS:
        if value is None and hint is None:
            c = ws.cell(r, 1, name)
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=C_HEAD2)
            for cc in (2, 3):
                ws.cell(r, cc).fill = PatternFill("solid", fgColor=C_HEAD2)
        else:
            ws.cell(r, 1, name).font = F_TXT
            v = ws.cell(r, 2, value)
            v.font, v.alignment = F_INPUT, A_CTR
            h = ws.cell(r, 3, hint)
            h.font, h.alignment = F_HINT, A_TXT
        for cc in (1, 2, 3):
            ws.cell(r, cc).border = BORDER
        ws.row_dimensions[r].height = 20
        r += 1
    add_list_dv(ws, "B", '"Да,Нет"', FIRST, r - 1, allow_blank=False)
    dv = DataValidation(type="list", formula1='"Авто,Только флешка,Только сеть"',
                        allow_blank=False, showErrorMessage=False)
    ws.add_data_validation(dv)
    for rr in range(FIRST, r):
        if str(ws.cell(rr, 1).value) == "Режим обмена":
            dv.add(ws.cell(rr, 2))
    # именованный диапазон, чтобы формулы и макросы не зависели от номеров строк
    return ws


# ══════════════════════════════════════════════════════════════════════
# Списки — источник всех выпадающих
# ══════════════════════════════════════════════════════════════════════
LISTS = [
    ("Методы контроля", METHODS),
    ("Статус записи", ["Черновик", "Зарегистрирован", "Аннулирован"]),
    ("Уровень квалификации", ["I", "II", "III"]),
    ("Вид метрологического контроля", ["Поверка", "Калибровка", "Аттестация", "Не подлежит"]),
    ("Вид отсутствия", ["Отпуск", "Больничный", "Командировка", "Выезд на измерения", "Прочее"]),
    ("Доступен для работ", ["Да", "Нет"]),
    ("Категория ТО", ["ТО-1", "ТО-2", "ТО-3", "ТО-4", "ТР", "СР"]),
    ("Вид работ по обслуживанию", ["ТО", "Текущий ремонт", "Средний ремонт",
                                    "Внеплановое", "Проверка после транспортировки"]),
    ("Результат обслуживания", ["Исправен", "Неисправен", "Выведен из эксплуатации"]),
    ("Раздел Паспорта", ["Средства НК", "Меры и образцы", "Дозиметрия", "Вспомогательное"]),
    ("Статус задачи", ["Запланировано", "В работе", "Выполнено", "Отменено"]),
    ("Приоритет", ["Высокий", "Средний", "Низкий"]),
    ("Основание привлечения", ["Трудовой договор", "Договор ГПХ", "Совместительство"]),
    ("Решение по рекламации", ["Повторный контроль", "Прекращение работ",
                                "Внеплановая внутренняя проверка", "Мотивированный отказ"]),
    ("Тип объекта", ["Сосуд", "Трубопровод", "Металлоконструкция", "Ёмкость",
                      "Здание, сооружение", "Механизм", "Прочее"]),
]


def build_lists(wb):
    ws = wb.create_sheet("Списки")
    sheet_head(ws, "Списки для выпадающих полей",
               "Отсюда берутся все выпадающие списки книги. Добавили строку в конец колонки — "
               "значение сразу появилось в списке. Названия колонок не переименовывать.", len(LISTS))
    table_head(ws, HDR, [col(name, max(16, min(30, len(name) + 4))) for name, _ in LISTS])
    for i, (name, values) in enumerate(LISTS, start=1):
        for j, v in enumerate(values):
            c = ws.cell(FIRST + j, i, v)
            c.font, c.border, c.alignment = F_INPUT, BORDER, A_CTR
    ws.sheet_properties.tabColor = "7F7F7F"
    return ws


def list_ref(name, n=40):
    """Ссылка на колонку списка с запасом на дописывание значений."""
    idx = [i for i, (nm, _) in enumerate(LISTS, start=1) if nm == name][0]
    L = get_column_letter(idx)
    return "=Списки!$%s$%d:$%s$%d" % (L, FIRST, L, FIRST + n)


# ══════════════════════════════════════════════════════════════════════
# Журнал — широкий формат
# ══════════════════════════════════════════════════════════════════════
J = {}  # имя поля → номер колонки


def journal_cols():
    cols = [col("ID записи", 16), col("№ п/п", 7), col("Статус", 15),
            col("Дата контроля", 13), col("Год", 7),
            col("Заказывающее\nподразделение", 22),
            col("ID объекта", 12), col("Наименование объекта\nв документе", 44),
            col("Заводской №\nобъекта", 16), col("Рег. / уч. №\nобъекта", 14)]
    # Заголовок не менять: по нему код ищет колонку (ColReq в модулях).
    for m in METHODS:
        cols += [col("%s\n№" % m, 7), col("%s\nисполнитель" % m, 17)]
    cols += [col("№ экз., дата передачи\nв архив ЛНК", 20),
             col("№ экз., исходящий,\nдата отправки", 20),
             col("Отправитель", 16),
             col("Документ\nсформирован", 14),
             col("Ориг. № старого\nжурнала (сверка)", 15),
             col("Кем внесено", 14), col("Изменено", 14), col("Примечание", 26)]
    return cols


def build_journal(wb, wide, objects):
    ws = wb.create_sheet("Журнал")
    cols = journal_cols()
    n = len(cols)
    global J
    names = ["ID", "NP", "STATUS", "DATE", "YEAR", "CUST", "OBJID", "OBJNAME", "OBJZAV", "OBJREG"]
    for m in METHODS:
        names += ["N_" + m, "E_" + m]
    names += ["ARCH", "OUT", "SENDER", "GEN", "ORIG", "BY", "CHG", "NOTE"]
    J = {nm: i for i, nm in enumerate(names, start=1)}

    sheet_head(ws, "Журнал учёта заключений (протоколов, актов) ЛНК",
               "Одна строка — один объект контроля за одну дату; применённые методы стоят колонками. "
               "Метод считается применённым, если по нему выбран ИСПОЛНИТЕЛЬ. Колонка «№» серая — "
               "номер туда впишет система при регистрации, руками её не заполнять. "
               "Печатные журналы по методам строятся отсюда макросом. "
               "Мигрировано %d записей из журнала 2026 года (было %d строк плоского вида). "
               "Серым — расчётные колонки, их не трогать." % (len(wide), 1001), n)
    table_head(ws, HDR, cols)

    for i, w in enumerate(wide):
        r = FIRST + i
        ws.cell(r, J["ID"], "МИГР-2026-%04d" % (i + 1))
        ws.cell(r, J["STATUS"], "Зарегистрирован")
        set_date(ws.cell(r, J["DATE"]), w["date"])
        ws.cell(r, J["CUST"], w["cust"] or None)
        okey = (w["obj"][0].lower(), w["obj"][1].lower(), w["obj"][2].lower(),
                w["obj"][3].lower(), w["obj"][4].lower())
        ws.cell(r, J["OBJID"], objects[okey]["id"])
        ws.cell(r, J["OBJNAME"], w["obj"][0])
        for m, f in w["methods"].items():
            ws.cell(r, J["E_" + m], f["exe"] or None)
            # номер из старого журнала; «?» — номер там был пустым, его допишут с бумаги
            ws.cell(r, J["N_" + m], f["orig"] if f["orig"] not in (None, "") else "?")
        ws.cell(r, J["ARCH"], w["arch"])
        ws.cell(r, J["OUT"], w["out"])
        ws.cell(r, J["SENDER"], w["sender"] or None)
        ws.cell(r, J["ORIG"], w["origs"] or None)
        ws.cell(r, J["BY"], "миграция")

    last = FIRST + len(wide) - 1
    for r in range(FIRST, MAXR + 1):
        ws.cell(r, J["NP"], '=IF($D{r}="","",COUNT($D${f}:$D{r}))'.format(r=r, f=FIRST))
        ws.cell(r, J["YEAR"], '=IF($D{r}="","",YEAR($D{r}))'.format(r=r))
        ws.cell(r, J["OBJZAV"],
                '=IF($G{r}="","",IFERROR(INDEX(Объекты!$C:$C,MATCH($G{r},Объекты!$A:$A,0)),""))'.format(r=r))
        ws.cell(r, J["OBJREG"],
                '=IF($G{r}="","",IFERROR(INDEX(Объекты!$D:$D,MATCH($G{r},Объекты!$A:$A,0)),""))'.format(r=r))
        for c in range(1, n + 1):
            ws.cell(r, c).border = BORDER
            ws.cell(r, c).alignment = A_TXT
        ws.row_dimensions[r].height = 28

    mark_input(ws, FIRST, MAXR, J["STATUS"], J["OBJID"])
    mark_input(ws, FIRST, MAXR, J["OBJNAME"], J["OBJNAME"])
    mark_input(ws, FIRST, MAXR, J["ARCH"], n)
    # Исполнитель — поле ввода: именно он отмечает, что метод применялся.
    # Колонка «№» — серая: номер туда впишет система при регистрации,
    # руками её заполнять не надо, и выглядеть она должна не как поле.
    for m in METHODS:
        mark_input(ws, FIRST, MAXR, J["E_" + m], J["E_" + m])
    mark_calc(ws, FIRST, MAXR, [J["N_" + m] for m in METHODS])
    mark_calc(ws, FIRST, MAXR, [J["NP"], J["YEAR"], J["OBJZAV"], J["OBJREG"]])
    mark_calc(ws, FIRST, MAXR, [J["ID"]])
    for r in range(FIRST, MAXR + 1):
        ws.cell(r, J["DATE"]).number_format = DATE_FMT
        ws.cell(r, J["GEN"]).number_format = DATE_FMT
        ws.cell(r, J["DATE"]).alignment = A_CTR
        for m in METHODS:
            ws.cell(r, J["N_" + m]).alignment = A_CTR

    add_list_dv(ws, get_column_letter(J["STATUS"]), list_ref("Статус записи", 5))
    add_list_dv(ws, get_column_letter(J["OBJID"]), "=Объекты!$A$%d:$A$%d" % (FIRST, MAXR))
    for m in METHODS:
        add_list_dv(ws, get_column_letter(J["E_" + m]), "=Специалисты!$C$%d:$C$%d" % (FIRST, 200))
    add_list_dv(ws, get_column_letter(J["SENDER"]), "=Специалисты!$C$%d:$C$%d" % (FIRST, 200))

    # черновики видны сразу
    ws.conditional_formatting.add(
        "A%d:%s%d" % (FIRST, get_column_letter(n), MAXR),
        FormulaRule(formula=['$C%d="Черновик"' % FIRST],
                    fill=PatternFill("solid", fgColor="FFF2CC")))
    ws.conditional_formatting.add(
        "A%d:%s%d" % (FIRST, get_column_letter(n), MAXR),
        FormulaRule(formula=['$C%d="Аннулирован"' % FIRST],
                    font=Font(color="9C0006", strike=True)))
    ws.sheet_properties.tabColor = "1F3864"
    return ws, last


def finish(ws, ncols, r1=FIRST, r2=MAXR, rowh=24):
    for r in range(r1, r2 + 1):
        for c in range(1, ncols + 1):
            ws.cell(r, c).border = BORDER
            if not ws.cell(r, c).alignment.wrap_text:
                ws.cell(r, c).alignment = A_TXT
        ws.row_dimensions[r].height = rowh


# ══════════════════════════════════════════════════════════════════════
# Объекты контроля
# ══════════════════════════════════════════════════════════════════════
def build_objects_sheet(wb, objects):
    ws = wb.create_sheet("Объекты")
    cols = [col("ID", 11), col("Каноническое наименование", 46), col("Заводской №", 18),
            col("Регистрационный №", 15), col("Учётный №", 13), col("Инвентарный №", 13),
            col("Тип / группа", 20), col("Заказывающее\nподразделение", 20),
            col("Место расположения", 22), col("Прежние наименования", 30),
            col("Работ в журнале", 10), col("Кем внесено", 13), col("Изменено", 13),
            col("Примечание", 24)]
    sheet_head(ws, "Справочник объектов контроля",
               "Собран из журнала 2026 года: %d объектов, номера вынуты из наименований. "
               "Наименование в протоколе может отличаться от канонического — оно правится "
               "прямо в журнале, связь с объектом при этом не рвётся." % len(objects), len(cols))
    table_head(ws, HDR, cols)
    for i, o in enumerate(objects.values()):
        r = FIRST + i
        ws.cell(r, 1, o["id"])
        ws.cell(r, 2, o["name"])
        ws.cell(r, 3, o["zav"] or None)
        ws.cell(r, 4, o["reg"] or None)
        ws.cell(r, 5, o["uch"] or None)
        ws.cell(r, 6, o["inv"] or None)
        ws.cell(r, 8, o["cust"] or None)
        ws.cell(r, 12, "миграция")
    for r in range(FIRST, MAXR + 1):
        ws.cell(r, 11, '=IF($A{r}="","",COUNTIF(Журнал!$G:$G,$A{r}))'.format(r=r)).alignment = A_CTR
    mark_input(ws, FIRST, MAXR, 2, 10)
    mark_input(ws, FIRST, MAXR, 14, 14)
    mark_calc(ws, FIRST, MAXR, [1, 11])
    add_list_dv(ws, "G", list_ref("Тип объекта", 20))
    finish(ws, len(cols))
    ws.sheet_properties.tabColor = "2E5496"
    return ws


# ══════════════════════════════════════════════════════════════════════
# Специалисты и аттестации
# ══════════════════════════════════════════════════════════════════════
STAFF = [
    ("Денисов Сергей Валерьевич", "С.В. Денисов", "Начальник ЛНК"),
    ("Ивакин Алексей Васильевич", "А.В. Ивакин", "Заместитель начальника ЛНК"),
    ("Сиволап Павел Сергеевич", "П.С. Сиволап", "Специалист по неразрушающему контролю"),
    ("Матяш Андрей Владимирович", "А.В. Матяш", "Специалист по неразрушающему контролю"),
    ("Бирюков Д.С.", "Д.С. Бирюков", "Специалист по неразрушающему контролю"),
]


def build_staff(wb):
    ws = wb.create_sheet("Специалисты")
    cols = [col("ID", 11), col("ФИО полностью", 34), col("ФИО в документах", 18),
            col("Должность", 34), col("Основание привлечения", 20),
            col("Функции в ЛНК", 30), col("Образование, повышение\nквалификации", 36),
            col("Стаж по НК с", 13), col("Телефон", 15),
            col("Кем внесено", 13), col("Изменено", 13), col("Примечание", 22)]
    sheet_head(ws, "Специалисты ЛНК",
               "Одна строка — один человек. Удостоверения по методам — на листе «Аттестации». "
               "Колонка «ФИО в документах» подставляется в журнал и печатные формы: пишите её "
               "ровно так, как должно быть в документе. Колонки 5-8 нужны разделу 2 Паспорта.", len(cols))
    table_head(ws, HDR, cols)
    for i, (full, short, post) in enumerate(STAFF):
        r = FIRST + i
        ws.cell(r, 1, "СП-%04d" % (i + 1))
        ws.cell(r, 2, full)
        ws.cell(r, 3, short)
        ws.cell(r, 4, post)
        ws.cell(r, 5, "Трудовой договор")
        ws.cell(r, 10, "миграция")
        ws.cell(r, 8).number_format = DATE_FMT
    for r in range(FIRST, MAXR_REG + 1):
        ws.cell(r, 8).number_format = DATE_FMT
    mark_input(ws, FIRST, MAXR_REG, 2, 9)
    mark_input(ws, FIRST, MAXR_REG, 12, 12)
    mark_calc(ws, FIRST, MAXR_REG, [1])
    add_list_dv(ws, "E", list_ref("Основание привлечения", 10))
    finish(ws, len(cols), rowh=26)
    ws.sheet_properties.tabColor = "2E5496"
    return ws


def build_attest(wb):
    ws = wb.create_sheet("Аттестации")
    cols = [col("ID", 11), col("ID специалиста", 13), col("ФИО", 26),
            col("Метод\nконтроля", 10), col("Уровень", 9),
            col("№ квалификационного\nудостоверения", 22), col("Дата аттестации", 13),
            col("Орган, выдавший\nудостоверение", 30), col("Действительно до", 14),
            col("Дней до\nистечения", 10), col("Статус", 20),
            col("Объекты контроля\n(№ по Паспорту ЛНК)", 24),
            col("Кем внесено", 13), col("Изменено", 13), col("Примечание", 22)]
    sheet_head(ws, "Аттестации специалистов (СДАНК-02-2020)",
               "Одна строка — одно удостоверение по одному методу. У специалиста столько строк, "
               "сколько у него методов. ФИО подставляется по ID специалиста. "
               "Перед формированием документа макрос проверяет срок по этому листу.", len(cols))
    table_head(ws, HDR, cols)
    for r in range(FIRST, MAXR_REG + 1):
        ws.cell(r, 3, '=IF($B{r}="","",IFERROR(INDEX(Специалисты!$C:$C,MATCH($B{r},Специалисты!$A:$A,0)),"нет такого ID"))'.format(r=r))
        ws.cell(r, 10, days_left("$I%d" % r)).alignment = A_CTR
        ws.cell(r, 11, status_formula("$J%d" % r))
        ws.cell(r, 7).number_format = DATE_FMT
        ws.cell(r, 9).number_format = DATE_FMT
    mark_input(ws, FIRST, MAXR_REG, 2, 2)
    mark_input(ws, FIRST, MAXR_REG, 4, 9)
    mark_input(ws, FIRST, MAXR_REG, 12, 12)
    mark_input(ws, FIRST, MAXR_REG, 15, 15)
    mark_calc(ws, FIRST, MAXR_REG, [1, 3, 10, 11])
    add_list_dv(ws, "B", "=Специалисты!$A$%d:$A$200" % FIRST)
    add_list_dv(ws, "D", list_ref("Методы контроля", 12))
    add_list_dv(ws, "E", list_ref("Уровень квалификации", 5))
    add_status_cf(ws, "K", FIRST, MAXR_REG)
    finish(ws, len(cols))
    ws.sheet_properties.tabColor = "2E5496"
    return ws


# ══════════════════════════════════════════════════════════════════════
# Средства измерений и контроля + обслуживание
# ══════════════════════════════════════════════════════════════════════
def read_ppr():
    """Читает пятилетний план-график ППР: приборы, виды ТО, помесячную раскладку."""
    wb = openpyxl.load_workbook(SRC_PPR, data_only=True)
    ws = wb["Пятилетний"]
    rows = []
    for r in range(4, 48):
        raw = ws.cell(r, 2).value
        if not raw:
            continue
        name, zav, reg, uch, inv = parse_object(raw)
        per = norm(ws.cell(r, 7).value)
        m = re.search(r"(ТО-\d|ТР|СР)", per)
        kind = m.group(1) if m else ""
        months = {}
        for yi, year in enumerate(PPR_YEARS):
            base = 13 + yi * 13
            months[year] = [ws.cell(r, base + k).value for k in range(12)]
        norm_h = ws.cell(r, 10).value
        rows.append({
            "row": r, "card": norm(ws.cell(r, 3).value), "name": name, "zav": zav,
            "unit": norm(ws.cell(r, 5).value), "qty": ws.cell(r, 6).value,
            "kind": kind, "per_raw": per,
            "norm": float(norm_h) if isinstance(norm_h, str) else norm_h,
            "norm_was_text": isinstance(norm_h, str),
            "months": months,
        })
    return rows


def build_si(wb, ppr):
    ws = wb.create_sheet("СИ")
    cols = [col("ID", 11), col("№ карточки\nучёта", 12), col("Наименование", 40),
            col("Тип / обозначение", 20), col("Заводской №", 18), col("Инвентарный №", 14),
            col("Год выпуска", 10), col("Дата ввода в\nэксплуатацию", 14),
            col("Изготовитель", 22), col("Владелец\nоборудования", 20),
            col("Основание владения", 20), col("Раздел Паспорта", 18),
            col("Метод\nконтроля", 10), col("Назначение", 34),
            col("Технические и метрологические\nхарактеристики", 44),
            col("Ед.\nизм.", 7), col("Кол-во", 8), col("Ответственный за\nэксплуатацию", 20),
            col("Вид метрологического\nконтроля", 18), col("Периодичность\nповерки, мес.", 12),
            col("Дата последней\nповерки", 14), col("№ свидетельства", 18),
            col("Организация,\nпроводившая поверку", 26), col("Действительно до", 14),
            col("Дней до\nистечения", 10), col("Статус поверки", 20),
            col("Место хранения /\nнахождения", 22), col("Группа\nобслуживания", 13),
            col("Состояние", 16), col("Кем внесено", 13), col("Изменено", 13),
            col("Примечание", 22)]
    seeded = len([p for p in ppr if p["card"]])
    sheet_head(ws, "Реестр средств измерений и контроля ЛНК",
               "Заведено %d позиций из пятилетнего план-графика ППР (605-2025.ПЛ): наименование, "
               "заводской номер, № карточки учёта, количество. Остальное — поверка, характеристики, "
               "владелец — заполняется на месте. № карточки учёта связывает прибор с планом ППР, "
               "поэтому его не меняйте без нужды." % seeded, len(cols))
    table_head(ws, HDR, cols)

    seen, i = {}, 0
    for p in ppr:
        key = (p["card"], p["name"])
        if key in seen:
            continue
        seen[key] = True
        r = FIRST + i
        i += 1
        ws.cell(r, 1, "СИ-%04d" % i)
        ws.cell(r, 2, p["card"] or None)
        ws.cell(r, 3, p["name"])
        ws.cell(r, 5, p["zav"] or None)
        ws.cell(r, 16, p["unit"] or None)
        ws.cell(r, 17, p["qty"])
        ws.cell(r, 29, "В работе")
        ws.cell(r, 30, "миграция")
    for r in range(FIRST, MAXR_REG + 1):
        ws.cell(r, 25, days_left("$X%d" % r)).alignment = A_CTR
        ws.cell(r, 26, status_formula("$Y%d" % r))
        for c in (8, 21, 24):
            ws.cell(r, c).number_format = DATE_FMT
    mark_input(ws, FIRST, MAXR_REG, 2, 24)
    mark_input(ws, FIRST, MAXR_REG, 27, 29)
    mark_input(ws, FIRST, MAXR_REG, 32, 32)
    mark_calc(ws, FIRST, MAXR_REG, [1, 25, 26])
    add_list_dv(ws, "L", list_ref("Раздел Паспорта", 8))
    add_list_dv(ws, "M", list_ref("Методы контроля", 12))
    add_list_dv(ws, "S", list_ref("Вид метрологического контроля", 8))
    add_list_dv(ws, "R", "=Специалисты!$C$%d:$C$200" % FIRST)
    add_list_dv(ws, "AC", '"В работе,На поверке,В ремонте,Выдан,Списан"')
    add_status_cf(ws, "Z", FIRST, MAXR_REG)
    finish(ws, len(cols), rowh=26)
    ws.sheet_properties.tabColor = "2E5496"
    return ws, i


def build_si_to(wb, ppr):
    ws = wb.create_sheet("СИ_Обслуживание")
    cols = [col("ID", 11), col("№ карточки\nучёта", 12), col("ID средства НК", 13),
            col("Наименование", 40), col("Категория", 11),
            col("Периодичность", 26), col("Интервал,\nдней", 11),
            col("Кол-во", 8), col("Норма трудозатрат\nна единицу, ч", 15),
            col("Трудозатраты за\nодну операцию, ч", 15),
            col("Кем внесено", 13), col("Изменено", 13), col("Примечание", 26)]
    sheet_head(ws, "Виды обслуживания средств НК",
               "Перенесено из пятилетнего план-графика ППР. У прибора столько строк, сколько "
               "у него категорий обслуживания: «Виброанализатор СД-23» идёт с ТО-2 и с ТО-4, "
               "весы — с ТО-1 и ТО-4. Расшифровка категорий подтверждена: ТО-1 ежедневно, "
               "ТО-2 подекадно, ТО-3 ежемесячно, ТО-4 ежеквартально.", len(cols))
    table_head(ws, HDR, cols)
    for i, p in enumerate(ppr):
        r = FIRST + i
        per_name, days = TO_KINDS.get(p["kind"], ("", None))
        ws.cell(r, 1, "ТО-строка-%04d" % (i + 1))
        ws.cell(r, 2, p["card"] or None)
        ws.cell(r, 5, p["kind"] or None)
        ws.cell(r, 6, per_name or None)
        ws.cell(r, 7, days)
        ws.cell(r, 8, p["qty"])
        ws.cell(r, 9, p["norm"])
        if p["norm_was_text"]:
            ws.cell(r, 13, "В исходном план-графике норма была записана текстом — исправлено на число")
        ws.cell(r, 11, "миграция")
    for r in range(FIRST, MAXR_REG + 1):
        ws.cell(r, 3, '=IF($B{r}="","",IFERROR(INDEX(СИ!$A:$A,MATCH($B{r},СИ!$B:$B,0)),"нет карточки"))'.format(r=r))
        ws.cell(r, 4, '=IF($C{r}="","",IFERROR(INDEX(СИ!$C:$C,MATCH($C{r},СИ!$A:$A,0)),""))'.format(r=r))
        ws.cell(r, 10, '=IF(OR($H{r}="",$I{r}=""),"",$H{r}*$I{r})'.format(r=r)).alignment = A_CTR
        for c in (5, 7, 8, 9):
            ws.cell(r, c).alignment = A_CTR
    mark_input(ws, FIRST, MAXR_REG, 2, 2)
    mark_input(ws, FIRST, MAXR_REG, 5, 9)
    mark_input(ws, FIRST, MAXR_REG, 13, 13)
    mark_calc(ws, FIRST, MAXR_REG, [1, 3, 4, 10])
    add_list_dv(ws, "E", list_ref("Категория ТО", 8))
    finish(ws, len(cols))
    ws.sheet_properties.tabColor = "2E5496"
    return ws


def build_ppr_plan(wb, ppr):
    ws = wb.create_sheet("ППР_План")
    cols = [col("№ п/п", 7), col("№ карточки\nучёта", 12), col("Наименование", 38),
            col("Категория", 11), col("Норма, ч", 9)]
    for y in PPR_YEARS:
        cols += [col("%d\n%s" % (y, m[:3]), 6) for m in MONTHS] + [col("%d\nВсего, ч" % y, 9)]
    sheet_head(ws, "Пятилетний план-график ППР на 2026-2030 гг.",
               "Перенесено из вашего план-графика 605-2025.ПЛ как есть, без пересчёта: документ "
               "подписан, и расчёт заново дал бы другие числа. В клетках — трудозатраты за месяц, "
               "в часах. Пересчитать по периодичности можно кнопкой — при составлении следующего "
               "пятилетнего графика.", len(cols))
    table_head(ws, HDR, cols)
    for i, p in enumerate(ppr):
        r = FIRST + i
        ws.cell(r, 1, i + 1).alignment = A_CTR
        ws.cell(r, 2, p["card"] or None).alignment = A_CTR
        ws.cell(r, 3, p["name"])
        ws.cell(r, 4, p["kind"] or None).alignment = A_CTR
        ws.cell(r, 5, p["norm"]).alignment = A_CTR
        for yi, y in enumerate(PPR_YEARS):
            base = 6 + yi * 13
            for k in range(12):
                v = p["months"][y][k]
                if v not in (None, ""):
                    c = ws.cell(r, base + k, v)
                    c.alignment, c.number_format = A_CTR, "0.00"
            tot = ws.cell(r, base + 12,
                          "=IF(COUNT({a}{r}:{b}{r})=0,\"\",SUM({a}{r}:{b}{r}))".format(
                              a=get_column_letter(base), b=get_column_letter(base + 11), r=r))
            tot.alignment, tot.number_format = A_CTR, "0.00"
    last = FIRST + len(ppr) - 1
    r = last + 1
    ws.cell(r, 3, "ИТОГО").font = Font(bold=True)
    for yi in range(len(PPR_YEARS)):
        base = 6 + yi * 13
        for k in range(13):
            L = get_column_letter(base + k)
            c = ws.cell(r, base + k, "=SUM({L}{f}:{L}{l})".format(L=L, f=FIRST, l=last))
            c.font, c.alignment, c.number_format = Font(bold=True), A_CTR, "0.00"
    for c in range(1, len(cols) + 1):
        ws.cell(r, c).fill = PatternFill("solid", fgColor="D9E2F3")
        ws.cell(r, c).border = BORDER
    mark_input(ws, FIRST, last, 5, len(cols))
    mark_calc(ws, FIRST, last, [1, 2, 3, 4])
    finish(ws, len(cols), FIRST, last)
    ws.sheet_properties.tabColor = "2E5496"
    return ws


# ══════════════════════════════════════════════════════════════════════
# Пустые реестры
# ══════════════════════════════════════════════════════════════════════
def simple_sheet(wb, name, title, hint, cols, inputs, calcs=(), dates=(),
                 dv=(), formulas=(), tab="404040", rowh=24, ctr=(), maxr=None):
    maxr = maxr or MAXR_REG
    ws = wb.create_sheet(name)
    sheet_head(ws, title, hint, len(cols))
    table_head(ws, HDR, cols)
    for r in range(FIRST, maxr + 1):
        for c, f in formulas:
            ws.cell(r, c, f.format(r=r))
        for c in dates:
            ws.cell(r, c).number_format = DATE_FMT
        for c in ctr:
            ws.cell(r, c).alignment = A_CTR
    for a, b in inputs:
        mark_input(ws, FIRST, maxr, a, b)
    if calcs:
        mark_calc(ws, FIRST, maxr, list(calcs))
    for letter, source in dv:
        add_list_dv(ws, letter, source, FIRST, maxr)
    finish(ws, len(cols), r2=maxr, rowh=rowh)
    ws.sheet_properties.tabColor = tab
    return ws


def build_rest(wb):
    SP = "=Специалисты!$C$%d:$C$200" % FIRST
    SI = "=СИ!$A$%d:$A$%d" % (FIRST, MAXR_REG)

    simple_sheet(wb, "НД", "Реестр нормативной и методической документации",
        "Периодичность проверки на актуальность по умолчанию 12 мес. — п.5 Правил учёта "
        "документации. Дата следующей проверки и статус считаются формулами.",
        [col("ID", 11), col("Наименование документа", 50), col("Обозначение", 22),
         col("Издательство / разработчик", 26), col("Год издания", 10),
         col("Дата последней проверки\nна актуальность", 16),
         col("Периодичность\nпроверки, мес.", 12), col("Дата следующей\nпроверки", 14),
         col("Дней до\nистечения", 10), col("Статус", 20),
         col("Место хранения\n(контрольный экземпляр)", 24),
         col("Кем внесено", 13), col("Изменено", 13), col("Примечание", 22)],
        inputs=[(2, 7), (11, 11), (14, 14)], calcs=[1, 8, 9, 10], dates=[6, 8],
        formulas=[(8, '=IF(OR($F{r}="",$G{r}=""),"",EDATE($F{r},$G{r}))'),
                  (9, days_left("$H{r}")), (10, status_formula("$I{r}"))],
        ctr=[5, 7, 9], tab="2E5496")
    add_status_cf(wb["НД"], "J", FIRST, MAXR)

    simple_sheet(wb, "Кто_где", "Кто где: отпуска, больничные, выезды",
        "«Выезд на измерения» — не отсутствие: человек работает, документы по нему оформлять "
        "можно, но для планирования он занят. Поэтому вид и признак «доступен» разведены. "
        "Всё нетиповое — вид «Прочее» с пояснением в основании.",
        [col("ID", 11), col("ID специалиста", 13), col("ФИО", 26), col("Вид", 22),
         col("Дата с", 12), col("Дата по", 12), col("Дней", 8),
         col("Доступен\nдля работ", 11), col("Основание", 30),
         col("Кем внесено", 13), col("Изменено", 13), col("Примечание", 24)],
        inputs=[(2, 2), (4, 6), (8, 9), (12, 12)], calcs=[1, 3, 7], dates=[5, 6],
        formulas=[(3, '=IF($B{r}="","",IFERROR(INDEX(Специалисты!$C:$C,MATCH($B{r},Специалисты!$A:$A,0)),"нет такого ID"))'),
                  (7, '=IF(OR($E{r}="",$F{r}=""),"",$F{r}-$E{r}+1)')],
        dv=[("B", "=Специалисты!$A$%d:$A$200" % FIRST),
            ("D", list_ref("Вид отсутствия", 12)),
            ("H", list_ref("Доступен для работ", 4))],
        ctr=[5, 6, 7, 8], tab="7030A0")

    simple_sheet(wb, "График_отпусков", "График отпусков",
        "Плановый, на год. По кнопке из него создаются записи в «Кто где», чтобы не вносить "
        "одно и то же дважды.",
        [col("ID", 11), col("ID специалиста", 13), col("ФИО", 26), col("Год", 8),
         col("Период 1: с", 12), col("Период 1: по", 12),
         col("Период 2: с", 12), col("Период 2: по", 12),
         col("Период 3: с", 12), col("Период 3: по", 12),
         col("Всего дней", 11), col("Основание (№ приказа)", 24),
         col("Перенесено\nв «Кто где»", 12), col("Примечание", 22)],
        inputs=[(2, 2), (4, 10), (12, 12), (14, 14)], calcs=[1, 3, 11],
        dates=[5, 6, 7, 8, 9, 10],
        formulas=[(3, '=IF($B{r}="","",IFERROR(INDEX(Специалисты!$C:$C,MATCH($B{r},Специалисты!$A:$A,0)),"нет такого ID"))'),
                  (11, '=IF($B{r}="","",IF($E{r}="",0,$F{r}-$E{r}+1)+IF($G{r}="",0,$H{r}-$G{r}+1)+IF($I{r}="",0,$J{r}-$I{r}+1))')],
        dv=[("B", "=Специалисты!$A$%d:$A$200" % FIRST),
            ("M", '"Да,Нет"')],
        ctr=[4, 5, 6, 7, 8, 9, 10, 11, 13], tab="7030A0")

    simple_sheet(wb, "Обслуживание_факт", "Обслуживание средств НК — выполненные работы",
        "Одна строка — одно выполненное мероприятие. Отсюда же печатается подтаблица карточки "
        "учёта прибора и считается дата следующего обслуживания.",
        [col("ID", 11), col("ID средства НК", 13), col("Наименование", 34),
         col("Дата проведения", 14), col("Вид работ", 22), col("Категория", 11),
         col("Содержание работ", 40), col("Исполнитель", 20), col("Результат", 20),
         col("Трудозатраты, ч", 12), col("Израсходовано\nспирта, г", 12),
         col("Основание (№ договора, акта)", 24),
         col("Кем внесено", 13), col("Изменено", 13), col("Примечание", 22)],
        inputs=[(2, 2), (4, 12), (15, 15)], calcs=[1, 3], dates=[4],
        formulas=[(3, '=IF($B{r}="","",IFERROR(INDEX(СИ!$C:$C,MATCH($B{r},СИ!$A:$A,0)),"нет такого ID"))')],
        dv=[("B", SI), ("E", list_ref("Вид работ по обслуживанию", 8)),
            ("F", list_ref("Категория ТО", 8)), ("H", SP),
            ("I", list_ref("Результат обслуживания", 6))],
        ctr=[4, 6, 10, 11], tab="2E5496")

    simple_sheet(wb, "Выдача_оборудования", "Журнал выдачи оборудования",
        "Строка закрывается при возврате. Пока дата возврата пуста, прибор числится выданным — "
        "это видно на Дашборде и мешает выдать его второй раз.",
        [col("ID", 11), col("ID средства НК", 13), col("Наименование", 34),
         col("Заводской №", 16), col("Дата выдачи", 13), col("Кому выдано", 20),
         col("Основание / цель", 26), col("Состояние при выдаче", 20),
         col("Дата возврата", 13), col("Состояние при возврате", 20),
         col("Дней у получателя", 12), col("Выдал", 18),
         col("Принял", 18), col("Примечание", 22)],
        inputs=[(2, 2), (5, 10), (12, 14)], calcs=[1, 3, 4, 11],
        dates=[5, 9],
        formulas=[(3, '=IF($B{r}="","",IFERROR(INDEX(СИ!$C:$C,MATCH($B{r},СИ!$A:$A,0)),"нет такого ID"))'),
                  (4, '=IF($B{r}="","",IFERROR(INDEX(СИ!$E:$E,MATCH($B{r},СИ!$A:$A,0)),""))'),
                  (11, '=IF($E{r}="","",IF($I{r}="",TODAY()-$E{r},$I{r}-$E{r}))')],
        dv=[("B", SI), ("F", SP), ("L", SP), ("M", SP),
            ("H", '"Исправен,Исправен, есть замечания,Неисправен"'),
            ("J", '"Исправен,Исправен, есть замечания,Неисправен,Не возвращён"')],
        ctr=[5, 9, 11], tab="2E5496")
    ws = wb["Выдача_оборудования"]
    ws.conditional_formatting.add("A%d:N%d" % (FIRST, MAXR), FormulaRule(
        formula=['AND($E%d<>"",$I%d="")' % (FIRST, FIRST)],
        fill=PatternFill("solid", fgColor="FFF2CC")))

    simple_sheet(wb, "Рекламации", "Журнал учёта рекламаций, претензий",
        "Графы 1-8 — по Приложению №1 к Правилам рассмотрения рекламаций (Ред. 1). "
        "Остальные колонки служебные: они связывают претензию с аннулированным заключением "
        "и с корректирующим действием, как того требует раздел 3 Правил.",
        [col("ID", 11), col("№ п/п", 7), col("Реквизиты документа", 30),
         col("Дата регистрации\nв ЛНК", 14),
         col("Подразделение, от кого\nпоступил документ", 26),
         col("Срок выполнения", 14),
         col("ФИО кому выдан,\nроспись о получении", 24),
         col("Отметка об выполненных\nмероприятиях", 34), col("Примечание", 22),
         col("Дней до срока", 12), col("Состояние", 18),
         col("ID записи журнала\n(аннулируемое заключение)", 22),
         col("Номер нового\nзаключения", 18), col("Решение", 26),
         col("Состав комиссии", 26), col("Ссылка на докладную", 26),
         col("Кем внесено", 13), col("Изменено", 13)],
        inputs=[(3, 9), (12, 16)], calcs=[1, 2, 10, 11], dates=[4, 6],
        formulas=[(2, '=IF($D{r}="","",COUNT($D$%d:$D{r}))' % FIRST),
                  (10, '=IF($F{r}="","",IF($H{r}<>"","",$F{r}-TODAY()))'),
                  (11, '=IF($D{r}="","",IF($H{r}<>"","Закрыта",IF($F{r}="","Срок не задан",'
                       'IF($J{r}<0,"ПРОСРОЧЕНА",IF($J{r}<=7,"Истекает ≤ 7 дн.","В работе")))))')],
        dv=[("G", SP), ("N", list_ref("Решение по рекламации", 8))],
        ctr=[2, 4, 6, 10], tab="C00000", rowh=28)
    ws = wb["Рекламации"]
    for text, fill, font in (("ПРОСРОЧЕНА", C_BAD, C_BAD_T), ("Истекает", C_WARN, C_WARN_T),
                             ("В работе", C_SOON, C_SOON_T), ("Закрыта", C_OK, C_OK_T)):
        ws.conditional_formatting.add("K%d:K%d" % (FIRST, MAXR), FormulaRule(
            formula=['ISNUMBER(SEARCH("%s",K%d))' % (text, FIRST)],
            fill=PatternFill("solid", fgColor=fill),
            font=Font(color=font, bold=(text == "ПРОСРОЧЕНА"))))

    simple_sheet(wb, "План_график", "План-график работ по неразрушающему контролю",
        "Наименование объекта здесь может отличаться от того, что попадёт в протокол: "
        "оно уточняется в процессе контроля. Связь держится по ID объекта.",
        [col("ID", 11), col("Год", 8), col("Месяц", 12), col("ID объекта", 12),
         col("Наименование объекта\nв плане", 40), col("Заказывающее\nподразделение", 22),
         col("Методы контроля", 22), col("Плановая дата", 13),
         col("Ответственный", 20), col("Статус", 18), col("Фактическая дата", 14),
         col("ID записи журнала", 18), col("Кем внесено", 13), col("Изменено", 13),
         col("Примечание", 22)],
        inputs=[(2, 12), (15, 15)], calcs=[1], dates=[8, 11],
        dv=[("D", "=Объекты!$A$%d:$A$%d" % (FIRST, MAXR)), ("I", SP),
            ("J", list_ref("Статус задачи", 6))],
        ctr=[2, 3, 8, 11], tab="7030A0")

    simple_sheet(wb, "Задачи", "Задачи и напоминания",
        "Всё, что не попадает в остальные реестры: подготовка к проверке, заказ расходных "
        "материалов, обучение. На Дашборде показываются просроченные и ближайшие.",
        [col("ID", 11), col("Категория", 22), col("Задача", 46), col("Ответственный", 20),
         col("Срок", 13), col("Приоритет", 12), col("Статус", 18), col("Дней до срока", 12),
         col("Выполнено", 13), col("Кем внесено", 13), col("Изменено", 13),
         col("Примечание", 22)],
        inputs=[(2, 7), (9, 9), (12, 12)], calcs=[1, 8], dates=[5, 9],
        formulas=[(8, '=IF($E{r}="","",IF($G{r}="Выполнено","",$E{r}-TODAY()))')],
        dv=[("D", SP), ("F", list_ref("Приоритет", 5)), ("G", list_ref("Статус задачи", 6))],
        ctr=[5, 6, 8, 9], tab="7030A0")

    simple_sheet(wb, "Электроинструмент",
        "Журнал учёта, проверки и испытаний электроинструмента и вспомогательного оборудования",
        "Форма по Приложению 4 к РД 34.03.204 — 16 граф, они идут первыми и в печать выводятся "
        "как есть. Дальше служебные колонки системы: связь с реестром СИ, срок и статус. "
        "Периодичность по умолчанию 6 мес.; дата следующей проверки и статус считаются формулами.",
        [col("ID", 11),
         col("1\n№ п/п", 7),
         col("2\nНаименование инструмента,\nвспомогательного оборудования", 40),
         col("3\nИнвентарный\n№", 13),
         col("4\nДата последней\nпроверки, испытания", 15),
         col("5\nПричина проверки,\nиспытания", 20),
         col("6\nИспытание изоляции\nповышенным напряжением:\nдата", 15),
         col("7\nИспытание изоляции\nповышенным напряжением:\nрезультат", 20),
         col("8\nИзмерение сопротивления\nизоляции: дата", 15),
         col("9\nИзмерение сопротивления\nизоляции: результат, МОм", 17),
         col("10\nПроверка цепи\nзаземления: дата", 15),
         col("11\nПроверка цепи\nзаземления: результат", 20),
         col("12\nВнешний осмотр и работа\nна холостом ходу: дата", 15),
         col("13\nВнешний осмотр и работа\nна холостом ходу: результат", 22),
         col("14\nЗаключение\nо пригодности", 22),
         col("15\nДата следующей\nпроверки, испытания", 15),
         col("16\nФамилия, инициалы и подпись\nпроводившего проверку", 28),
         col("ID средства НК", 13),
         col("Класс\nзащиты", 9),
         col("Периодичность,\nмес.", 12),
         col("Дней до\nистечения", 10),
         col("Статус", 20),
         col("Кем внесено", 13), col("Изменено", 13), col("Примечание", 22)],
        inputs=[(3, 17), (18, 20), (25, 25)], calcs=[1, 2, 16, 21, 22],
        dates=[5, 7, 9, 11, 13, 16],
        formulas=[(2, '=IF($C{r}="","",COUNTA($C$%d:$C{r}))' % FIRST),
                  (16, '=IF(OR($E{r}="",$T{r}=""),"",EDATE($E{r},$T{r}))'),
                  (21, days_left("$P{r}")),
                  (22, status_formula("$U{r}", warn=30, soon=14))],
        dv=[("R", "=СИ!$A$%d:$A$%d" % (FIRST, MAXR_REG)),
            ("S", '"0,I,II,III"'),
            ("F", '"Периодическая,После ремонта,Перед выдачей,Внеплановая"'),
            ("O", '"Пригоден к эксплуатации,Не пригоден, изъят,Направлен в ремонт"')],
        ctr=[2, 4, 5, 7, 9, 10, 11, 13, 16, 19, 20, 21], tab="C55A11", rowh=26)
    add_status_cf(wb["Электроинструмент"], "V", FIRST, MAXR_REG)

    simple_sheet(wb, "Журнал_копирования", "Журнал резервного копирования",
        "Заполняется макросом. Вручную вносится только отметка о ежегодной проверке "
        "восстановления. Без этих записей утверждение «мы делаем копии» ничем не подтверждается.",
        [col("Дата и время", 18), col("Что копировалось", 30), col("Куда", 44),
         col("Размер, МБ", 11), col("Результат", 18), col("Кто выполнял", 18),
         col("Примечание", 34)],
        inputs=[(7, 7)], dates=[], ctr=[4, 5], tab="808080")


# ══════════════════════════════════════════════════════════════════════
# Паспорт — постоянная часть
# ══════════════════════════════════════════════════════════════════════
PASSPORT = [
    ("Раздел 1. Общие данные о лаборатории", None, None),
    ("Полное наименование ЛНК", "", "Как в свидетельстве об аттестации."),
    ("Сокращённое наименование", "", ""),
    ("Почтовый адрес", "", "Индекс, адрес одной строкой."),
    ("Телефон, электронная почта", "", ""),
    ("Наименование организации, в составе которой ЛНК", "", ""),
    ("№ свидетельства об аттестации ЛНК", "", "Идёт и в Паспорт, и на плитку Дашборда."),
    ("Дата выдачи свидетельства", "", "Формат ДД.ММ.ГГГГ."),
    ("Свидетельство действительно до", "", "Дашборд предупреждает за 180 и за 90 дней."),
    ("Орган по аттестации", "", ""),
    ("Начальник ЛНК: должность, ФИО, телефон", "", ""),
    ("Ответственный за функционирование СМ: ФИО", "", ""),
    ("Ответственный за метрологическое обеспечение: ФИО", "", ""),
    ("Ответственный за фонд НД: ФИО", "", ""),
    ("Раздел 3. Объекты неразрушающего контроля", None, None),
    ("Область аттестации: объекты контроля", "", "Перечень по свидетельству, по одному в строке ниже — "
                                                  "строки можно добавлять, макрос читает до первой пустой."),
    ("Раздел 4. Виды (методы) неразрушающего контроля", None, None),
    ("Область аттестации: методы", "", "Методы по свидетельству: ВИК, УЗД, УЗТ, АЭ, ТВ, ТК, ВД, МК."),
    ("Раздел 10. Помещения", None, None),
    ("Помещения ЛНК: номер, площадь, назначение", "", "По одному помещению в строке."),
    ("Условия: температура, влажность, освещённость", "", ""),
    ("Раздел 11. Передвижные лаборатории", None, None),
    ("Передвижные лаборатории", "", "Если нет — так и напишите «отсутствуют»."),
    ("Служебное", None, None),
    ("Номер редакции Паспорта", "", "Проставляется при выпуске."),
    ("Дата редакции", "", ""),
    ("Кем утверждён", "", ""),
]


def build_passport(wb):
    ws = wb.create_sheet("Паспорт_постоянная")
    sheet_head(ws, "Паспорт ЛНК — постоянная часть",
               "То, что меняется раз в несколько лет. Разделы 2, 5-9 в Паспорт не вносятся "
               "руками — они собираются из реестров СИ, Специалисты, Аттестации и НД. "
               "Обязательные поля подсвечены: пока они пустые, Паспорт не соберётся.", 4)
    table_head(ws, HDR, [col("Поле", 50), col("Значение", 52), col("Подсказка", 60),
                         col("Обязательное", 13)])
    required = {"Полное наименование ЛНК", "Почтовый адрес",
                "№ свидетельства об аттестации ЛНК", "Свидетельство действительно до",
                "Начальник ЛНК: должность, ФИО, телефон",
                "Область аттестации: объекты контроля", "Область аттестации: методы"}
    r = FIRST
    for name, value, hint in PASSPORT:
        if value is None and hint is None:
            c = ws.cell(r, 1, name)
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            for cc in range(1, 5):
                ws.cell(r, cc).fill = PatternFill("solid", fgColor=C_HEAD2)
        else:
            ws.cell(r, 1, name).font = F_TXT
            ws.cell(r, 2, value).font = F_INPUT
            h = ws.cell(r, 3, hint)
            h.font, h.alignment = F_HINT, A_TXT
            if name in required:
                ws.cell(r, 4, "да").alignment = A_CTR
                ws.cell(r, 2).fill = PatternFill("solid", fgColor="FFF2CC")
        for cc in range(1, 5):
            ws.cell(r, cc).border = BORDER
        ws.cell(r, 1).alignment = A_TXT
        ws.cell(r, 2).alignment = A_TXT
        ws.row_dimensions[r].height = 22
        r += 1
    # запас пустых строк — перечни объектов и помещений идут списком
    for k in range(24):
        for cc in range(1, 5):
            ws.cell(r + k, cc).border = BORDER
        ws.cell(r + k, 2).font = F_INPUT
    ws.sheet_properties.tabColor = "1F3864"
    return ws


# ══════════════════════════════════════════════════════════════════════
# Сверка миграции
# ══════════════════════════════════════════════════════════════════════
def collect_issues(flat, wide, objects, ppr, warnings):
    """Всё, что при переносе требует человеческого решения."""
    out = []
    for w in warnings:
        out.append(("Журнал", "Одинаковые объекты в один день", w,
                    "Проверьте по бумажному журналу, что это действительно две работы"))
    for i, w in enumerate(wide):
        r = FIRST + i
        for m, f in w["methods"].items():
            if f["orig"] in (None, ""):
                out.append(("Журнал", "Нет номера в серии",
                            "строка %d, метод %s, объект «%s», %s" % (
                                r, m, w["obj"][0][:40],
                                w["date"].strftime("%d.%m.%Y") if w["date"] else "без даты"),
                            "В старом журнале номер был пустым — впишите номер с бумаги"))
            if not f["exe"]:
                out.append(("Журнал", "Нет исполнителя",
                            "строка %d, метод %s, объект «%s»" % (r, m, w["obj"][0][:40]),
                            "Выберите исполнителя из справочника специалистов"))
        if not w["date"]:
            out.append(("Журнал", "Нет даты", "строка %d" % r, "Проставьте дату с бумаги"))
    noid = [o for o in objects.values()
            if not (o["zav"] or o["reg"] or o["uch"] or o["inv"])]
    out.append(("Объекты", "Объект без единого номера",
                "%d из %d объектов" % (len(noid), len(objects)),
                "Здания и агрегаты без заводского номера — это нормально; "
                "для сосудов номер лучше проставить, по нему собирается история"))
    # Одноимённые объекты, у которых ни одного номера: различить их нечем.
    # Те, что отличаются рег. или уч. номером, — разные объекты, и это нормально.
    dup = collections.Counter(o["name"].lower() for o in noid)
    for nm, k in dup.items():
        if k > 1:
            out.append(("Объекты", "Одноимённые объекты, различить нечем",
                        "«%s» — %d записей, ни у одной нет номера" % (nm[:50], k),
                        "Проверьте, один это объект или разные; если разные — проставьте номера"))
    for p in ppr:
        if p["norm_was_text"]:
            out.append(("ППР", "Норма трудозатрат была текстом",
                        "строка %d исходного графика, «%s»" % (p["row"], p["name"][:40]),
                        "Перенесено числом. В исходном документе стоит поправить: "
                        "текст не попадает в суммы"))
        if p["qty"] in (None, ""):
            out.append(("ППР", "Не заполнено количество",
                        "строка %d исходного графика, «%s»" % (p["row"], p["name"][:40]),
                        "Проставьте количество в реестре СИ"))
        cnts = {y: sum(1 for v in p["months"][y] if v not in (None, "")) for y in PPR_YEARS}
        if len(set(cnts.values())) > 1:
            out.append(("ППР", "Число отметок по годам расходится",
                        "строка %d, «%s»: %s" % (p["row"], p["name"][:40],
                                                 ", ".join("%d — %d" % (y, cnts[y]) for y in PPR_YEARS)),
                        "Периодичность одна, а отметок разное число — похоже на случайную протяжку"))
    return out


def build_migration_report(wb, issues):
    ws = wb.create_sheet("Сверка_миграции", 1)
    cols = [col("№", 6), col("Где", 12), col("Что", 34), col("Подробности", 70),
            col("Что с этим делать", 60), col("Разобрано", 12)]
    sheet_head(ws, "Сверка после миграции",
               "Всё, что при переносе данных требует вашего решения. Разобрали строку — "
               "поставьте «да» в последней колонке. Переход считается выполненным, когда "
               "здесь не осталось неразобранного. Найдено записей: %d." % len(issues), len(cols))
    table_head(ws, HDR, cols)
    for i, (where, what, detail, todo) in enumerate(issues):
        r = FIRST + i
        ws.cell(r, 1, i + 1).alignment = A_CTR
        ws.cell(r, 2, where).alignment = A_CTR
        ws.cell(r, 3, what)
        ws.cell(r, 4, detail)
        ws.cell(r, 5, todo).font = F_HINT
    last = max(FIRST + len(issues) - 1, FIRST + 20)
    add_list_dv(ws, "F", '"да,нет"', FIRST, last)
    mark_input(ws, FIRST, last, 6, 6)
    finish(ws, len(cols), FIRST, last, rowh=26)
    ws.conditional_formatting.add("A%d:F%d" % (FIRST, last), FormulaRule(
        formula=['$F%d="да"' % FIRST], font=Font(color="A6A6A6", strike=True)))
    ws.sheet_properties.tabColor = "C00000"
    return ws


# ══════════════════════════════════════════════════════════════════════
def main():
    src = openpyxl.load_workbook(SRC_V2, data_only=True)
    flat = read_flat(src["Журнал"])
    wide, warnings = to_wide(flat)
    objects = build_objects(wide)
    ppr = read_ppr()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_settings(wb)
    build_journal(wb, wide, objects)
    build_objects_sheet(wb, objects)
    build_staff(wb)
    build_attest(wb)
    si_ws, n_si = build_si(wb, ppr)
    build_si_to(wb, ppr)
    build_ppr_plan(wb, ppr)
    build_rest(wb)
    build_passport(wb)
    build_lists(wb)
    issues = collect_issues(flat, wide, objects, ppr, warnings)
    build_migration_report(wb, issues)

    wb.properties.title = "ЛНК — данные"
    wb.properties.creator = "ЛНК"
    wb.save(OUT)

    print("Сохранено: %s" % OUT)
    print("  плоских строк журнала : %d" % len(flat))
    print("  широких строк         : %d" % len(wide))
    print("  объектов контроля     : %d" % len(objects))
    print("  средств НК из ППР     : %d" % n_si)
    print("  строк обслуживания    : %d" % len(ppr))
    print("  листов в книге        : %d" % len(wb.sheetnames))
    print("  строк в сверке        : %d" % len(issues))
    for w in warnings:
        print("  ! %s" % w)


if __name__ == "__main__":
    main()
