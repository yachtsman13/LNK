# -*- coding: utf-8 -*-
"""Общие настройки оформления книги ЛНК."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── палитра ───────────────────────────────────────────────────────────
C_HEAD   = "1F3864"   # шапка таблиц
C_HEAD2  = "2E5496"   # шапка второго уровня
C_TITLE  = "1F3864"   # заголовок листа
C_HINT   = "595959"   # пояснения
C_INPUT  = "0033A0"   # цвет текста в полях ручного ввода
C_CALC   = "808080"   # цвет текста в расчётных полях
C_BAND   = "F2F5FA"   # заливка чётных строк
C_LOCK   = "EDEDED"   # заливка расчётных колонок
C_OK     = "C6EFCE"; C_OK_T   = "006100"
C_WARN   = "FFEB9C"; C_WARN_T = "9C6500"
C_SOON   = "FFD8A8"; C_SOON_T = "974706"
C_BAD    = "FFC7CE"; C_BAD_T  = "9C0006"

THIN = Side(style="thin", color="B4C6E7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

F_TITLE = Font(name="Calibri", size=14, bold=True, color=C_TITLE)
F_HINT  = Font(name="Calibri", size=9,  italic=True, color=C_HINT)
F_HEAD  = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_INPUT = Font(name="Calibri", size=10, color=C_INPUT)
F_CALC  = Font(name="Calibri", size=10, color=C_CALC)
F_TXT   = Font(name="Calibri", size=10)

FILL_HEAD = PatternFill("solid", fgColor=C_HEAD)
FILL_LOCK = PatternFill("solid", fgColor=C_LOCK)

A_HEAD = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_TXT  = Alignment(vertical="top", wrap_text=True)
A_CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)

DATE_FMT = "DD.MM.YYYY"

# ── формулы статусов сроков ──────────────────────────────────────────
def days_left(date_cell):
    """Дней до истечения: пусто, если дата не заполнена."""
    return '=IF({d}="","",{d}-TODAY())'.format(d=date_cell)

def status_formula(days_cell, warn=90, soon=30):
    return ('=IF({d}="","",IF({d}<0,"ПРОСРОЧЕНО",'
            'IF({d}<={s},"Истекает ≤ {s} дн.",'
            'IF({d}<={w},"Истекает ≤ {w} дн.","Действует"))))'
            ).format(d=days_cell, s=soon, w=warn)

STATUS_COLORS = [
    ("ПРОСРОЧЕНО", C_BAD,  C_BAD_T),
    ("Истекает",   C_WARN, C_WARN_T),
    ("Действует",  C_OK,   C_OK_T),
]

def sheet_head(ws, title, hint, ncols):
    """Заголовок листа + строка-пояснение. Таблица начинается со строки 3."""
    ws.cell(1, 1, title).font = F_TITLE
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 2))
    ws.cell(2, 1, hint).font = F_HINT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(ncols, 2))
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 26

def table_head(ws, row, cols):
    """cols: список словарей {t: заголовок, w: ширина}."""
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row, i, c["t"])
        cell.font, cell.fill, cell.alignment, cell.border = F_HEAD, FILL_HEAD, A_HEAD, BORDER
        ws.column_dimensions[get_column_letter(i)].width = c["w"]
    ws.row_dimensions[row].height = 42
    ws.freeze_panes = ws.cell(row + 1, 1)
    ws.auto_filter.ref = "A%d:%s%d" % (row, get_column_letter(len(cols)), row)
