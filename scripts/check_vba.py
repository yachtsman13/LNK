# -*- coding: utf-8 -*-
"""Статическая проверка модулей VBA: блоки, имена процедур, заголовки колонок.

Выполнить VBA здесь нельзя, поэтому проверяется то, что проверяется чтением:
парность блоков, существование вызываемых процедур и — главное — что каждый
заголовок колонки, который код ищет через ColReq, действительно есть в книге.
"""
import io, os, re, sys, glob, collections

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAC = os.path.join(ROOT, "Макросы")
DATA = os.path.join(ROOT, "Реестры", "ЛНК_Данные.xlsx")
HDR_ROW = 3

errors, oks = [], []


def check(cond, msg):
    (oks if cond else errors).append(("OK   " if cond else "ОШИБКА ") + msg)


def strip_code(line):
    """Убирает строковые литералы и комментарии — чтобы не ловить слова из текста."""
    out, in_str = [], False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_str = not in_str
            out.append(" ")
        elif in_str:
            out.append(" ")
        elif ch == "'":
            break
        else:
            out.append(ch)
        i += 1
    return "".join(out)


BLOCKS = [
    (re.compile(r"^\s*(?:Public |Private )?(?:Static )?Sub\b", re.I), re.compile(r"^\s*End Sub\b", re.I), "Sub"),
    (re.compile(r"^\s*(?:Public |Private )?(?:Static )?Function\b", re.I), re.compile(r"^\s*End Function\b", re.I), "Function"),
    (re.compile(r"^\s*(?:Public |Private )?Property\b", re.I), re.compile(r"^\s*End Property\b", re.I), "Property"),
    (re.compile(r"^\s*With\b", re.I), re.compile(r"^\s*End With\b", re.I), "With"),
    (re.compile(r"^\s*Select Case\b", re.I), re.compile(r"^\s*End Select\b", re.I), "Select"),
]
RE_FOR = re.compile(r"^\s*For\b", re.I)
RE_NEXT = re.compile(r"^\s*Next\b", re.I)
RE_DO = re.compile(r"^\s*Do\b", re.I)
RE_LOOP = re.compile(r"^\s*Loop\b", re.I)
# Многострочный If: Then в конце строки (или его нет вовсе)
RE_IF = re.compile(r"^\s*If\b.*\bThen\s*$", re.I)
RE_ENDIF = re.compile(r"^\s*End If\b", re.I)

RE_DECL = re.compile(r"^\s*(?:Public |Private )?(?:Static )?(Sub|Function|Property\s+\w+)\s+([^\s(]+)", re.I)
# Public/Private на уровне модуля объявляют переменную так же, как Dim
# внутри процедуры; Sub, Function и прочее разбирает RE_DECL.
RE_LOCAL = re.compile(
    r"^\s*(?:Dim|Const|ReDim|Static|Public|Private)\s+"
    r"(?:Const\s+|Static\s+)?"
    r"(?!(?:Sub|Function|Property|Type|Enum|Declare|Const|Static|WithEvents)\b)"
    r"([^\s,]+)", re.I)
RE_LABEL = re.compile(r"^\s*([А-Яа-яЁёA-Za-z_]\w*):\s*$")
# Set ws = Лист("Журнал") / ЛистТихо(Л_СИ) — по этому узнаём, какой лист
# стоит за переменной, и сверяем заголовок именно с ним.
RE_SETSHEET = re.compile(r'^\s*Set\s+(\w+)\s*=\s*Лист(?:Тихо)?\(\s*(?:"([^"]+)"|(Л_\w+))\s*\)')
RE_PARAMS = re.compile(r"\(([^)]*)\)")


def read_lines(path):
    """Склеивает продолжения строк (символ _ в конце)."""
    raw = io.open(path, encoding="utf-8").read().split("\n")
    out, buf, start = [], "", 0
    for i, line in enumerate(raw, 1):
        s = line.rstrip()
        if s.endswith(" _"):
            if not buf:
                start = i
            buf += s[:-1]
            continue
        if buf:
            out.append((start, buf + s))
            buf = ""
        else:
            out.append((i, s))
    if buf:
        out.append((start, buf))
    return out


def main():
    files = sorted(glob.glob(os.path.join(MAC, "*.bas")))
    check(bool(files), "найдены модули: %d" % len(files))

    # Константы имён листов из modConfig
    cfg = io.open(os.path.join(MAC, "modConfig.bas"), encoding="utf-8").read()
    SHEET_CONST = dict(re.findall(r'Public Const (Л_\w+) As String = "([^"]+)"', cfg))

    declared = {}          # имя процедуры → модуль
    known_locals = set()
    per_module = collections.defaultdict(set)
    col_uses = []          # (файл, строка, лист-переменная, заголовок)

    for path in files:
        name = os.path.basename(path)[:-4]
        lines = read_lines(path)
        text_first = lines[0][1] if lines else ""
        check(text_first.strip() == 'Attribute VB_Name = "%s"' % name,
              "%s: первая строка Attribute VB_Name совпадает с именем файла" % name)
        check(any(l.strip() == "Option Explicit" for _, l in lines[:5]),
              "%s: Option Explicit на месте" % name)

        stack = []
        var_sheet = {}
        текущая_процедура = ""
        for ln, raw in lines:
            code = strip_code(raw)
            m = RE_DECL.match(code)
            if m:
                if m.group(2) in per_module[name]:
                    errors.append("ОШИБКА %s:%d процедура %s объявлена дважды в одном модуле"
                                  % (name, ln, m.group(2)))
                declared[m.group(2)] = name
                per_module[name].add(m.group(2))
                p = RE_PARAMS.search(code)
                if p:
                    for arg in p.group(1).split(","):
                        a = arg.strip()
                        a = re.sub(r"^(?:ByVal|ByRef|Optional)\s+", "", a, flags=re.I)
                        a = re.split(r"\s+As\s+|\s*=", a)[0].strip()
                        if a:
                            known_locals.add(a)
            m = RE_LOCAL.match(code)
            if m:
                known_locals.add(m.group(1).rstrip("()"))
            m = RE_LABEL.match(code)          # метка перехода: Сбой:
            if m:
                known_locals.add(m.group(1))

            for op, cl, kind in BLOCKS:
                if cl.match(code):
                    if not stack or stack[-1][0] != kind:
                        errors.append("ОШИБКА %s:%d лишний End %s" % (name, ln, kind))
                    else:
                        stack.pop()
                elif op.match(code) and not re.match(r"^\s*(?:Declare|End)\b", code, re.I):
                    stack.append((kind, ln))
            if RE_FOR.match(code):
                stack.append(("For", ln))
            elif RE_NEXT.match(code):
                if not stack or stack[-1][0] != "For":
                    errors.append("ОШИБКА %s:%d Next без For" % (name, ln))
                else:
                    stack.pop()
            elif RE_DO.match(code):
                stack.append(("Do", ln))
            elif RE_LOOP.match(code):
                if not stack or stack[-1][0] != "Do":
                    errors.append("ОШИБКА %s:%d Loop без Do" % (name, ln))
                else:
                    stack.pop()
            elif RE_IF.match(code):
                stack.append(("If", ln))
            elif RE_ENDIF.match(code):
                if not stack or stack[-1][0] != "If":
                    errors.append("ОШИБКА %s:%d End If без If" % (name, ln))
                else:
                    stack.pop()

            m = RE_SETSHEET.match(code)
            if m:
                var_sheet[m.group(1)] = m.group(2) or SHEET_CONST.get(m.group(3))
            d2 = RE_DECL.match(code)
            if d2:
                var_sheet.clear()          # новая процедура — свои переменные
                текущая_процедура = d2.group(2)
            # Заголовок, склеенный из частей (ColReq(ws, "Замер " & c)),
            # проверить нельзя: имя известно только в момент выполнения.
            for mm in re.finditer(
                    r"Col(?:Req|ByHeader)\s*\(\s*(\w+)\s*,\s*\"([^\"]*)\"([^)]*)\)", raw):
                if "&" in mm.group(3):
                    continue
                col_uses.append((name, ln, mm.group(1), var_sheet.get(mm.group(1)),
                                 mm.group(2), текущая_процедура))

        check(not stack, "%s: все блоки закрыты%s" % (
            name, "" if not stack else " (не закрыт %s со строки %d)" % (stack[-1][0], stack[-1][1])))

    # ── квалифицированные вызовы modX.Имя ────────────────────────────
    # Код листа «Меню» и «ЭтаКнига» лежит в .txt (его вставляют руками,
    # импортировать нельзя), но на модули он ссылается так же.
    bad = []
    for path in files + sorted(glob.glob(os.path.join(MAC, "*.txt"))):
        name = os.path.basename(path)[:-4]
        for ln, raw in read_lines(path):
            for mm in re.finditer(r"\b(mod\w+)\.([А-Яа-яЁё\w]+)", strip_code(raw)):
                if mm.group(2) not in per_module.get(mm.group(1), set()):
                    bad.append("%s:%d %s.%s" % (name, ln, mm.group(1), mm.group(2)))
    check(not bad, "все вызовы вида modX.Имя ведут в существующую процедуру%s" %
          ("" if not bad else ": " + ", ".join(bad)))

    # ── кириллические вызовы: объявлены ли ───────────────────────────
    builtins = {"Данные", "Лист", "Настройка", "Путь"}
    unknown = collections.Counter()
    for path in files:
        name = os.path.basename(path)[:-4]
        for ln, raw in read_lines(path):
            code = strip_code(raw)
            if RE_DECL.match(code) or RE_LOCAL.match(code):
                continue
            for ident in re.findall(r"\b([А-ЯЁ][А-Яа-яЁё_]{1,})\b", code):
                if ident in declared or ident in known_locals or ident in builtins:
                    continue
                if ident in ("МЕТОДЫ", "ПЕРВАЯ_СТРОКА", "СТРОКА_ЗАГОЛОВКОВ", "ИМЯ_ФАЙЛА_ДАННЫХ"):
                    continue
                if ident.startswith("Л_") or ident.isupper():
                    continue
                unknown[ident] += 1
    check(not unknown, "нет обращений к необъявленным именам%s" %
          ("" if not unknown else ": " + ", ".join("%s (%d)" % (k, v) for k, v in unknown.most_common(8))))

    # Кириллица и латиница неразличимы на вид: «pусто» с латинской p и
    # «пусто» — разные имена, а Option Explicit сообщит об этом только при
    # запуске. Смешанное имя само по себе не ошибка (cФИО — обычный
    # префикс), ошибка — когда два РАЗНЫХ имени выглядят одинаково.
    CONFUSABLE = str.maketrans("аеорсухАЕОРСТХВКМН", "aeopcyxAEOPCTXBKMH")
    skeleton = collections.defaultdict(set)
    where = {}
    for path in files:
        name = os.path.basename(path)[:-4]
        for ln, raw in read_lines(path):
            for ident in re.findall(r"\b(\w{2,})\b", strip_code(raw)):
                skeleton[ident.translate(CONFUSABLE)].add(ident)
                where.setdefault(ident, "%s:%d" % (name, ln))
    collisions = ["/".join("%s (%s)" % (i, where[i]) for i in sorted(v))
                  for v in skeleton.values() if len(v) > 1]
    check(not collisions, "нет пар имён, различающихся только алфавитом%s" %
          ("" if not collisions else ": " + "; ".join(collisions[:5])))

    # ── циклы в вызовах ──────────────────────────────────────────────
    # «Out of stack space» — это цикл в графе вызовов, и увидеть его
    # чтением одного модуля нельзя: ПапкаОбмена звала поиск носителя,
    # тот — чтение списка, а оно снова ПапкаОбмена. Строим граф и ищем
    # замкнутые пути.
    edges = collections.defaultdict(set)
    for path in files + sorted(glob.glob(os.path.join(MAC, "*.txt"))):
        proc = ""
        for ln, raw in read_lines(path):
            code = strip_code(raw)
            d = RE_DECL.match(code)
            if d:
                proc = d.group(2)
                continue
            if not proc:
                continue
            for callee in declared:
                if callee == proc:
                    continue
                # Вызов бывает и квалифицированным: modMedia.Имя. Без учёта
                # префикса ребро графа терялось — и петля не находилась.
                if re.search(r"(?<![\wА-Яа-яЁё])(?:mod\w+\.)?" + re.escape(callee) +
                             r"(?![\wА-Яа-яЁё])", code):
                    edges[proc].add(callee)

    cycles, seen_cycle = [], set()

    def walk(node, path_so_far, visiting):
        if len(cycles) >= 5:
            return
        for nxt in sorted(edges.get(node, ())):
            if nxt in visiting:
                cyc = path_so_far[path_so_far.index(nxt):] + [nxt]
                key = tuple(sorted(set(cyc)))
                if key not in seen_cycle:
                    seen_cycle.add(key)
                    cycles.append(" → ".join(cyc))
            elif len(path_so_far) < 12:
                walk(nxt, path_so_far + [nxt], visiting | {nxt})

    for start in sorted(edges):
        walk(start, [start], {start})

    check(not cycles, "нет циклов в вызовах процедур%s" %
          ("" if not cycles else ": " + "; ".join(cycles[:3])))

    # ── Dir для проверки существования — только в modConfig ──────────
    # Dir на недоступном сетевом пути падает с ошибкой 52, а не
    # возвращает пустую строку. Проверки идут через ПапкаСуществует /
    # ФайлСуществует / ПервыйФайл, которые эту ошибку гасят.
    bare = []
    for path in files:
        name = os.path.basename(path)[:-4]
        if name == "modConfig":
            continue
        for ln, raw in read_lines(path):
            code = strip_code(raw)
            if re.search(r'Dir\([^)]*\)\s*(?:=|<>)\s*""', code) or "vbDirectory" in code:
                bare.append("%s:%d" % (name, ln))
    check(not bare, "Dir для проверки путей вызывается только в modConfig%s" %
          ("" if not bare else ": " + ", ".join(bare[:8])))

    # ── лист, переданный параметром, узнаём с места вызова ───────────
    params = {}            # процедура → список имён параметров
    for path in files:
        for ln, raw in read_lines(path):
            code = strip_code(raw)
            m = RE_DECL.match(code)
            if m:
                pp = RE_PARAMS.search(code)
                names = []
                if pp:
                    for arg in pp.group(1).split(","):
                        a = re.sub(r"^\s*(?:ByVal|ByRef|Optional)\s+", "", arg.strip(), flags=re.I)
                        names.append(re.split(r"\s+As\s+|\s*=", a)[0].strip())
                params[m.group(2)] = names

    param_sheet = {}       # (процедура, имя параметра) → лист
    for path in files:
        var_sheet, proc = {}, ""
        for ln, raw in read_lines(path):
            code = strip_code(raw)
            m = RE_SETSHEET.match(code)
            if m:
                var_sheet[m.group(1)] = m.group(2) or SHEET_CONST.get(m.group(3))
            d = RE_DECL.match(code)
            if d:
                var_sheet, proc = {}, d.group(2)
                continue
            for callee, names in params.items():
                if not names:
                    continue
                mm = re.search(r"(?:^|[=\s])" + re.escape(callee) + r"[ (]([^)\n]*)", code)
                if not mm:
                    continue
                for i, arg in enumerate(mm.group(1).split(",")):
                    a = arg.strip()
                    if i < len(names) and a in var_sheet and var_sheet[a]:
                        param_sheet[(callee, names[i])] = var_sheet[a]

    # ── пункты меню и их обработчик ──────────────────────────────────
    # Пункт запускается сравнением текста строки со списком Case.
    # Разошлись на один символ — пункт молча ничего не делает.
    menu = os.path.join(MAC, "modMenu.bas")
    if os.path.exists(menu):
        src = io.open(menu, encoding="utf-8").read()
        items = set(re.findall(r'Пункт\(ws, r, "([^"]+)"', src))
        cases = set(re.findall(r'Case "([^"]+)":\s*mod', src))
        check(items == cases,
              "каждый пункт меню разбирается обработчиком (пунктов %d, ветвей %d)%s"
              % (len(items), len(cases),
                 "" if items == cases else
                 "; без обработчика: %s; лишние ветви: %s" % (sorted(items - cases), sorted(cases - items))))

    # ── заголовки колонок существуют в книге ─────────────────────────
    if not os.path.exists(DATA):
        errors.append("ОШИБКА не найден %s — проверка заголовков пропущена" % DATA)
    else:
        import openpyxl
        wb = openpyxl.load_workbook(DATA, read_only=True)

        def norm(s):
            t = re.sub(r"\s+", " ", str(s or "").replace("\n", " "))
            return t.strip().lower()

        heads = {}
        for ws in wb.worksheets:
            heads[ws.title] = {norm(c.value) for c in next(ws.iter_rows(min_row=HDR_ROW, max_row=HDR_ROW))}
        all_heads = set()
        for s in heads.values():
            all_heads |= s
        methods = "ВИК УЗД УЗТ АЭ ТВ ТК ВД МК".split()
        missing, loose = [], []
        for mod, ln, var, sheet, title, proc in col_uses:
            if not sheet:
                sheet = param_sheet.get((proc, var))
            if sheet and sheet in heads:
                if norm(title) not in heads[sheet]:
                    missing.append("%s:%d %s → «%s» нет на листе «%s»" % (mod, ln, var, title, sheet))
            elif norm(title) not in all_heads:
                missing.append("%s:%d «%s» нет ни на одном листе" % (mod, ln, title))
            elif not sheet:
                loose.append("%s:%d %s" % (mod, ln, var))
        check(not missing, "заголовки колонок сверены с конкретным листом%s" %
              ("" if not missing else ": " + "; ".join(missing[:6])))
        check(True, "обращений с неизвестным листом (сверены со всеми листами): %d" % len(loose))
        # динамические заголовки вида «<метод> №» / «<метод> исполнитель»
        dyn_missing = [
            "%s %s" % (m, suf) for m in methods for suf in ("№", "исполнитель")
            if norm("%s %s" % (m, suf)) not in heads["Журнал"]
        ]
        check(not dyn_missing, "колонки методов «<метод> №» и «<метод> исполнитель» есть по всем 8 методам%s"
              % ("" if not dyn_missing else ": " + ", ".join(dyn_missing)))

    print("\n".join(oks))
    if errors:
        print("\n" + "\n".join(errors))
        print("\nПРОВЕРКА НЕ ПРОЙДЕНА: %d" % len(errors))
        return 1
    print("\nПроверка пройдена: %d контролей." % len(oks))
    return 0


if __name__ == "__main__":
    sys.exit(main())
