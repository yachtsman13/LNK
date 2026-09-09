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
    private_of = collections.defaultdict(set)   # модуль → его Private-процедуры
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
                if re.match(r"^\s*Private\b", code, re.I):
                    private_of[name].add(m.group(2))
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

    # ── объявления уровня модуля — только до первой процедуры ────────
    # Excel 2016 читает Private/Dim/Const уровня модуля лишь в секции
    # объявлений, до первой Sub/Function. Объявление ниже по файлу он
    # молча не видит, а обращение к нему падает «Variable not defined» —
    # с подсветкой самой константы, которая в файле есть и написана
    # верно. Проверка родилась на живой установке: три константы
    # modDashboard стояли посреди файла, и пульт не собирался.
    поздние = []
    for path in files:
        name = os.path.basename(path)[:-4]
        начались_процедуры = False
        внутри = False
        for ln, raw in read_lines(path):
            code = strip_code(raw)
            if not code.strip():
                continue
            if re.match(r"^\s*End\s+(Sub|Function|Property)\b", code, re.I):
                внутри = False
                continue
            if RE_DECL.match(code):
                начались_процедуры = True
                внутри = True
                continue
            if внутри:
                continue
            if начались_процедуры and re.match(
                    r"^\s*(?:Public|Private|Dim|Const|Type|Enum|Declare)\b", code, re.I):
                поздние.append("%s:%d %s" % (name, ln, code.strip()[:50]))
    check(not поздние, "объявления уровня модуля стоят до первой процедуры%s" %
          ("" if not поздние else ": " + "; ".join(поздние[:5])))

    # ── квалифицированные вызовы modX.Имя ────────────────────────────
    # Код листа «Меню» и «ЭтаКнига» лежит в .txt (его вставляют руками,
    # импортировать нельзя), но на модули он ссылается так же.
    bad = []
    закрытые = []
    for path in files + sorted(glob.glob(os.path.join(MAC, "*.txt"))):
        name = os.path.basename(path)[:-4]
        for ln, raw in read_lines(path):
            for mm in re.finditer(r"\b(mod\w+)\.([А-Яа-яЁё\w]+)", strip_code(raw)):
                if mm.group(2) not in per_module.get(mm.group(1), set()):
                    bad.append("%s:%d %s.%s" % (name, ln, mm.group(1), mm.group(2)))
                elif mm.group(1) != name and mm.group(2) in private_of.get(mm.group(1), set()):
                    закрытые.append("%s:%d %s.%s" % (name, ln, mm.group(1), mm.group(2)))
    check(not bad, "все вызовы вида modX.Имя ведут в существующую процедуру%s" %
          ("" if not bad else ": " + ", ".join(bad)))
    # Private видна только внутри своего модуля: снаружи это ошибка
    # компиляции, а не предупреждение.
    check(not закрытые, "вызываемые снаружи процедуры объявлены Public%s" %
          ("" if not закрытые else ": " + ", ".join(закрытые)))

    # ── кириллические вызовы: объявлены ли ───────────────────────────
    builtins = {"Данные", "Лист", "Настройка", "Путь"}
    unknown = collections.Counter()
    for path in files:
        name = os.path.basename(path)[:-4]
        for ln, raw in read_lines(path):
            code = strip_code(raw)
            if RE_DECL.match(code) or RE_LOCAL.match(code):
                continue
            # После точки стоит не имя проекта, а свойство или метод
            # объекта: у формы поиска это ф.Список, ф.Запрос. Проверять
            # их здесь нельзя — они объявлены не в модулях, а на самой
            # форме. Квалифицированные вызовы modX.Имя разбирает
            # отдельная проверка выше, так что ничего не теряется.
            for ident in re.findall(r"(?<![.\w])([А-ЯЁ][А-Яа-яЁё_]{1,})\b", code):
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

    # ── локальная переменная, затеняющая процедуру ───────────────────
    # VBA не различает регистр: Dim строкаСИ и Function СтрокаСИ — это
    # одно имя. Внутри такой процедуры вызов СтрокаСИ(...) обращается к
    # переменной, и книга не компилируется. Ошибка тихая: находится не
    # при написании, а при первом запуске у человека.
    затенения = []
    DECL_FULL = re.compile(
        r"^\s*(?:Public |Private )?(?:Static )?(Sub|Function)\s+([^\s(]+)\s*(?:\(([^)]*)\))?", re.I)
    RE_END = re.compile(r"^\s*End (Sub|Function)\b", re.I)
    процедуры = {имя.lower(): имя for имя in declared}
    for path in files:
        имяМодуля = os.path.basename(path)[:-4]
        строки = read_lines(path)
        k = 0
        while k < len(строки):
            d = DECL_FULL.match(strip_code(строки[k][1]))
            if not d:
                k += 1
                continue
            локальные, тело = {}, []
            for a in (d.group(3) or "").split(","):
                a = re.sub(r"^\s*(?:ByVal|ByRef|Optional)\s+", "", a.strip(), flags=re.I)
                части = a.split()
                if части:
                    локальные[части[0].lower()] = части[0]
            k += 1
            while k < len(строки) and not RE_END.match(strip_code(строки[k][1])):
                code = strip_code(строки[k][1])
                m = RE_LOCAL.match(code)
                if m:
                    локальные[m.group(1).lower()] = m.group(1)
                тело.append(code)
                k += 1
            текст = "\n".join(тело)
            for низ, как in локальные.items():
                if низ not in процедуры:
                    continue
                имя = re.escape(процедуры[низ])
                # Вызвать процедуру можно двумя записями, и затеняет их
                # обе одинаково: со скобками — Имя(арг) — и без скобок,
                # оператором: Имя арг1, арг2. Вторую форму проверка
                # раньше не видела, и «Подсказка ws, r, подсказка»
                # доехало до заказчика: VBA читает там строковую
                # переменную вместо вызова и говорит «Expected Sub,
                # Function, or Property».
                со_скобками = r"(?<![.\w])" + имя + r"\s*\("
                оператором = r"(?im)^[ \t]*(?:Call[ \t]+)?" + имя + r"[ \t]+(?![=<>])"
                if re.search(со_скобками, текст, re.I) or re.search(оператором, текст):
                    затенения.append("%s:%s локальная «%s» затеняет %s" %
                                     (имяМодуля, d.group(2), как, процедуры[низ]))
    check(not затенения, "локальные имена не затеняют вызываемые процедуры%s" %
          ("" if not затенения else ": " + "; ".join(затенения)))

    # ── число аргументов в вызовах ───────────────────────────────────
    # «Wrong number of arguments» — тоже ошибка компиляции, то есть до
    # первого запуска её не видно, а после — не работает ничего. Считать
    # аргументы приходится с оглядкой: strip_code стирает строки вместе
    # с запятыми внутри них, поэтому здесь строковый литерал заменяется
    # на «"S"» и остаётся одним аргументом.
    #
    # Область видимости важна: ПересоздатьЛист без параметров в modCard
    # и ПересоздатьЛист(имя) в modSM — разные Private-процедуры, и
    # сравнивать вызов надо со своей, а не с чужой.
    def строка_с_литералами(line):
        out, i, instr = [], 0, False
        while i < len(line):
            ch = line[i]
            if ch == '"':
                out.append('"S' if not instr else '"')
                instr = not instr
                i += 1
                continue
            if instr:
                i += 1
                continue
            if ch == "'":
                break
            out.append(ch)
            i += 1
        return "".join(out)

    def тело_скобок(s, откр):
        depth, i, buf = 0, откр, ""
        while i < len(s):
            ch = s[i]
            if ch == "(":
                depth += 1
                if depth == 1:
                    i += 1
                    continue
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return buf
            if depth >= 1:
                buf += ch
            i += 1
        return buf

    def сколько(body):
        if not body.strip():
            return 0
        depth, n = 0, 1
        for ch in body:
            if ch in "([":
                depth += 1
            elif ch in ")]":
                depth -= 1
            elif ch == "," and depth == 0:
                n += 1
        return n

    ГОЛОВА = re.compile(
        r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+"
        r"([А-Яа-яЁёA-Za-z_]\w*)\s*\(", re.I)
    свои, общие = {}, {}
    for path in files:
        мод = os.path.basename(path)[:-4]
        for ln, raw in read_lines(path):
            code = строка_с_литералами(raw)
            m = ГОЛОВА.match(code)
            if not m:
                continue
            тело = тело_скобок(code, code.index("(", m.end(3)))
            части, depth, cur = [], 0, ""
            for ch in тело:
                if ch in "([":
                    depth += 1
                elif ch in ")]":
                    depth -= 1
                if ch == "," and depth == 0:
                    части.append(cur)
                    cur = ""
                else:
                    cur += ch
            if cur.strip():
                части.append(cur)
            опц = sum(1 for a in части if re.match(r"^\s*Optional\b", a.strip(), re.I))
            свободных = any(re.search(r"\bParamArray\b", a, re.I) for a in части)
            запись = (m.group(2).lower(), len(части) - опц,
                      999 if свободных else len(части), m.group(3))
            свои[(мод, m.group(3).lower())] = запись
            if (m.group(1) or "Public").lower() != "private":
                общие[m.group(3).lower()] = запись

    def чья(мод, имя):
        return свои.get((мод, имя)) or общие.get(имя)

    арность, скобки_у_sub = [], []
    for path in files:
        мод = os.path.basename(path)[:-4]
        внутри = False
        for ln, raw in read_lines(path):
            code = строка_с_литералами(raw)
            if ГОЛОВА.match(code):
                внутри = True
                continue
            if re.match(r"^\s*End\s+(Sub|Function)\b", code, re.I):
                внутри = False
                continue
            if not внутри:
                continue

            # вызов оператором: Имя арг1, арг2
            m = re.match(r"^[ \t]*(?:Call[ \t]+)?([А-Яа-яЁёA-Za-z_]\w*)[ \t]+(?![=<>])(.*)$", code)
            if m:
                p = чья(мод, m.group(1).lower())
                хвост = m.group(2).strip()
                if p and not re.match(r"^(As|Then|=|To\b)", хвост, re.I):
                    n = сколько(хвост)
                    if not (p[1] <= n <= p[2]):
                        арность.append("%s:%d %s — передано %d, объявлено %d..%d"
                                       % (мод, ln, p[3], n, p[1], p[2]))

            # вызов со скобками: Имя(арг1, арг2)
            for mm in re.finditer(r"(?<![.\w])([А-Яа-яЁёA-Za-z_]\w*)\s*\(", code):
                p = чья(мод, mm.group(1).lower())
                if not p:
                    continue
                n = сколько(тело_скобок(code, code.index("(", mm.end(1))))
                перед = code[:mm.start()].rstrip().lower()
                # Sub со скобками и несколькими аргументами — синтаксическая
                # ошибка VBA: либо без скобок, либо через Call.
                if p[0] == "sub" and n > 1 and not перед.endswith("call"):
                    скобки_у_sub.append("%s:%d %s (%d аргументов)" % (мод, ln, p[3], n))
                if not (p[1] <= n <= p[2]):
                    арность.append("%s:%d %s — передано %d, объявлено %d..%d"
                                   % (мод, ln, p[3], n, p[1], p[2]))

    арность = sorted(set(арность))
    check(not арность, "число аргументов в вызовах совпадает с объявлением%s" %
          ("" if not арность else ": " + "; ".join(арность[:5])))
    check(not скобки_у_sub, "Sub не вызывается со скобками и несколькими аргументами%s" %
          ("" if not скобки_у_sub else ": " + "; ".join(скобки_у_sub[:5])))

    # ── пункты меню и их обработчик ──────────────────────────────────
    # Пункт запускается сравнением текста строки со списком Case.
    # Разошлись на один символ — пункт молча ничего не делает.
    menu = os.path.join(MAC, "modMenu.bas")
    if os.path.exists(menu):
        src = io.open(menu, encoding="utf-8").read()
        items = set(re.findall(r'\bКнопка\(ws, r, "([^"]+)"', src))
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

    # Код «ЭтаКнига» не должен звать модули по имени напрямую.
    # Он попадает в книгу отдельно от модулей — текстом, а не импортом,
    # — поэтому возможно состояние «код есть, модулей нет». Прямой вызов
    # modМодуль.Процедура в этом состоянии даёт Compile error: Variable
    # not defined, и не собирается ВЕСЬ проект: не работает ни лента, ни
    # пульт, ни правая кнопка, а причина из сообщения не следует никак.
    # Один раз это уже стоило заказчику нескольких дней. Только
    # Application.Run: имя ищется во время выполнения.
    книга = io.open(os.path.join(MAC, "ЭтаКнига_код.txt"), encoding="utf-8").read()
    прямые = []
    for n, строка in enumerate(книга.splitlines(), 1):
        # Сначала снимаются строковые литералы, и только потом
        # комментарий: имена в кавычках — это как раз то, что нужно,
        # они и есть Application.Run.
        голая = re.sub(r'"[^"]*"', '""', строка).split("'")[0]
        for м in re.finditer(r'\b(mod[A-Za-zА-Яа-я]\w*)\s*\.', голая):
            прямые.append("строка %d: %s" % (n, м.group(1)))
    check(not прямые,
          "«ЭтаКнига» зовёт модули только через Application.Run%s"
          % ("" if not прямые else ": " + ", ".join(прямые)))

    # Каждая процедура, которую «ЭтаКнига» зовёт по имени, должна
    # существовать. Application.Run опечатку не ловит при сборке —
    # значит, ловим здесь.
    нет_такой = []
    for имя in sorted(set(re.findall(r'"(mod\w+\.[^"]+)"', книга))):
        модуль, процедура = имя.split(".", 1)
        путь = os.path.join(MAC, модуль + ".bas")
        if not os.path.exists(путь):
            нет_такой.append(имя + " (нет модуля)")
        elif not re.search(r'^\s*Public\s+(Sub|Function)\s+%s\b' % re.escape(процедура),
                           io.open(путь, encoding="utf-8").read(), re.M):
            нет_такой.append(имя)
    check(not нет_такой,
          "все процедуры, которые зовёт «ЭтаКнига», существуют и Public%s"
          % ("" if not нет_такой else ": " + ", ".join(нет_такой)))

    # Типы из внешних библиотек (Office, Word, Scripting, MSForms).
    # Тип проверяется при СБОРКЕ проекта: нет ссылки на библиотеку — не
    # собирается весь проект, с ошибкой «User-defined type not defined»
    # на этой строке, и не работает уже ничего. Ссылки на Word и Office
    # в Excel есть по умолчанию, но «по умолчанию» — не «всегда».
    # Позднее связывание (As Object + CreateObject) этой проверки не
    # требует: имена ищутся во время выполнения. Так сделано во всей
    # системе, и одно исключение однажды встало поперёк всего.
    чужие_типы = []
    ОПАСНЫЕ = (r"IRibbon\w*", r"Word\.\w+", r"Scripting\.\w+", r"MSForms\.\w+",
               r"Office\.\w+", r"FileDialog", r"CommandBar\b", r"Outlook\.\w+")
    for файл in sorted(glob.glob(os.path.join(MAC, "*.bas"))):
        for n, строка in enumerate(io.open(файл, encoding="utf-8"), 1):
            голая = строка.split("'")[0]
            for тип in ОПАСНЫЕ:
                if re.search(r"\bAs\s+(New\s+)?%s" % тип, голая):
                    чужие_типы.append("%s:%d" % (os.path.basename(файл), n))
                    break
    check(not чужие_типы,
          "нет типов из внешних библиотек (только As Object)%s"
          % ("" if not чужие_типы else ": " + ", ".join(чужие_типы)))

    # Отладочные операторы в присланном коде. Stop и Debug.Assert ставят
    # Excel заказчика на паузу в редакторе VBA — ровно то состояние,
    # в котором перестаёт работать вообще всё, а сообщение при этом
    # говорит только «Can't execute code in break mode», ни словом не
    # называя причину. Debug.Print безвреден, но в чужом Excel окно
    # отладки никто не смотрит: значит, это забытый след отладки.
    отладочные = []
    for файл in sorted(glob.glob(os.path.join(MAC, "*.bas"))):
        for n, строка in enumerate(io.open(файл, encoding="utf-8"), 1):
            голая = строка.split("'")[0]
            if re.search(r'^\s*Stop\s*$', голая) or re.search(r'\bDebug\.(Assert|Print)\b', голая):
                отладочные.append("%s:%d" % (os.path.basename(файл), n))
    check(not отладочные,
          "нет отладочных Stop и Debug.* (они останавливают Excel заказчика)%s"
          % ("" if not отладочные else ": " + ", ".join(отладочные)))

    # Имя, отличающееся от константы модуля только регистром. Редактор
    # VBA приводит такие имена к одному написанию по всему проекту: на
    # экране КОНСТАНТА превращается в «константу» и читается переменной,
    # хотя код при этом работает. Один раз это уже сбило с толку прямо
    # на скриншоте от заказчика.
    регистр = []
    for файл in sorted(glob.glob(os.path.join(MAC, "*.bas"))):
        текст = io.open(файл, encoding="utf-8").read()
        консты = {}
        for м in re.finditer(r'^\s*(?:Public|Private)\s+Const\s+(\w+)', текст, re.M):
            консты[м.group(1).lower()] = м.group(1)
        имена = set(re.findall(r'^\s*Dim\s+(\w+)', текст, re.M))
        служебные = ("Optional", "ByVal", "ByRef", "ParamArray")
        for группа in re.findall(r'(?:Sub|Function)\s+\w+\s*\(([^)]*)\)', текст):
            for аргумент in группа.split(","):
                слова = [с for с in аргумент.split() if с not in служебные]
                if слова:
                    имена.add(слова[0])
        for имя in имена:
            эталон = консты.get(имя.lower())
            if эталон and эталон != имя:
                регистр.append("%s: %s ↔ Const %s" % (os.path.basename(файл), имя, эталон))
    check(not регистр,
          "нет имён, отличающихся от константы модуля только регистром%s"
          % ("" if not регистр else ": " + "; ".join(регистр)))

    # Версия кода и версия в книге данных приходят одной присылкой и
    # обязаны совпадать. Разойдясь, они дают самую неприятную породу
    # ошибок: код зовёт колонку, которой в старой книге ещё нет, и
    # ругается на колонку — про забытое обновление ни слова.
    m = re.search(r'ВЕРСИЯ_КОДА\s+As\s+String\s*=\s*"([^"]+)"',
                  io.open(os.path.join(MAC, "modConfig.bas"), encoding="utf-8").read())
    сборка = io.open(os.path.join(ROOT, "scripts", "build_data.py"), encoding="utf-8").read()
    m2 = re.search(r'\("Версия программы",\s*"([^"]+)"', сборка)
    в_коде = m.group(1) if m else None
    в_книге = m2.group(1) if m2 else None
    check(в_коде is not None and в_коде == в_книге,
          "версия в modConfig и в книге данных одна (%s / %s)" % (в_коде, в_книге))

    print("\n".join(oks))
    if errors:
        print("\n" + "\n".join(errors))
        print("\nПРОВЕРКА НЕ ПРОЙДЕНА: %d" % len(errors))
        return 1
    print("\nПроверка пройдена: %d контролей." % len(oks))
    return 0


if __name__ == "__main__":
    sys.exit(main())
