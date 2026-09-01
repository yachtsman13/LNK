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
RE_LOCAL = re.compile(r"^\s*(?:Dim|Const|ReDim|Static)\s+([^\s,]+)", re.I)
RE_LABEL = re.compile(r"^\s*([А-Яа-яЁёA-Za-z_]\w*):\s*$")
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
        for ln, raw in lines:
            code = strip_code(raw)
            m = RE_DECL.match(code)
            if m:
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

            for mm in re.finditer(r"Col(?:Req|ByHeader)\s*\(\s*(\w+)\s*,\s*\"([^\"]+)\"", raw):
                col_uses.append((name, ln, mm.group(1), mm.group(2)))

        check(not stack, "%s: все блоки закрыты%s" % (
            name, "" if not stack else " (не закрыт %s со строки %d)" % (stack[-1][0], stack[-1][1])))

    # ── квалифицированные вызовы modX.Имя ────────────────────────────
    bad = []
    for path in files:
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
            for ident in re.findall(r"\b([А-ЯЁ][А-Яа-яЁё_]{3,})\b", code):
                if ident in declared or ident in known_locals or ident in builtins:
                    continue
                if ident in ("МЕТОДЫ", "ПЕРВАЯ_СТРОКА", "СТРОКА_ЗАГОЛОВКОВ", "ИМЯ_ФАЙЛА_ДАННЫХ"):
                    continue
                if ident.startswith("Л_") or ident.isupper():
                    continue
                unknown[ident] += 1
    check(not unknown, "нет обращений к необъявленным именам%s" %
          ("" if not unknown else ": " + ", ".join("%s (%d)" % (k, v) for k, v in unknown.most_common(8))))

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
        # какой лист стоит за переменной в коде — по имени переменной
        var2sheet = {"ws": None, "att": "Аттестации", "src": "Журнал"}
        all_heads = set()
        for s in heads.values():
            all_heads |= s
        methods = "ВИК УЗД УЗТ АЭ ТВ ТК ВД МК".split()
        missing = []
        for mod, ln, var, title in col_uses:
            cands = [title]
            if "метод &" in title or "методы(" in title:
                continue
            for h in [title]:
                if norm(h) in all_heads:
                    break
            else:
                missing.append("%s:%d «%s»" % (mod, ln, title))
        check(not missing, "все заголовки колонок из ColReq есть в книге%s" %
              ("" if not missing else ": " + "; ".join(missing)))
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
