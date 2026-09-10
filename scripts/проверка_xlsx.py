# -*- coding: utf-8 -*-
"""Проверка сгенерированного .xlsx: порядок элементов по схеме ECMA-376.

Excel, как и Word, не говорит «переставлен элемент» — он объявляет книгу
повреждённой. Порядок проверяется здесь, а не на глаз.
"""
import sys, zipfile, posixpath, re, xml.etree.ElementTree as ET

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

ПОРЯДОК = {
 "worksheet": ["sheetPr","dimension","sheetViews","sheetFormatPr","cols","sheetData",
   "sheetCalcPr","sheetProtection","protectedRanges","scenarios","autoFilter","sortState",
   "dataConsolidate","customSheetViews","mergeCells","phoneticPr","conditionalFormatting",
   "dataValidations","hyperlinks","printOptions","pageMargins","pageSetup","headerFooter",
   "rowBreaks","colBreaks","customProperties","cellWatches","ignoredErrors","smartTags",
   "drawing","legacyDrawing","picture","tableParts","extLst"],
 "workbook": ["fileVersion","fileSharing","workbookPr","workbookProtection","bookViews",
   "sheets","functionGroups","externalReferences","definedNames","calcPr","oleSize",
   "customWorkbookViews","pivotCaches","smartTagPr","smartTagTypes","webPublishing",
   "fileRecoveryPr","webPublishObjects","extLst"],
 "styleSheet": ["numFmts","fonts","fills","borders","cellStyleXfs","cellXfs","cellStyles",
   "dxfs","tableStyles","colors","extLst"],
 "row": ["c","extLst"],
 "c": ["f","v","is","extLst"],
}

def проверить(путь):
    z = zipfile.ZipFile(путь); беды = []; имена = set(z.namelist())
    if z.testzip() is not None: беды.append("ZIP повреждён")

    ct = z.read("[Content_Types].xml").decode("utf-8")
    for часть in имена:
        if часть.endswith(".rels") or часть == "[Content_Types].xml": continue
        ext = часть.rsplit(".", 1)[-1]
        if ('PartName="/' + часть + '"') not in ct and ('Extension="' + ext + '"') not in ct:
            беды.append("нет типа содержимого: " + часть)
    for rels in [n for n in имена if n.endswith(".rels")]:
        база = rels.rsplit("_rels/", 1)[0]
        for m in re.finditer(r'Target="([^"]+)"', z.read(rels).decode("utf-8")):
            цель = m.group(1)
            if цель.startswith("http"): continue
            путь_цели = posixpath.normpath(posixpath.join(база, цель)).lstrip("/")
            if путь_цели not in имена:
                беды.append("связь ведёт в никуда: " + путь_цели)

    for часть in sorted(n for n in имена if n.endswith(".xml") and n.startswith("xl/")):
        root = ET.fromstring(z.read(часть))
        for эл in root.iter():
            имя = эл.tag[len(S):] if эл.tag.startswith(S) else None
            if имя not in ПОРЯДОК: continue
            ожид = ПОРЯДОК[имя]
            дети = [c.tag[len(S):] for c in эл if c.tag.startswith(S)]
            прошлое = -1
            for д in дети:
                if д not in ожид:
                    беды.append("%s: <%s> не бывает внутри <%s>" % (часть, д, имя)); continue
                i = ожид.index(д)
                if i < прошлое:
                    беды.append("%s: в <%s> элемент <%s> стоит не на своём месте (%s)"
                                % (часть, имя, д, " → ".join(дети[:8])))
                    break
                прошлое = i
    return беды

if __name__ == "__main__":
    беды = проверить(sys.argv[1])
    if беды:
        print("НАЙДЕНЫ НАРУШЕНИЯ (%d):" % len(беды))
        for б in беды[:20]: print("  ✗", б)
        sys.exit(1)
    print("Проверка пройдена: архив цел, все части объявлены, связи ведут к существующим "
          "частям, порядок элементов соответствует схеме.")
