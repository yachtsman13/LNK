# -*- coding: utf-8 -*-
"""Проверка сгенерированного .docx: порядок элементов по схеме ECMA-376.

Word не сообщает «переставлен элемент» — он объявляет файл повреждённым.
Поэтому порядок проверяется здесь, а не на глаз.
"""
import sys, zipfile, re, posixpath, xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Последовательности из схемы (CT_PPr, CT_RPr, CT_SectPr, CT_TblPr, CT_TcPr, CT_TrPr)
ПОРЯДОК = {
 "pPr": ["pStyle","keepNext","keepLines","pageBreakBefore","framePr","widowControl","numPr",
   "suppressLineNumbers","pBdr","shd","tabs","suppressAutoHyphens","kinsoku","wordWrap",
   "overflowPunct","topLinePunct","autoSpaceDE","autoSpaceDN","bidi","adjustRightInd",
   "snapToGrid","spacing","ind","contextualSpacing","mirrorIndents","suppressOverlap","jc",
   "textDirection","textAlignment","textboxTightWrap","outlineLvl","divId","cnfStyle","rPr",
   "sectPr","pPrChange"],
 "rPr": ["rStyle","rFonts","b","bCs","i","iCs","caps","smallCaps","strike","dstrike","outline",
   "shadow","emboss","imprint","noProof","snapToGrid","vanish","webHidden","color","spacing",
   "w","kern","position","sz","szCs","highlight","u","effect","bdr","shd","fitText","vertAlign",
   "rtl","cs","em","lang","eastAsianLayout","specVanish","oMath","rPrChange"],
 "sectPr": ["headerReference","footerReference","footnotePr","endnotePr","type","pgSz","pgMar",
   "paperSrc","pgBorders","lnNumType","pgNumType","cols","formProt","vAlign","noEndnote",
   "titlePg","textDirection","bidi","rtlGutter","docGrid","printerSettings","sectPrChange"],
 "tblPr": ["tblStyle","tblpPr","tblOverlap","bidiVisual","tblStyleRowBandSize",
   "tblStyleColBandSize","tblW","jc","tblCellSpacing","tblInd","tblBorders","shd","tblLayout",
   "tblCellMar","tblLook","tblCaption","tblDescription","tblPrChange"],
 "tcPr": ["cnfStyle","tcW","gridSpan","hMerge","vMerge","tcBorders","shd","noWrap","tcMar",
   "textDirection","tcFitText","vAlign","hideMark","headers","tcPrChange"],
 "trPr": ["cnfStyle","divId","gridBefore","gridAfter","wBefore","wAfter","cantSplit","trHeight",
   "tblHeader","tblCellSpacing","jc","hidden","ins","del","trPrChange"],
}

def проверить(путь):
    z = zipfile.ZipFile(путь)
    беды = []
    имена = set(z.namelist())

    # 1. Целостность архива
    if z.testzip() is not None:
        беды.append("ZIP повреждён")

    # 2. Каждая часть из связей существует; у каждой есть тип содержимого
    ct = z.read("[Content_Types].xml").decode("utf-8")
    for часть in имена:
        if часть.endswith(".rels") or часть == "[Content_Types].xml":
            continue
        ext = часть.rsplit(".", 1)[-1]
        if ('PartName="/' + часть + '"') not in ct and ('Extension="' + ext + '"') not in ct:
            беды.append("нет типа содержимого: " + часть)
    for rels in [n for n in имена if n.endswith(".rels")]:
        база = rels.rsplit("_rels/", 1)[0]
        for m in re.finditer(r'Target="([^"]+)"[^>]*?(TargetMode="External")?/>', z.read(rels).decode("utf-8")):
            цель, внешняя = m.group(1), m.group(2)
            if внешняя or цель.startswith("http"):
                continue
            путь_цели = posixpath.normpath(posixpath.join(база, цель)).lstrip("/")
            if путь_цели not in имена:
                беды.append("связь ведёт в никуда: " + путь_цели)

    # 3. Порядок элементов
    for часть in sorted(n for n in имена if n.endswith(".xml") and n.startswith("word/")):
        root = ET.fromstring(z.read(часть))
        for эл in root.iter():
            имя = эл.tag[len(W):] if эл.tag.startswith(W) else None
            if имя not in ПОРЯДОК:
                continue
            ожид = ПОРЯДОК[имя]
            дети = [c.tag[len(W):] for c in эл if c.tag.startswith(W)]
            позиции, прошлое = [], -1
            for д in дети:
                if д not in ожид:
                    беды.append("%s: <%s> не бывает внутри <%s>" % (часть, д, имя)); continue
                i = ожид.index(д)
                if i < прошлое:
                    беды.append("%s: в <%s> элемент <%s> стоит не на своём месте (%s)"
                                % (часть, имя, д, " → ".join(дети)))
                    break
                прошлое = i
    return беды

if __name__ == "__main__":
    путь = sys.argv[1]
    беды = проверить(путь)
    if беды:
        print("НАЙДЕНЫ НАРУШЕНИЯ (%d):" % len(беды))
        for б in беды[:20]: print("  ✗", б)
        sys.exit(1)
    print("Проверка пройдена: архив цел, все части объявлены, связи ведут к существующим "
          "частям, порядок элементов соответствует схеме.")
