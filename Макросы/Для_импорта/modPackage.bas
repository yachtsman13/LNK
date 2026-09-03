Attribute VB_Name = "modPackage"
Option Explicit

' =======================================================================
'  —борка и печать пакета документов по одной работе.
'
'  ѕор€док в пакете задан вами и не мен€етс€:
'    акт -> протоколы -> свидетельства лаборатории -> квалификационные
'    удостоверени€ -> свидетельства о поверке -> пустой лист дл€ сшива.
'
'  —квозна€ нумераци€ внутри экземпл€ра, на первой странице номера нет.
'  “ри экземпл€ра: каждый нумеруетс€ заново, поэтому печатаютс€ копии
'  одного документа, а не три склеенных подр€д.
'
'  —видетельства и удостоверени€ Ч сканы, часть в JPG/PNG/TIFF, часть в
'  PDF (сканируют разные люди по-разному, и это нормально).  артинки
'  Word вставл€ет сам Ч их страницы пр€мо в своде, в общей нумерации.
'  PDF Word вставить не может; их вместо этого объедин€ет через Acrobat
'  modScans в один файл Ч печатаетс€ и подшиваетс€ отдельно, без
'  сквозного номера (см. modScans про то, почему номер туда программно
'  не ставитс€).
' =======================================================================

Private Const wdPageBreak As Long = 7
Private Const wdStory As Long = 6
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignPageNumberCenter As Long = 1
Private Const wdFieldNumPages As Long = 26
Private Const wdStatisticPages As Long = 2
Private Const wdFieldEmpty As Long = -1
Private Const wdFieldPage As Long = 33
Private Const wdReplaceOne As Long = 1
Private Const wdFindStop As Long = 0
Private Const wdAlignParagraphCenter As Long = 1


Public Sub Ќапечататьѕакет()
    Dim ws As Worksheet
    Dim r As Long
    Dim файлы As Collection
    Dim строки As Collection
    Dim акт As String
    Dim сканов As Long
    Dim wdApp As Object
    Dim свод As Object
    Dim страниц—вода As Long
    Dim страниц—канов As Long
    Dim путь—канов As String
    Dim путьPDF As String
    Dim ошибкаPDF As String
    Dim нумеровать As Boolean
    Dim всего As Long
    Dim ответ As VbMsgBoxResult

    Set ws = Ћист(Ћ_∆”–ЌјЋ)
    If ws Is Nothing Then Exit Sub
    r = “екуща€—трока(ws)
    If r = 0 Then Exit Sub

    ' ѕакет собираетс€ по ј “” целиком: акт один на работу, а объектов
    ' в ней несколько, и протоколы по всем из них идут в один пакет.
    ' ≈сли акта нет Ч печатаетс€ пакет по одной строке.
    акт = Ќомерјкта(ws, r)
    Set строки = —трокијкта(ws, r, акт)
    Set файлы = —обрать‘айлыјкта(ws, строки)
    If файлы.Count = 0 Then
        MsgBox "Ќет сформированных документов" & _
               IIf(акт = "", " по этой строке.", " по акту " & акт & ".") & vbCrLf & vbCrLf & _
               "—начала Ђ—формировать документыї Ч они по€в€тс€ в папке Ђ√отовыеї. " & _
               "≈сли в акте несколько объектов, документы нужны по каждому.", _
               vbExclamation, "ѕечать пакета"
        Exit Sub
    End If

    Application.StatusBar = "—обираю пакетЕ"
    On Error GoTo —бой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    ' —каны-картинки идут в свод как обычные страницы: свидетельства
    ' лаборатории, удостоверени€, свидетельства о поверке Ч в этом
    ' пор€дке, после протоколов. ¬кладывать их руками и держать вне
    ' нумерации больше не нужно.
    путь—канов = modScans.—обрать—каныѕакета(wdApp, строки, страниц—канов, _
                                             путьPDF, сканов, ошибкаPDF)
    If путь—канов <> "" Then файлы.Add путь—канов

    ' PDF-сканы объедин€ютс€ той же процедурой в один файл, и Acrobat
    ' сам сообщает число страниц Ч спрашивать человека не нужно, он
    ' его и не знает, пока не откроет файл. —просим только если
    ' объединение не удалось (Acrobat не найден, все файлы повреждены):
    ' тогда считать приходитс€ по старинке, как раньше.
    If путьPDF = "" And ошибкаPDF <> "" Then
        сканов = —просить„исло—канов()
        If сканов < 0 Then
            wdApp.Quit
            Exit Sub
        End If
    End If

    Set свод = —клеитьƒокументы(wdApp, файлы)

    ƒобавитьѕустойЋист свод
    страниц—вода = свод.ComputeStatistics(wdStatisticPages)
    нумеровать = ѕоЋистам()
    ѕроставитьЌомера свод, нумеровать
    If нумеровать Then
        ' ƒвусторонн€€ печать: лист Ч это две страницы, и последн€€
        ' нечЄтна€ страница всЄ равно занимает целый лист.
        всего = (страниц—вода + 1) \ 2 + сканов
    Else
        всего = страниц—вода + сканов
    End If

    Application.StatusBar = False
    ответ = MsgBox("ѕакет собран." & vbCrLf & vbCrLf & _
                   IIf(акт = "", "Ѕез акта, одна строка журнала." & vbCrLf, _
                       акт & ", объектов в нЄм: " & строки.Count & vbCrLf) & _
                   "ƒокументов Word: " & файлы.Count & vbCrLf & _
                   "—траниц в них (с пустым листом): " & страниц—вода & vbCrLf & _
                   IIf(страниц—канов = 0, "", "»з них страниц сканов-картинок: " & страниц—канов & _
                       " Ч они уже в своде и в общей нумерации" & vbCrLf) & _
                   IIf(путьPDF = "", "", "PDF-сканы объединены Acrobat в " & сканов & _
                       " стр.: " & путьPDF & vbCrLf) & _
                   IIf(путьPDF = "" And сканов > 0, "—траниц PDF, вкладываемых вручную: " & _
                       сканов & vbCrLf, "") & _
                   "Ќумераци€: " & IIf(нумеровать, "по листам, номер на лицевой стороне", _
                                       "по страницам, сквозна€") & vbCrLf & _
                   "¬сего листов в экземпл€ре: " & всего & vbCrLf & vbCrLf & _
                   "ѕечатать три экземпл€ра?" & vbCrLf & _
                   "ƒвусторонний режим включаетс€ в свойствах принтера Ч из макроса " & _
                   "им управл€ть нельз€, это настройка драйвера.", _
                   vbYesNoCancel + vbQuestion, "ѕечать пакета")

    If ответ = vbCancel Then
        свод.Close SaveChanges:=False
        wdApp.Quit
        Exit Sub
    End If

    —охранить—вод свод, ws, r
    If ответ = vbYes Then
        ' Copies:=3 печатает три отдельных экземпл€ра, и нумераци€ в
        ' каждом начинаетс€ заново Ч именно это и нужно.
        свод.PrintOut Copies:=3, Collate:=True
        ЌапечататьЅирку wdApp, всего, ws, r
    End If

    wdApp.Visible = True
    MsgBox "√отово." & vbCrLf & vbCrLf & _
           IIf(страниц—канов > 0, "—каны-картинки вошли в пакет и пронумерованы вместе " & _
               "со всем остальным." & vbCrLf & vbCrLf, "") & _
           IIf(путьPDF <> "", "PDF-сканы объединены в один файл Ч откройте его в Acrobat, " & _
               "напечатайте (тоже 3 экз., двусторонне) и вложите между протоколами и " & _
               "пустым листом: " & путьPDF & _
               IIf(ошибкаPDF <> "", " (" & ошибкаPDF & ")", "") & vbCrLf & vbCrLf, "") & _
           IIf(путьPDF = "" And ошибкаPDF <> "", "PDF-сканы найдены, но Acrobat их не " & _
               "объединил: " & ошибкаPDF & vbCrLf & "¬ложите их по одному, как раньше." & _
               vbCrLf & vbCrLf, "") & _
           IIf(страниц—канов = 0 And путьPDF = "" And ошибкаPDF = "", _
               "—канов не нашлось Ч проверьте папку сканов в настройках.", ""), _
           vbInformation, "ѕечать пакета"
    Exit Sub

—бой:
    Application.StatusBar = False
    On Error Resume Next
    If Not свод Is Nothing Then свод.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Ќе удалось собрать пакет." & vbCrLf & "ќшибка: " & Err.Description, _
           vbCritical, "ѕечать пакета"
End Sub


' -- —бор файлов -------------------------------------------------------

' ¬се строки одного акта. Ѕез номера акта Ч только текуща€ строка.
Public Function —трокијкта(ws As Worksheet, r As Long, акт As String) As Collection
    Dim список As Collection
    Dim i As Long, last As Long

    Set список = New Collection
    Set —трокијкта = список
    If акт = "" Then
        список.Add r
        Exit Function
    End If

    last = ѕоследн€€—трока(ws, ColReq(ws, "ƒата контрол€"))
    For i = ѕ≈–¬јя_—“–ќ ј To last
        If StrComp(Ќомерјкта(ws, i), акт, vbTextCompare) = 0 Then список.Add i
    Next i
End Function


' ѕор€док в пакете: сначала акт ¬»  по всем объектам, затем протоколы
' остальных методов Ч как задано в вашем пор€дке сшива.
Private Function —обрать‘айлыјкта(ws As Worksheet, строки As Collection) As Collection
    Dim список As Collection
    Dim методы As Variant
    Dim i As Long, k As Long
    Dim метод As String
    Dim путь As String

    Set список = New Collection
    Set —обрать‘айлыјкта = список

    методы = Split(ћ≈“ќƒџ, ",")
    For i = LBound(методы) To UBound(методы)
        метод = CStr(методы(i))
        For k = 1 To строки.Count
            If ћетодѕрименЄн(ws, CLng(строки(k)), метод) Then
                путь = Ќайтиƒокумент(ws, CLng(строки(k)), метод)
                If путь <> "" Then список.Add путь
            End If
        Next k
    Next i
End Function


Private Function Ќайтиƒокумент(ws As Worksheet, r As Long, метод As String) As String
    Dim папка As String
    Dim маска As String
    Dim файл As String
    Dim дата As Variant
    Dim год As Long

    дата = ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))
    папка = ѕапка√отовых() & CStr(год) & Application.PathSeparator & метод & Application.PathSeparator
    If Not ѕапка—уществует(папка) Then Exit Function

    маска = ќчиститьƒл€»мени‘айла(–егЌомер(ws, r, метод)) & "*.docx"
    файл = ѕервый‘айл(папка & маска)
    If файл <> "" Then Ќайтиƒокумент = папка & файл
End Function


' -- —борка документа --------------------------------------------------

Private Function —клеитьƒокументы(wdApp As Object, файлы As Collection) As Object
    Dim свод As Object
    Dim i As Long

    Set свод = wdApp.Documents.Add
    For i = 1 To файлы.Count
        If i > 1 Then
            свод.Content.InsertParagraphAfter
            свод.Paragraphs(свод.Paragraphs.Count).Range.InsertBreak Type:=wdPageBreak
        End If
        свод.Content.InsertParagraphAfter
        свод.Paragraphs(свод.Paragraphs.Count).Range.InsertFile fileName:=CStr(файлы(i))
    Next i
    Set —клеитьƒокументы = свод
End Function


' ѕоследний лист пакета Ч абсолютно пустой, кроме номера страницы:
' он нужен, чтобы прошить пакет и наклеить бирку.
Private Sub ƒобавитьѕустойЋист(свод As Object)
    свод.Content.InsertParagraphAfter
    свод.Paragraphs(свод.Paragraphs.Count).Range.InsertBreak Type:=wdPageBreak
    свод.Content.InsertParagraphAfter
End Sub


Private Function ѕоЋистам() As Boolean
    Dim s As String
    s = Ќастройка("Ќумеровать")
    ѕоЋистам = (s = "" Or StrComp(Left$(s, 4), "Ћист", vbTextCompare) = 0)
End Function


' Ќомер по центру нижнего колонтитула, на первой странице номера нет.
'
' ѕо листам Ч это нумераци€ подшивки, а не документа: при двусторонней
' печати лист бумаги несЄт две страницы, а номер у него один. ѕоэтому
' номер ставитс€ только на нечЄтных (лицевых) страницах и считает не
' страницы, а листы: 1, 3, 5-€ страницы получают 1, 2, 3.
'
' —читает это поле {=INT((PAGE+1)/2)} Ч вложенное поле внутри формулы.
' —обрать его из VBA одной строкой нельз€, поэтому собираетс€ в два
' приЄма и результат провер€етс€: если Word его не пон€л, ставитс€
' обычна€ нумераци€ страниц и об этом говоритс€ вслух Ч молча
' напечатать пакет не с той нумерацией хуже, чем сказать.
Private Sub ѕроставитьЌомера(свод As Object, режимЋистов As Boolean)
    Dim s As Object
    Dim получилось As Boolean

    получилось = режимЋистов
    If режимЋистов Then
        For Each s In свод.Sections
            s.PageSetup.DifferentFirstPageHeaderFooter = True
            s.PageSetup.OddAndEvenPagesHeaderFooter = True
            ' Primary при включЄнных чЄтных/нечЄтных Ч это нечЄтные, то
            ' есть лицевые стороны листов. „Єтный колонтитул пуст.
            If Not ЌомерЋиста(s.Footers(wdHeaderFooterPrimary)) Then
                получилось = False
                Exit For
            End If
        Next s
    End If

    If получилось Then Exit Sub

    For Each s In свод.Sections
        s.PageSetup.DifferentFirstPageHeaderFooter = True
        s.PageSetup.OddAndEvenPagesHeaderFooter = False
        With s.Footers(wdHeaderFooterPrimary)
            .Range.Text = ""
            .PageNumbers.RestartNumberingAtSection = False
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=False
        End With
    Next s

    If режимЋистов Then
        MsgBox "Ќумераци€ по листам не собралась Ч Word не прин€л поле " & _
               "{=INT((PAGE+1)/2)}." & vbCrLf & vbCrLf & _
               "ѕакет пронумерован по страницам: печатать так можно, но это " & _
               "не то, что просили. —кажите Ч разберЄмс€, почему Word отказалс€.", _
               vbExclamation, "ѕечать пакета"
    End If
End Sub


' ¬кладывает поле PAGE внутрь формулы. True Ч получилось.
Private Function ЌомерЋиста(колонтитул As Object) As Boolean
    Dim rng As Object
    Dim поле As Object
    Dim код As Object
    Dim результат As String

    On Error GoTo Ќет
    Set rng = колонтитул.Range
    rng.Text = ""
    Set rng = колонтитул.Range
    Set поле = rng.Fields.Add(rng, wdFieldEmpty, "=INT((—“–јЌ»÷ј+1)/2)", False)

    Set код = поле.Code
    With код.Find
        .Text = "—“–јЌ»÷ј"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceOne
    End With
    ' ѕосле замены диапазон схлопнулс€ там, где сто€ло слово, Ч туда и
    ' вкладываетс€ насто€щее поле PAGE.
    код.Fields.Add код, wdFieldPage, , False
    поле.Update

    колонтитул.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    результат = Trim$(поле.Result.Text)
    ЌомерЋиста = (результат <> "" And IsNumeric(результат))
    Exit Function
Ќет:
    колонтитул.Range.Text = ""
End Function


Private Sub —охранить—вод(свод As Object, ws As Worksheet, r As Long)
    Dim папка As String
    Dim им€ As String
    Dim дата As Variant
    Dim год As Long

    On Error Resume Next
    дата = ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))
    папка = ѕапка√отовых() & CStr(год) & Application.PathSeparator & "ѕакеты" & Application.PathSeparator
    MkPath папка
    им€ = Ќомерјкта(ws, r)
    If им€ = "" Then
        им€ = CStr(ws.Cells(r, ColReq(ws, "ID записи")).Value) & " " & _
              Left$(CStr(ws.Cells(r, ColReq(ws, "Ќаименование объекта в документе")).Value), 50)
    End If
    им€ = ќчиститьƒл€»мени‘айла(им€)
    свод.SaveAs2 fileName:=папка & им€ & ".docx"
End Sub


' -- Ѕирка -------------------------------------------------------------

' „исло листов на бирке пишетс€ цифрами и прописью. ѕрописью его чаще
' всего и пишут с ошибкой, поэтому склон€етс€ автоматически.
Private Sub ЌапечататьЅирку(wdApp As Object, листов As Long, ws As Worksheet, r As Long)
    Dim док As Object
    Dim текст As String
    Dim шаблон As String

    On Error Resume Next
    шаблон = Ќастройка("“екст бирки")
    If шаблон = "" Then
        шаблон = "ѕрошнуровано, пронумеровано и скреплено печатью" & vbCrLf & _
                 "#„»—Ћќ (#ѕ–ќѕ»—№ё) #Ћ»—“ќ¬" & vbCrLf & vbCrLf & _
                 "ƒолжность ______________________" & vbCrLf & vbCrLf & _
                 "ѕодпись, ‘»ќ ___________________" & vbCrLf & vbCrLf & _
                 "Ђ____ї ________________ 20____ г."
    End If

    текст = Replace(шаблон, "#„»—Ћќ", CStr(листов))
    текст = Replace(текст, "#ѕ–ќѕ»—№ё", „ислоѕрописью(листов))
    текст = Replace(текст, "#Ћ»—“ќ¬", —клонениеЋистов(листов))
    текст = Replace(текст, "#ќЅЏ≈ “", CStr(ws.Cells(r, ColReq(ws, "Ќаименование объекта в документе")).Value))

    Set док = wdApp.Documents.Add
    док.Content.Text = текст
    док.Content.ParagraphFormat.SpaceAfter = 6
    док.PrintOut Copies:=3
End Sub


Public Function „ислоѕрописью(n As Long) As String
    Dim сотни As Variant
    Dim дес€тки As Variant
    Dim до20 As Variant
    Dim с As Long, д As Long, е As Long
    Dim s As String

    If n <= 0 Then Exit Function
    If n > 999 Then
        „ислоѕрописью = CStr(n)          ' пакетов такой толщины не бывает
        Exit Function
    End If

    сотни = Array("", "сто", "двести", "триста", "четыреста", "п€тьсот", _
                  "шестьсот", "семьсот", "восемьсот", "дев€тьсот")
    дес€тки = Array("", "", "двадцать", "тридцать", "сорок", "п€тьдес€т", _
                    "шестьдес€т", "семьдес€т", "восемьдес€т", "дев€носто")
    до20 = Array("", "один", "два", "три", "четыре", "п€ть", "шесть", "семь", _
                 "восемь", "дев€ть", "дес€ть", "одиннадцать", "двенадцать", _
                 "тринадцать", "четырнадцать", "п€тнадцать", "шестнадцать", _
                 "семнадцать", "восемнадцать", "дев€тнадцать")

    с = n \ 100
    д = (n - с * 100) \ 10
    е = n - с * 100 - д * 10

    If с > 0 Then s = сотни(с)
    If д = 1 Then
        s = Trim$(s & " " & до20(10 + е))
    Else
        If д > 1 Then s = Trim$(s & " " & дес€тки(д))
        If е > 0 Then s = Trim$(s & " " & до20(е))
    End If
    „ислоѕрописью = Trim$(s)
End Function


Public Function —клонениеЋистов(n As Long) As String
    Dim последн€€ As Long
    Dim две As Long
    последн€€ = n Mod 10
    две = n Mod 100
    If две >= 11 And две <= 14 Then
        —клонениеЋистов = "листов"
    ElseIf последн€€ = 1 Then
        —клонениеЋистов = "лист"
    ElseIf последн€€ >= 2 And последн€€ <= 4 Then
        —клонениеЋистов = "листа"
    Else
        —клонениеЋистов = "листов"
    End If
End Function


Private Function —просить„исло—канов() As Long
    Dim s As String
    s = InputBox("—колько листов вкладываетс€ в пакет ¬–”„Ќ”ё?" & vbCrLf & vbCrLf & _
                 "—каны-картинки (JPG, PNG, TIFF) программа вставл€ет в пакет сама Ч " & _
                 "их здесь считать не надо. –уками остаютс€ только сканы в PDF: " & _
                 "объединить PDF без Acrobat нельз€." & vbCrLf & vbCrLf & _
                 "Ёто число нужно, чтобы на бирке сошлось количество листов.", _
                 "ѕечать пакета", "0")
    If StrPtr(s) = 0 Then                ' нажали Ђќтменаї
        —просить„исло—канов = -1
        Exit Function
    End If
    If Trim$(s) = "" Then Exit Function
    If Not IsNumeric(s) Then
        MsgBox "Ќужно число.", vbExclamation, "ѕечать пакета"
        —просить„исло—канов = -1
        Exit Function
    End If
    —просить„исло—канов = CLng(s)
End Function
