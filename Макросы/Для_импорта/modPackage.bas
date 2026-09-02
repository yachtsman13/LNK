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
'  —видетельства и удостоверени€ Ч сканы. — одним Adobe Reader свести их
'  в общий файл со сквозной нумерацией нельз€: программно объедин€ть PDF
'  он не умеет, а установить что-то ещЄ нельз€. ѕоэтому их страницы
'  учитываютс€ в нумерации и на бирке, а печатаютс€ и вкладываютс€
'  отдельно. Ёто единственное место во всЄм процессе, где остаЄтс€
'  ручна€ операци€.
' =======================================================================

Private Const wdPageBreak As Long = 7
Private Const wdStory As Long = 6
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignPageNumberCenter As Long = 1
Private Const wdFieldNumPages As Long = 26
Private Const wdStatisticPages As Long = 2


Public Sub Ќапечататьѕакет()
    Dim ws As Worksheet
    Dim r As Long
    Dim файлы As Collection
    Dim сканов As Long
    Dim wdApp As Object
    Dim свод As Object
    Dim страниц—вода As Long
    Dim всего As Long
    Dim ответ As VbMsgBoxResult

    Set ws = Ћист(Ћ_∆”–ЌјЋ)
    If ws Is Nothing Then Exit Sub
    r = “екуща€—трока(ws)
    If r = 0 Then Exit Sub

    Set файлы = —обрать‘айлы–аботы(ws, r)
    If файлы.Count = 0 Then
        MsgBox "ѕо этой строке нет сформированных документов." & vbCrLf & vbCrLf & _
               "—начала Ђ—формировать документыї Ч они по€в€тс€ в папке Ђ√отовыеї.", _
               vbExclamation, "ѕечать пакета"
        Exit Sub
    End If

    сканов = —просить„исло—канов()
    If сканов < 0 Then Exit Sub

    Application.StatusBar = "—обираю пакетЕ"
    On Error GoTo —бой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set свод = —клеитьƒокументы(wdApp, файлы)

    ƒобавитьѕустойЋист свод
    страниц—вода = свод.ComputeStatistics(wdStatisticPages)
    ѕроставитьЌомера свод, сканов
    всего = страниц—вода + сканов

    Application.StatusBar = False
    ответ = MsgBox("ѕакет собран." & vbCrLf & vbCrLf & _
                   "ƒокументов Word: " & файлы.Count & vbCrLf & _
                   "—траниц в них (с пустым листом): " & страниц—вода & vbCrLf & _
                   "—траниц сканов, вкладываемых вручную: " & сканов & vbCrLf & _
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
           "ќсталось вручную: напечатать сканы (свидетельства лаборатории, " & _
           "удостоверени€, свидетельства о поверке) и вложить их в пакет " & _
           "между протоколами и пустым листом.", vbInformation, "ѕечать пакета"
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

' јкт ¬»  идЄт первым, остальные протоколы Ч в пор€дке методов.
Private Function —обрать‘айлы–аботы(ws As Worksheet, r As Long) As Collection
    Dim список As Collection
    Dim методы As Variant
    Dim i As Long
    Dim метод As String
    Dim путь As String

    Set список = New Collection
    Set —обрать‘айлы–аботы = список

    методы = Split(ћ≈“ќƒџ, ",")
    For i = LBound(методы) To UBound(методы)
        метод = CStr(методы(i))
        If ћетодѕрименЄн(ws, r, метод) Then
            путь = Ќайтиƒокумент(ws, r, метод)
            If путь <> "" Then список.Add путь
        End If
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


' Ќомера страниц по центру нижнего колонтитула, на первой странице
' номера нет Ч как в ваших правилах оформлени€.
Private Sub ѕроставитьЌомера(свод As Object, сканов As Long)
    Dim s As Object
    For Each s In свод.Sections
        s.PageSetup.DifferentFirstPageHeaderFooter = True
        With s.Footers(wdHeaderFooterPrimary)
            .PageNumbers.RestartNumberingAtSection = False
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=False
        End With
    Next s
End Sub


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
    им€ = ќчиститьƒл€»мени‘айла(CStr(ws.Cells(r, ColReq(ws, "ID записи")).Value)) & " " & _
          ќчиститьƒл€»мени‘айла(Left$(CStr(ws.Cells(r, ColReq(ws, "Ќаименование объекта в документе")).Value), 50))
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
    s = InputBox("—колько страниц сканов вкладываетс€ в пакет вручную?" & vbCrLf & vbCrLf & _
                 "—читаютс€ свидетельства лаборатории, квалификационные удостоверени€ " & _
                 "исполнителей и свидетельства о поверке применЄнного оборудовани€." & vbCrLf & vbCrLf & _
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
