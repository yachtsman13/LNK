Attribute VB_Name = "modPackage"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Сборка и печать пакета документов по одной работе.
'
'  Порядок в пакете задан вами и не меняется:
'    акт → протоколы → свидетельства лаборатории → квалификационные
'    удостоверения → свидетельства о поверке → пустой лист для сшива.
'
'  Сквозная нумерация внутри экземпляра, на первой странице номера нет.
'  Три экземпляра: каждый нумеруется заново, поэтому печатаются копии
'  одного документа, а не три склеенных подряд.
'
'  Свидетельства и удостоверения — сканы. С одним Adobe Reader свести их
'  в общий файл со сквозной нумерацией нельзя: программно объединять PDF
'  он не умеет, а установить что-то ещё нельзя. Поэтому их страницы
'  учитываются в нумерации и на бирке, а печатаются и вкладываются
'  отдельно. Это единственное место во всём процессе, где остаётся
'  ручная операция.
' ═══════════════════════════════════════════════════════════════════════

Private Const wdPageBreak As Long = 7
Private Const wdStory As Long = 6
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignPageNumberCenter As Long = 1
Private Const wdFieldNumPages As Long = 26
Private Const wdStatisticPages As Long = 2


Public Sub НапечататьПакет()
    Dim ws As Worksheet
    Dim r As Long
    Dim файлы As Collection
    Dim строки As Collection
    Dim акт As String
    Dim сканов As Long
    Dim wdApp As Object
    Dim свод As Object
    Dim страницСвода As Long
    Dim всего As Long
    Dim ответ As VbMsgBoxResult

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ws)
    If r = 0 Then Exit Sub

    ' Пакет собирается по АКТУ целиком: акт один на работу, а объектов
    ' в ней несколько, и протоколы по всем из них идут в один пакет.
    ' Если акта нет — печатается пакет по одной строке.
    акт = НомерАкта(ws, r)
    Set строки = СтрокиАкта(ws, r, акт)
    Set файлы = СобратьФайлыАкта(ws, строки)
    If файлы.Count = 0 Then
        MsgBox "Нет сформированных документов" & _
               IIf(акт = "", " по этой строке.", " по акту " & акт & ".") & vbCrLf & vbCrLf & _
               "Сначала «Сформировать документы» — они появятся в папке «Готовые». " & _
               "Если в акте несколько объектов, документы нужны по каждому.", _
               vbExclamation, "Печать пакета"
        Exit Sub
    End If

    сканов = СпроситьЧислоСканов()
    If сканов < 0 Then Exit Sub

    Application.StatusBar = "Собираю пакет…"
    On Error GoTo Сбой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set свод = СклеитьДокументы(wdApp, файлы)

    ДобавитьПустойЛист свод
    страницСвода = свод.ComputeStatistics(wdStatisticPages)
    ПроставитьНомера свод, сканов
    всего = страницСвода + сканов

    Application.StatusBar = False
    ответ = MsgBox("Пакет собран." & vbCrLf & vbCrLf & _
                   IIf(акт = "", "Без акта, одна строка журнала." & vbCrLf, _
                       акт & ", объектов в нём: " & строки.Count & vbCrLf) & _
                   "Документов Word: " & файлы.Count & vbCrLf & _
                   "Страниц в них (с пустым листом): " & страницСвода & vbCrLf & _
                   "Страниц сканов, вкладываемых вручную: " & сканов & vbCrLf & _
                   "Всего листов в экземпляре: " & всего & vbCrLf & vbCrLf & _
                   "Печатать три экземпляра?" & vbCrLf & _
                   "Двусторонний режим включается в свойствах принтера — из макроса " & _
                   "им управлять нельзя, это настройка драйвера.", _
                   vbYesNoCancel + vbQuestion, "Печать пакета")

    If ответ = vbCancel Then
        свод.Close SaveChanges:=False
        wdApp.Quit
        Exit Sub
    End If

    СохранитьСвод свод, ws, r
    If ответ = vbYes Then
        ' Copies:=3 печатает три отдельных экземпляра, и нумерация в
        ' каждом начинается заново — именно это и нужно.
        свод.PrintOut Copies:=3, Collate:=True
        НапечататьБирку wdApp, всего, ws, r
    End If

    wdApp.Visible = True
    MsgBox "Готово." & vbCrLf & vbCrLf & _
           "Осталось вручную: напечатать сканы (свидетельства лаборатории, " & _
           "удостоверения, свидетельства о поверке) и вложить их в пакет " & _
           "между протоколами и пустым листом.", vbInformation, "Печать пакета"
    Exit Sub

Сбой:
    Application.StatusBar = False
    On Error Resume Next
    If Not свод Is Nothing Then свод.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Не удалось собрать пакет." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Печать пакета"
End Sub


' ── Сбор файлов ───────────────────────────────────────────────────────

' Все строки одного акта. Без номера акта — только текущая строка.
Private Function СтрокиАкта(ws As Worksheet, r As Long, акт As String) As Collection
    Dim список As Collection
    Dim i As Long, last As Long

    Set список = New Collection
    Set СтрокиАкта = список
    If акт = "" Then
        список.Add r
        Exit Function
    End If

    last = ПоследняяСтрока(ws, ColReq(ws, "Дата контроля"))
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(НомерАкта(ws, i), акт, vbTextCompare) = 0 Then список.Add i
    Next i
End Function


' Порядок в пакете: сначала акт ВИК по всем объектам, затем протоколы
' остальных методов — как задано в вашем порядке сшива.
Private Function СобратьФайлыАкта(ws As Worksheet, строки As Collection) As Collection
    Dim список As Collection
    Dim методы As Variant
    Dim i As Long, k As Long
    Dim метод As String
    Dim путь As String

    Set список = New Collection
    Set СобратьФайлыАкта = список

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        метод = CStr(методы(i))
        For k = 1 To строки.Count
            If МетодПрименён(ws, CLng(строки(k)), метод) Then
                путь = НайтиДокумент(ws, CLng(строки(k)), метод)
                If путь <> "" Then список.Add путь
            End If
        Next k
    Next i
End Function


Private Function НайтиДокумент(ws As Worksheet, r As Long, метод As String) As String
    Dim папка As String
    Dim маска As String
    Dim файл As String
    Dim дата As Variant
    Dim год As Long

    дата = ws.Cells(r, ColReq(ws, "Дата контроля")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))
    папка = ПапкаГотовых() & CStr(год) & Application.PathSeparator & метод & Application.PathSeparator
    If Not ПапкаСуществует(папка) Then Exit Function

    маска = ОчиститьДляИмениФайла(РегНомер(ws, r, метод)) & "*.docx"
    файл = ПервыйФайл(папка & маска)
    If файл <> "" Then НайтиДокумент = папка & файл
End Function


' ── Сборка документа ──────────────────────────────────────────────────

Private Function СклеитьДокументы(wdApp As Object, файлы As Collection) As Object
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
    Set СклеитьДокументы = свод
End Function


' Последний лист пакета — абсолютно пустой, кроме номера страницы:
' он нужен, чтобы прошить пакет и наклеить бирку.
Private Sub ДобавитьПустойЛист(свод As Object)
    свод.Content.InsertParagraphAfter
    свод.Paragraphs(свод.Paragraphs.Count).Range.InsertBreak Type:=wdPageBreak
    свод.Content.InsertParagraphAfter
End Sub


' Номера страниц по центру нижнего колонтитула, на первой странице
' номера нет — как в ваших правилах оформления.
Private Sub ПроставитьНомера(свод As Object, сканов As Long)
    Dim s As Object
    For Each s In свод.Sections
        s.PageSetup.DifferentFirstPageHeaderFooter = True
        With s.Footers(wdHeaderFooterPrimary)
            .PageNumbers.RestartNumberingAtSection = False
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=False
        End With
    Next s
End Sub


Private Sub СохранитьСвод(свод As Object, ws As Worksheet, r As Long)
    Dim папка As String
    Dim имя As String
    Dim дата As Variant
    Dim год As Long

    On Error Resume Next
    дата = ws.Cells(r, ColReq(ws, "Дата контроля")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))
    папка = ПапкаГотовых() & CStr(год) & Application.PathSeparator & "Пакеты" & Application.PathSeparator
    MkPath папка
    имя = НомерАкта(ws, r)
    If имя = "" Then
        имя = CStr(ws.Cells(r, ColReq(ws, "ID записи")).Value) & " " & _
              Left$(CStr(ws.Cells(r, ColReq(ws, "Наименование объекта в документе")).Value), 50)
    End If
    имя = ОчиститьДляИмениФайла(имя)
    свод.SaveAs2 fileName:=папка & имя & ".docx"
End Sub


' ── Бирка ─────────────────────────────────────────────────────────────

' Число листов на бирке пишется цифрами и прописью. Прописью его чаще
' всего и пишут с ошибкой, поэтому склоняется автоматически.
Private Sub НапечататьБирку(wdApp As Object, листов As Long, ws As Worksheet, r As Long)
    Dim док As Object
    Dim текст As String
    Dim шаблон As String

    On Error Resume Next
    шаблон = Настройка("Текст бирки")
    If шаблон = "" Then
        шаблон = "Прошнуровано, пронумеровано и скреплено печатью" & vbCrLf & _
                 "#ЧИСЛО (#ПРОПИСЬЮ) #ЛИСТОВ" & vbCrLf & vbCrLf & _
                 "Должность ______________________" & vbCrLf & vbCrLf & _
                 "Подпись, ФИО ___________________" & vbCrLf & vbCrLf & _
                 "«____» ________________ 20____ г."
    End If

    текст = Replace(шаблон, "#ЧИСЛО", CStr(листов))
    текст = Replace(текст, "#ПРОПИСЬЮ", ЧислоПрописью(листов))
    текст = Replace(текст, "#ЛИСТОВ", СклонениеЛистов(листов))
    текст = Replace(текст, "#ОБЪЕКТ", CStr(ws.Cells(r, ColReq(ws, "Наименование объекта в документе")).Value))

    Set док = wdApp.Documents.Add
    док.Content.Text = текст
    док.Content.ParagraphFormat.SpaceAfter = 6
    док.PrintOut Copies:=3
End Sub


Public Function ЧислоПрописью(n As Long) As String
    Dim сотни As Variant
    Dim десятки As Variant
    Dim до20 As Variant
    Dim с As Long, д As Long, е As Long
    Dim s As String

    If n <= 0 Then Exit Function
    If n > 999 Then
        ЧислоПрописью = CStr(n)          ' пакетов такой толщины не бывает
        Exit Function
    End If

    сотни = Array("", "сто", "двести", "триста", "четыреста", "пятьсот", _
                  "шестьсот", "семьсот", "восемьсот", "девятьсот")
    десятки = Array("", "", "двадцать", "тридцать", "сорок", "пятьдесят", _
                    "шестьдесят", "семьдесят", "восемьдесят", "девяносто")
    до20 = Array("", "один", "два", "три", "четыре", "пять", "шесть", "семь", _
                 "восемь", "девять", "десять", "одиннадцать", "двенадцать", _
                 "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", _
                 "семнадцать", "восемнадцать", "девятнадцать")

    с = n \ 100
    д = (n - с * 100) \ 10
    е = n - с * 100 - д * 10

    If с > 0 Then s = сотни(с)
    If д = 1 Then
        s = Trim$(s & " " & до20(10 + е))
    Else
        If д > 1 Then s = Trim$(s & " " & десятки(д))
        If е > 0 Then s = Trim$(s & " " & до20(е))
    End If
    ЧислоПрописью = Trim$(s)
End Function


Public Function СклонениеЛистов(n As Long) As String
    Dim последняя As Long
    Dim две As Long
    последняя = n Mod 10
    две = n Mod 100
    If две >= 11 And две <= 14 Then
        СклонениеЛистов = "листов"
    ElseIf последняя = 1 Then
        СклонениеЛистов = "лист"
    ElseIf последняя >= 2 And последняя <= 4 Then
        СклонениеЛистов = "листа"
    Else
        СклонениеЛистов = "листов"
    End If
End Function


Private Function СпроситьЧислоСканов() As Long
    Dim s As String
    s = InputBox("Сколько страниц сканов вкладывается в пакет вручную?" & vbCrLf & vbCrLf & _
                 "Считаются свидетельства лаборатории, квалификационные удостоверения " & _
                 "исполнителей и свидетельства о поверке применённого оборудования." & vbCrLf & vbCrLf & _
                 "Это число нужно, чтобы на бирке сошлось количество листов.", _
                 "Печать пакета", "0")
    If StrPtr(s) = 0 Then                ' нажали «Отмена»
        СпроситьЧислоСканов = -1
        Exit Function
    End If
    If Trim$(s) = "" Then Exit Function
    If Not IsNumeric(s) Then
        MsgBox "Нужно число.", vbExclamation, "Печать пакета"
        СпроситьЧислоСканов = -1
        Exit Function
    End If
    СпроситьЧислоСканов = CLng(s)
End Function
