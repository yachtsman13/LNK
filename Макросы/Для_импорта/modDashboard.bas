Attribute VB_Name = "modDashboard"
Option Explicit

' =======================================================================
'  Дашборд: что требует внимания сегодня.
'
'  Правило одно — плитка показывается, только если по ней есть что
'  делать. Дашборд, на котором всегда всё зелёное, перестают читать
'  через неделю.
' =======================================================================

Private Const Л_ДАШБОРД As String = "Дашборд"

Private Const ЦВ_ОК As Long = 12379352        ' спокойный зелёный
Private Const ЦВ_ВНИМАНИЕ As Long = 10284031  ' жёлтый
Private Const ЦВ_СРОЧНО As Long = 13551615    ' красный
Private Const ЦВ_ТЕКСТ_ОК As Long = 25600
Private Const ЦВ_ТЕКСТ_ВН As Long = 26012
Private Const ЦВ_ТЕКСТ_СР As Long = 2500


Public Sub ПостроитьДашборд()
    Dim ws As Worksheet
    Dim r As Long

    Application.ScreenUpdating = False
    On Error GoTo Сбой
    Set ws = ПересоздатьДашборд()

    ws.Range("B2").Value = "ЧТО ТРЕБУЕТ ВНИМАНИЯ"
    ws.Range("B3").Value = "на " & Format$(Date, "dd.mm.yyyy") & _
                           IIf(КтоМастер() = "", "", "  ·  мастер: " & КтоМастер())
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(31, 56, 100)
    ws.Range("B3").Font.Italic = True
    ws.Range("B3").Font.Color = RGB(89, 89, 89)

    ' Каждая плитка со своей защитой: не собралась одна — Дашборд
    ' всё равно показывает остальные. Наполовину собранный Дашборд
    ' полезнее, чем сообщение об ошибке вместо него.
    r = 5
    On Error Resume Next
    r = ПлиткаАттестацияЛНК(ws, r)
    r = ПлиткаСрокиПоЛисту(ws, r, Л_АТТЕСТАЦИИ, "Действительно до", "Удостоверения специалистов", _
                           "ФИО", "Метод контроля", 90)
    r = ПлиткаСрокиПоЛисту(ws, r, Л_СИ, "Действительно до", "Поверка средств измерений", _
                           "Наименование", "Заводской №", 90)
    r = ПлиткаСрокиПоЛисту(ws, r, "НД", "Дата следующей проверки", "Проверка актуальности НД", _
                           "Наименование документа", "Обозначение", 60)
    r = ПлиткаЧерновиков(ws, r)
    r = ПлиткаВыданного(ws, r)
    r = ПлиткаРекламаций(ws, r)
    r = ПлиткаОбмена(ws, r)
    If Err.Number <> 0 Then
        r = Плитка(ws, r, "Часть плиток не собралась", "сбой", _
                   "Ошибка " & Err.Number & ": " & Err.Description & _
                   ". Остальное на Дашборде верно.", 1)
        Err.Clear
    End If
    On Error GoTo Сбой

    If r = 5 Then
        ws.Cells(r, 2).Value = "Просроченного и ближайшего нет — всё в сроках."
        ws.Cells(r, 2).Font.Color = RGB(0, 97, 0)
        ws.Cells(r, 2).Font.Size = 12
    End If

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 46
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 76
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    Exit Sub
Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось построить Дашборд." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Дашборд"
End Sub


' -- Плитки ------------------------------------------------------------

' Общая плитка «сроки по колонке с датой»: считает просроченное и
' ближайшее, перечисляет до пяти позиций поимённо.
Private Function ПлиткаСрокиПоЛисту(ws As Worksheet, r As Long, имяЛиста As String, _
                                    колонкаДаты As String, заголовок As String, _
                                    поле1 As String, поле2 As String, порог As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long
    Dim cДата As Long, c1 As Long, c2 As Long, cСост As Long
    Dim просрочено As Long, скоро As Long
    Dim список As String
    Dim д As Date
    Dim осталось As Long

    ПлиткаСрокиПоЛисту = r
    Set src = ЛистТихо(имяЛиста)
    If src Is Nothing Then Exit Function
    cДата = ColByHeader(src, колонкаДаты)
    If cДата = 0 Then Exit Function
    c1 = ColByHeader(src, поле1)
    c2 = ColByHeader(src, поле2)
    cСост = ColByHeader(src, "Состояние")
    last = ПоследняяСтрока(src, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If cСост = 0 Or StrComp(CStr(src.Cells(i, cСост).Value), "Списан", vbTextCompare) <> 0 Then
            If IsDate(src.Cells(i, cДата).Value) Then
                д = CDate(src.Cells(i, cДата).Value)
                осталось = DateDiff("d", Date, д)
                If осталось < 0 Then
                    просрочено = просрочено + 1
                ElseIf осталось <= порог Then
                    скоро = скоро + 1
                End If
                If осталось <= порог Then
                    If Len(список) < 400 Then
                        список = список & IIf(список = "", "", "; ") & _
                                 CStr(src.Cells(i, c1).Value) & _
                                 IIf(c2 = 0, "", " " & CStr(src.Cells(i, c2).Value)) & _
                                 " — " & IIf(осталось < 0, "просрочено " & -осталось & " дн.", _
                                             "осталось " & осталось & " дн.")
                    End If
                End If
            End If
        End If
    Next i

    If просрочено = 0 And скоро = 0 Then Exit Function
    ПлиткаСрокиПоЛисту = Плитка(ws, r, заголовок, _
        IIf(просрочено > 0, просрочено & " просрочено", скоро & " истекает"), _
        список, IIf(просрочено > 0, 2, 1))
End Function


' Срок аттестации самой лаборатории. Пороги здесь другие — переаттестация
' готовится месяцами, поэтому предупреждаем за полгода.
Private Function ПлиткаАттестацияЛНК(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long
    Dim значение As String
    Dim д As Date
    Dim осталось As Long

    ПлиткаАттестацияЛНК = r
    Set src = ЛистТихо("Паспорт_постоянная")
    If src Is Nothing Then Exit Function
    last = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    For i = ПЕРВАЯ_СТРОКА To last
        If InStr(1, CStr(src.Cells(i, 1).Value), "Свидетельство действительно до", vbTextCompare) > 0 Then
            значение = Trim$(CStr(src.Cells(i, 2).Value))
            Exit For
        End If
    Next i

    If значение = "" Then
        ПлиткаАттестацияЛНК = Плитка(ws, r, "Аттестация лаборатории", "срок не заполнен", _
            "Впишите срок свидетельства на листе «Паспорт_постоянная» — это самый дорогой " & _
            "по последствиям срок во всей системе: с его истечением лаборатория " & _
            "не может работать вообще.", 1)
        Exit Function
    End If
    If Not IsDate(значение) Then Exit Function

    д = CDate(значение)
    осталось = DateDiff("d", Date, д)
    If осталось > 180 Then Exit Function
    ПлиткаАттестацияЛНК = Плитка(ws, r, "Аттестация лаборатории", _
        IIf(осталось < 0, "ПРОСРОЧЕНА", осталось & " дн."), _
        "Свидетельство действительно до " & Format$(д, "dd.mm.yyyy") & _
        ". Переаттестация готовится месяцами — начинать пора.", _
        IIf(осталось <= 90, 2, 1))
End Function


Private Function ПлиткаЧерновиков(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long, n As Long
    Dim cСтатус As Long, cДата As Long
    Dim старейший As Date

    ПлиткаЧерновиков = r
    Set src = ЛистТихо(Л_ЖУРНАЛ)
    If src Is Nothing Then Exit Function
    cСтатус = ColByHeader(src, "Статус")
    cДата = ColByHeader(src, "Дата контроля")
    If cСтатус = 0 Then Exit Function
    last = ПоследняяСтрока(src, cДата)

    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(CStr(src.Cells(i, cСтатус).Value), "Черновик", vbTextCompare) = 0 Then
            n = n + 1
            If IsDate(src.Cells(i, cДата).Value) Then
                If старейший = 0 Or CDate(src.Cells(i, cДата).Value) < старейший Then
                    старейший = CDate(src.Cells(i, cДата).Value)
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    ПлиткаЧерновиков = Плитка(ws, r, "Черновики в журнале", n & " шт.", _
        "Незарегистрированные записи: номера им не выданы, в бумажный журнал они не идут." & _
        IIf(старейший = 0, "", " Самая ранняя от " & Format$(старейший, "dd.mm.yyyy") & "."), _
        IIf(старейший > 0 And DateDiff("d", старейший, Date) > 14, 2, 1))
End Function


Private Function ПлиткаВыданного(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long, n As Long
    Dim cВыдачи As Long, cВозврат As Long, cКому As Long, cНаим As Long
    Dim список As String
    Dim дней As Long

    ПлиткаВыданного = r
    Set src = ЛистТихо("Выдача_оборудования")
    If src Is Nothing Then Exit Function
    cВыдачи = ColByHeader(src, "Дата выдачи")
    cВозврат = ColByHeader(src, "Дата возврата")
    cКому = ColByHeader(src, "Кому выдано")
    cНаим = ColByHeader(src, "Наименование")
    If cВыдачи = 0 Then Exit Function
    last = ПоследняяСтрока(src, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If IsDate(src.Cells(i, cВыдачи).Value) And Not IsDate(src.Cells(i, cВозврат).Value) Then
            n = n + 1
            дней = DateDiff("d", CDate(src.Cells(i, cВыдачи).Value), Date)
            If Len(список) < 300 Then
                список = список & IIf(список = "", "", "; ") & _
                         CStr(src.Cells(i, cНаим).Value) & " — " & _
                         CStr(src.Cells(i, cКому).Value) & ", " & дней & " дн."
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    ПлиткаВыданного = Плитка(ws, r, "Оборудование на руках", n & " шт.", список, 0)
End Function


Private Function ПлиткаРекламаций(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long
    Dim cСрок As Long, cОтметка As Long, cРекв As Long
    Dim просрочено As Long, вРаботе As Long
    Dim список As String

    ПлиткаРекламаций = r
    Set src = ЛистТихо("Рекламации")
    If src Is Nothing Then Exit Function
    cСрок = ColByHeader(src, "Срок выполнения")
    cОтметка = ColByHeader(src, "Отметка об выполненных мероприятиях")
    cРекв = ColByHeader(src, "Реквизиты документа")
    If cСрок = 0 Then Exit Function
    last = ПоследняяСтрока(src, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If Trim$(CStr(src.Cells(i, cОтметка).Value)) = "" Then
            If IsDate(src.Cells(i, cСрок).Value) Then
                вРаботе = вРаботе + 1
                If CDate(src.Cells(i, cСрок).Value) < Date Then
                    просрочено = просрочено + 1
                    If Len(список) < 300 Then
                        список = список & IIf(список = "", "", "; ") & _
                                 CStr(src.Cells(i, cРекв).Value) & " — срок был " & _
                                 Format$(src.Cells(i, cСрок).Value, "dd.mm.yyyy")
                    End If
                End If
            End If
        End If
    Next i

    If вРаботе = 0 Then Exit Function
    ПлиткаРекламаций = Плитка(ws, r, "Рекламации", _
        IIf(просрочено > 0, просрочено & " просрочено", вРаботе & " в работе"), _
        IIf(список = "", "Мероприятия ещё не отмечены выполненными.", список), _
        IIf(просрочено > 0, 2, 1))
End Function


Private Function ПлиткаОбмена(ws As Worksheet, r As Long) As Long
    Dim папка As String
    Dim файл As String
    Dim n As Long

    ПлиткаОбмена = r
    папка = ПапкаОбмена() & "Входящие" & Application.PathSeparator
    If Not ПапкаСуществует(папка) Then Exit Function
    файл = ПервыйФайл(папка & "*.csv")
    Do While файл <> ""
        n = n + 1
        файл = Dir
    Loop
    If n = 0 Then Exit Function
    If Not ЭтоМастер() Then Exit Function
    ПлиткаОбмена = Плитка(ws, r, "Дельты ждут сборки", n & " шт.", _
        "Пока свод не собран, остальные видят данные на момент прошлой сборки.", 1)
End Function


' -- Отрисовка ---------------------------------------------------------

Private Function Плитка(ws As Worksheet, r As Long, заголовок As String, _
                        число As String, пояснение As String, уровень As Long) As Long
    Dim фон As Long
    Dim цвет As Long

    Select Case уровень
        Case 2: фон = ЦВ_СРОЧНО: цвет = ЦВ_ТЕКСТ_СР
        Case 1: фон = ЦВ_ВНИМАНИЕ: цвет = ЦВ_ТЕКСТ_ВН
        Case Else: фон = ЦВ_ОК: цвет = ЦВ_ТЕКСТ_ОК
    End Select

    ws.Cells(r, 2).Value = заголовок
    ws.Cells(r, 3).Value = число
    ws.Cells(r, 4).Value = пояснение
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Interior.Color = фон
        .Font.Color = цвет
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
        .Borders.Weight = xlMedium
    End With
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Font.Bold = True
    ws.Cells(r, 3).HorizontalAlignment = xlCenter
    ws.Cells(r, 4).Font.Size = 9
    ws.Rows(r).RowHeight = 34
    Плитка = r + 1
End Function


Private Function ПересоздатьДашборд() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Л_ДАШБОРД).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
    ws.Name = Л_ДАШБОРД
    ws.Tab.Color = RGB(46, 84, 150)
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Cells.Font.Name = "Calibri"
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Set ПересоздатьДашборд = ws
End Function
