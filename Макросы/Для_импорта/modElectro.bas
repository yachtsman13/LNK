Attribute VB_Name = "modElectro"
Option Explicit

' =======================================================================
'  Журнал учёта, проверки и испытаний электроинструмента и
'  вспомогательного оборудования.
'
'  Форма утверждена и не наша, поэтому в печать идут ровно её 16 граф и
'  в том же порядке. Служебные колонки системы — связь с реестром СИ,
'  срок, статус — на рабочем листе стоят после них и в печатную форму
'  не попадают.
'
'  Периодичность по умолчанию 6 месяцев. Дата следующей проверки
'  считается формулой от даты последней, а не вводится руками: именно
'  на ручном вводе такие журналы и разъезжаются.
' =======================================================================

Private Const Л_ЭИ As String = "Электроинструмент"
Private Const Л_ПЕЧАТЬ As String = "Печать_Электроинструмент"
Private Const ГРАФ As Long = 16          ' столько граф в утверждённой форме


Public Sub ПостроитьЖурналЭлектроинструмента()
    Dim src As Worksheet
    Dim ws As Worksheet
    Dim r As Long, last As Long, out As Long, c As Long
    Dim год As Long
    Dim записей As Long

    Set src = Лист(Л_ЭИ)
    If src Is Nothing Then Exit Sub

    год = CLng(Val(InputBox("За какой год печатать журнал?" & vbCrLf & vbCrLf & _
                            "Берутся записи по дате последней проверки. " & _
                            "Пусто или 0 — все записи.", _
                            "Журнал электроинструмента", CStr(Year(Date)))))

    Application.ScreenUpdating = False
    On Error GoTo Сбой
    Set ws = ПересоздатьЛист()

    ws.Cells(1, 1).Value = "ЖУРНАЛ"
    ws.Cells(2, 1).Value = "учёта, проверки и испытаний электроинструмента и " & _
                           "вспомогательного оборудования к нему"
    ws.Cells(3, 1).Value = Настройка("Полное наименование ЛНК") & _
                           IIf(год = 0, "", "     за " & год & " год")
    ws.Cells(4, 1).Value = "Учётный номер журнала: " & _
                           IIf(Настройка("Учётный номер журнала электроинструмента") = "", _
                               "____________", Настройка("Учётный номер журнала электроинструмента"))
    ws.Range(ws.Cells(1, 1), ws.Cells(1, ГРАФ)).Merge
    ws.Range(ws.Cells(2, 1), ws.Cells(2, ГРАФ)).Merge
    ws.Range(ws.Cells(3, 1), ws.Cells(3, ГРАФ)).Merge
    ws.Range(ws.Cells(4, 1), ws.Cells(4, ГРАФ)).Merge
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(4, 1).Font.Italic = True
    ws.Range(ws.Cells(1, 1), ws.Cells(4, 1)).HorizontalAlignment = xlCenter

    ' Шапка — те же графы и в том же порядке, что на рабочем листе:
    ' колонки 2…17, то есть графы 1…16 утверждённой формы.
    For c = 1 To ГРАФ
        ws.Cells(6, c).Value = src.Cells(СТРОКА_ЗАГОЛОВКОВ, c + 1).Value
        ws.Columns(c).ColumnWidth = src.Columns(c + 1).ColumnWidth
    Next c
    With ws.Range(ws.Cells(6, 1), ws.Cells(6, ГРАФ))
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With
    ws.Rows(6).RowHeight = 56

    out = 7
    last = ПоследняяСтрока(src, 3)
    For r = ПЕРВАЯ_СТРОКА To last
        If Trim$(CStr(src.Cells(r, 3).Value)) <> "" Then
            If год = 0 Or ГодЗаписи(src, r) = год Then
                For c = 1 To ГРАФ
                    ws.Cells(out, c).Value = src.Cells(r, c + 1).Value
                    If IsDate(src.Cells(r, c + 1).Value) Then
                        ws.Cells(out, c).NumberFormat = "DD.MM.YYYY"
                    End If
                Next c
                ws.Cells(out, 1).Value = out - 6        ' сквозная нумерация в печати
                ОформитьСтроку ws, out
                записей = записей + 1
                out = out + 1
            End If
        End If
    Next r

    If записей = 0 Then
        ws.Cells(out, 1).Value = "Записей нет."
        ws.Range(ws.Cells(out, 1), ws.Cells(out, ГРАФ)).Merge
        ws.Cells(out, 1).Font.Italic = True
    End If

    НастроитьПечать ws
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True

    MsgBox "Журнал электроинструмента построен: записей " & записей & "." & vbCrLf & vbCrLf & _
           "В печать вышли только 16 граф утверждённой формы, в её порядке. " & _
           "Служебные колонки системы остались на рабочем листе.", _
           vbInformation, "Журнал электроинструмента"
    Exit Sub

Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось построить журнал электроинструмента." & vbCrLf & _
           "Ошибка: " & Err.Description, vbCritical, "Журнал электроинструмента"
End Sub


Private Function ГодЗаписи(src As Worksheet, r As Long) As Long
    If IsDate(src.Cells(r, 5).Value) Then ГодЗаписи = Year(CDate(src.Cells(r, 5).Value))
End Function


' Что просрочено или вот-вот просрочится. Здесь пороги короче, чем у
' поверки: периодичность всего полгода, и месяц запаса — это уже много.
Public Function ПроверитьСрокиЭлектроинструмента() As String
    Dim ws As Worksheet
    Dim r As Long, last As Long
    Dim д As Date
    Dim осталось As Long
    Dim просрочено As Long, скоро As Long
    Dim список As String

    Set ws = ЛистТихо(Л_ЭИ)
    If ws Is Nothing Then Exit Function
    last = ПоследняяСтрока(ws, 3)

    For r = ПЕРВАЯ_СТРОКА To last
        If Trim$(CStr(ws.Cells(r, 3).Value)) <> "" Then
            If IsDate(ws.Cells(r, 16).Value) Then
                д = CDate(ws.Cells(r, 16).Value)
                осталось = DateDiff("d", Date, д)
                If осталось < 0 Then
                    просрочено = просрочено + 1
                ElseIf осталось <= 30 Then
                    скоро = скоро + 1
                End If
                If осталось <= 30 And Len(список) < 320 Then
                    список = список & IIf(список = "", "", "; ") & _
                             CStr(ws.Cells(r, 3).Value) & _
                             IIf(Trim$(CStr(ws.Cells(r, 4).Value)) = "", "", _
                                 " инв. № " & CStr(ws.Cells(r, 4).Value)) & _
                             " — " & IIf(осталось < 0, "просрочен " & -осталось & " дн.", _
                                         "осталось " & осталось & " дн.")
                End If
            Else
                ' Без даты последней проверки инструмент к работе не допущен,
                ' и молчать об этом нельзя.
                If Not IsDate(ws.Cells(r, 5).Value) Then
                    просрочено = просрочено + 1
                    If Len(список) < 320 Then
                        список = список & IIf(список = "", "", "; ") & _
                                 CStr(ws.Cells(r, 3).Value) & " — проверка не отмечена ни разу"
                    End If
                End If
            End If
        End If
    Next r

    If просрочено = 0 And скоро = 0 Then Exit Function
    ПроверитьСрокиЭлектроинструмента = список
End Function


Public Function СколькоПросроченоЭлектроинструмента() As Long
    Dim ws As Worksheet
    Dim r As Long, last As Long
    Dim n As Long

    Set ws = ЛистТихо(Л_ЭИ)
    If ws Is Nothing Then Exit Function
    last = ПоследняяСтрока(ws, 3)
    For r = ПЕРВАЯ_СТРОКА To last
        If Trim$(CStr(ws.Cells(r, 3).Value)) <> "" Then
            If IsDate(ws.Cells(r, 16).Value) Then
                If CDate(ws.Cells(r, 16).Value) < Date Then n = n + 1
            ElseIf Not IsDate(ws.Cells(r, 5).Value) Then
                n = n + 1
            End If
        End If
    Next r
    СколькоПросроченоЭлектроинструмента = n
End Function


' -- Оформление --------------------------------------------------------

Private Sub ОформитьСтроку(ws As Worksheet, r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, ГРАФ))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(140, 140, 140)
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 9
    End With
    ws.Cells(r, 1).HorizontalAlignment = xlCenter
End Sub


Private Function ПересоздатьЛист() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Л_ПЕЧАТЬ).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = Л_ПЕЧАТЬ
    ws.Tab.Color = RGB(197, 90, 17)
    Set ПересоздатьЛист = ws
End Function


Private Sub НастроитьПечать(ws As Worksheet)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$6:$6"
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .CenterFooter = "&P"
        .LeftFooter = "Журнал электроинструмента"
    End With
End Sub
