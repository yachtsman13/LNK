Attribute VB_Name = "modYearEnd"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Закрытие года: печать всех журналов одной процедурой.
'
'  В декабре журналы закрываются все сразу, и делать это по одному —
'  верный способ что-нибудь забыть. Процедура строит титульные листы,
'  сами журналы и листы заверения, а в конце показывает опись: что
'  напечатано, сколько в чём листов и что осталось сделать руками.
' ═══════════════════════════════════════════════════════════════════════

Private Const ПРЕФИКС_ГОД As String = "Год_"


Public Sub ЗакрытьГод()
    Dim год As Long
    Dim опись As Collection
    Dim методы As Variant
    Dim i As Long
    Dim записей As Long
    Dim текст As String
    Dim з As Variant

    If Данные() Is Nothing Then Exit Sub
    год = CLng(Val(InputBox("Какой год закрываем?" & vbCrLf & vbCrLf & _
                            "Будут построены титульные листы, журналы по методам и " & _
                            "листы заверения — всё за один раз.", _
                            "Закрытие года", CStr(Year(Date)))))
    If год = 0 Then Exit Sub

    Set опись = New Collection
    Application.ScreenUpdating = False
    On Error GoTo Сбой

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        записей = modPrintJournals.ПостроитьЖурналМетода(CStr(методы(i)), год)
        If записей > 0 Then
            ДобавитьТитулИЗаверение CStr(методы(i)), год, записей
            опись.Add Array("Журнал " & методы(i), записей)
        End If
    Next i

    ПостроитьОписьПакета опись, год
    Application.ScreenUpdating = True

    For Each з In опись
        текст = текст & vbCrLf & "  " & з(0) & ": " & з(1) & " записей"
    Next з
    If текст = "" Then
        MsgBox "За " & год & " год записей в журнале нет — закрывать нечего.", _
               vbInformation, "Закрытие года"
        Exit Sub
    End If

    MsgBox "Год " & год & " подготовлен к закрытию." & текст & vbCrLf & vbCrLf & _
           "Что построено: титульный лист, журнал и лист заверения по каждому методу, " & _
           "плюс общая опись на листе «" & ПРЕФИКС_ГОД & год & "»." & vbCrLf & vbCrLf & _
           "Печатать листы по очереди из каждого листа «Печать_*». Число листов " & _
           "в заверении проставляется после печати — до неё оно неизвестно.", _
           vbInformation, "Закрытие года"
    Exit Sub

Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось закрыть год." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Закрытие года"
End Sub


' Титульный лист и лист заверения добавляются к уже построенному
' журналу метода: первый — перед таблицей, второй — после неё.
Private Sub ДобавитьТитулИЗаверение(метод As String, год As Long, записей As Long)
    Dim ws As Worksheet
    Dim последняя As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Печать_" & метод)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    последняя = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 3

    ws.Cells(последняя, 1).Value = "ЛИСТ ЗАВЕРЕНИЯ"
    ws.Cells(последняя + 1, 1).Value = _
        "В настоящем журнале пронумеровано, прошнуровано и скреплено печатью " & _
        "______ (____________________________) листов."
    ws.Cells(последняя + 2, 1).Value = "Записей за " & год & " год: " & записей & "."
    ws.Cells(последняя + 4, 1).Value = "Должность ______________________"
    ws.Cells(последняя + 5, 1).Value = "Подпись, ФИО ___________________"
    ws.Cells(последняя + 6, 1).Value = "«____» ________________ " & (год + 1) & " г."

    ws.Range(ws.Cells(последняя, 1), ws.Cells(последняя + 6, 1)).Font.Size = 11
    ws.Cells(последняя, 1).Font.Bold = True
    ws.Range(ws.Cells(последняя, 1), ws.Cells(последняя + 6, 10)).WrapText = False

    ' Заверение начинается с новой страницы: подпись не должна попасть
    ' на один лист с последними записями.
    ws.HPageBreaks.Add Before:=ws.Cells(последняя, 1)
End Sub


' Общая опись — то, что кладётся сверху пакета журналов за год.
Private Sub ПостроитьОписьПакета(опись As Collection, год As Long)
    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long
    Dim всего As Long

    Set ws = ПересоздатьЛист(ПРЕФИКС_ГОД & год)
    ws.Cells(1, 1).Value = "ОПИСЬ"
    ws.Cells(2, 1).Value = "журналов лаборатории неразрушающего контроля за " & год & " год"
    ws.Cells(3, 1).Value = Настройка("Полное наименование ЛНК")
    ws.Range("A1:D1").Merge
    ws.Range("A2:D2").Merge
    ws.Range("A3:D3").Merge
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:A3").HorizontalAlignment = xlCenter

    r = 5
    ws.Cells(r, 1).Value = "№ п/п"
    ws.Cells(r, 2).Value = "Наименование журнала"
    ws.Cells(r, 3).Value = "Учётный номер"
    ws.Cells(r, 4).Value = "Записей за год"
    ws.Cells(r, 5).Value = "Листов"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 44
    ws.Columns(3).ColumnWidth = 20
    ws.Columns(4).ColumnWidth = 15
    ws.Columns(5).ColumnWidth = 12
    r = r + 1

    For i = 1 To опись.Count
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = опись(i)(0)
        ws.Cells(r, 3).Value = Настройка("Учётный номер журнала " & _
                                         Replace(CStr(опись(i)(0)), "Журнал ", ""))
        ws.Cells(r, 4).Value = опись(i)(1)
        всего = всего + CLng(опись(i)(1))
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 1)).HorizontalAlignment = xlCenter
        ws.Cells(r, 4).HorizontalAlignment = xlCenter
        r = r + 1
    Next i

    ws.Cells(r, 2).Value = "ИТОГО"
    ws.Cells(r, 4).Value = всего
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Font.Bold = True
        .Interior.Color = RGB(226, 234, 246)
        .Borders.LineStyle = xlContinuous
    End With
    ws.Cells(r, 4).HorizontalAlignment = xlCenter

    r = r + 3
    ws.Cells(r, 1).Value = "Колонка «Листов» заполняется после печати: до неё число листов неизвестно."
    ws.Cells(r, 1).Font.Italic = True
    ws.Cells(r, 1).Font.Color = RGB(120, 120, 120)
    r = r + 2
    ws.Cells(r, 1).Value = "Начальник ЛНК  ______________________  " & _
                           ЗначениеПостоянной("Начальник ЛНК: должность, ФИО, телефон")
    r = r + 2
    ws.Cells(r, 1).Value = "«____» ________________ " & (год + 1) & " г."

    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "&P"
    End With
End Sub


Private Function ЗначениеПостоянной(поле As String) As String
    Dim ws As Worksheet
    Dim r As Long, last As Long
    Set ws = ЛистТихо("Паспорт_постоянная")
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), поле, vbTextCompare) = 0 Then
            ЗначениеПостоянной = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function


Private Function ПересоздатьЛист(имя As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(имя).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = имя
    ws.Tab.Color = RGB(112, 48, 160)
    Set ПересоздатьЛист = ws
End Function
