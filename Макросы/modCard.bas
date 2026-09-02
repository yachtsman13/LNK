Attribute VB_Name = "modCard"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Карточка объекта контроля.
'
'  В справочнике объектов 35 колонок, из них два десятка — паспортные.
'  Заполнять такое построчно в таблице неудобно: строка не помещается на
'  экран, и половина полей заполняется вслепую. Карточка показывает один
'  объект вертикально, полями сверху вниз, с подсказкой к каждому.
'
'  Паспортные данные заполняются один раз и подставляются во все
'  протоколы по этому объекту. Если по конкретной работе значение другое,
'  оно переопределяется на листе «Данные_протоколов» — карточка при этом
'  остаётся историей объекта, а не работы.
' ═══════════════════════════════════════════════════════════════════════

Private Const Л_КАРТОЧКА As String = "Карточка_объекта"
Private Const СТРОКА_ID As Long = 2       ' здесь карточка помнит, чей это объект


' Открыть карточку объекта из текущей строки журнала или справочника.
Public Sub ОткрытьКарточкуОбъекта()
    Dim ид As String

    ид = ИДИзТекущейСтроки()
    If ид = "" Then Exit Sub
    ПоказатьКарточку ид
End Sub


Private Function ИДИзТекущейСтроки() As String
    Dim ws As Worksheet
    Dim r As Long

    r = ActiveCell.Row
    If r < ПЕРВАЯ_СТРОКА Then
        MsgBox "Встаньте на строку с записью.", vbInformation, "Карточка объекта"
        Exit Function
    End If

    If ActiveSheet.Name = Л_ЖУРНАЛ Then
        Set ws = Лист(Л_ЖУРНАЛ)
        ИДИзТекущейСтроки = Trim$(CStr(ws.Cells(r, ColReq(ws, "ID объекта")).Value))
        If ИДИзТекущейСтроки = "" Then
            MsgBox "В этой строке журнала не выбран объект контроля.", _
                   vbExclamation, "Карточка объекта"
        End If
    ElseIf ActiveSheet.Name = Л_ОБЪЕКТЫ Then
        Set ws = Лист(Л_ОБЪЕКТЫ)
        ИДИзТекущейСтроки = Trim$(CStr(ws.Cells(r, 1).Value))
    Else
        MsgBox "Встаньте на строку журнала или справочника объектов и повторите.", _
               vbInformation, "Карточка объекта"
    End If
End Function


Private Sub ПоказатьКарточку(ид As String)
    Dim src As Worksheet
    Dim ws As Worksheet
    Dim строкаОбъекта As Long
    Dim c As Long, r As Long, колонок As Long

    Set src = Лист(Л_ОБЪЕКТЫ)
    If src Is Nothing Then Exit Sub
    строкаОбъекта = НайтиОбъект(src, ид)
    If строкаОбъекта = 0 Then
        MsgBox "В справочнике нет объекта с ID «" & ид & "».", vbExclamation, "Карточка объекта"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set ws = ПересоздатьЛист()

    ws.Cells(1, 1).Value = "КАРТОЧКА ОБЪЕКТА КОНТРОЛЯ"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(СТРОКА_ID, 1).Value = "ID объекта"
    ws.Cells(СТРОКА_ID, 2).Value = ид
    ws.Cells(СТРОКА_ID, 2).Font.Bold = True
    ws.Cells(СТРОКА_ID, 3).Value = "Не менять: по нему карточка сохраняется обратно"
    ws.Cells(СТРОКА_ID, 3).Font.Italic = True
    ws.Cells(СТРОКА_ID, 3).Font.Color = RGB(120, 120, 120)

    колонок = src.Cells(СТРОКА_ЗАГОЛОВКОВ, src.Columns.Count).End(xlToLeft).Column
    r = 4
    For c = 2 To колонок
        If Не_Служебная(CStr(src.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value)) Then
            ws.Cells(r, 1).Value = Replace(CStr(src.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value), vbLf, " ")
            ws.Cells(r, 2).Value = src.Cells(строкаОбъекта, c).Value
            ws.Cells(r, 3).Value = Подсказка(CStr(src.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value))
            ws.Cells(r, 4).Value = c            ' служебная колонка: куда писать обратно
            ОформитьСтроку ws, r
            r = r + 1
        End If
    Next c

    ws.Cells(r + 1, 1).Value = "После правки нажмите «Сохранить карточку» в меню."
    ws.Cells(r + 1, 1).Font.Italic = True
    ws.Cells(r + 1, 1).Font.Color = RGB(0, 51, 160)

    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 46
    ws.Columns(3).ColumnWidth = 58
    ws.Columns(4).Hidden = True
    ws.Activate
    ws.Cells(4, 2).Select
    Application.ScreenUpdating = True
End Sub


' Сохранить карточку обратно в справочник объектов.
Public Sub СохранитьКарточкуОбъекта()
    Dim ws As Worksheet
    Dim src As Worksheet
    Dim ид As String
    Dim строкаОбъекта As Long
    Dim r As Long
    Dim c As Long
    Dim изменено As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Л_КАРТОЧКА)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Карточка не открыта.", vbInformation, "Карточка объекта"
        Exit Sub
    End If

    ид = Trim$(CStr(ws.Cells(СТРОКА_ID, 2).Value))
    Set src = Лист(Л_ОБЪЕКТЫ)
    If src Is Nothing Then Exit Sub
    строкаОбъекта = НайтиОбъект(src, ид)
    If строкаОбъекта = 0 Then
        MsgBox "Объект «" & ид & "» больше не найден в справочнике. " & _
               "Возможно, строку удалили — карточка не сохранена.", _
               vbExclamation, "Карточка объекта"
        Exit Sub
    End If

    For r = 4 To ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        c = CLng(Val(ws.Cells(r, 4).Value))
        If c > 0 Then
            If CStr(src.Cells(строкаОбъекта, c).Value) <> CStr(ws.Cells(r, 2).Value) Then
                src.Cells(строкаОбъекта, c).Value = ws.Cells(r, 2).Value
                изменено = изменено + 1
            End If
        End If
    Next r

    If изменено > 0 Then
        c = ColByHeader(src, "Изменено")
        If c > 0 Then src.Cells(строкаОбъекта, c).Value = Now
        c = ColByHeader(src, "Кем внесено")
        If c > 0 Then
            If Trim$(CStr(src.Cells(строкаОбъекта, c).Value)) = "миграция" Then
                src.Cells(строкаОбъекта, c).Value = ТекущийСотрудник()
            End If
        End If
    End If

    MsgBox IIf(изменено = 0, "Изменений не было.", _
               "Сохранено в справочник, изменённых полей: " & изменено & "."), _
           vbInformation, "Карточка объекта"
End Sub


' ── Вспомогательное ───────────────────────────────────────────────────

Private Function НайтиОбъект(src As Worksheet, ид As String) As Long
    Dim r As Long, last As Long
    last = ПоследняяСтрока(src, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(src.Cells(r, 1).Value)), ид, vbTextCompare) = 0 Then
            НайтиОбъект = r
            Exit Function
        End If
    Next r
End Function


' Расчётные и служебные поля в карточке не показываем: их не правят.
Private Function Не_Служебная(заголовок As String) As Boolean
    Dim h As String
    h = Replace(заголовок, vbLf, " ")
    Select Case True
        Case InStr(1, h, "Работ в журнале", vbTextCompare) > 0: Не_Служебная = False
        Case InStr(1, h, "Кем внесено", vbTextCompare) > 0: Не_Служебная = False
        Case InStr(1, h, "Изменено", vbTextCompare) > 0: Не_Служебная = False
        Case Else: Не_Служебная = True
    End Select
End Function


' Подсказка к полю: куда оно попадёт в документе.
Private Function Подсказка(заголовок As String) As String
    Dim h As String
    h = Replace(заголовок, vbLf, " ")
    Select Case True
        Case InStr(1, h, "Каноническое", vbTextCompare) > 0
            Подсказка = "Как объект называется в справочнике. В протоколе может стоять " & _
                        "уточнённое название — оно правится в журнале"
        Case InStr(1, h, "Прежние наименования", vbTextCompare) > 0
            Подсказка = "Чтобы поиск по старым журналам продолжал находить объект"
        Case InStr(1, h, "Марка материала", vbTextCompare) > 0
            Подсказка = "Подставляется в #Марка_материала протокола АЭ"
        Case InStr(1, h, "Рабочее давление", vbTextCompare) > 0
            Подсказка = "#Рабочее_давление. С единицами, как пишете в протоколе"
        Case InStr(1, h, "Испытательное давление", vbTextCompare) > 0
            Подсказка = "#Испытательное_давление"
        Case InStr(1, h, "Толщины", vbTextCompare) > 0
            Подсказка = "#Толщины — номинальные по паспорту, не результат замера"
        Case InStr(1, h, "Диаметр", vbTextCompare) > 0
            Подсказка = "Подставляется в протокол АЭ"
        Case InStr(1, h, "Магнитные свойства", vbTextCompare) > 0
            Подсказка = "#Магнитные_свойства: имеются / отсутствуют"
        Case InStr(1, h, "Потенциально опасные", vbTextCompare) > 0
            Подсказка = "#Потенциально_опасные_участки"
        Case InStr(1, h, "нагружающего", vbTextCompare) > 0
            Подсказка = "#Тип_нагружающего_оборудования"
        Case InStr(1, h, "авариях", vbTextCompare) > 0
            Подсказка = "#Сведения_об_авариях_и_ремонтах"
        Case InStr(1, h, "Заводской", vbTextCompare) > 0 Or _
             InStr(1, h, "Регистрационный", vbTextCompare) > 0 Or _
             InStr(1, h, "Учётный", vbTextCompare) > 0 Or _
             InStr(1, h, "Инвентарный", vbTextCompare) > 0
            Подсказка = "По этим номерам объект отличается от одноимённых"
        Case Else
            Подсказка = ""
    End Select
End Function


Private Sub ОформитьСтроку(ws As Worksheet, r As Long)
    ws.Cells(r, 1).Font.Bold = False
    ws.Cells(r, 2).Font.Color = RGB(0, 51, 160)
    ws.Cells(r, 3).Font.Italic = True
    ws.Cells(r, 3).Font.Size = 9
    ws.Cells(r, 3).Font.Color = RGB(120, 120, 120)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    ws.Rows(r).RowHeight = 22
End Sub


Private Function ПересоздатьЛист() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Л_КАРТОЧКА).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
    ws.Name = Л_КАРТОЧКА
    ws.Tab.Color = RGB(46, 84, 150)
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    Set ПересоздатьЛист = ws
End Function
