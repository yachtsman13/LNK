Attribute VB_Name = "modSM"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Документооборот системы менеджмента ЛНК: номера документов СМ и
'  лист ознакомления сотрудников.
'
'  Лист ознакомления — общая форма для всех документов СМ (Приложение №2
'  к Правилам рассмотрения рекламаций и одноимённые приложения к другим
'  Правилам): документ, редакция, должность, ФИО, подпись, дата. ФИО и
'  должности берутся из реестра специалистов — переписывать их в каждый
'  документ вручную незачем, а подпись и дата всё равно ставятся от руки.
' ═══════════════════════════════════════════════════════════════════════

Private Const Л_ДОК_СМ As String = "Документы_СМ"
Private Const Л_ОЗНАКОМЛЕНИЕ As String = "Лист_ознакомления"


' Выдаёт номер документу СМ на текущей строке листа «Документы_СМ».
' Год берётся из даты введения в действие — как и везде в системе, год
' в номере значит «когда выпущен», а не входит в ключ счётчика.
Public Sub ВыдатьНомерДокСМ()
    Dim ws As Worksheet
    Dim r As Long
    Dim вид As String
    Dim год As Long
    Dim номер As String

    Set ws = Лист(Л_ДОК_СМ)
    If ws Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ws)
    If r = 0 Then Exit Sub

    If Trim$(CStr(ws.Cells(r, ColReq(ws, "Номер")).Value)) <> "" Then
        MsgBox "У этой строки уже есть номер: " & ws.Cells(r, ColReq(ws, "Номер")).Value & _
               "." & vbCrLf & vbCrLf & "Повторная выдача дала бы второй номер тому же " & _
               "документу.", vbExclamation, "Номер документа СМ"
        Exit Sub
    End If

    вид = Trim$(CStr(ws.Cells(r, ColReq(ws, "Вид документа")).Value))
    If вид = "" Then
        MsgBox "Не заполнен вид документа — по нему строится буквенная часть номера.", _
               vbExclamation, "Номер документа СМ"
        Exit Sub
    End If

    If Not IsDate(ws.Cells(r, ColReq(ws, "Дата введения в действие")).Value) Then
        MsgBox "Не заполнена дата введения в действие — по ней определяется год в номере.", _
               vbExclamation, "Номер документа СМ"
        Exit Sub
    End If
    год = Year(CDate(ws.Cells(r, ColReq(ws, "Дата введения в действие")).Value))

    номер = modNumbers.ВыдатьНомерДокументаСМ(вид, год)
    If номер = "" Then Exit Sub          ' modNumbers уже объяснил причину

    ws.Cells(r, ColReq(ws, "Номер")).Value = номер
    ws.Cells(r, ColReq(ws, "Изменено")).Value = Now
    If Trim$(CStr(ws.Cells(r, ColReq(ws, "Кем внесено")).Value)) = "" Then
        ws.Cells(r, ColReq(ws, "Кем внесено")).Value = ТекущийСотрудник()
    End If

    MsgBox "Выдан номер " & номер & ".", vbInformation, "Номер документа СМ"
End Sub


' ── Лист ознакомления ─────────────────────────────────────────────────

Public Sub ПостроитьЛистОзнакомления()
    Dim сп As Worksheet
    Dim ws As Worksheet
    Dim документ As String
    Dim редакция As String
    Dim i As Long, last As Long, out As Long
    Dim фио As String, должность As String

    Set сп = Лист(Л_СПЕЦИАЛИСТЫ)
    If сп Is Nothing Then Exit Sub

    документ = Trim$(InputBox("Какой документ?" & vbCrLf & vbCrLf & _
                              "Например: «Положение о лаборатории», «Инструкция по " & _
                              "охране труда 05-2026.И».", "Лист ознакомления"))
    If документ = "" Then Exit Sub

    редакция = Trim$(InputBox("Редакция / дата редакции документа:", _
                              "Лист ознакомления", Format$(Date, "dd.mm.yyyy")))
    If StrPtr(редакция) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Сбой
    Set ws = ПересоздатьЛист()

    ws.Cells(1, 1).Value = "ЛИСТ ОЗНАКОМЛЕНИЯ"
    ws.Cells(2, 1).Value = документ & IIf(редакция = "", "", ", ред. " & редакция)
    ws.Cells(3, 1).Value = Настройка("Полное наименование ЛНК")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 5)).Merge
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 5)).Merge
    ws.Range(ws.Cells(3, 1), ws.Cells(3, 5)).Merge
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Italic = True
    ws.Range(ws.Cells(1, 1), ws.Cells(3, 1)).HorizontalAlignment = xlCenter

    out = 5
    ws.Cells(out, 1).Value = "№ п/п"
    ws.Cells(out, 2).Value = "Должность"
    ws.Cells(out, 3).Value = "Фамилия, инициалы"
    ws.Cells(out, 4).Value = "Подпись"
    ws.Cells(out, 5).Value = "Дата"
    With ws.Range(ws.Cells(out, 1), ws.Cells(out, 5))
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(out).RowHeight = 22
    out = out + 1

    last = ПоследняяСтрока(сп, ColReq(сп, "ФИО полностью"))
    For i = ПЕРВАЯ_СТРОКА To last
        фио = Trim$(CStr(сп.Cells(i, ColReq(сп, "ФИО в документах")).Value))
        If фио = "" Then фио = Trim$(CStr(сп.Cells(i, ColReq(сп, "ФИО полностью")).Value))
        If фио <> "" Then
            должность = Trim$(CStr(сп.Cells(i, ColReq(сп, "Должность")).Value))
            ws.Cells(out, 1).Value = out - 5
            ws.Cells(out, 2).Value = должность
            ws.Cells(out, 3).Value = фио
            With ws.Range(ws.Cells(out, 1), ws.Cells(out, 5))
                .Borders.LineStyle = xlContinuous
                .RowHeight = 20
            End With
            ws.Cells(out, 1).HorizontalAlignment = xlCenter
            out = out + 1
        End If
    Next i

    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 26
    ws.Columns(3).ColumnWidth = 26
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(5).ColumnWidth = 14
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False

    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True

    MsgBox "Лист ознакомления построен на листе «" & Л_ОЗНАКОМЛЕНИЕ & "», строк: " & _
           (out - 6) & "." & vbCrLf & vbCrLf & "Подпись и дата — от руки при печати; " & _
           "лист пересобирается заново при каждом запуске, править его нет смысла.", _
           vbInformation, "Лист ознакомления"
    Exit Sub

Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось построить лист ознакомления." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Лист ознакомления"
End Sub


Private Function ПересоздатьЛист() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Л_ОЗНАКОМЛЕНИЕ).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = Л_ОЗНАКОМЛЕНИЕ
    ws.Tab.Color = RGB(197, 90, 17)
    Set ПересоздатьЛист = ws
End Function
