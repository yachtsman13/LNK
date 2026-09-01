Attribute VB_Name = "modPrintJournals"
Option Explicit

' =======================================================================
'  Печатные журналы по методам.
'
'  Листы «Печать_ВИК», «Печать_УЗТ» и так далее строятся заново из
'  широкого журнала, а не тянутся формулами ИНДЕКС/ПОИСКПОЗ. В прежнем
'  файле такие формулы были протянуты на 2000 строк, и подпись года в
'  них разъехалась с данными — формула этого не замечает. Перестроение
'  снимает целый класс таких ошибок.
' =======================================================================

Private Const ПРЕФИКС As String = "Печать_"


Public Sub ПостроитьВсеПечатныеЖурналы()
    Dim методы As Variant
    Dim i As Long
    Dim год As Long
    Dim всего As Long

    год = СпроситьГод()
    If год = 0 Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Сбой
    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        всего = всего + ПостроитьЖурналМетода(CStr(методы(i)), год)
    Next i
    Application.ScreenUpdating = True

    MsgBox "Печатные журналы за " & год & " год перестроены." & vbCrLf & _
           "Всего записей: " & всего & ".", vbInformation, "Печатные журналы"
    Exit Sub
Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось построить печатные журналы." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Печатные журналы"
End Sub


' Строит один лист и возвращает число попавших в него записей.
Public Function ПостроитьЖурналМетода(метод As String, год As Long) As Long
    Dim src As Worksheet
    Dim ws As Worksheet
    Dim r As Long, last As Long, out As Long
    Dim cНомер As Long, cИсп As Long, cДата As Long, cСтатус As Long
    Dim записи() As Variant
    Dim n As Long
    Dim i As Long, j As Long
    Dim tmp As Variant

    Set src = Лист(Л_ЖУРНАЛ)
    If src Is Nothing Then Exit Function

    cНомер = ColReq(src, метод & " №")
    cИсп = ColReq(src, метод & " исполнитель")
    cДата = ColReq(src, "Дата контроля")
    cСтатус = ColReq(src, "Статус")
    last = ПоследняяСтрока(src, cДата)

    ReDim записи(1 To 10, 1 To WorksheetFunction.Max(last - ПЕРВАЯ_СТРОКА + 1, 1))

    For r = ПЕРВАЯ_СТРОКА To last
        If Trim$(CStr(src.Cells(r, cНомер).Value)) <> "" Then
            If IsDate(src.Cells(r, cДата).Value) Then
                If Year(CDate(src.Cells(r, cДата).Value)) = год Then
                    n = n + 1
                    записи(1, n) = СортКлюч(src.Cells(r, cНомер).Value)
                    записи(2, n) = РегНомер(src, r, метод)
                    записи(3, n) = src.Cells(r, cДата).Value
                    записи(4, n) = src.Cells(r, ColReq(src, "Заказывающее подразделение")).Value
                    записи(5, n) = ОписаниеОбъекта(src, r)
                    записи(6, n) = src.Cells(r, cИсп).Value
                    записи(7, n) = src.Cells(r, ColReq(src, "№ экз., дата передачи в архив ЛНК")).Value
                    записи(8, n) = src.Cells(r, ColReq(src, "№ экз., исходящий, дата отправки")).Value
                    записи(9, n) = src.Cells(r, ColReq(src, "Отправитель")).Value
                    записи(10, n) = ПримечаниеСоСтатусом(src, r, cСтатус)
                End If
            End If
        End If
    Next r

    ' Сортировка по номеру: журнал печатается в порядке нумерации,
    ' а строки широкого журнала идут в порядке дат.
    For i = 1 To n - 1
        For j = 1 To n - i
            If записи(1, j) > записи(1, j + 1) Then
                For out = 1 To 10
                    tmp = записи(out, j)
                    записи(out, j) = записи(out, j + 1)
                    записи(out, j + 1) = tmp
                Next out
            End If
        Next j
    Next i

    Set ws = ПодготовитьЛист(ПРЕФИКС & метод, метод, год)
    For i = 1 To n
        ws.Cells(4 + i, 1).Value = i
        For j = 2 To 10
            ws.Cells(4 + i, j).Value = записи(j, i)
        Next j
        ws.Cells(4 + i, 4).NumberFormat = "DD.MM.YYYY"
    Next i
    ОформитьТаблицу ws, 5, 4 + n, 10
    НастроитьПечать ws, метод, год
    ПостроитьЖурналМетода = n
End Function


' Ключ сортировки: числовые номера идут по возрастанию, черновики и «?» —
' в конец, чтобы не мешались и сразу были заметны.
Private Function СортКлюч(значение As Variant) As Double
    If IsNumeric(значение) Then
        СортКлюч = CDbl(значение)
    Else
        СортКлюч = 1000000
    End If
End Function


Private Function ОписаниеОбъекта(src As Worksheet, r As Long) As String
    Dim имя As String
    Dim зав As String
    Dim рег As String
    имя = Trim$(CStr(src.Cells(r, ColReq(src, "Наименование объекта в документе")).Value))
    зав = Trim$(CStr(src.Cells(r, ColReq(src, "Заводской № объекта")).Value))
    рег = Trim$(CStr(src.Cells(r, ColReq(src, "Рег. / уч. № объекта")).Value))
    If зав <> "" Then имя = имя & ", зав. № " & зав
    If рег <> "" Then имя = имя & ", рег. № " & рег
    ОписаниеОбъекта = имя
End Function


Private Function ПримечаниеСоСтатусом(src As Worksheet, r As Long, cСтатус As Long) As String
    Dim s As String
    s = Trim$(CStr(src.Cells(r, ColReq(src, "Примечание")).Value))
    If StrComp(CStr(src.Cells(r, cСтатус).Value), "Аннулирован", vbTextCompare) = 0 Then
        s = Trim$("АННУЛИРОВАНО. " & s)
    ElseIf StrComp(CStr(src.Cells(r, cСтатус).Value), "Черновик", vbTextCompare) = 0 Then
        s = Trim$("ЧЕРНОВИК, не зарегистрировано. " & s)
    End If
    ПримечаниеСоСтатусом = s
End Function


' Лист создаётся заново каждый раз: так в нём заведомо нет остатков
' прошлой сборки — ни лишних строк, ни старой подписи года.
Private Function ПодготовитьЛист(имя As String, метод As String, год As Long) As Worksheet
    Dim ws As Worksheet
    Dim заголовки As Variant
    Dim ширины As Variant
    Dim i As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(имя).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = имя
    ws.Tab.Color = RGB(31, 56, 100)

    ws.Range("A1").Value = "ЖУРНАЛ"
    ws.Range("A2").Value = "учёта заключений по результатам контроля методом " & метод & _
                           " за " & год & " год"
    ws.Range("A3").Value = "Учётный номер журнала: " & _
                           IIf(Настройка("Учётный номер журнала " & метод) = "", _
                               "____________", Настройка("Учётный номер журнала " & метод))
    ws.Range("A1:J1").Merge
    ws.Range("A2:J2").Merge
    ws.Range("A3:J3").Merge
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:A3").HorizontalAlignment = xlCenter
    ws.Range("A3").Font.Italic = True

    заголовки = Array("№ п/п", "Регистрационный" & vbLf & "номер", "Дата" & vbLf & "регистрации", _
                      "Заказывающее" & vbLf & "подразделение", "Объект контроля", _
                      "Исполнитель", "№ экз., дата" & vbLf & "в архив ЛНК", _
                      "№ экз., исходящий," & vbLf & "дата отправки", "Отправитель", "Примечание")
    ширины = Array(6, 17, 12, 20, 46, 18, 18, 20, 16, 24)
    For i = 0 To 9
        ws.Cells(4, i + 1).Value = заголовки(i)
        ws.Columns(i + 1).ColumnWidth = ширины(i)
    Next i
    With ws.Range(ws.Cells(4, 1), ws.Cells(4, 10))
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(4).RowHeight = 38
    Set ПодготовитьЛист = ws
End Function


Private Sub ОформитьТаблицу(ws As Worksheet, перваяСтрока As Long, последняя As Long, колонок As Long)
    Dim rng As Range
    If последняя < перваяСтрока Then Exit Sub
    Set rng = ws.Range(ws.Cells(4, 1), ws.Cells(последняя, колонок))
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(120, 120, 120)
    With ws.Range(ws.Cells(перваяСтрока, 1), ws.Cells(последняя, колонок))
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 9
    End With
    ws.Range(ws.Cells(перваяСтрока, 1), ws.Cells(последняя, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(перваяСтрока, 3), ws.Cells(последняя, 3)).HorizontalAlignment = xlCenter
    ws.Rows(перваяСтрока & ":" & последняя).AutoFit
End Sub


' Печать журнала: альбомная, шапка повторяется на каждом листе,
' номер страницы в нижнем колонтитуле.
Private Sub НастроитьПечать(ws As Worksheet, метод As String, год As Long)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$4:$4"
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .CenterFooter = "&P"
        .LeftFooter = "Журнал " & метод & " за " & год & " год"
    End With
End Sub


Private Function СпроситьГод() As Long
    Dim s As String
    s = InputBox("За какой год строить печатные журналы?", "Печатные журналы", CStr(Year(Date)))
    If Trim$(s) = "" Then Exit Function
    If Not IsNumeric(s) Then
        MsgBox "Год должен быть числом.", vbExclamation, "Печатные журналы"
        Exit Function
    End If
    СпроситьГод = CLng(s)
End Function
