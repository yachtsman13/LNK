Attribute VB_Name = "modYearEnd"
Option Explicit

' =======================================================================
'  Закрытие года: печать всех журналов одной процедурой.
'
'  В декабре журналы закрываются все сразу, и делать это по одному —
'  верный способ что-нибудь забыть. Процедура строит титульные листы,
'  сами журналы и листы заверения, а в конце показывает опись: что
'  напечатано, сколько в чём листов и что осталось сделать руками.
' =======================================================================

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
            ЗаписатьВЖурналЖурналов CStr(методы(i)), год, записей
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
' Строка в «Журнал журналов» на этот метод и год — если уже есть
' (повторное закрытие того же года), обновляется она же, а не заводится
' вторая: иначе один и тот же журнал числился бы дважды в собственном
' учёте журналов, что для документа учёта документов особенно нелепо.
Private Sub ЗаписатьВЖурналЖурналов(метод As String, год As Long, записей As Long)
    Dim жж As Worksheet
    Dim i As Long, last As Long
    Dim строка As Long

    Set жж = ЛистТихо("Журнал_журналов")
    If жж Is Nothing Then Exit Sub

    last = ПоследняяСтрока(жж, ColReq(жж, "Метод"))
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(жж.Cells(i, ColReq(жж, "Метод")).Value)), метод, vbTextCompare) = 0 Then
            If CLng(Val(жж.Cells(i, ColReq(жж, "Год")).Value)) = год Then
                строка = i
                Exit For
            End If
        End If
    Next i
    If строка = 0 Then строка = last + 1

    жж.Cells(строка, ColReq(жж, "ID")).Value = "ЖЖ-" & метод & "-" & год
    жж.Cells(строка, ColReq(жж, "Наименование журнала")).Value = _
        "Журнал учёта заключений " & метод
    жж.Cells(строка, ColReq(жж, "Метод")).Value = метод
    жж.Cells(строка, ColReq(жж, "Год")).Value = год
    жж.Cells(строка, ColReq(жж, "Учётный номер")).Value = _
        Настройка("Учётный номер журнала " & метод)
    жж.Cells(строка, ColReq(жж, "Записей за год")).Value = записей
    жж.Cells(строка, ColReq(жж, "Дата закрытия")).Value = Date
    жж.Cells(строка, ColReq(жж, "Кем закрыт")).Value = ТекущийСотрудник()
End Sub


' -- Передача в архив ------------------------------------------------

' Отмечает выделенные строки журнала как переданные в архив: пишет
' в саму строку журнала (там для этого есть колонка «№ экз., дата
' передачи в архив ЛНК») и заводит по строке в «Архивное_хранение»,
' Приложение №5 к Правилам учёта документации. Строки, уже отмеченные
' раньше, пропускаются молча — по счётчику пропусков видно, что часть
' выделенного уже была в архиве, а не задваивается.
Public Sub ОтметитьПередачуВАрхив()
    Dim ж As Worksheet
    Dim ар As Worksheet
    Dim строки As Range
    Dim ячейка As Range
    Dim номерЭкз As String
    Dim строкаАр As Long
    Dim cПоле As Long
    Dim заведено As Long, пропущено As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    Set ар = Лист("Архивное_хранение")
    If ж Is Nothing Or ар Is Nothing Then Exit Sub
    If ActiveSheet.Name <> ж.Name Then
        MsgBox "Выделите на листе «" & ж.Name & "» строки, которые передаёте в архив.", _
               vbInformation, "Передача в архив"
        Exit Sub
    End If

    номерЭкз = Trim$(InputBox("Номер переданного экземпляра:", "Передача в архив", "1"))
    If StrPtr(номерЭкз) = 0 Then Exit Sub

    Set строки = Selection
    cПоле = ColReq(ж, "№ экз., дата передачи в архив ЛНК")
    строкаАр = ПоследняяСтрока(ар, 1) + 1

    For Each ячейка In строки.Columns(1).Cells
        If ячейка.Row >= ПЕРВАЯ_СТРОКА Then
            If Trim$(CStr(ж.Cells(ячейка.Row, ColReq(ж, "ID записи")).Value)) <> "" Then
                If Trim$(CStr(ж.Cells(ячейка.Row, cПоле).Value)) <> "" Then
                    пропущено = пропущено + 1
                Else
                    ж.Cells(ячейка.Row, cПоле).Value = _
                        "№" & номерЭкз & ", " & Format$(Date, "dd.mm.yyyy")

                    ар.Cells(строкаАр, ColReq(ар, "ID")).Value = НовыйIDАрхива(ар)
                    ар.Cells(строкаАр, ColReq(ар, "ID записи журнала")).Value = _
                        ж.Cells(ячейка.Row, ColReq(ж, "ID записи")).Value
                    ар.Cells(строкаАр, ColReq(ар, "Обозначение записи")).Value = _
                        ж.Cells(ячейка.Row, ColReq(ж, "Обозначение записи (для выбора)")).Value
                    ар.Cells(строкаАр, ColReq(ар, "Дата передачи в архив")).Value = Date
                    ар.Cells(строкаАр, ColReq(ар, "Кем передано")).Value = ТекущийСотрудник()
                    строкаАр = строкаАр + 1
                    заведено = заведено + 1
                End If
            End If
        End If
    Next ячейка

    If заведено = 0 And пропущено = 0 Then
        MsgBox "Не выделено ни одной строки записи.", vbExclamation, "Передача в архив"
        Exit Sub
    End If

    MsgBox "Передано в архив: " & заведено & "." & _
           IIf(пропущено > 0, vbCrLf & "Пропущено (были отмечены раньше): " & пропущено & ".", ""), _
           vbInformation, "Передача в архив"
End Sub


' -- Опись хранения ----------------------------------------------------

' Опись, которая вкладывается в папку с результатами контроля —
' Приложение №4 к Правилам учёта документации. Собирается из журнала
' заключений за период, который физически лёг в одну папку: не «за
' год», как закрытие журналов, а «что в этой конкретной папке» —
' диапазон может быть любым, от недели до нескольких лет.
Public Sub ПостроитьОписьХранения()
    Dim ж As Worksheet
    Dim ws As Worksheet
    Dim сНач As String, сКон As String
    Dim дНач As Date, дКон As Date
    Dim i As Long, last As Long, n As Long
    Dim cДата As Long, cАкт As Long, cОбъект As Long, cАрхив As Long
    Dim методы As Variant, m As Long
    Dim список As String
    Dim д As Date

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub

    сНач = Trim$(InputBox("Опись за какой период? Дата начала, ДД.ММ.ГГГГ:", _
                          "Опись хранения", Format$(DateSerial(Year(Date), 1, 1), "dd.mm.yyyy")))
    If сНач = "" Or Not IsDate(сНач) Then Exit Sub
    дНач = CDate(сНач)
    сКон = Trim$(InputBox("Дата окончания, ДД.ММ.ГГГГ:", "Опись хранения", _
                          Format$(Date, "dd.mm.yyyy")))
    If сКон = "" Or Not IsDate(сКон) Then Exit Sub
    дКон = CDate(сКон)

    cДата = ColReq(ж, "Дата контроля")
    cАкт = ColReq(ж, "№ акта")
    cОбъект = ColReq(ж, "Наименование объекта в документе")
    cАрхив = ColReq(ж, "№ экз., дата передачи в архив ЛНК")
    методы = Split(МЕТОДЫ, ",")

    Application.ScreenUpdating = False
    Set ws = ПересоздатьЛистОписи()
    ws.Cells(1, 1).Value = "ОПИСЬ ХРАНЕНИЯ"
    ws.Cells(2, 1).Value = Настройка("Полное наименование ЛНК")
    ws.Cells(3, 1).Value = "за период с " & Format$(дНач, "dd.mm.yyyy") & _
                           " по " & Format$(дКон, "dd.mm.yyyy")
    ws.Range("A1:F1").Merge: ws.Range("A2:F2").Merge: ws.Range("A3:F3").Merge
    ws.Cells(1, 1).Font.Size = 13: ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Italic = True
    ws.Range("A1:A3").HorizontalAlignment = xlCenter

    n = 5
    ws.Cells(n, 1).Value = "№ п/п": ws.Cells(n, 2).Value = "Дата"
    ws.Cells(n, 3).Value = "№ акта": ws.Cells(n, 4).Value = "Объект"
    ws.Cells(n, 5).Value = "Методы": ws.Cells(n, 6).Value = "Передано в архив"
    With ws.Range(ws.Cells(n, 1), ws.Cells(n, 6))
        .Font.Bold = True: .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255): .Borders.LineStyle = xlContinuous
    End With
    n = n + 1

    last = ПоследняяСтрока(ж, cДата)
    For i = ПЕРВАЯ_СТРОКА To last
        If IsDate(ж.Cells(i, cДата).Value) Then
            д = CDate(ж.Cells(i, cДата).Value)
            If д >= дНач And д <= дКон Then
                список = ""
                For m = LBound(методы) To UBound(методы)
                    If МетодПрименён(ж, i, CStr(методы(m))) Then
                        список = список & IIf(список = "", "", ", ") & методы(m)
                    End If
                Next m
                ws.Cells(n, 1).Value = n - 5
                ws.Cells(n, 2).Value = д: ws.Cells(n, 2).NumberFormat = "dd.mm.yyyy"
                ws.Cells(n, 3).Value = ж.Cells(i, cАкт).Value
                ws.Cells(n, 4).Value = ж.Cells(i, cОбъект).Value
                ws.Cells(n, 5).Value = список
                ws.Cells(n, 6).Value = ж.Cells(i, cАрхив).Value
                n = n + 1
            End If
        End If
    Next i

    ws.Columns(1).ColumnWidth = 7: ws.Columns(2).ColumnWidth = 13
    ws.Columns(3).ColumnWidth = 16: ws.Columns(4).ColumnWidth = 44
    ws.Columns(5).ColumnWidth = 20: ws.Columns(6).ColumnWidth = 20
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.Activate: ws.Range("A1").Select
    Application.ScreenUpdating = True

    If n = 6 Then
        MsgBox "За этот период в журнале нет записей.", vbInformation, "Опись хранения"
    Else
        MsgBox "Опись построена, строк: " & (n - 6) & ".", vbInformation, "Опись хранения"
    End If
End Sub


Private Function ПересоздатьЛистОписи() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Опись_хранения").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "Опись_хранения"
    ws.Tab.Color = RGB(128, 128, 128)
    Set ПересоздатьЛистОписи = ws
End Function


Private Function НовыйIDАрхива(ар As Worksheet) As String
    Dim r As Long, last As Long, n As Long, m As Long
    Dim s As String
    last = ПоследняяСтрока(ар, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        s = Trim$(CStr(ар.Cells(r, 1).Value))
        If InStr(1, s, "АРХ-", vbTextCompare) = 1 Then
            m = CLng(Val(Mid$(s, 5)))
            If m > n Then n = m
        End If
    Next r
    НовыйIDАрхива = "АРХ-" & Format$(n + 1, "0000")
End Function


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
