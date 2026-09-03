Attribute VB_Name = "modPlan"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Плановые и внеплановые работы.
'
'  Разделение не вводится руками: работа считается плановой, если в
'  плане-графике есть строка, привязанная к этой записи журнала. Ставить
'  галочку «плановая» самому — значит рано или поздно поставить её там,
'  где плана не было; а вопрос «сколько работ мы сделали сверх плана»
'  задают именно тогда, когда галочкам уже не верят.
'
'  Связь ставится одной кнопкой на строке журнала: система сама ищет
'  подходящую строку плана по объекту и месяцу.
' ═══════════════════════════════════════════════════════════════════════

Private Const Л_ПЛАН As String = "План_график"
Private Const Л_ОТЧЁТ As String = "Отчёт_план"


' Привязать текущую запись журнала к строке плана-графика.
Public Sub ОтметитьВыполнениеПоПлану()
    Dim ж As Worksheet
    Dim пл As Worksheet
    Dim r As Long
    Dim идЗаписи As String, идОбъекта As String
    Dim дата As Date
    Dim строкаПлана As Long
    Dim сколько As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    Set пл = Лист(Л_ПЛАН)
    If ж Is Nothing Or пл Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub

    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))
    идОбъекта = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID объекта")).Value))
    If идЗаписи = "" Then
        MsgBox "У строки нет ID записи — привязывать нечего.", vbExclamation, "План"
        Exit Sub
    End If
    If Not IsDate(ж.Cells(r, ColReq(ж, "Дата контроля")).Value) Then
        MsgBox "Не заполнена дата контроля — по ней ищется строка плана.", _
               vbExclamation, "План"
        Exit Sub
    End If
    дата = CDate(ж.Cells(r, ColReq(ж, "Дата контроля")).Value)

    If УжеПривязана(пл, идЗаписи) > 0 Then
        MsgBox "Эта запись уже отмечена в плане-графике (строка " & _
               УжеПривязана(пл, идЗаписи) & ")." & vbCrLf & vbCrLf & _
               "В журнале она числится плановой.", vbInformation, "План"
        Exit Sub
    End If

    строкаПлана = НайтиСтрокуПлана(пл, идОбъекта, дата, сколько)
    If строкаПлана = 0 Then
        MsgBox "В плане-графике нет незакрытой строки по этому объекту " & _
               "на " & Format$(дата, "mmmm yyyy") & " или раньше." & vbCrLf & vbCrLf & _
               "Значит, работа внеплановая — так она и будет считаться в отчёте. " & _
               "Если план всё-таки был, заведите строку на листе «План-график работ» " & _
               "и повторите.", vbInformation, "План"
        Exit Sub
    End If

    If сколько > 1 Then
        If MsgBox("Подходящих строк плана: " & сколько & ". Берётся ближайшая по " & _
                  "плановой дате — строка " & строкаПлана & ", " & _
                  CStr(пл.Cells(строкаПлана, ColReq(пл, "Наименование объекта в плане")).Value) & _
                  "." & vbCrLf & vbCrLf & "Отметить её выполненной?", _
                  vbYesNo + vbQuestion, "План") = vbNo Then Exit Sub
    End If

    пл.Cells(строкаПлана, ColReq(пл, "Запись журнала (выбор)")).Value = _
        ж.Cells(r, ColReq(ж, "Обозначение записи (для выбора)")).Value
    пл.Cells(строкаПлана, ColReq(пл, "Фактическая дата")).Value = дата
    пл.Cells(строкаПлана, ColReq(пл, "Статус")).Value = "Выполнено"
    пл.Cells(строкаПлана, ColReq(пл, "Изменено")).Value = Now
    пл.Cells(строкаПлана, ColReq(пл, "ID записи журнала")).Calculate

    MsgBox "Строка плана " & строкаПлана & " отмечена выполненной." & vbCrLf & vbCrLf & _
           "В журнале работа теперь числится плановой — колонка «Вид работ» " & _
           "считается сама.", vbInformation, "План"
End Sub


Private Function УжеПривязана(пл As Worksheet, идЗаписи As String) As Long
    Dim i As Long, last As Long
    Dim c As Long
    c = ColReq(пл, "ID записи журнала")
    last = ПоследняяСтрока(пл, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(пл.Cells(i, c).Value)), идЗаписи, vbTextCompare) = 0 Then
            УжеПривязана = i
            Exit Function
        End If
    Next i
End Function


' Незакрытая строка плана по тому же объекту, не позже месяца контроля.
' Из нескольких берётся ближайшая по плановой дате: план на март,
' выполненный в апреле, — это выполнение мартовского плана, а не
' досрочное выполнение майского.
Private Function НайтиСтрокуПлана(пл As Worksheet, идОбъекта As String, _
                                  дата As Date, ByRef сколько As Long) As Long
    Dim i As Long, last As Long
    Dim cОб As Long, cГод As Long, cМес As Long, cДата As Long, cЗапись As Long
    Dim лучшая As Long
    Dim лучшаяДата As Date
    Dim планДата As Date

    сколько = 0
    If идОбъекта = "" Then Exit Function
    cОб = ColReq(пл, "ID объекта")
    cГод = ColReq(пл, "Год")
    cМес = ColReq(пл, "Месяц")
    cДата = ColReq(пл, "Плановая дата")
    cЗапись = ColReq(пл, "ID записи журнала")
    last = ПоследняяСтрока(пл, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(пл.Cells(i, cОб).Value)), идОбъекта, vbTextCompare) = 0 Then
            If Trim$(CStr(пл.Cells(i, cЗапись).Value)) = "" Then
                планДата = ПлановаяДата(пл, i, cДата, cГод, cМес)
                If планДата > 0 Then
                    If Year(планДата) * 12 + Month(планДата) <= Year(дата) * 12 + Month(дата) Then
                        сколько = сколько + 1
                        If лучшая = 0 Or планДата > лучшаяДата Then
                            лучшая = i
                            лучшаяДата = планДата
                        End If
                    End If
                End If
            End If
        End If
    Next i
    НайтиСтрокуПлана = лучшая
End Function


' Плановая дата, а если её нет — первое число планового месяца.
Private Function ПлановаяДата(пл As Worksheet, i As Long, cДата As Long, _
                              cГод As Long, cМес As Long) As Date
    Dim г As Long, м As Long
    If IsDate(пл.Cells(i, cДата).Value) Then
        ПлановаяДата = CDate(пл.Cells(i, cДата).Value)
        Exit Function
    End If
    г = CLng(Val(пл.Cells(i, cГод).Value))
    м = НомерМесяца(CStr(пл.Cells(i, cМес).Value))
    If г > 2000 And м > 0 Then ПлановаяДата = DateSerial(г, м, 1)
End Function


Private Function НомерМесяца(текст As String) As Long
    Dim i As Long
    Dim s As String
    s = Trim$(текст)
    If s = "" Then Exit Function
    If IsNumeric(s) Then
        НомерМесяца = CLng(Val(s))
        If НомерМесяца > 12 Then НомерМесяца = 0
        Exit Function
    End If
    For i = 1 To 12
        If StrComp(Left$(s, 3), Left$(Format$(DateSerial(2000, i, 1), "mmmm"), 3), _
                   vbTextCompare) = 0 Then
            НомерМесяца = i
            Exit Function
        End If
    Next i
End Function


' ── Отчёт: план и внеплановые работы ──────────────────────────────────

Public Sub ОтчётПланИВнеплановые()
    Dim ж As Worksheet
    Dim пл As Worksheet
    Dim ws As Worksheet
    Dim s As String
    Dim год As Long, месяц As Long
    Dim r As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    Set пл = Лист(Л_ПЛАН)
    If ж Is Nothing Or пл Is Nothing Then Exit Sub

    s = InputBox("За какой месяц отчёт? Формат ММ.ГГГГ." & vbCrLf & vbCrLf & _
                 "Пустой месяц — за весь год: 2026.", "План и внеплановые", _
                 Format$(Date, "mm.yyyy"))
    If StrPtr(s) = 0 Then Exit Sub
    s = Trim$(s)
    If s = "" Then Exit Sub
    If InStr(s, ".") > 0 Then
        месяц = CLng(Val(Left$(s, InStr(s, ".") - 1)))
        год = CLng(Val(Mid$(s, InStr(s, ".") + 1)))
    Else
        год = CLng(Val(s))
    End If
    If год < 2000 Then
        MsgBox "Не разобрал год. Нужно ММ.ГГГГ или ГГГГ.", vbExclamation, "План и внеплановые"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set ws = ПересоздатьОтчёт()
    ws.Cells(1, 1).Value = "ПЛАНОВЫЕ И ВНЕПЛАНОВЫЕ РАБОТЫ ПО НЕРАЗРУШАЮЩЕМУ КОНТРОЛЮ"
    ws.Cells(2, 1).Value = Настройка("Полное наименование ЛНК")
    ws.Cells(3, 1).Value = "за " & IIf(месяц = 0, год & " год", _
                                       Format$(DateSerial(год, месяц, 1), "mmmm yyyy") & " г.")
    ws.Cells(1, 1).Font.Size = 13
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Italic = True

    r = 5
    r = ЧастьПлана(ws, r, пл, год, месяц)
    r = ЧастьЖурнала(ws, r, ж, год, месяц)

    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 20
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True

    MsgBox "Отчёт построен на листе «" & Л_ОТЧЁТ & "»." & vbCrLf & vbCrLf & _
           "Он пересобирается каждый раз заново — править его нет смысла, " & _
           "правьте план-график и журнал.", vbInformation, "План и внеплановые"
End Sub


Private Function ЧастьПлана(ws As Worksheet, r As Long, пл As Worksheet, _
                            год As Long, месяц As Long) As Long
    Dim i As Long, last As Long
    Dim cГод As Long, cМес As Long, cОбъ As Long, cМет As Long
    Dim cСтатус As Long, cФакт As Long, cОтв As Long
    Dim всего As Long, выполнено As Long
    Dim м As Long

    cГод = ColReq(пл, "Год")
    cМес = ColReq(пл, "Месяц")
    cОбъ = ColReq(пл, "Наименование объекта в плане")
    cМет = ColReq(пл, "Методы контроля")
    cСтатус = ColReq(пл, "Статус")
    cФакт = ColReq(пл, "Фактическая дата")
    cОтв = ColReq(пл, "Ответственный")

    r = Заголовок(ws, r, "По плану-графику")
    ws.Cells(r, 1).Value = "№"
    ws.Cells(r, 2).Value = "Месяц плана"
    ws.Cells(r, 3).Value = "Объект"
    ws.Cells(r, 4).Value = "Методы"
    ws.Cells(r, 5).Value = "Статус"
    ws.Cells(r, 6).Value = "Фактическая дата"
    ШапкаТаблицы ws, r
    r = r + 1

    last = ПоследняяСтрока(пл, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If CLng(Val(пл.Cells(i, cГод).Value)) = год Then
            м = НомерМесяца(CStr(пл.Cells(i, cМес).Value))
            If месяц = 0 Or м = месяц Then
                всего = всего + 1
                ws.Cells(r, 1).Value = всего
                ws.Cells(r, 2).Value = пл.Cells(i, cМес).Value
                ws.Cells(r, 3).Value = пл.Cells(i, cОбъ).Value
                ws.Cells(r, 4).Value = пл.Cells(i, cМет).Value
                ws.Cells(r, 5).Value = пл.Cells(i, cСтатус).Value
                ws.Cells(r, 6).Value = пл.Cells(i, cФакт).Value
                ws.Cells(r, 6).NumberFormat = "dd.mm.yyyy"
                If StrComp(CStr(пл.Cells(i, cСтатус).Value), "Выполнено", vbTextCompare) = 0 Then
                    выполнено = выполнено + 1
                Else
                    ws.Cells(r, 5).Font.Color = RGB(192, 0, 0)
                End If
                r = r + 1
            End If
        End If
    Next i

    If всего = 0 Then
        ws.Cells(r, 2).Value = "плановых работ на этот период не заведено"
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    Else
        ws.Cells(r, 2).Value = "Итого по плану:"
        ws.Cells(r, 3).Value = "запланировано " & всего & ", выполнено " & выполнено & _
                               ", осталось " & (всего - выполнено)
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Bold = True
        r = r + 1
    End If
    ЧастьПлана = r + 1
End Function


Private Function ЧастьЖурнала(ws As Worksheet, r As Long, ж As Worksheet, _
                              год As Long, месяц As Long) As Long
    Dim i As Long, last As Long
    Dim cДата As Long, cВид As Long, cОбъ As Long, cАкт As Long, cСтатус As Long
    Dim д As Date
    Dim всего As Long, плановых As Long, вне As Long
    Dim n As Long

    cДата = ColReq(ж, "Дата контроля")
    cВид = ColReq(ж, "Вид работ")
    cОбъ = ColReq(ж, "Наименование объекта в документе")
    cАкт = ColReq(ж, "№ акта")
    cСтатус = ColReq(ж, "Статус")

    r = Заголовок(ws, r, "Внеплановые работы")
    ws.Cells(r, 1).Value = "№"
    ws.Cells(r, 2).Value = "Дата"
    ws.Cells(r, 3).Value = "Объект"
    ws.Cells(r, 4).Value = "№ акта"
    ws.Cells(r, 5).Value = "Статус записи"
    ШапкаТаблицы ws, r
    r = r + 1

    last = ПоследняяСтрока(ж, cДата)
    For i = ПЕРВАЯ_СТРОКА To last
        If IsDate(ж.Cells(i, cДата).Value) Then
            д = CDate(ж.Cells(i, cДата).Value)
            If Year(д) = год Then
                If месяц = 0 Or Month(д) = месяц Then
                    If StrComp(CStr(ж.Cells(i, cСтатус).Value), "Аннулирован", vbTextCompare) <> 0 Then
                        всего = всего + 1
                        If StrComp(Trim$(CStr(ж.Cells(i, cВид).Value)), "По плану", vbTextCompare) = 0 Then
                            плановых = плановых + 1
                        Else
                            вне = вне + 1
                            n = n + 1
                            ws.Cells(r, 1).Value = n
                            ws.Cells(r, 2).Value = д
                            ws.Cells(r, 2).NumberFormat = "dd.mm.yyyy"
                            ws.Cells(r, 3).Value = ж.Cells(i, cОбъ).Value
                            ws.Cells(r, 4).Value = ж.Cells(i, cАкт).Value
                            ws.Cells(r, 5).Value = ж.Cells(i, cСтатус).Value
                            r = r + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If вне = 0 Then
        ws.Cells(r, 2).Value = "внеплановых работ за период нет"
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    End If

    r = r + 1
    ws.Cells(r, 2).Value = "Итого работ за период:"
    ws.Cells(r, 3).Value = всего & " — из них по плану " & плановых & ", внеплановых " & вне
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Bold = True
    ЧастьЖурнала = r + 2
End Function


Private Function Заголовок(ws As Worksheet, r As Long, текст As String) As Long
    ws.Cells(r, 1).Value = текст
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
    Заголовок = r + 1
End Function


Private Sub ШапкаТаблицы(ws As Worksheet, r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 6))
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
End Sub


Private Function ПересоздатьОтчёт() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Л_ОТЧЁТ).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = Л_ОТЧЁТ
    ws.Tab.Color = RGB(112, 48, 160)
    Set ПересоздатьОтчёт = ws
End Function
