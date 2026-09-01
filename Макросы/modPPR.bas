Attribute VB_Name = "modPPR"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  План-графики обслуживания и заявка на поверку.
'
'  Расшифровка категорий подтверждена: ТО-1 ежедневно (по рабочим дням),
'  ТО-2 подекадно, ТО-3 ежемесячно, ТО-4 ежеквартально.
'
'  Пятилетний план перенесён из вашего графика 605-2025.ПЛ как есть и
'  заново НЕ пересчитывается: документ подписан, а расчёт по периодичности
'  дал бы другие числа. Пересчёт — отдельная кнопка, для следующего
'  пятилетнего графика.
' ═══════════════════════════════════════════════════════════════════════

Private Const Л_ОБСЛУЖИВАНИЕ As String = "СИ_Обслуживание"
Private Const Л_ФАКТ As String = "Обслуживание_факт"
Private Const Л_МЕСЯЧНЫЙ As String = "ППР_Месяц"
Private Const Л_ЗАЯВКА As String = "Заявка_на_поверку"


' ── Месячный план-график ──────────────────────────────────────────────

Public Sub ПостроитьМесячныйПлан()
    Dim ws As Worksheet
    Dim об As Worksheet
    Dim si As Worksheet
    Dim год As Long, месяц As Long
    Dim i As Long, last As Long, r As Long, n As Long
    Dim началоМесяца As Date, конецМесяца As Date
    Dim интервал As Long
    Dim последнее As Date
    Dim следующее As Date
    Dim сколько As Long
    Dim трудозатраты As Double
    Dim итого As Double

    If Данные() Is Nothing Then Exit Sub
    If Not СпроситьПериод(год, месяц) Then Exit Sub

    началоМесяца = DateSerial(год, месяц, 1)
    конецМесяца = DateSerial(год, месяц + 1, 0)

    Set об = ЛистТихо(Л_ОБСЛУЖИВАНИЕ)
    Set si = ЛистТихо(Л_СИ)
    If об Is Nothing Or si Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Сбой
    Set ws = ПересоздатьЛист(Л_МЕСЯЧНЫЙ)

    ws.Cells(1, 1).Value = "ПЛАН-ГРАФИК"
    ws.Cells(2, 1).Value = "технического обслуживания средств неразрушающего контроля на " & _
                           НазваниеМесяца(месяц) & " " & год & " года"
    ws.Range("A1:H1").Merge
    ws.Range("A2:H2").Merge
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:A2").HorizontalAlignment = xlCenter

    r = Шапка(ws, 4, Array("№ п/п", "№ карточки", "Наименование средства НК", "Категория", _
                           "Периодичность", "Последнее обслуживание", "Плановая дата", _
                           "Трудозатраты, ч", "Отметка о выполнении"), _
              Array(6, 11, 46, 10, 24, 18, 14, 13, 26))

    last = ПоследняяСтрока(об, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        интервал = CLng(Val(об.Cells(i, ColReq(об, "Интервал, дней")).Value))
        If интервал > 0 Then
            последнее = ДатаПоследнего(CStr(об.Cells(i, ColReq(об, "ID средства НК")).Value), _
                                       CStr(об.Cells(i, ColReq(об, "Категория")).Value))
            сколько = СколькоВМесяце(интервал, последнее, началоМесяца, конецМесяца, следующее)
            If сколько > 0 Then
                n = n + 1
                трудозатраты = сколько * ЧислоИли0(об.Cells(i, ColReq(об, "Трудозатраты за одну операцию, ч")).Value)
                итого = итого + трудозатраты
                ws.Cells(r, 1).Value = n
                ws.Cells(r, 2).Value = об.Cells(i, ColReq(об, "№ карточки учёта")).Value
                ws.Cells(r, 3).Value = об.Cells(i, ColReq(об, "Наименование")).Value
                ws.Cells(r, 4).Value = об.Cells(i, ColReq(об, "Категория")).Value
                ws.Cells(r, 5).Value = об.Cells(i, ColReq(об, "Периодичность")).Value & _
                                       IIf(сколько > 1, " — " & сколько & " раз за месяц", "")
                ws.Cells(r, 6).Value = IIf(последнее = 0, "не проводилось", _
                                           Format$(последнее, "dd.mm.yyyy"))
                ws.Cells(r, 7).Value = Format$(следующее, "dd.mm.yyyy")
                ws.Cells(r, 8).Value = Round(трудозатраты, 2)
                ОформитьСтроку ws, r, 9
                If последнее = 0 Then ws.Cells(r, 6).Font.Color = RGB(156, 0, 6)
                r = r + 1
            End If
        End If
    Next i

    If n = 0 Then
        ws.Cells(r, 1).Value = "На этот месяц обслуживание не запланировано."
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Merge
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    Else
        ws.Cells(r, 3).Value = "ИТОГО"
        ws.Cells(r, 8).Value = Round(итого, 2)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Font.Bold = True
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = RGB(226, 234, 246)
        ОформитьСтроку ws, r, 9
    End If

    НастроитьПечать ws, "План-график ТО на " & НазваниеМесяца(месяц) & " " & год, 4
    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "Месячный план-график построен: позиций " & n & _
           ", трудозатраты " & Round(итого, 2) & " ч." & vbCrLf & vbCrLf & _
           "Красным помечено оборудование, по которому обслуживание не проводилось ни разу — " & _
           "для него план считается от даты ввода в эксплуатацию.", _
           vbInformation, "План-график ППР"
    Exit Sub
Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось построить план-график." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "План-график ППР"
End Sub


' Сколько раз мероприятие попадает в месяц и когда первый раз.
' Считаем от последнего фактического, а если его не было — от начала
' месяца: иначе прибор, купленный три года назад, дал бы сотню
' «просроченных» операций задним числом.
Private Function СколькоВМесяце(интервал As Long, последнее As Date, _
                                началоМесяца As Date, конецМесяца As Date, _
                                ByRef первое As Date) As Long
    Dim д As Date
    Dim n As Long

    If последнее = 0 Then
        д = началоМесяца
    Else
        д = последнее + интервал
        If д < началоМесяца Then
            ' Просрочено: показываем один раз в начале месяца, а не всю
            ' пропущенную серию — план должен быть выполнимым.
            д = началоМесяца
        End If
    End If

    первое = 0
    Do While д <= конецМесяца
        If д >= началоМесяца Then
            n = n + 1
            If первое = 0 Then первое = д
        End If
        If интервал = 1 Then
            д = СледующийРабочийДень(д)
        Else
            д = д + интервал
        End If
        If n > 40 Then Exit Do          ' ТО-1: больше рабочих дней в месяце не бывает
    Loop
    СколькоВМесяце = n
End Function


Private Function СледующийРабочийДень(д As Date) As Date
    Dim x As Date
    x = д + 1
    Do While Weekday(x, vbMonday) > 5
        x = x + 1
    Loop
    СледующийРабочийДень = x
End Function


Private Function ДатаПоследнего(идСредства As String, категория As String) As Date
    Dim ws As Worksheet
    Dim i As Long, last As Long
    Dim д As Date

    If Trim$(идСредства) = "" Then Exit Function
    Set ws = ЛистТихо(Л_ФАКТ)
    If ws Is Nothing Then Exit Function
    last = ПоследняяСтрока(ws, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(i, ColReq(ws, "ID средства НК")).Value)), _
                   Trim$(идСредства), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(i, ColReq(ws, "Категория")).Value)), _
                       Trim$(категория), vbTextCompare) = 0 Then
                If IsDate(ws.Cells(i, ColReq(ws, "Дата проведения")).Value) Then
                    д = CDate(ws.Cells(i, ColReq(ws, "Дата проведения")).Value)
                    If д > ДатаПоследнего Then ДатаПоследнего = д
                End If
            End If
        End If
    Next i
End Function


' ── Заявка на поверку ─────────────────────────────────────────────────

Public Sub ЗаявкаНаПоверку()
    Dim ws As Worksheet
    Dim si As Worksheet
    Dim год As Long
    Dim i As Long, last As Long, r As Long, n As Long
    Dim д As Date
    Dim без As Long
    Dim спискеБез As String

    If Данные() Is Nothing Then Exit Sub
    год = CLng(Val(InputBox("Заявка на поверку на какой год?", "Заявка на поверку", _
                            CStr(Year(Date) + 1))))
    If год = 0 Then Exit Sub

    Set si = ЛистТихо(Л_СИ)
    If si Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Сбой
    Set ws = ПересоздатьЛист(Л_ЗАЯВКА)

    ws.Cells(1, 1).Value = "ЗАЯВКА"
    ws.Cells(2, 1).Value = "на поверку средств измерений на " & год & " год"
    ws.Cells(3, 1).Value = Настройка("Полное наименование ЛНК")
    ws.Range("A1:H1").Merge
    ws.Range("A2:H2").Merge
    ws.Range("A3:H3").Merge
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:A3").HorizontalAlignment = xlCenter

    r = Шапка(ws, 5, Array("№ п/п", "Наименование, тип", "Заводской №", "Год выпуска", _
                           "Дата последней поверки", "Действительно до", "Квартал подачи", _
                           "Примечание"), _
              Array(6, 44, 18, 11, 18, 16, 13, 26))

    last = ПоследняяСтрока(si, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "Состояние")).Value)), "Списан", vbTextCompare) <> 0 Then
            If StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "Вид метрологического контроля")).Value)), _
                       "Не подлежит", vbTextCompare) <> 0 Then
                If IsDate(si.Cells(i, ColReq(si, "Действительно до")).Value) Then
                    д = CDate(si.Cells(i, ColReq(si, "Действительно до")).Value)
                    If Year(д) = год Then
                        n = n + 1
                        ws.Cells(r, 1).Value = n
                        ws.Cells(r, 2).Value = Trim$(CStr(si.Cells(i, ColReq(si, "Наименование")).Value) & _
                                                     " " & CStr(si.Cells(i, ColReq(si, "Тип / обозначение")).Value))
                        ws.Cells(r, 3).Value = si.Cells(i, ColReq(si, "Заводской №")).Value
                        ws.Cells(r, 4).Value = si.Cells(i, ColReq(si, "Год выпуска")).Value
                        ws.Cells(r, 5).Value = si.Cells(i, ColReq(si, "Дата последней поверки")).Value
                        ws.Cells(r, 6).Value = д
                        ws.Cells(r, 7).Value = "IV кв. " & (год - 1) & " — " & Квартал(д) & " кв. " & год
                        ws.Cells(r, 5).NumberFormat = "DD.MM.YYYY"
                        ws.Cells(r, 6).NumberFormat = "DD.MM.YYYY"
                        ОформитьСтроку ws, r, 8
                        ws.Cells(r, 9).Value = Квартал(д)      ' служебное, для сортировки
                        r = r + 1
                    End If
                Else
                    без = без + 1
                    If Len(спискеБез) < 300 Then
                        спискеБез = спискеБез & vbCrLf & "  • " & _
                                    CStr(si.Cells(i, ColReq(si, "Наименование")).Value) & _
                                    " зав. № " & CStr(si.Cells(i, ColReq(si, "Заводской №")).Value)
                    End If
                End If
            End If
        End If
    Next i

    ' Сортировка по кварталу: заявка так и читается как график.
    If n > 1 Then
        ws.Range(ws.Cells(6, 1), ws.Cells(r - 1, 9)).Sort _
            Key1:=ws.Cells(6, 9), Order1:=xlAscending, Header:=xlNo
        For i = 6 To r - 1
            ws.Cells(i, 1).Value = i - 5
        Next i
    End If
    ws.Columns(9).Delete

    НастроитьПечать ws, "Заявка на поверку на " & год & " год", 5
    ws.Activate
    Application.ScreenUpdating = True

    MsgBox "Заявка на " & год & " год готова: приборов " & n & "." & _
           IIf(без = 0, "", vbCrLf & vbCrLf & "Не попали в заявку — не заполнена дата поверки (" & _
               без & "):" & спискеБез & vbCrLf & _
               "Проверьте их: молча выпасть из заявки они не должны."), _
           vbInformation, "Заявка на поверку"
    Exit Sub
Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось построить заявку." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Заявка на поверку"
End Sub


' ── Пересчёт пятилетнего плана ────────────────────────────────────────

' Только для НОВОГО пятилетнего графика. Действующий 605-2025.ПЛ подписан,
' и пересчёт разошёлся бы с подписанным документом.
Public Sub ПересчитатьПятилетний()
    Dim первыйГод As Long
    If MsgBox("Пересчёт заменит помесячную раскладку на листе «ППР_План» " & _
              "расчётной." & vbCrLf & vbCrLf & _
              "Действующий график 605-2025.ПЛ подписан, и расчёт даст другие числа: " & _
              "например, декадному ТО он даст 36 операций в год вместо ваших 32." & vbCrLf & vbCrLf & _
              "Делать это стоит только когда составляете НОВЫЙ пятилетний график. " & _
              "Продолжить?", vbYesNo + vbExclamation, "Пересчёт плана") = vbNo Then Exit Sub

    первыйГод = CLng(Val(InputBox("С какого года начинается новый пятилетний график?", _
                                  "Пересчёт плана", CStr(Year(Date) + 1))))
    If первыйГод = 0 Then Exit Sub
    ПересчитатьЛистПлана первыйГод
End Sub


Private Sub ПересчитатьЛистПлана(первыйГод As Long)
    Dim ws As Worksheet
    Dim об As Worksheet
    Dim i As Long, last As Long, r As Long
    Dim y As Long, m As Long
    Dim интервал As Long
    Dim заОперацию As Double
    Dim вМесяц As Long
    Dim колонка As Long
    Dim карточка As String

    Set ws = ЛистТихо("ППР_План")
    Set об = ЛистТихо(Л_ОБСЛУЖИВАНИЕ)
    If ws Is Nothing Or об Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    last = ПоследняяСтрока(об, 1)
    For r = ПЕРВАЯ_СТРОКА To ПоследняяСтрока(ws, 2)
        карточка = Trim$(CStr(ws.Cells(r, 2).Value))
        интервал = 0
        For i = ПЕРВАЯ_СТРОКА To last
            If StrComp(Trim$(CStr(об.Cells(i, ColReq(об, "№ карточки учёта")).Value)), карточка, vbTextCompare) = 0 Then
                If StrComp(Trim$(CStr(об.Cells(i, ColReq(об, "Категория")).Value)), _
                           Trim$(CStr(ws.Cells(r, 4).Value)), vbTextCompare) = 0 Then
                    интервал = CLng(Val(об.Cells(i, ColReq(об, "Интервал, дней")).Value))
                    заОперацию = ЧислоИли0(об.Cells(i, ColReq(об, "Трудозатраты за одну операцию, ч")).Value)
                    Exit For
                End If
            End If
        Next i
        If интервал > 0 Then
            For y = 0 To 4
                For m = 1 To 12
                    колонка = 6 + y * 13 + (m - 1)
                    вМесяц = ОперацийВМесяце(интервал, первыйГод + y, m)
                    If вМесяц > 0 Then
                        ws.Cells(r, колонка).Value = Round(вМесяц * заОперацию, 2)
                    Else
                        ws.Cells(r, колонка).ClearContents
                    End If
                Next m
            Next y
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox "Помесячная раскладка пересчитана по периодичности начиная с " & первыйГод & " года.", _
           vbInformation, "Пересчёт плана"
End Sub


Private Function ОперацийВМесяце(интервал As Long, год As Long, месяц As Long) As Long
    Dim дней As Long
    Dim д As Date
    Dim n As Long

    дней = Day(DateSerial(год, месяц + 1, 0))
    Select Case интервал
        Case 1                          ' ТО-1: по рабочим дням
            For n = 1 To дней
                д = DateSerial(год, месяц, n)
                If Weekday(д, vbMonday) <= 5 Then ОперацийВМесяце = ОперацийВМесяце + 1
            Next n
        Case Is >= 85                   ' ТО-4: раз в квартал
            If месяц Mod 3 = 1 Then ОперацийВМесяце = 1
        Case Is >= 25                   ' ТО-3: раз в месяц
            ОперацийВМесяце = 1
        Case Else                       ' ТО-2 и прочее: по интервалу
            ОперацийВМесяце = Int(дней / интервал)
            If ОперацийВМесяце = 0 Then ОперацийВМесяце = 1
    End Select
End Function


' ── Вспомогательное ───────────────────────────────────────────────────

Private Function СпроситьПериод(ByRef год As Long, ByRef месяц As Long) As Boolean
    Dim s As String
    s = InputBox("На какой месяц строить план-график?" & vbCrLf & _
                 "Формат: ММ.ГГГГ", "План-график ППР", Format$(DateAdd("m", 1, Date), "mm.yyyy"))
    If Trim$(s) = "" Then Exit Function
    If InStr(s, ".") = 0 Then
        MsgBox "Нужен формат ММ.ГГГГ, например 10.2026.", vbExclamation, "План-график ППР"
        Exit Function
    End If
    месяц = CLng(Val(Split(s, ".")(0)))
    год = CLng(Val(Split(s, ".")(1)))
    If месяц < 1 Or месяц > 12 Or год < 2000 Then
        MsgBox "Не разобрал период «" & s & "». Нужен формат ММ.ГГГГ.", _
               vbExclamation, "План-график ППР"
        Exit Function
    End If
    СпроситьПериод = True
End Function


Private Function НазваниеМесяца(месяц As Long) As String
    НазваниеМесяца = Split("январь,февраль,март,апрель,май,июнь,июль,август," & _
                           "сентябрь,октябрь,ноябрь,декабрь", ",")(месяц - 1)
End Function


Private Function Квартал(д As Date) As Long
    Квартал = Int((Month(д) - 1) / 3) + 1
End Function


Private Function ЧислоИли0(значение As Variant) As Double
    If IsNumeric(значение) Then
        If Trim$(CStr(значение)) <> "" Then ЧислоИли0 = CDbl(значение)
    End If
End Function



Private Function Шапка(ws As Worksheet, r As Long, заголовки As Variant, ширины As Variant) As Long
    Dim i As Long
    For i = LBound(заголовки) To UBound(заголовки)
        ws.Cells(r, i + 1).Value = заголовки(i)
        ws.Columns(i + 1).ColumnWidth = ширины(i)
    Next i
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(заголовки) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With
    ws.Rows(r).RowHeight = 34
    Шапка = r + 1
End Function


Private Sub ОформитьСтроку(ws As Worksheet, r As Long, колонок As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, колонок))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 9
    End With
    ws.Cells(r, 1).HorizontalAlignment = xlCenter
End Sub


Private Function ПересоздатьЛист(имя As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(имя).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = имя
    ws.Tab.Color = RGB(46, 84, 150)
    Set ПересоздатьЛист = ws
End Function


Private Sub НастроитьПечать(ws As Worksheet, подпись As String, строкаШапки As Long)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$" & строкаШапки & ":$" & строкаШапки
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(1)
        .CenterFooter = "&P"
        .LeftFooter = подпись
    End With
End Sub
