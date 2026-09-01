Attribute VB_Name = "modPPR"
Option Explicit

' =======================================================================
'  ѕлан-графики обслуживани€ и за€вка на поверку.
'
'  –асшифровка категорий подтверждена: “ќ-1 ежедневно (по рабочим дн€м),
'  “ќ-2 подекадно, “ќ-3 ежемес€чно, “ќ-4 ежеквартально.
'
'  ѕ€тилетний план перенесЄн из вашего графика 605-2025.ѕЋ как есть и
'  заново Ќ≈ пересчитываетс€: документ подписан, а расчЄт по периодичности
'  дал бы другие числа. ѕересчЄт Ч отдельна€ кнопка, дл€ следующего
'  п€тилетнего графика.
' =======================================================================

Private Const Ћ_ќЅ—Ћ”∆»¬јЌ»≈ As String = "—»_ќбслуживание"
Private Const Ћ_‘ј “ As String = "ќбслуживание_факт"
Private Const Ћ_ћ≈—я„Ќџ… As String = "ѕѕ–_ћес€ц"
Private Const Ћ_«јя¬ ј As String = "«а€вка_на_поверку"


' -- ћес€чный план-график ----------------------------------------------

Public Sub ѕостроитьћес€чныйѕлан()
    Dim ws As Worksheet
    Dim об As Worksheet
    Dim si As Worksheet
    Dim год As Long, мес€ц As Long
    Dim i As Long, last As Long, r As Long, n As Long
    Dim началоћес€ца As Date, конецћес€ца As Date
    Dim интервал As Long
    Dim последнее As Date
    Dim следующее As Date
    Dim сколько As Long
    Dim трудозатраты As Double
    Dim итого As Double

    If ƒанные() Is Nothing Then Exit Sub
    If Not —проситьѕериод(год, мес€ц) Then Exit Sub

    началоћес€ца = DateSerial(год, мес€ц, 1)
    конецћес€ца = DateSerial(год, мес€ц + 1, 0)

    Set об = Ћист“ихо(Ћ_ќЅ—Ћ”∆»¬јЌ»≈)
    Set si = Ћист“ихо(Ћ_—»)
    If об Is Nothing Or si Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo —бой
    Set ws = ѕересоздатьЋист(Ћ_ћ≈—я„Ќџ…)

    ws.Cells(1, 1).Value = "ѕЋјЌ-√–ј‘» "
    ws.Cells(2, 1).Value = "технического обслуживани€ средств неразрушающего контрол€ на " & _
                           Ќазваниећес€ца(мес€ц) & " " & год & " года"
    ws.Range("A1:H1").Merge
    ws.Range("A2:H2").Merge
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:A2").HorizontalAlignment = xlCenter

    r = Ўапка(ws, 4, Array("є п/п", "є карточки", "Ќаименование средства Ќ ", " атегори€", _
                           "ѕериодичность", "ѕоследнее обслуживание", "ѕланова€ дата", _
                           "“рудозатраты, ч", "ќтметка о выполнении"), _
              Array(6, 11, 46, 10, 24, 18, 14, 13, 26))

    last = ѕоследн€€—трока(об, 1)
    For i = ѕ≈–¬јя_—“–ќ ј To last
        интервал = CLng(Val(об.Cells(i, ColReq(об, "»нтервал, дней")).Value))
        If интервал > 0 Then
            последнее = ƒатаѕоследнего(CStr(об.Cells(i, ColReq(об, "ID средства Ќ ")).Value), _
                                       CStr(об.Cells(i, ColReq(об, " атегори€")).Value))
            сколько = —колько¬ћес€це(интервал, последнее, началоћес€ца, конецћес€ца, следующее)
            If сколько > 0 Then
                n = n + 1
                трудозатраты = сколько * „исло»ли0(об.Cells(i, ColReq(об, "“рудозатраты за одну операцию, ч")).Value)
                итого = итого + трудозатраты
                ws.Cells(r, 1).Value = n
                ws.Cells(r, 2).Value = об.Cells(i, ColReq(об, "є карточки учЄта")).Value
                ws.Cells(r, 3).Value = об.Cells(i, ColReq(об, "Ќаименование")).Value
                ws.Cells(r, 4).Value = об.Cells(i, ColReq(об, " атегори€")).Value
                ws.Cells(r, 5).Value = об.Cells(i, ColReq(об, "ѕериодичность")).Value & _
                                       IIf(сколько > 1, " Ч " & сколько & " раз за мес€ц", "")
                ws.Cells(r, 6).Value = IIf(последнее = 0, "не проводилось", _
                                           Format$(последнее, "dd.mm.yyyy"))
                ws.Cells(r, 7).Value = Format$(следующее, "dd.mm.yyyy")
                ws.Cells(r, 8).Value = Round(трудозатраты, 2)
                ќформить—троку ws, r, 9
                If последнее = 0 Then ws.Cells(r, 6).Font.Color = RGB(156, 0, 6)
                r = r + 1
            End If
        End If
    Next i

    If n = 0 Then
        ws.Cells(r, 1).Value = "Ќа этот мес€ц обслуживание не запланировано."
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Merge
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    Else
        ws.Cells(r, 3).Value = "»“ќ√ќ"
        ws.Cells(r, 8).Value = Round(итого, 2)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Font.Bold = True
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = RGB(226, 234, 246)
        ќформить—троку ws, r, 9
    End If

    Ќастроитьѕечать ws, "ѕлан-график “ќ на " & Ќазваниећес€ца(мес€ц) & " " & год, 4
    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "ћес€чный план-график построен: позиций " & n & _
           ", трудозатраты " & Round(итого, 2) & " ч." & vbCrLf & vbCrLf & _
           " расным помечено оборудование, по которому обслуживание не проводилось ни разу Ч " & _
           "дл€ него план считаетс€ от даты ввода в эксплуатацию.", _
           vbInformation, "ѕлан-график ѕѕ–"
    Exit Sub
—бой:
    Application.ScreenUpdating = True
    MsgBox "Ќе удалось построить план-график." & vbCrLf & "ќшибка: " & Err.Description, _
           vbCritical, "ѕлан-график ѕѕ–"
End Sub


' —колько раз меропри€тие попадает в мес€ц и когда первый раз.
' —читаем от последнего фактического, а если его не было Ч от начала
' мес€ца: иначе прибор, купленный три года назад, дал бы сотню
' Ђпросроченныхї операций задним числом.
Private Function —колько¬ћес€це(интервал As Long, последнее As Date, _
                                началоћес€ца As Date, конецћес€ца As Date, _
                                ByRef первое As Date) As Long
    Dim д As Date
    Dim n As Long

    If последнее = 0 Then
        д = началоћес€ца
    Else
        д = последнее + интервал
        If д < началоћес€ца Then
            ' ѕросрочено: показываем один раз в начале мес€ца, а не всю
            ' пропущенную серию Ч план должен быть выполнимым.
            д = началоћес€ца
        End If
    End If

    первое = 0
    Do While д <= конецћес€ца
        If д >= началоћес€ца Then
            n = n + 1
            If первое = 0 Then первое = д
        End If
        If интервал = 1 Then
            д = —ледующий–абочийƒень(д)
        Else
            д = д + интервал
        End If
        If n > 40 Then Exit Do          ' “ќ-1: больше рабочих дней в мес€це не бывает
    Loop
    —колько¬ћес€це = n
End Function


Private Function —ледующий–абочийƒень(д As Date) As Date
    Dim x As Date
    x = д + 1
    Do While Weekday(x, vbMonday) > 5
        x = x + 1
    Loop
    —ледующий–абочийƒень = x
End Function


Private Function ƒатаѕоследнего(ид—редства As String, категори€ As String) As Date
    Dim ws As Worksheet
    Dim i As Long, last As Long
    Dim д As Date

    If Trim$(ид—редства) = "" Then Exit Function
    Set ws = Ћист“ихо(Ћ_‘ј “)
    If ws Is Nothing Then Exit Function
    last = ѕоследн€€—трока(ws, 1)
    For i = ѕ≈–¬јя_—“–ќ ј To last
        If StrComp(Trim$(CStr(ws.Cells(i, ColReq(ws, "ID средства Ќ ")).Value)), _
                   Trim$(ид—редства), vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(ws.Cells(i, ColReq(ws, " атегори€")).Value)), _
                       Trim$(категори€), vbTextCompare) = 0 Then
                If IsDate(ws.Cells(i, ColReq(ws, "ƒата проведени€")).Value) Then
                    д = CDate(ws.Cells(i, ColReq(ws, "ƒата проведени€")).Value)
                    If д > ƒатаѕоследнего Then ƒатаѕоследнего = д
                End If
            End If
        End If
    Next i
End Function


' -- «а€вка на поверку -------------------------------------------------

Public Sub «а€вкаЌаѕоверку()
    Dim ws As Worksheet
    Dim si As Worksheet
    Dim год As Long
    Dim i As Long, last As Long, r As Long, n As Long
    Dim д As Date
    Dim без As Long
    Dim спискеЅез As String

    If ƒанные() Is Nothing Then Exit Sub
    год = CLng(Val(InputBox("«а€вка на поверку на какой год?", "«а€вка на поверку", _
                            CStr(Year(Date) + 1))))
    If год = 0 Then Exit Sub

    Set si = Ћист“ихо(Ћ_—»)
    If si Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo —бой
    Set ws = ѕересоздатьЋист(Ћ_«јя¬ ј)

    ws.Cells(1, 1).Value = "«јя¬ ј"
    ws.Cells(2, 1).Value = "на поверку средств измерений на " & год & " год"
    ws.Cells(3, 1).Value = Ќастройка("ѕолное наименование ЋЌ ")
    ws.Range("A1:H1").Merge
    ws.Range("A2:H2").Merge
    ws.Range("A3:H3").Merge
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:A3").HorizontalAlignment = xlCenter

    r = Ўапка(ws, 5, Array("є п/п", "Ќаименование, тип", "«аводской є", "√од выпуска", _
                           "ƒата последней поверки", "ƒействительно до", " вартал подачи", _
                           "ѕримечание"), _
              Array(6, 44, 18, 11, 18, 16, 13, 26))

    last = ѕоследн€€—трока(si, 1)
    For i = ѕ≈–¬јя_—“–ќ ј To last
        If StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "—осто€ние")).Value)), "—писан", vbTextCompare) <> 0 Then
            If StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "¬ид метрологического контрол€")).Value)), _
                       "Ќе подлежит", vbTextCompare) <> 0 Then
                If IsDate(si.Cells(i, ColReq(si, "ƒействительно до")).Value) Then
                    д = CDate(si.Cells(i, ColReq(si, "ƒействительно до")).Value)
                    If Year(д) = год Then
                        n = n + 1
                        ws.Cells(r, 1).Value = n
                        ws.Cells(r, 2).Value = Trim$(CStr(si.Cells(i, ColReq(si, "Ќаименование")).Value) & _
                                                     " " & CStr(si.Cells(i, ColReq(si, "“ип / обозначение")).Value))
                        ws.Cells(r, 3).Value = si.Cells(i, ColReq(si, "«аводской є")).Value
                        ws.Cells(r, 4).Value = si.Cells(i, ColReq(si, "√од выпуска")).Value
                        ws.Cells(r, 5).Value = si.Cells(i, ColReq(si, "ƒата последней поверки")).Value
                        ws.Cells(r, 6).Value = д
                        ws.Cells(r, 7).Value = "IV кв. " & (год - 1) & " Ч " &  вартал(д) & " кв. " & год
                        ws.Cells(r, 5).NumberFormat = "DD.MM.YYYY"
                        ws.Cells(r, 6).NumberFormat = "DD.MM.YYYY"
                        ќформить—троку ws, r, 8
                        ws.Cells(r, 9).Value =  вартал(д)      ' служебное, дл€ сортировки
                        r = r + 1
                    End If
                Else
                    без = без + 1
                    If Len(спискеЅез) < 300 Then
                        спискеЅез = спискеЅез & vbCrLf & "  - " & _
                                    CStr(si.Cells(i, ColReq(si, "Ќаименование")).Value) & _
                                    " зав. є " & CStr(si.Cells(i, ColReq(si, "«аводской є")).Value)
                    End If
                End If
            End If
        End If
    Next i

    ' —ортировка по кварталу: за€вка так и читаетс€ как график.
    If n > 1 Then
        ws.Range(ws.Cells(6, 1), ws.Cells(r - 1, 9)).Sort _
            Key1:=ws.Cells(6, 9), Order1:=xlAscending, Header:=xlNo
        For i = 6 To r - 1
            ws.Cells(i, 1).Value = i - 5
        Next i
    End If
    ws.Columns(9).Delete

    Ќастроитьѕечать ws, "«а€вка на поверку на " & год & " год", 5
    ws.Activate
    Application.ScreenUpdating = True

    MsgBox "«а€вка на " & год & " год готова: приборов " & n & "." & _
           IIf(без = 0, "", vbCrLf & vbCrLf & "Ќе попали в за€вку Ч не заполнена дата поверки (" & _
               без & "):" & спискеЅез & vbCrLf & _
               "ѕроверьте их: молча выпасть из за€вки они не должны."), _
           vbInformation, "«а€вка на поверку"
    Exit Sub
—бой:
    Application.ScreenUpdating = True
    MsgBox "Ќе удалось построить за€вку." & vbCrLf & "ќшибка: " & Err.Description, _
           vbCritical, "«а€вка на поверку"
End Sub


' -- ѕересчЄт п€тилетнего плана ----------------------------------------

' “олько дл€ Ќќ¬ќ√ќ п€тилетнего графика. ƒействующий 605-2025.ѕЋ подписан,
' и пересчЄт разошЄлс€ бы с подписанным документом.
Public Sub ѕересчитатьѕ€тилетний()
    Dim первый√од As Long
    If MsgBox("ѕересчЄт заменит помес€чную раскладку на листе Ђѕѕ–_ѕланї " & _
              "расчЄтной." & vbCrLf & vbCrLf & _
              "ƒействующий график 605-2025.ѕЋ подписан, и расчЄт даст другие числа: " & _
              "например, декадному “ќ он даст 36 операций в год вместо ваших 32." & vbCrLf & vbCrLf & _
              "ƒелать это стоит только когда составл€ете Ќќ¬џ… п€тилетний график. " & _
              "ѕродолжить?", vbYesNo + vbExclamation, "ѕересчЄт плана") = vbNo Then Exit Sub

    первый√од = CLng(Val(InputBox("— какого года начинаетс€ новый п€тилетний график?", _
                                  "ѕересчЄт плана", CStr(Year(Date) + 1))))
    If первый√од = 0 Then Exit Sub
    ѕересчитатьЋистѕлана первый√од
End Sub


Private Sub ѕересчитатьЋистѕлана(первый√од As Long)
    Dim ws As Worksheet
    Dim об As Worksheet
    Dim i As Long, last As Long, r As Long
    Dim y As Long, m As Long
    Dim интервал As Long
    Dim заќперацию As Double
    Dim вћес€ц As Long
    Dim колонка As Long
    Dim карточка As String

    Set ws = Ћист“ихо("ѕѕ–_ѕлан")
    Set об = Ћист“ихо(Ћ_ќЅ—Ћ”∆»¬јЌ»≈)
    If ws Is Nothing Or об Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    last = ѕоследн€€—трока(об, 1)
    For r = ѕ≈–¬јя_—“–ќ ј To ѕоследн€€—трока(ws, 2)
        карточка = Trim$(CStr(ws.Cells(r, 2).Value))
        интервал = 0
        For i = ѕ≈–¬јя_—“–ќ ј To last
            If StrComp(Trim$(CStr(об.Cells(i, ColReq(об, "є карточки учЄта")).Value)), карточка, vbTextCompare) = 0 Then
                If StrComp(Trim$(CStr(об.Cells(i, ColReq(об, " атегори€")).Value)), _
                           Trim$(CStr(ws.Cells(r, 4).Value)), vbTextCompare) = 0 Then
                    интервал = CLng(Val(об.Cells(i, ColReq(об, "»нтервал, дней")).Value))
                    заќперацию = „исло»ли0(об.Cells(i, ColReq(об, "“рудозатраты за одну операцию, ч")).Value)
                    Exit For
                End If
            End If
        Next i
        If интервал > 0 Then
            For y = 0 To 4
                For m = 1 To 12
                    колонка = 6 + y * 13 + (m - 1)
                    вћес€ц = ќпераций¬ћес€це(интервал, первый√од + y, m)
                    If вћес€ц > 0 Then
                        ws.Cells(r, колонка).Value = Round(вћес€ц * заќперацию, 2)
                    Else
                        ws.Cells(r, колонка).ClearContents
                    End If
                Next m
            Next y
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox "ѕомес€чна€ раскладка пересчитана по периодичности начина€ с " & первый√од & " года.", _
           vbInformation, "ѕересчЄт плана"
End Sub


Private Function ќпераций¬ћес€це(интервал As Long, год As Long, мес€ц As Long) As Long
    Dim дней As Long
    Dim д As Date
    Dim n As Long

    дней = Day(DateSerial(год, мес€ц + 1, 0))
    Select Case интервал
        Case 1                          ' “ќ-1: по рабочим дн€м
            For n = 1 To дней
                д = DateSerial(год, мес€ц, n)
                If Weekday(д, vbMonday) <= 5 Then ќпераций¬ћес€це = ќпераций¬ћес€це + 1
            Next n
        Case Is >= 85                   ' “ќ-4: раз в квартал
            If мес€ц Mod 3 = 1 Then ќпераций¬ћес€це = 1
        Case Is >= 25                   ' “ќ-3: раз в мес€ц
            ќпераций¬ћес€це = 1
        Case Else                       ' “ќ-2 и прочее: по интервалу
            ќпераций¬ћес€це = Int(дней / интервал)
            If ќпераций¬ћес€це = 0 Then ќпераций¬ћес€це = 1
    End Select
End Function


' -- ¬спомогательное ---------------------------------------------------

Private Function —проситьѕериод(ByRef год As Long, ByRef мес€ц As Long) As Boolean
    Dim s As String
    s = InputBox("Ќа какой мес€ц строить план-график?" & vbCrLf & _
                 "‘ормат: ћћ.√√√√", "ѕлан-график ѕѕ–", Format$(DateAdd("m", 1, Date), "mm.yyyy"))
    If Trim$(s) = "" Then Exit Function
    If InStr(s, ".") = 0 Then
        MsgBox "Ќужен формат ћћ.√√√√, например 10.2026.", vbExclamation, "ѕлан-график ѕѕ–"
        Exit Function
    End If
    мес€ц = CLng(Val(Split(s, ".")(0)))
    год = CLng(Val(Split(s, ".")(1)))
    If мес€ц < 1 Or мес€ц > 12 Or год < 2000 Then
        MsgBox "Ќе разобрал период Ђ" & s & "ї. Ќужен формат ћћ.√√√√.", _
               vbExclamation, "ѕлан-график ѕѕ–"
        Exit Function
    End If
    —проситьѕериод = True
End Function


Private Function Ќазваниећес€ца(мес€ц As Long) As String
    Ќазваниећес€ца = Split("€нварь,февраль,март,апрель,май,июнь,июль,август," & _
                           "сент€брь,окт€брь,но€брь,декабрь", ",")(мес€ц - 1)
End Function


Private Function  вартал(д As Date) As Long
     вартал = Int((Month(д) - 1) / 3) + 1
End Function


Private Function „исло»ли0(значение As Variant) As Double
    If IsNumeric(значение) Then
        If Trim$(CStr(значение)) <> "" Then „исло»ли0 = CDbl(значение)
    End If
End Function



Private Function Ўапка(ws As Worksheet, r As Long, заголовки As Variant, ширины As Variant) As Long
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
    Ўапка = r + 1
End Function


Private Sub ќформить—троку(ws As Worksheet, r As Long, колонок As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, колонок))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 9
    End With
    ws.Cells(r, 1).HorizontalAlignment = xlCenter
End Sub


Private Function ѕересоздатьЋист(им€ As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(им€).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = им€
    ws.Tab.Color = RGB(46, 84, 150)
    Set ѕересоздатьЋист = ws
End Function


Private Sub Ќастроитьѕечать(ws As Worksheet, подпись As String, строкаЎапки As Long)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$" & строкаЎапки & ":$" & строкаЎапки
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(1)
        .CenterFooter = "&P"
        .LeftFooter = подпись
    End With
End Sub
