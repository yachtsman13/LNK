Attribute VB_Name = "modBackup"
Option Explicit

' =======================================================================
'  –езервное копирование, уровень 1 Ч локальные копии.
'
'   опи€ данных делаетс€ при закрытии книги, одна в день.  опи€ папки
'  счЄтчиков Ч при каждой выдаче номера: файлы там по дес€тку байт,
'  а цена потери максимальна€.
'
'  —чЄтчики после восстановлени€ Ќ≈ берутс€ из копии: копи€ недельной
'  давности вернЄт номер, который за неделю уже кому-то выдали.
'  »х пересчитывает ¬осстановить—чЄтчики»з∆урнала.
' =======================================================================


Public Sub —делатьЋокальную опию(Optional тихо As Boolean = False)
    Dim источник As String
    Dim папка As String
    Dim цель As String
    Dim размер As Double

    On Error GoTo —бой
    If ƒанные() Is Nothing Then Exit Sub
    источник = ƒанные().FullName
    If источник = "" Then Exit Sub

    папка = ѕапка опий()
    MkPath папка
    цель = папка & "ЋЌ _ƒанные_" & Format$(Date, "yyyy-mm-dd") & ".xlsx"

    ƒанные().Save
    FileCopy источник, цель
    размер = FileLen(цель) / 1048576

    «аписать¬∆урнал опировани€ "ƒанные (лист-книга)", цель, размер, "”спешно"
    ѕочистить—тарые опии папка

    If Not тихо Then
        MsgBox " опи€ сохранена:" & vbCrLf & цель, vbInformation, "–езервное копирование"
    End If
    Exit Sub

—бой:
    «аписать¬∆урнал опировани€ "ƒанные (лист-книга)", цель, 0, "ќЎ»Ѕ ј: " & Err.Description
    If Not тихо Then
        MsgBox "Ќе удалось сделать копию." & vbCrLf & "ќшибка: " & Err.Description, _
               vbExclamation, "–езервное копирование"
    End If
End Sub


' —чЄтчики копируютс€ при каждой выдаче номера Ч вызываетс€ из modNumbers.
Public Sub —копировать—чЄтчики()
    Dim откуда As String
    Dim куда As String
    Dim файл As String

    On Error Resume Next
    откуда = ѕапка—чЄтчиков()
    If Not ѕапка—уществует(откуда) Then Exit Sub
    куда = ѕапка опий() & "—чЄтчики" & Application.PathSeparator
    MkPath куда

    файл = ѕервый‘айл(откуда & "*.txt")
    Do While файл <> ""
        FileCopy откуда & файл, куда & файл
        файл = Dir
    Loop
End Sub


' ≈жедневные копии живут заданный срок; копии за первое число мес€ца
' остаютс€ дольше Ч по ним восстанавливают состо€ние на давнюю дату.
Private Sub ѕочистить—тарые опии(папка As String)
    Dim файл As String
    Dim дата As Date
    Dim s As String
    Dim держатьƒней As Long
    Dim держатьћес€цев As Long
    Dim удалено As Long

    держатьƒней = CLng(Val(Ќастройка("’ранить ежедневных копий, дней")))
    держатьћес€цев = CLng(Val(Ќастройка("’ранить помес€чных копий, мес.")))
    If держатьƒней <= 0 Then держатьƒней = 30
    If держатьћес€цев <= 0 Then держатьћес€цев = 12

    On Error Resume Next
    файл = ѕервый‘айл(папка & "ЋЌ _ƒанные_*.xlsx")
    Do While файл <> ""
        s = Mid$(файл, Len("ЋЌ _ƒанные_") + 1, 10)
        If IsDate(s) Then
            дата = CDate(s)
            If Day(дата) = 1 Then
                If DateDiff("m", дата, Date) > держатьћес€цев Then
                    Kill папка & файл
                    удалено = удалено + 1
                End If
            ElseIf DateDiff("d", дата, Date) > держатьƒней Then
                Kill папка & файл
                удалено = удалено + 1
            End If
        End If
        файл = Dir
    Loop
End Sub


Public Sub «аписать¬∆урнал опировани€(что As String, куда As String, _
                                      размерћЅ As Double, результат As String)
    Dim ws As Worksheet
    Dim r As Long

    On Error Resume Next
    Set ws = Ћист(Ћ_ ќѕ»–ќ¬јЌ»≈)
    If ws Is Nothing Then Exit Sub
    r = ѕоследн€€—трока(ws, 1) + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "DD.MM.YYYY hh:mm"
    ws.Cells(r, 2).Value = что
    ws.Cells(r, 3).Value = куда
    ws.Cells(r, 4).Value = Round(размерћЅ, 2)
    ws.Cells(r, 5).Value = результат
    ws.Cells(r, 6).Value = “екущий—отрудник()
End Sub


' -- ¬осстановление счЄтчиков ------------------------------------------

' —чЄтчики восстанавливаютс€ из журнала, а не из копии: журнал первичен,
' счЄтчик идЄт следом. “от же механизм используетс€ при переходе на
' систему Ч стартовые значени€ берутс€ отсюда же.
Public Sub ¬осстановить—чЄтчики»з∆урнала()
    Dim ws As Worksheet
    Dim методы As Variant
    Dim i As Long
    Dim метод As String
    Dim год As Long
    Dim r As Long, last As Long
    Dim макс As Long
    Dim v As Variant
    Dim было As Long
    Dim отчЄт As String
    Dim пропуски As String

    Set ws = Ћист(Ћ_∆”–ЌјЋ)
    If ws Is Nothing Then Exit Sub

    год = CLng(Val(InputBox("«а какой год пересчитать счЄтчики?", _
                            "—чЄтчики", CStr(Year(Date)))))
    If год = 0 Then Exit Sub

    методы = Split(ћ≈“ќƒџ, ",")
    last = ѕоследн€€—трока(ws, ColReq(ws, "ƒата контрол€"))

    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        макс = 0
        For r = ѕ≈–¬јя_—“–ќ ј To last
            If IsDate(ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value) Then
                If Year(CDate(ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value)) = год Then
                    v = ws.Cells(r, ColReq(ws, метод & " є")).Value
                    If IsNumeric(v) Then
                        If Trim$(CStr(v)) <> "" Then
                            If CLng(v) > макс Then макс = CLng(v)
                        End If
                    End If
                End If
            End If
        Next r
        было = ѕосмотреть—чЄтчик(метод, год)
        отчЄт = отчЄт & vbCrLf & "  " & метод & ": в журнале максимум " & макс & _
                ", в счЄтчике " & было & " -> станет " & макс
        пропуски = пропуски & ѕропускиѕоћетоду(ws, метод, год, макс, last)
    Next i

    If MsgBox("ѕересчЄт счЄтчиков за " & год & " год по журналу:" & vbCrLf & отчЄт & _
              IIf(пропуски = "", "", vbCrLf & vbCrLf & "ѕропуски в нумерации:" & пропуски) & _
              vbCrLf & vbCrLf & "—верьте с бумажными журналами и подтвердите. " & _
              "”становить счЄтчики?", vbYesNo + vbQuestion, "—чЄтчики") = vbNo Then Exit Sub

    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        макс = 0
        For r = ѕ≈–¬јя_—“–ќ ј To last
            If IsDate(ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value) Then
                If Year(CDate(ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value)) = год Then
                    v = ws.Cells(r, ColReq(ws, метод & " є")).Value
                    If IsNumeric(v) Then
                        If Trim$(CStr(v)) <> "" Then
                            If CLng(v) > макс Then макс = CLng(v)
                        End If
                    End If
                End If
            End If
        Next r
        ”становить—чЄтчик метод, год, макс
    Next i

    MsgBox "—чЄтчики установлены. —ледующа€ регистраци€ продолжит нумерацию с них.", _
           vbInformation, "—чЄтчики"
End Sub


' ѕропуск Ч это либо потер€нна€ запись, либо аннулированный на бумаге
' номер. » то и другое разбираетс€ до запуска, а не в декабре.
Private Function ѕропускиѕоћетоду(ws As Worksheet, метод As String, год As Long, _
                                  макс As Long, last As Long) As String
    Dim есть() As Boolean
    Dim r As Long, n As Long
    Dim v As Variant
    Dim список As String
    Dim сколько As Long

    If макс = 0 Then Exit Function
    ReDim есть(1 To макс)
    For r = ѕ≈–¬јя_—“–ќ ј To last
        If IsDate(ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value) Then
            If Year(CDate(ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value)) = год Then
                v = ws.Cells(r, ColReq(ws, метод & " є")).Value
                If IsNumeric(v) Then
                    If Trim$(CStr(v)) <> "" Then
                        n = CLng(v)
                        If n >= 1 And n <= макс Then есть(n) = True
                    End If
                End If
            End If
        End If
    Next r

    For n = 1 To макс
        If Not есть(n) Then
            сколько = сколько + 1
            If сколько <= 12 Then список = список & IIf(список = "", "", ", ") & CStr(n)
        End If
    Next n
    If сколько > 12 Then список = список & " Е всего " & сколько
    If список <> "" Then ѕропускиѕоћетоду = vbCrLf & "  " & метод & ": " & список
End Function
