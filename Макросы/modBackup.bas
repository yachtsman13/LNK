Attribute VB_Name = "modBackup"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Резервное копирование, уровень 1 — локальные копии.
'
'  Копия данных делается при закрытии книги, одна в день. Копия папки
'  счётчиков — при каждой выдаче номера: файлы там по десятку байт,
'  а цена потери максимальная.
'
'  Счётчики после восстановления НЕ берутся из копии: копия недельной
'  давности вернёт номер, который за неделю уже кому-то выдали.
'  Их пересчитывает ВосстановитьСчётчикиИзЖурнала.
' ═══════════════════════════════════════════════════════════════════════


Public Sub СделатьЛокальнуюКопию(Optional тихо As Boolean = False)
    Dim источник As String
    Dim папка As String
    Dim цель As String
    Dim размер As Double

    On Error GoTo Сбой
    If Данные() Is Nothing Then Exit Sub
    источник = Данные().FullName
    If источник = "" Then Exit Sub

    папка = ПапкаКопий()
    MkPath папка
    цель = папка & "ЛНК_Данные_" & Format$(Date, "yyyy-mm-dd") & ".xlsx"

    Данные().Save
    FileCopy источник, цель
    размер = FileLen(цель) / 1048576

    ЗаписатьВЖурналКопирования "Данные (лист-книга)", цель, размер, "Успешно"
    ПочиститьСтарыеКопии папка

    If Not тихо Then
        MsgBox "Копия сохранена:" & vbCrLf & цель, vbInformation, "Резервное копирование"
    End If
    Exit Sub

Сбой:
    ЗаписатьВЖурналКопирования "Данные (лист-книга)", цель, 0, "ОШИБКА: " & Err.Description
    If Not тихо Then
        MsgBox "Не удалось сделать копию." & vbCrLf & "Ошибка: " & Err.Description, _
               vbExclamation, "Резервное копирование"
    End If
End Sub


' Счётчики копируются при каждой выдаче номера — вызывается из modNumbers.
Public Sub СкопироватьСчётчики()
    Dim откуда As String
    Dim куда As String
    Dim файл As String

    On Error Resume Next
    откуда = ПапкаСчётчиков()
    If Dir(откуда, vbDirectory) = "" Then Exit Sub
    куда = ПапкаКопий() & "Счётчики" & Application.PathSeparator
    MkPath куда

    файл = Dir(откуда & "*.txt")
    Do While файл <> ""
        FileCopy откуда & файл, куда & файл
        файл = Dir
    Loop
End Sub


' Ежедневные копии живут заданный срок; копии за первое число месяца
' остаются дольше — по ним восстанавливают состояние на давнюю дату.
Private Sub ПочиститьСтарыеКопии(папка As String)
    Dim файл As String
    Dim дата As Date
    Dim s As String
    Dim держатьДней As Long
    Dim держатьМесяцев As Long
    Dim удалено As Long

    держатьДней = CLng(Val(Настройка("Хранить ежедневных копий, дней")))
    держатьМесяцев = CLng(Val(Настройка("Хранить помесячных копий, мес.")))
    If держатьДней <= 0 Then держатьДней = 30
    If держатьМесяцев <= 0 Then держатьМесяцев = 12

    On Error Resume Next
    файл = Dir(папка & "ЛНК_Данные_*.xlsx")
    Do While файл <> ""
        s = Mid$(файл, Len("ЛНК_Данные_") + 1, 10)
        If IsDate(s) Then
            дата = CDate(s)
            If Day(дата) = 1 Then
                If DateDiff("m", дата, Date) > держатьМесяцев Then
                    Kill папка & файл
                    удалено = удалено + 1
                End If
            ElseIf DateDiff("d", дата, Date) > держатьДней Then
                Kill папка & файл
                удалено = удалено + 1
            End If
        End If
        файл = Dir
    Loop
End Sub


Public Sub ЗаписатьВЖурналКопирования(что As String, куда As String, _
                                      размерМБ As Double, результат As String)
    Dim ws As Worksheet
    Dim r As Long

    On Error Resume Next
    Set ws = Лист(Л_КОПИРОВАНИЕ)
    If ws Is Nothing Then Exit Sub
    r = ПоследняяСтрока(ws, 1) + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "DD.MM.YYYY hh:mm"
    ws.Cells(r, 2).Value = что
    ws.Cells(r, 3).Value = куда
    ws.Cells(r, 4).Value = Round(размерМБ, 2)
    ws.Cells(r, 5).Value = результат
    ws.Cells(r, 6).Value = ТекущийСотрудник()
End Sub


' ── Восстановление счётчиков ──────────────────────────────────────────

' Счётчики восстанавливаются из журнала, а не из копии: журнал первичен,
' счётчик идёт следом. Тот же механизм используется при переходе на
' систему — стартовые значения берутся отсюда же.
Public Sub ВосстановитьСчётчикиИзЖурнала()
    Dim ws As Worksheet
    Dim методы As Variant
    Dim i As Long
    Dim метод As String
    Dim год As Long
    Dim r As Long, last As Long
    Dim макс As Long
    Dim v As Variant
    Dim было As Long
    Dim отчёт As String
    Dim пропуски As String

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub

    год = CLng(Val(InputBox("За какой год пересчитать счётчики?", _
                            "Счётчики", CStr(Year(Date)))))
    If год = 0 Then Exit Sub

    методы = Split(МЕТОДЫ, ",")
    last = ПоследняяСтрока(ws, ColReq(ws, "Дата контроля"))

    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        макс = 0
        For r = ПЕРВАЯ_СТРОКА To last
            If IsDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value) Then
                If Year(CDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value)) = год Then
                    v = ws.Cells(r, ColReq(ws, метод & " №")).Value
                    If IsNumeric(v) Then
                        If Trim$(CStr(v)) <> "" Then
                            If CLng(v) > макс Then макс = CLng(v)
                        End If
                    End If
                End If
            End If
        Next r
        было = ПосмотретьСчётчик(метод, год)
        отчёт = отчёт & vbCrLf & "  " & метод & ": в журнале максимум " & макс & _
                ", в счётчике " & было & " → станет " & макс
        пропуски = пропуски & ПропускиПоМетоду(ws, метод, год, макс, last)
    Next i

    If MsgBox("Пересчёт счётчиков за " & год & " год по журналу:" & vbCrLf & отчёт & _
              IIf(пропуски = "", "", vbCrLf & vbCrLf & "Пропуски в нумерации:" & пропуски) & _
              vbCrLf & vbCrLf & "Сверьте с бумажными журналами и подтвердите. " & _
              "Установить счётчики?", vbYesNo + vbQuestion, "Счётчики") = vbNo Then Exit Sub

    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        макс = 0
        For r = ПЕРВАЯ_СТРОКА To last
            If IsDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value) Then
                If Year(CDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value)) = год Then
                    v = ws.Cells(r, ColReq(ws, метод & " №")).Value
                    If IsNumeric(v) Then
                        If Trim$(CStr(v)) <> "" Then
                            If CLng(v) > макс Then макс = CLng(v)
                        End If
                    End If
                End If
            End If
        Next r
        УстановитьСчётчик метод, год, макс
    Next i

    MsgBox "Счётчики установлены. Следующая регистрация продолжит нумерацию с них.", _
           vbInformation, "Счётчики"
End Sub


' Пропуск — это либо потерянная запись, либо аннулированный на бумаге
' номер. И то и другое разбирается до запуска, а не в декабре.
Private Function ПропускиПоМетоду(ws As Worksheet, метод As String, год As Long, _
                                  макс As Long, last As Long) As String
    Dim есть() As Boolean
    Dim r As Long, n As Long
    Dim v As Variant
    Dim список As String
    Dim сколько As Long

    If макс = 0 Then Exit Function
    ReDim есть(1 To макс)
    For r = ПЕРВАЯ_СТРОКА To last
        If IsDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value) Then
            If Year(CDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value)) = год Then
                v = ws.Cells(r, ColReq(ws, метод & " №")).Value
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
    If сколько > 12 Then список = список & " … всего " & сколько
    If список <> "" Then ПропускиПоМетоду = vbCrLf & "  " & метод & ": " & список
End Function
