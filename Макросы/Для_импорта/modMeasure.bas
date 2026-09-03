Attribute VB_Name = "modMeasure"
Option Explicit

' =======================================================================
'  Ввод результатов замеров.
'
'  Одна строка листа «Замеры» — одна точка. Для УЗТ заполняются пять
'  замеров, минимум считается формулой; для ТВ — твёрдость в МПа, НВ
'  пересчитывается по таблице, снятой с вашего файла.
'
'  В шаблоне протокола четыре места под точки: #Т1С1…#Т4С6 у УЗТ и
'  #Т1_HB…#Т4_Расч у ТВ. Точек может быть больше — тогда в документ идут
'  первые четыре, а об остальных программа честно говорит, а не молча
'  их теряет.
' =======================================================================

Public Const Л_ЗАМЕРЫ As String = "Замеры"
Public Const Л_ДАННЫЕ_ПРОТОКОЛОВ As String = "Данные_протоколов"
Public Const Л_ДАННЫЕ_АЭ As String = "Данные_АЭ"
Private Const МЕСТ_В_ШАБЛОНЕ As Long = 4


' Добавить точки замеров к текущей строке журнала.
Public Sub ДобавитьТочкиЗамеров()
    Dim ж As Worksheet
    Dim ws As Worksheet
    Dim r As Long
    Dim идЗаписи As String
    Dim метод As String
    Dim сколько As Long
    Dim i As Long
    Dim строка As Long
    Dim было As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub

    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))
    If идЗаписи = "" Then
        MsgBox "У этой строки нет ID записи — замеры не к чему привязать.", _
               vbExclamation, "Замеры"
        Exit Sub
    End If

    метод = СпроситьМетод(ж, r)
    If метод = "" Then Exit Sub

    сколько = CLng(Val(InputBox( _
        "Сколько точек замера добавить по методу " & метод & "?" & vbCrLf & vbCrLf & _
        "В шаблон протокола попадут первые " & МЕСТ_В_ШАБЛОНЕ & " — столько в нём мест. " & _
        "Точек можно завести больше: они останутся в системе и будут видны в журнале замеров.", _
        "Замеры", "4")))
    If сколько <= 0 Then Exit Sub
    If сколько > 100 Then
        MsgBox "Больше ста точек за раз — похоже на опечатку.", vbExclamation, "Замеры"
        Exit Sub
    End If

    Set ws = Лист(Л_ЗАМЕРЫ)
    If ws Is Nothing Then Exit Sub
    было = СколькоТочек(ws, идЗаписи, метод)
    строка = ПоследняяСтрока(ws, 1) + 1

    For i = 1 To сколько
        ws.Cells(строка, 1).Value = НовыйID(ws, "ЗМ-")
        ws.Cells(строка, ColReq(ws, "ID записи журнала")).Value = идЗаписи
        ws.Cells(строка, ColReq(ws, "Метод")).Value = метод
        ws.Cells(строка, ColReq(ws, "№ точки")).Value = было + i
        ws.Cells(строка, ColReq(ws, "Кем внесено")).Value = ТекущийСотрудник()
        ws.Cells(строка, ColReq(ws, "Изменено")).Value = Now
        строка = строка + 1
    Next i

    ws.Activate
    ws.Cells(строка - сколько, ColReq(ws, "Замер 1")).Select
    MsgBox "Добавлено точек: " & сколько & " (номера " & (было + 1) & "…" & (было + сколько) & ")." & _
           vbCrLf & vbCrLf & _
           IIf(метод = "УЗТ", "Заполняйте «Замер 1…5», минимум посчитается сам.", _
               "Заполняйте «Твёрдость, МПа» — НВ пересчитается по таблице."), _
           vbInformation, "Замеры"
End Sub


' Открыть замеры текущей строки журнала: показать, что уже введено.
Public Sub ПоказатьЗамерыЗаписи()
    Dim ж As Worksheet
    Dim ws As Worksheet
    Dim r As Long, i As Long, last As Long
    Dim идЗаписи As String
    Dim найдено As Long
    Dim перваяСтрока As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub
    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))

    Set ws = Лист(Л_ЗАМЕРЫ)
    If ws Is Nothing Then Exit Sub
    last = ПоследняяСтрока(ws, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(i, ColReq(ws, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            найдено = найдено + 1
            If перваяСтрока = 0 Then перваяСтрока = i
        End If
    Next i

    If найдено = 0 Then
        MsgBox "По записи " & идЗаписи & " замеров ещё нет." & vbCrLf & vbCrLf & _
               "Добавьте их пунктом «Добавить точки замеров».", vbInformation, "Замеры"
        Exit Sub
    End If

    ws.Activate
    ws.Cells(перваяСтрока, 1).Select
    ActiveWindow.ScrollRow = WorksheetFunction.Max(перваяСтрока - 3, 1)
End Sub


' Настройки аппаратуры АЭ от работы к работе меняются мало: те же
' преобразователи, та же полоса частот, то же усиление. Переписывать
' двенадцать полей каждый раз — верный способ ошибиться в одном из них,
' поэтому они копируются из последней заполненной работы, а потом
' правится только то, что действительно изменилось.
Public Sub СкопироватьНастройкиАЭ()
    Dim ж As Worksheet
    Dim ws As Worksheet
    Dim r As Long
    Dim идЗаписи As String
    Dim строка As Long
    Dim источник As Long
    Dim c As Long, колонок As Long
    Dim скопировано As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub

    If Not МетодПрименён(ж, r, "АЭ") Then
        MsgBox "В этой строке журнала не отмечен метод АЭ.", vbExclamation, "Настройки АЭ"
        Exit Sub
    End If

    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))
    Set ws = Лист(Л_ДАННЫЕ_АЭ)
    If ws Is Nothing Then Exit Sub

    строка = НайтиПоЗаписи(ws, идЗаписи)
    источник = ПоследняяЗаполненная(ws, строка)
    If источник = 0 Then
        If строка = 0 Then строка = ЗавестиСтрокуАЭ(ws, идЗаписи)
        ws.Activate
        ws.Cells(строка, 3).Select
        MsgBox "Копировать пока не с чего — это первая работа по АЭ." & vbCrLf & vbCrLf & _
               "Заполните настройки здесь; в следующий раз они скопируются сами.", _
               vbInformation, "Настройки АЭ"
        Exit Sub
    End If

    If строка = 0 Then строка = ЗавестиСтрокуАЭ(ws, идЗаписи)
    колонок = ws.Cells(СТРОКА_ЗАГОЛОВКОВ, ws.Columns.Count).End(xlToLeft).Column
    For c = 3 To колонок - 3          ' до служебных «Кем внесено», «Изменено», «Примечание»
        If Trim$(CStr(ws.Cells(строка, c).Value)) = "" Then
            If Trim$(CStr(ws.Cells(источник, c).Value)) <> "" Then
                ws.Cells(строка, c).Value = ws.Cells(источник, c).Value
                скопировано = скопировано + 1
            End If
        End If
    Next c
    ws.Cells(строка, ColReq(ws, "Изменено")).Value = Now

    ws.Activate
    ws.Cells(строка, 3).Select
    MsgBox "Скопировано полей: " & скопировано & " из работы " & _
           CStr(ws.Cells(источник, ColReq(ws, "ID записи журнала")).Value) & "." & vbCrLf & vbCrLf & _
           "Уже заполненные поля не затирались. Проверьте, что настройки " & _
           "действительно те же, и поправьте расхождения.", _
           vbInformation, "Настройки АЭ"
End Sub


Private Function НайтиПоЗаписи(ws As Worksheet, идЗаписи As String) As Long
    Dim i As Long, last As Long
    last = ПоследняяСтрока(ws, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(i, ColReq(ws, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            НайтиПоЗаписи = i
            Exit Function
        End If
    Next i
End Function


' Последняя строка с настройками, кроме текущей: из неё и копируем.
Private Function ПоследняяЗаполненная(ws As Worksheet, кроме As Long) As Long
    Dim i As Long, last As Long
    last = ПоследняяСтрока(ws, 1)
    For i = last To ПЕРВАЯ_СТРОКА Step -1
        If i <> кроме Then
            If Trim$(CStr(ws.Cells(i, 3).Value)) <> "" Then
                ПоследняяЗаполненная = i
                Exit Function
            End If
        End If
    Next i
End Function


Private Function ЗавестиСтрокуАЭ(ws As Worksheet, идЗаписи As String) As Long
    Dim r As Long
    r = ПоследняяСтрока(ws, 1) + 1
    ws.Cells(r, 1).Value = НовыйID(ws, "АЭ-")
    ws.Cells(r, ColReq(ws, "ID записи журнала")).Value = идЗаписи
    ws.Cells(r, ColReq(ws, "Кем внесено")).Value = ТекущийСотрудник()
    ws.Cells(r, ColReq(ws, "Изменено")).Value = Now
    ЗавестиСтрокуАЭ = r
End Function


' -- Чтение замеров для подстановки в документ -------------------------

' Заполняет плейсхолдеры точек в открытом документе Word.
' Возвращает число точек, не поместившихся в шаблон.
Public Function ПодставитьЗамеры(doc As Object, идЗаписи As String, метод As String) As Long
    Dim ws As Worksheet
    Dim i As Long, last As Long
    Dim точка As Long
    Dim всего As Long

    Set ws = ЛистТихо(Л_ЗАМЕРЫ)
    If ws Is Nothing Then Exit Function
    last = ПоследняяСтрока(ws, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If Совпадает(ws, i, идЗаписи, метод) Then
            всего = всего + 1
            If всего <= МЕСТ_В_ШАБЛОНЕ Then
                точка = всего
                If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
                    ПодставитьУЗТ doc, ws, i, точка
                ElseIf StrComp(метод, "ТВ", vbTextCompare) = 0 Then
                    ПодставитьТВ doc, ws, i, точка
                End If
            End If
        End If
    Next i

    ' Незанятые места в шаблоне очищаем, иначе в документе останутся
    ' видимые #Т3С1 и прочие — хуже, чем пустая клетка.
    For точка = всего + 1 To МЕСТ_В_ШАБЛОНЕ
        If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
            ОчиститьУЗТ doc, точка
        ElseIf StrComp(метод, "ТВ", vbTextCompare) = 0 Then
            ОчиститьТВ doc, точка
        End If
    Next точка

    If всего > МЕСТ_В_ШАБЛОНЕ Then ПодставитьЗамеры = всего - МЕСТ_В_ШАБЛОНЕ
End Function


Private Sub ПодставитьУЗТ(doc As Object, ws As Worksheet, r As Long, точка As Long)
    Dim c As Long
    ' С1…С5 — замеры, С6 — минимальное значение: так устроен ваш шаблон.
    For c = 1 To 5
        ЗаменитьВДокументе doc, "Т" & точка & "С" & c, _
                           Текст(ws.Cells(r, ColReq(ws, "Замер " & c)).Value)
    Next c
    ЗаменитьВДокументе doc, "Т" & точка & "С6", Текст(ws.Cells(r, ColReq(ws, "Минимум")).Value)
End Sub


Private Sub ПодставитьТВ(doc As Object, ws As Worksheet, r As Long, точка As Long)
    ЗаменитьВДокументе doc, "Т" & точка & "_HB", Текст(ws.Cells(r, ColReq(ws, "НВ (пересчёт)")).Value)
    ЗаменитьВДокументе doc, "Т" & точка & "_Расч", Текст(ws.Cells(r, ColReq(ws, "Твёрдость, МПа")).Value)
    ЗаменитьВДокументе doc, "Т" & точка & "_Норм", Текст(ws.Cells(r, ColReq(ws, "Нормируемое значение")).Value)
    ЗаменитьВДокументе doc, "Т" & точка & "_ЗавНомер", Текст(ws.Cells(r, ColReq(ws, "Зав. № прибора")).Value)
End Sub


Private Sub ОчиститьУЗТ(doc As Object, точка As Long)
    Dim c As Long
    For c = 1 To 6
        ЗаменитьВДокументе doc, "Т" & точка & "С" & c, ""
    Next c
End Sub


Private Sub ОчиститьТВ(doc As Object, точка As Long)
    ЗаменитьВДокументе doc, "Т" & точка & "_HB", ""
    ЗаменитьВДокументе doc, "Т" & точка & "_Расч", ""
    ЗаменитьВДокументе doc, "Т" & точка & "_Норм", ""
    ЗаменитьВДокументе doc, "Т" & точка & "_ЗавНомер", ""
End Sub


Private Function Совпадает(ws As Worksheet, r As Long, идЗаписи As String, метод As String) As Boolean
    If StrComp(Trim$(CStr(ws.Cells(r, ColReq(ws, "ID записи журнала")).Value)), _
               идЗаписи, vbTextCompare) <> 0 Then Exit Function
    Совпадает = (StrComp(Trim$(CStr(ws.Cells(r, ColReq(ws, "Метод")).Value)), _
                         метод, vbTextCompare) = 0)
End Function


Private Function СколькоТочек(ws As Worksheet, идЗаписи As String, метод As String) As Long
    Dim i As Long, last As Long
    last = ПоследняяСтрока(ws, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If Совпадает(ws, i, идЗаписи, метод) Then СколькоТочек = СколькоТочек + 1
    Next i
End Function


Private Function СпроситьМетод(ж As Worksheet, r As Long) As String
    Dim доступные As String
    Dim ответ As String

    If МетодПрименён(ж, r, "УЗТ") Then доступные = "УЗТ"
    If МетодПрименён(ж, r, "ТВ") Then доступные = доступные & IIf(доступные = "", "", ", ") & "ТВ"

    If доступные = "" Then
        MsgBox "В этой строке не отмечены ни УЗТ, ни ТВ." & vbCrLf & vbCrLf & _
               "Пока замеры вводятся только по этим двум методам: именно у них в " & _
               "шаблонах есть таблицы под точки. По остальным методам результаты " & _
               "пишутся текстом на листе «" & Л_ДАННЫЕ_ПРОТОКОЛОВ & "».", _
               vbExclamation, "Замеры"
        Exit Function
    End If
    If InStr(доступные, ",") = 0 Then
        СпроситьМетод = доступные
        Exit Function
    End If

    ответ = UCase$(Trim$(InputBox("По какому методу вводим замеры: " & доступные & "?", _
                                  "Замеры", "УЗТ")))
    If ответ = "УЗТ" Or ответ = "ТВ" Then СпроситьМетод = ответ
End Function


Private Function НовыйID(ws As Worksheet, префикс As String) As String
    Dim r As Long, last As Long, n As Long, m As Long
    Dim s As String
    last = ПоследняяСтрока(ws, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        s = Trim$(CStr(ws.Cells(r, 1).Value))
        If InStr(1, s, префикс, vbTextCompare) = 1 Then
            m = CLng(Val(Mid$(s, Len(префикс) + 1)))
            If m > n Then n = m
        End If
    Next r
    НовыйID = префикс & Format$(n + 1, "0000")
End Function


Private Function Текст(значение As Variant) As String
    If IsError(значение) Then Exit Function
    Текст = Trim$(CStr(значение))
End Function
