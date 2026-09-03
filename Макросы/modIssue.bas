Attribute VB_Name = "modIssue"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Журнал выдачи оборудования.
'
'  Строка закрывается датой возврата. Пока её нет, прибор числится
'  выданным — это видно на Дашборде и мешает выдать его второй раз.
' ═══════════════════════════════════════════════════════════════════════

Private Const Л_ВЫДАЧА As String = "Выдача_оборудования"


Public Sub ВыдатьОборудование()
    Dim ws As Worksheet
    Dim si As Worksheet
    Dim идСИ As String
    Dim строкаСИ As Long
    Dim r As Long
    Dim кому As String
    Dim предупреждения As String

    Set ws = Лист(Л_ВЫДАЧА)
    Set si = Лист(Л_СИ)
    If ws Is Nothing Or si Is Nothing Then Exit Sub
    If Not ПроверитьКодРабочегоМеста() Then Exit Sub

    ' Прибор берётся со строки, на которой стоит курсор в реестре СИ.
    ' Вводить ID руками не нужно: он ничего не говорит человеку, а
    ' ошибиться в нём легко.
    If ActiveSheet.Name <> Л_СИ Then
        MsgBox "Откройте лист «СИ», встаньте на строку нужного прибора и " & _
               "повторите." & vbCrLf & vbCrLf & _
               "Выбирать прибор по ID вручную не нужно — достаточно курсора.", _
               vbInformation, "Выдача оборудования"
        Exit Sub
    End If
    строкаСИ = ActiveCell.Row
    If строкаСИ < ПЕРВАЯ_СТРОКА Then
        MsgBox "Курсор в шапке таблицы. Встаньте на строку прибора.", _
               vbInformation, "Выдача оборудования"
        Exit Sub
    End If
    идСИ = Trim$(CStr(si.Cells(строкаСИ, 1).Value))
    If идСИ = "" Then
        MsgBox "В этой строке реестра нет прибора.", vbExclamation, "Выдача оборудования"
        Exit Sub
    End If

    предупреждения = ПроверитьПередВыдачей(ws, si, идСИ, строкаСИ)
    If предупреждения <> "" Then
        If MsgBox(предупреждения & vbCrLf & vbCrLf & "Выдать всё равно?", _
                  vbYesNo + vbExclamation, "Выдача оборудования") = vbNo Then Exit Sub
    End If

    кому = Trim$(InputBox("Кому выдаётся (ФИО как в справочнике специалистов):", _
                          "Выдача оборудования", ТекущийСотрудник()))
    If кому = "" Then Exit Sub

    r = ПоследняяСтрока(ws, 1) + 1
    ws.Cells(r, 1).Value = НовыйIDВыдачи(ws)
    ' ID, наименование и заводской номер в журнале выдачи — формулы по
    ' обозначению прибора, писать в них нельзя: затрём формулу.
    ws.Cells(r, ColReq(ws, "Средство НК (выбор)")).Value = _
        si.Cells(строкаСИ, ColReq(si, "Обозначение (для выбора)")).Value
    ws.Cells(r, ColReq(ws, "Дата выдачи")).Value = Date
    ws.Cells(r, ColReq(ws, "Кому выдано")).Value = кому
    ws.Cells(r, ColReq(ws, "Состояние при выдаче")).Value = "Исправен"
    ws.Cells(r, ColReq(ws, "Выдал")).Value = ТекущийСотрудник()

    si.Cells(строкаСИ, ColReq(si, "Состояние")).Value = "Выдан"

    ws.Activate
    ws.Cells(r, ColReq(ws, "Основание / цель")).Select
    MsgBox "Записано. Допишите основание и цель выдачи." & vbCrLf & vbCrLf & _
           "Прибор помечен в реестре как выданный — вторично его выдать не дадут, " & _
           "пока не будет отмечен возврат.", vbInformation, "Выдача оборудования"
End Sub


Public Sub ПринятьВозврат()
    Dim ws As Worksheet
    Dim si As Worksheet
    Dim r As Long
    Dim идСИ As String
    Dim строкаСИ As Long

    Set ws = Лист(Л_ВЫДАЧА)
    Set si = Лист(Л_СИ)
    If ws Is Nothing Or si Is Nothing Then Exit Sub

    r = ТекущаяСтрока(ws)
    If r = 0 Then Exit Sub

    If Not IsDate(ws.Cells(r, ColReq(ws, "Дата выдачи")).Value) Then
        MsgBox "В этой строке нет даты выдачи — возвращать нечего. " & _
               "Встаньте на строку выдачи.", vbExclamation, "Возврат оборудования"
        Exit Sub
    End If
    If IsDate(ws.Cells(r, ColReq(ws, "Дата возврата")).Value) Then
        MsgBox "По этой строке возврат уже отмечен " & _
               Format$(ws.Cells(r, ColReq(ws, "Дата возврата")).Value, "dd.mm.yyyy") & ".", _
               vbExclamation, "Возврат оборудования"
        Exit Sub
    End If

    ws.Cells(r, ColReq(ws, "Дата возврата")).Value = Date
    ws.Cells(r, ColReq(ws, "Состояние при возврате")).Value = "Исправен"
    ws.Cells(r, ColReq(ws, "Принял")).Value = ТекущийСотрудник()

    идСИ = Trim$(CStr(ws.Cells(r, ColReq(ws, "ID средства НК")).Value))
    строкаСИ = НайтиСИ(si, идСИ)
    If строкаСИ > 0 Then
        si.Cells(строкаСИ, ColReq(si, "Состояние")).Value = "В работе"
    End If

    ws.Cells(r, ColReq(ws, "Состояние при возврате")).Select
    MsgBox "Возврат отмечен. Если прибор вернулся с замечаниями — поправьте " & _
           "состояние при возврате: по нему потом разбираются причины отказов.", _
           vbInformation, "Возврат оборудования"
End Sub


' Проверки перед выдачей: не выдан ли уже, в порядке ли поверка,
' не списан ли. Ни одна не запрещает выдачу — решает человек.
Private Function ПроверитьПередВыдачей(ws As Worksheet, si As Worksheet, _
                                       идСИ As String, строкаСИ As Long) As String
    Dim r As Long, last As Long
    Dim текст As String
    Dim состояние As String
    Dim поверка As Variant

    last = ПоследняяСтрока(ws, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, ColReq(ws, "ID средства НК")).Value)), идСИ, vbTextCompare) = 0 Then
            If IsDate(ws.Cells(r, ColReq(ws, "Дата выдачи")).Value) And _
               Not IsDate(ws.Cells(r, ColReq(ws, "Дата возврата")).Value) Then
                текст = текст & vbCrLf & "• Прибор уже выдан " & _
                        Format$(ws.Cells(r, ColReq(ws, "Дата выдачи")).Value, "dd.mm.yyyy") & _
                        " — " & CStr(ws.Cells(r, ColReq(ws, "Кому выдано")).Value) & _
                        " (строка " & r & "). Возврат не отмечен."
            End If
        End If
    Next r

    состояние = Trim$(CStr(si.Cells(строкаСИ, ColReq(si, "Состояние")).Value))
    If StrComp(состояние, "Списан", vbTextCompare) = 0 Then
        текст = текст & vbCrLf & "• Прибор помечен как списанный."
    ElseIf StrComp(состояние, "В ремонте", vbTextCompare) = 0 Then
        текст = текст & vbCrLf & "• Прибор помечен как находящийся в ремонте."
    ElseIf StrComp(состояние, "На поверке", vbTextCompare) = 0 Then
        текст = текст & vbCrLf & "• Прибор помечен как находящийся на поверке."
    End If

    поверка = si.Cells(строкаСИ, ColReq(si, "Действительно до")).Value
    If IsDate(поверка) Then
        If CDate(поверка) < Date Then
            текст = текст & vbCrLf & "• Поверка просрочена: действовала до " & _
                    Format$(поверка, "dd.mm.yyyy") & ". Измерения этим прибором " & _
                    "недействительны."
        ElseIf DateDiff("d", Date, CDate(поверка)) <= 14 Then
            текст = текст & vbCrLf & "• Поверка истекает " & Format$(поверка, "dd.mm.yyyy") & _
                    " — через " & DateDiff("d", Date, CDate(поверка)) & " дн."
        End If
    Else
        текст = текст & vbCrLf & "• Не заполнена дата поверки — проверьте по свидетельству."
    End If

    If текст <> "" Then
        ПроверитьПередВыдачей = "Перед выдачей «" & _
            CStr(si.Cells(строкаСИ, ColReq(si, "Наименование")).Value) & "»:" & текст
    End If
End Function


Private Function НайтиСИ(si As Worksheet, идСИ As String) As Long
    Dim r As Long, last As Long
    last = ПоследняяСтрока(si, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(si.Cells(r, 1).Value)), идСИ, vbTextCompare) = 0 Then
            НайтиСИ = r
            Exit Function
        End If
    Next r
End Function


Private Function НовыйIDВыдачи(ws As Worksheet) As String
    Dim r As Long, last As Long, n As Long, m As Long
    Dim s As String
    last = ПоследняяСтрока(ws, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        s = Trim$(CStr(ws.Cells(r, 1).Value))
        If InStr(1, s, "ВЫД-", vbTextCompare) = 1 Then
            m = CLng(Val(Mid$(s, 5)))
            If m > n Then n = m
        End If
    Next r
    НовыйIDВыдачи = "ВЫД-" & Format$(n + 1, "0000")
End Function


' ═══════════════════════════════════════════════════════════════════════
'  Выдача, вытекающая из контроля.
'
'  Прибор, которым провели контроль, к этому моменту был у исполнителя
'  на руках — иначе контроля бы не было. Заставлять вписывать тот же
'  факт второй раз, отдельной строкой «выдал — принял», значит просить
'  человека доказывать очевидное; на третий раз это перестают делать, и
'  журнал выдачи расходится с журналом заключений.
'
'  Поэтому строки по работам НК заводятся сами — из отмеченных
'  применённых средств. Если прибор на эту дату уже числится выданным,
'  новой строки не появляется: выдача на неделю выездных работ
'  закрывает все её дни, а не плодит строку на каждый день.
' ═══════════════════════════════════════════════════════════════════════

Public Function ВыдачаПоПрименённым(ж As Worksheet, r As Long, _
                                    ByRef создано As Long) As String
    Dim вд As Worksheet
    Dim пр As Worksheet
    Dim si As Worksheet
    Dim i As Long, last As Long
    Dim идЗаписи As String, идСИ As String, метод As String
    Dim исполнитель As String, держатель As String
    Dim дата As Date
    Dim строкаВыдачи As Long, строкаСИ As Long
    Dim конфликты As String
    Dim основание As String

    Set вд = ЛистТихо(Л_ВЫДАЧА)
    Set пр = ЛистТихо(Л_ПРИМЕНЁННЫЕ)
    Set si = ЛистТихо(Л_СИ)
    If вд Is Nothing Or пр Is Nothing Or si Is Nothing Then Exit Function

    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))
    If идЗаписи = "" Then Exit Function
    If Not IsDate(ж.Cells(r, ColReq(ж, "Дата контроля")).Value) Then Exit Function
    дата = CDate(ж.Cells(r, ColReq(ж, "Дата контроля")).Value)

    основание = "Контроль"
    If Trim$(CStr(ж.Cells(r, ColReq(ж, "№ акта")).Value)) <> "" Then
        основание = основание & " по акту " & Trim$(CStr(ж.Cells(r, ColReq(ж, "№ акта")).Value))
    End If

    last = ПоследняяСтрока(пр, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(пр.Cells(i, ColReq(пр, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            идСИ = Trim$(CStr(пр.Cells(i, ColReq(пр, "ID средства НК")).Value))
            метод = Trim$(CStr(пр.Cells(i, ColReq(пр, "Метод")).Value))
            строкаСИ = НайтиСИ(si, идСИ)
            If идСИ <> "" And строкаСИ > 0 Then
                исполнитель = ИсполнительМетода(ж, r, метод)
                держатель = ""
                строкаВыдачи = ВыдачаНаДату(вд, идСИ, дата, держатель)
                If строкаВыдачи > 0 Then
                    ' Прибор на эту дату числится за другим человеком.
                    ' Это не мелочь: либо ошиблись исполнителем, либо
                    ' прибор передали из рук в руки, не отметив возврат.
                    If исполнитель <> "" And держатель <> "" Then
                        If StrComp(держатель, исполнитель, vbTextCompare) <> 0 Then
                            конфликты = конфликты & vbCrLf & "  • " & _
                                CStr(si.Cells(строкаСИ, ColReq(si, "Наименование")).Value) & _
                                " (" & метод & "): на " & Format$(дата, "dd.mm.yyyy") & _
                                " числится за " & держатель & ", а исполнитель — " & исполнитель & _
                                " (строка выдачи " & строкаВыдачи & ")."
                        End If
                    End If
                Else
                    ЗавестиВыдачуПоКонтролю вд, si, строкаСИ, дата, исполнитель, _
                                            основание & ", метод " & метод, идЗаписи
                    создано = создано + 1
                End If
            End If
        End If
    Next i

    If конфликты <> "" Then
        ВыдачаПоПрименённым = "Расхождения с журналом выдачи:" & конфликты
    End If
End Function


' Строка выдачи, покрывающая эту дату: выдан не позже — возвращён не
' раньше (или ещё не возвращён). Заодно отдаёт, за кем числится.
Private Function ВыдачаНаДату(вд As Worksheet, идСИ As String, дата As Date, _
                              ByRef держатель As String) As Long
    Dim i As Long, last As Long
    Dim cИД As Long, cВыдачи As Long, cВозврат As Long, cКому As Long
    Dim выдан As Date

    cИД = ColReq(вд, "ID средства НК")
    cВыдачи = ColReq(вд, "Дата выдачи")
    cВозврат = ColReq(вд, "Дата возврата")
    cКому = ColReq(вд, "Кому выдано")
    last = ПоследняяСтрока(вд, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(вд.Cells(i, cИД).Value)), идСИ, vbTextCompare) = 0 Then
            If IsDate(вд.Cells(i, cВыдачи).Value) Then
                выдан = CDate(вд.Cells(i, cВыдачи).Value)
                If выдан <= дата Then
                    If Not IsDate(вд.Cells(i, cВозврат).Value) Then
                        держатель = Trim$(CStr(вд.Cells(i, cКому).Value))
                        ВыдачаНаДату = i
                        Exit Function
                    ElseIf CDate(вд.Cells(i, cВозврат).Value) >= дата Then
                        держатель = Trim$(CStr(вд.Cells(i, cКому).Value))
                        ВыдачаНаДату = i
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function


' Выдача и возврат одним днём: работа выполнена, прибор в лаборатории.
' Оставить строку открытой означало бы, что через неделю Дашборд
' покажет полсотни «невозвращённых» приборов, которые стоят на месте, —
' и на предупреждение про действительно невозвращённый перестанут
' смотреть. Если прибор остался на руках, его выдают обычным порядком,
' и тогда эта строка уже не заведётся.
Private Sub ЗавестиВыдачуПоКонтролю(вд As Worksheet, si As Worksheet, _
                                    строкаСИ As Long, дата As Date, _
                                    исполнитель As String, основание As String, _
                                    идЗаписи As String)
    Dim r As Long
    Dim кому As String

    кому = исполнитель
    If кому = "" Then кому = ТекущийСотрудник()

    r = ПоследняяСтрока(вд, 1) + 1
    вд.Cells(r, 1).Value = НовыйIDВыдачи(вд)
    вд.Cells(r, ColReq(вд, "Средство НК (выбор)")).Value = _
        si.Cells(строкаСИ, ColReq(si, "Обозначение (для выбора)")).Value
    ' ID считается формулой по обозначению, а следующая строка того же
    ' пакета уже ищет по нему выдачу. При ручном режиме пересчёта
    ' формула осталась бы пустой, и прибор попал бы в журнал дважды.
    вд.Cells(r, ColReq(вд, "ID средства НК")).Calculate
    вд.Cells(r, ColReq(вд, "Дата выдачи")).Value = дата
    вд.Cells(r, ColReq(вд, "Кому выдано")).Value = кому
    вд.Cells(r, ColReq(вд, "Основание / цель")).Value = основание
    вд.Cells(r, ColReq(вд, "Состояние при выдаче")).Value = "Исправен"
    вд.Cells(r, ColReq(вд, "Выдал")).Value = ТекущийСотрудник()
    вд.Cells(r, ColReq(вд, "Дата возврата")).Value = дата
    вд.Cells(r, ColReq(вд, "Состояние при возврате")).Value = "Исправен"
    вд.Cells(r, ColReq(вд, "Принял")).Value = ТекущийСотрудник()
    вд.Cells(r, ColReq(вд, "Примечание")).Value = _
        "Отмечено по применению на контроле, запись " & идЗаписи
End Sub


Private Function ИсполнительМетода(ж As Worksheet, r As Long, метод As String) As String
    Dim c As Long
    If метод = "" Then Exit Function
    c = ColByHeader(ж, метод & " исполнитель")
    If c > 0 Then ИсполнительМетода = Trim$(CStr(ж.Cells(r, c).Value))
End Function


' Отметить выдачу по уже зарегистрированным работам — для записей,
' заведённых до того, как связь появилась, и для правок задним числом.
Public Sub ОтметитьВыдачуПоЗаписи()
    Dim ж As Worksheet
    Dim r As Long
    Dim создано As Long
    Dim конфликты As String

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub

    конфликты = ВыдачаПоПрименённым(ж, r, создано)

    If создано = 0 And конфликты = "" Then
        MsgBox "Добавлять нечего: по этой записи либо не отмечены применённые " & _
               "приборы, либо все они уже числятся выданными на дату контроля." & vbCrLf & vbCrLf & _
               "Применённые приборы отмечаются кнопкой «Добавить применённые средства».", _
               vbInformation, "Выдача по контролю"
        Exit Sub
    End If

    MsgBox "Заведено строк в журнале выдачи: " & создано & "." & _
           IIf(конфликты = "", "", vbCrLf & vbCrLf & конфликты), _
           IIf(конфликты = "", vbInformation, vbExclamation), "Выдача по контролю"
End Sub
