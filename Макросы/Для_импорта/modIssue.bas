Attribute VB_Name = "modIssue"
Option Explicit

' =======================================================================
'  Журнал выдачи оборудования.
'
'  Строка закрывается датой возврата. Пока её нет, прибор числится
'  выданным — это видно на Дашборде и мешает выдать его второй раз.
' =======================================================================

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

    идСИ = Trim$(InputBox("ID средства НК (например СИ-0012):" & vbCrLf & vbCrLf & _
                          "Посмотреть можно на листе «СИ», первая колонка.", _
                          "Выдача оборудования"))
    If идСИ = "" Then Exit Sub

    строкаСИ = НайтиСИ(si, идСИ)
    If строкаСИ = 0 Then
        MsgBox "В реестре СИ нет записи с ID «" & идСИ & "».", vbExclamation, "Выдача оборудования"
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
    ws.Cells(r, ColReq(ws, "ID средства НК")).Value = идСИ
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
                текст = текст & vbCrLf & "- Прибор уже выдан " & _
                        Format$(ws.Cells(r, ColReq(ws, "Дата выдачи")).Value, "dd.mm.yyyy") & _
                        " — " & CStr(ws.Cells(r, ColReq(ws, "Кому выдано")).Value) & _
                        " (строка " & r & "). Возврат не отмечен."
            End If
        End If
    Next r

    состояние = Trim$(CStr(si.Cells(строкаСИ, ColReq(si, "Состояние")).Value))
    If StrComp(состояние, "Списан", vbTextCompare) = 0 Then
        текст = текст & vbCrLf & "- Прибор помечен как списанный."
    ElseIf StrComp(состояние, "В ремонте", vbTextCompare) = 0 Then
        текст = текст & vbCrLf & "- Прибор помечен как находящийся в ремонте."
    ElseIf StrComp(состояние, "На поверке", vbTextCompare) = 0 Then
        текст = текст & vbCrLf & "- Прибор помечен как находящийся на поверке."
    End If

    поверка = si.Cells(строкаСИ, ColReq(si, "Действительно до")).Value
    If IsDate(поверка) Then
        If CDate(поверка) < Date Then
            текст = текст & vbCrLf & "- Поверка просрочена: действовала до " & _
                    Format$(поверка, "dd.mm.yyyy") & ". Измерения этим прибором " & _
                    "недействительны."
        ElseIf DateDiff("d", Date, CDate(поверка)) <= 14 Then
            текст = текст & vbCrLf & "- Поверка истекает " & Format$(поверка, "dd.mm.yyyy") & _
                    " — через " & DateDiff("d", Date, CDate(поверка)) & " дн."
        End If
    Else
        текст = текст & vbCrLf & "- Не заполнена дата поверки — проверьте по свидетельству."
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
