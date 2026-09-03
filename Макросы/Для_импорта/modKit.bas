Attribute VB_Name = "modKit"
Option Explicit

' =======================================================================
'  Минимальный комплект средств НК на метод.
'
'  Позиция комплекта описывается не конкретным прибором, а признаком —
'  частью наименования или назначения («лупа», «штангенциркуль»,
'  «шаблон сварщика»). Приборы списывают и покупают новые, а требование
'  остаётся прежним; привязка к инвентарному номеру пережила бы первую
'  же замену прибора.
'
'  Проверка ничего не запрещает. Она называет, чего не хватает, и
'  спрашивает — решение за исполнителем, но принимается осознанно.
' =======================================================================

Public Const Л_КОМПЛЕКТЫ As String = "Комплекты_СИ"
Public Const Л_ПРИМЕНЁННЫЕ As String = "Применённые_СИ"


' Что не так с комплектом по методу для этой записи. "" — всё в порядке.
Public Function ПроверитьКомплект(идЗаписи As String, метод As String) As String
    Dim к As Worksheet
    Dim i As Long, last As Long
    Dim позиция As String
    Dim признак As String
    Dim обязательно As Boolean
    Dim нехватка As String
    Dim рекомендуемых As String
    Dim применено As Long

    Set к = ЛистТихо(Л_КОМПЛЕКТЫ)
    If к Is Nothing Then Exit Function

    применено = СколькоПрименено(идЗаписи, метод)
    ' Строки комплекта заводят руками, а колонка ID служебная и может
    ' остаться пустой — считаем по названию позиции.
    last = ПоследняяСтрока(к, ColReq(к, "Позиция комплекта"))

    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(к.Cells(i, ColReq(к, "Метод")).Value)), метод, vbTextCompare) = 0 Then
            позиция = Trim$(CStr(к.Cells(i, ColReq(к, "Позиция комплекта")).Value))
            признак = Trim$(CStr(к.Cells(i, ColReq(к, "Признак для поиска (часть наименования)")).Value))
            обязательно = (StrComp(Trim$(CStr(к.Cells(i, ColReq(к, "Обязательно")).Value)), _
                                   "Да", vbTextCompare) = 0)
            If позиция <> "" Then
                If Not ЕстьВПрименённых(идЗаписи, метод, признак) Then
                    If обязательно Then
                        нехватка = нехватка & vbCrLf & "  - " & позиция & _
                                   IIf(признак = "", "", " (искали «" & признак & "»)")
                    Else
                        рекомендуемых = рекомендуемых & vbCrLf & "  - " & позиция
                    End If
                End If
            End If
        End If
    Next i

    If нехватка = "" And рекомендуемых = "" Then Exit Function

    ПроверитьКомплект = "Комплект по методу " & метод & _
        " (применено средств: " & применено & "):"
    If нехватка <> "" Then
        ПроверитьКомплект = ПроверитьКомплект & vbCrLf & "Не хватает обязательного:" & нехватка
    End If
    If рекомендуемых <> "" Then
        ПроверитьКомплект = ПроверитьКомплект & vbCrLf & "Не отмечено (необязательно):" & рекомендуемых
    End If
End Function


' Есть ли среди применённых средств такое, чьё наименование или
' назначение содержит признак.
Private Function ЕстьВПрименённых(идЗаписи As String, метод As String, признак As String) As Boolean
    Dim пр As Worksheet
    Dim si As Worksheet
    Dim i As Long, last As Long
    Dim идСИ As String
    Dim строкаСИ As Long
    Dim текст As String

    If Trim$(признак) = "" Then Exit Function
    Set пр = ЛистТихо(Л_ПРИМЕНЁННЫЕ)
    Set si = ЛистТихо(Л_СИ)
    If пр Is Nothing Or si Is Nothing Then Exit Function

    last = ПоследняяСтрока(пр, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If Совпало(пр, i, идЗаписи, метод) Then
            идСИ = Trim$(CStr(пр.Cells(i, ColReq(пр, "ID средства НК")).Value))
            строкаСИ = НайтиСИ(si, идСИ)
            If строкаСИ > 0 Then
                текст = CStr(si.Cells(строкаСИ, ColReq(si, "Наименование")).Value) & " " & _
                        CStr(si.Cells(строкаСИ, ColReq(si, "Тип / обозначение")).Value) & " " & _
                        CStr(si.Cells(строкаСИ, ColReq(si, "Назначение")).Value)
                If InStr(1, текст, признак, vbTextCompare) > 0 Then
                    ЕстьВПрименённых = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function


' Просроченная поверка у применённого средства — отдельная беда:
' измерения таким прибором недействительны, и это не «предупреждение».
Public Function ПроверитьПоверкуПрименённых(идЗаписи As String, метод As String) As String
    Dim пр As Worksheet
    Dim si As Worksheet
    Dim i As Long, last As Long
    Dim идСИ As String
    Dim строкаСИ As Long
    Dim до As Variant
    Dim список As String

    Set пр = ЛистТихо(Л_ПРИМЕНЁННЫЕ)
    Set si = ЛистТихо(Л_СИ)
    If пр Is Nothing Or si Is Nothing Then Exit Function

    last = ПоследняяСтрока(пр, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If Совпало(пр, i, идЗаписи, метод) Then
            идСИ = Trim$(CStr(пр.Cells(i, ColReq(пр, "ID средства НК")).Value))
            строкаСИ = НайтиСИ(si, идСИ)
            If строкаСИ > 0 Then
                до = si.Cells(строкаСИ, ColReq(si, "Действительно до")).Value
                If IsDate(до) Then
                    If CDate(до) < Date Then
                        список = список & vbCrLf & "  - " & _
                                 CStr(si.Cells(строкаСИ, ColReq(si, "Наименование")).Value) & _
                                 " зав. № " & CStr(si.Cells(строкаСИ, ColReq(si, "Заводской №")).Value) & _
                                 " — поверка до " & Format$(до, "dd.mm.yyyy")
                    End If
                Else
                    список = список & vbCrLf & "  - " & _
                             CStr(si.Cells(строкаСИ, ColReq(si, "Наименование")).Value) & _
                             " — дата поверки не заполнена"
                End If
            End If
        End If
    Next i

    If список <> "" Then
        ПроверитьПоверкуПрименённых = "Применены средства с недействующей поверкой:" & список & _
            vbCrLf & "Измерения такими приборами недействительны."
    End If
End Function


' Добавить применённые средства к текущей строке журнала.
Public Sub ДобавитьПрименённыеСИ()
    Dim ж As Worksheet
    Dim пр As Worksheet
    Dim r As Long
    Dim строка As Long
    Dim метод As String
    Dim сколько As Long
    Dim i As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub

    метод = СпроситьПрименённыйМетод(ж, r)
    If метод = "" Then Exit Sub

    сколько = CLng(Val(InputBox("Сколько средств НК добавить по методу " & метод & "?" & vbCrLf & vbCrLf & _
                                "Прибор выбирается из списка по наименованию с заводским " & _
                                "номером — запоминать ID не нужно.", _
                                "Применённые средства", "3")))
    If сколько <= 0 Or сколько > 50 Then Exit Sub

    Set пр = Лист(Л_ПРИМЕНЁННЫЕ)
    If пр Is Nothing Then Exit Sub
    строка = ПоследняяСтрока(пр, 1) + 1

    For i = 1 To сколько
        пр.Cells(строка, 1).Value = НовыйIDПрименения(пр)
        пр.Cells(строка, ColReq(пр, "Запись журнала (выбор)")).Value = _
            ж.Cells(r, ColReq(ж, "Обозначение записи (для выбора)")).Value
        пр.Cells(строка, ColReq(пр, "Метод")).Value = метод
        пр.Cells(строка, ColReq(пр, "Кем внесено")).Value = ТекущийСотрудник()
        пр.Cells(строка, ColReq(пр, "Изменено")).Value = Now
        строка = строка + 1
    Next i

    пр.Activate
    пр.Cells(строка - сколько, ColReq(пр, "Средство НК (выбор)")).Select
    MsgBox "Добавлено строк: " & сколько & "." & vbCrLf & vbCrLf & _
           "Выберите приборы из списка. Срок поверки и его статус подтянутся сами.", _
           vbInformation, "Применённые средства"
End Sub


' Показать, чего не хватает по всем применённым методам строки.
Public Sub ПроверитьКомплектСтроки()
    Dim ж As Worksheet
    Dim r As Long
    Dim методы As Variant
    Dim i As Long
    Dim идЗаписи As String
    Dim текст As String
    Dim часть As String

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub
    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        If МетодПрименён(ж, r, CStr(методы(i))) Then
            часть = ПроверитьКомплект(идЗаписи, CStr(методы(i)))
            If часть <> "" Then текст = текст & vbCrLf & vbCrLf & часть
            часть = ПроверитьПоверкуПрименённых(идЗаписи, CStr(методы(i)))
            If часть <> "" Then текст = текст & vbCrLf & vbCrLf & часть
        End If
    Next i

    If текст = "" Then
        MsgBox "Комплект средств НК по всем применённым методам собран, " & _
               "поверка действует.", vbInformation, "Проверка комплекта"
    Else
        MsgBox Mid$(текст, 3), vbExclamation, "Проверка комплекта"
    End If
End Sub


' -- Вспомогательное ---------------------------------------------------

Private Function Совпало(пр As Worksheet, i As Long, идЗаписи As String, метод As String) As Boolean
    If StrComp(Trim$(CStr(пр.Cells(i, ColReq(пр, "ID записи журнала")).Value)), _
               идЗаписи, vbTextCompare) <> 0 Then Exit Function
    Совпало = (StrComp(Trim$(CStr(пр.Cells(i, ColReq(пр, "Метод")).Value)), _
                       метод, vbTextCompare) = 0)
End Function


Private Function СколькоПрименено(идЗаписи As String, метод As String) As Long
    Dim пр As Worksheet
    Dim i As Long, last As Long
    Set пр = ЛистТихо(Л_ПРИМЕНЁННЫЕ)
    If пр Is Nothing Then Exit Function
    last = ПоследняяСтрока(пр, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If Совпало(пр, i, идЗаписи, метод) Then СколькоПрименено = СколькоПрименено + 1
    Next i
End Function


Private Function НайтиСИ(si As Worksheet, идСИ As String) As Long
    Dim r As Long, last As Long
    If Trim$(идСИ) = "" Then Exit Function
    last = ПоследняяСтрока(si, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(si.Cells(r, 1).Value)), идСИ, vbTextCompare) = 0 Then
            НайтиСИ = r
            Exit Function
        End If
    Next r
End Function


Private Function СпроситьПрименённыйМетод(ж As Worksheet, r As Long) As String
    Dim методы As Variant
    Dim i As Long
    Dim список As String
    Dim ответ As String

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        If МетодПрименён(ж, r, CStr(методы(i))) Then
            список = список & IIf(список = "", "", ", ") & методы(i)
        End If
    Next i

    If список = "" Then
        MsgBox "В этой строке не отмечен ни один метод контроля.", _
               vbExclamation, "Применённые средства"
        Exit Function
    End If
    If InStr(список, ",") = 0 Then
        СпроситьПрименённыйМетод = список
        Exit Function
    End If

    ответ = UCase$(Trim$(InputBox("По какому методу добавляем средства: " & список & "?", _
                                  "Применённые средства", Split(список, ",")(0))))
    If InStr(1, список, ответ, vbTextCompare) > 0 And ответ <> "" Then
        СпроситьПрименённыйМетод = ответ
    End If
End Function


Private Function НовыйIDПрименения(пр As Worksheet) As String
    Dim r As Long, last As Long, n As Long, m As Long
    Dim s As String
    last = ПоследняяСтрока(пр, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        s = Trim$(CStr(пр.Cells(r, 1).Value))
        If InStr(1, s, "ПСИ-", vbTextCompare) = 1 Then
            m = CLng(Val(Mid$(s, 5)))
            If m > n Then n = m
        End If
    Next r
    НовыйIDПрименения = "ПСИ-" & Format$(n + 1, "0000")
End Function
