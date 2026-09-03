Attribute VB_Name = "modChecks"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Проверки перед регистрацией и формированием документов.
'
'  Каждая функция возвращает "" если всё в порядке, иначе — текст
'  предупреждения. Ни одна проверка не запрещает действие сама:
'  решение остаётся за человеком, но принимается осознанно.
' ═══════════════════════════════════════════════════════════════════════


' Действует ли удостоверение исполнителя по методу на дату контроля.
Public Function ПроверитьАттестацию(фио As String, метод As String, наДату As Variant) As String
    Dim ws As Worksheet
    Dim cФИО As Long, cМетод As Long, cСрок As Long, cУдост As Long
    Dim r As Long, last As Long
    Dim дата As Date
    Dim нашли As Boolean
    Dim лучшийСрок As Date

    If Trim$(фио) = "" Then Exit Function
    Set ws = Лист(Л_АТТЕСТАЦИИ)
    If ws Is Nothing Then Exit Function

    cФИО = ColReq(ws, "ФИО")
    cМетод = ColReq(ws, "Метод контроля")
    cСрок = ColReq(ws, "Действительно до")
    cУдост = ColReq(ws, "№ квалификационного удостоверения")

    If IsDate(наДату) Then дата = CDate(наДату) Else дата = Date
    last = ПоследняяСтрока(ws, ColReq(ws, "ID"))

    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, cФИО).Value)), Trim$(фио), vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(ws.Cells(r, cМетод).Value)), Trim$(метод), vbTextCompare) = 0 Then
            нашли = True
            If IsDate(ws.Cells(r, cСрок).Value) Then
                If CDate(ws.Cells(r, cСрок).Value) > лучшийСрок Then
                    лучшийСрок = CDate(ws.Cells(r, cСрок).Value)
                End If
            Else
                ПроверитьАттестацию = "У «" & фио & "» по методу " & метод & _
                    " не заполнен срок действия удостоверения " & _
                    CStr(ws.Cells(r, cУдост).Value) & "."
                Exit Function
            End If
        End If
    Next r

    If Not нашли Then
        ПроверитьАттестацию = "В реестре аттестаций нет записи «" & фио & "» по методу " & _
            метод & ". Проверьте квалификацию вручную — или заведите удостоверение в реестр."
        Exit Function
    End If

    If лучшийСрок < дата Then
        ПроверитьАттестацию = "Удостоверение «" & фио & "» по методу " & метод & _
            " просрочено: действовало до " & Format$(лучшийСрок, "dd.mm.yyyy") & _
            ", контроль датирован " & Format$(дата, "dd.mm.yyyy") & "."
    ElseIf DateDiff("d", дата, лучшийСрок) <= 30 Then
        ПроверитьАттестацию = "Удостоверение «" & фио & "» по методу " & метод & _
            " истекает " & Format$(лучшийСрок, "dd.mm.yyyy") & " — осталось " & _
            DateDiff("d", дата, лучшийСрок) & " дн."
    End If
End Function


' Не числится ли человек отсутствующим на дату работ.
' «Выезд на измерения» отсутствием не считается — человек работает.
Public Function ПроверитьЗанятость(фио As String, наДату As Variant) As String
    Dim ws As Worksheet
    Dim cФИО As Long, cВид As Long, cС As Long, cПо As Long, cДоступен As Long
    Dim r As Long, last As Long
    Dim дата As Date

    If Trim$(фио) = "" Then Exit Function
    If Not IsDate(наДату) Then Exit Function
    дата = CDate(наДату)

    Set ws = Лист(Л_КТО_ГДЕ)
    If ws Is Nothing Then Exit Function
    cФИО = ColReq(ws, "ФИО")
    cВид = ColReq(ws, "Вид")
    cС = ColReq(ws, "Дата с")
    cПо = ColReq(ws, "Дата по")
    cДоступен = ColReq(ws, "Доступен для работ")
    last = ПоследняяСтрока(ws, ColReq(ws, "ID"))

    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, cФИО).Value)), Trim$(фио), vbTextCompare) = 0 Then
            If IsDate(ws.Cells(r, cС).Value) And IsDate(ws.Cells(r, cПо).Value) Then
                If дата >= CDate(ws.Cells(r, cС).Value) And дата <= CDate(ws.Cells(r, cПо).Value) Then
                    If StrComp(Trim$(CStr(ws.Cells(r, cДоступен).Value)), "Нет", vbTextCompare) = 0 Then
                        ПроверитьЗанятость = "По реестру «Кто где» " & фио & " " & _
                            Format$(дата, "dd.mm.yyyy") & " отсутствует: " & _
                            CStr(ws.Cells(r, cВид).Value) & " с " & _
                            Format$(ws.Cells(r, cС).Value, "dd.mm.yyyy") & " по " & _
                            Format$(ws.Cells(r, cПо).Value, "dd.mm.yyyy") & "."
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r
End Function


' Приборы с истёкшей поверкой. Проверяется весь реестр: какой именно
' прибор применялся, в журнале не хранится, поэтому предупреждаем
' обо всём, что просрочено на дату контроля.
Public Function ПроверитьПоверкуСИ(наДату As Variant) As String
    Dim ws As Worksheet
    Dim cНаим As Long, cДо As Long, cСост As Long, cЗав As Long
    Dim r As Long, last As Long
    Dim дата As Date
    Dim список As String
    Dim сколько As Long

    If IsDate(наДату) Then дата = CDate(наДату) Else дата = Date
    Set ws = Лист(Л_СИ)
    If ws Is Nothing Then Exit Function
    cНаим = ColReq(ws, "Наименование")
    cДо = ColReq(ws, "Действительно до")
    cСост = ColReq(ws, "Состояние")
    cЗав = ColReq(ws, "Заводской №")
    last = ПоследняяСтрока(ws, ColReq(ws, "ID"))

    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, cСост).Value)), "Списан", vbTextCompare) <> 0 Then
            If IsDate(ws.Cells(r, cДо).Value) Then
                If CDate(ws.Cells(r, cДо).Value) < дата Then
                    сколько = сколько + 1
                    If сколько <= 5 Then
                        список = список & vbCrLf & "  • " & CStr(ws.Cells(r, cНаим).Value) & _
                                 " зав. № " & CStr(ws.Cells(r, cЗав).Value) & _
                                 " — до " & Format$(ws.Cells(r, cДо).Value, "dd.mm.yyyy")
                    End If
                End If
            End If
        End If
    Next r

    If сколько > 0 Then
        If сколько > 5 Then список = список & vbCrLf & "  • … и ещё " & (сколько - 5)
        ПроверитьПоверкуСИ = "На " & Format$(дата, "dd.mm.yyyy") & " просрочена поверка " & _
            "у " & сколько & " приборов:" & список & vbCrLf & vbCrLf & _
            "Если применялся любой из них, документ выпускать нельзя."
    End If
End Function


' Собирает предупреждения по всей строке журнала: по каждому
' заполненному методу — исполнитель, его аттестация и занятость.
Public Function ПроверитьСтрокуЖурнала(строка As Long) As String
    Dim ws As Worksheet
    Dim методы As Variant
    Dim i As Long
    Dim метод As String
    Dim фио As String
    Dim дата As Variant
    Dim текст As String
    Dim ответ As String
    Dim идЗаписи As String

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Function
    дата = ws.Cells(строка, ColReq(ws, "Дата контроля")).Value
    идЗаписи = Trim$(CStr(ws.Cells(строка, ColReq(ws, "ID записи")).Value))
    методы = Split(МЕТОДЫ, ",")

    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        If МетодПрименён(ws, строка, метод) Then
            ответ = modKit.ПроверитьКомплект(идЗаписи, метод)
            If ответ <> "" Then текст = текст & vbCrLf & "• " & Replace(ответ, vbCrLf, vbCrLf & "  ")
            ответ = modKit.ПроверитьПоверкуПрименённых(идЗаписи, метод)
            If ответ <> "" Then текст = текст & vbCrLf & "• " & Replace(ответ, vbCrLf, vbCrLf & "  ")
            фио = Trim$(CStr(ws.Cells(строка, ColReq(ws, метод & " исполнитель")).Value))
            If фио = "" Then
                текст = текст & vbCrLf & "• " & метод & ": не указан исполнитель."
            Else
                ответ = ПроверитьАттестацию(фио, метод, дата)
                If ответ <> "" Then текст = текст & vbCrLf & "• " & ответ
                ответ = ПроверитьЗанятость(фио, дата)
                If ответ <> "" Then текст = текст & vbCrLf & "• " & ответ
            End If
        End If
    Next i

    If Trim$(CStr(ws.Cells(строка, ColReq(ws, "Объект контроля (выбор из справочника)")).Value)) = "" Then
        текст = текст & vbCrLf & "• Не выбран объект контроля."
    End If
    If Not IsDate(дата) Then
        текст = текст & vbCrLf & "• Не заполнена дата контроля."
    End If

    ПроверитьСтрокуЖурнала = текст
End Function
