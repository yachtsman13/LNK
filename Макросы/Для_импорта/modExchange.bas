Attribute VB_Name = "modExchange"
Option Explicit

' =======================================================================
'  Обмен данными между рабочими местами.
'
'  Что изменилось с прошлой отправки, определяется сравнением со снимком
'  данных, сделанным при той отправке, — а не отметкой «Изменено».
'  Причина: книгу данных можно открыть и без программы, и тогда никакая
'  отметка не проставится. Снимок ловит любую правку, кем бы она ни была
'  сделана.
'
'  Расчётные колонки не переносятся: при применении дельты ячейка с
'  формулой пропускается. Иначе первое же слияние затёрло бы формулы
'  значениями и книга перестала бы считать.
' =======================================================================

' Листы, участвующие в обмене. Настройки у каждой машины свои,
' Списки и Паспорт правятся редко и вручную, ППР_План — печатная форма.
Private Const ЛИСТЫ_ОБМЕНА As String = _
    "Журнал,Объекты,Специалисты,Аттестации,СИ,СИ_Обслуживание,НД,Кто_где," & _
    "График_отпусков,Обслуживание_факт,Выдача_оборудования,Рекламации,План_график,Задачи"

Private Const ФАЙЛ_СНИМКА As String = "ЛНК_Данные_снимок.xlsx"


' -- Отправка изменений ------------------------------------------------

Public Sub ОтправитьМоиИзменения()
    Dim снимок As Workbook
    Dim путьСнимка As String
    Dim файлДельты As String
    Dim f As Integer
    Dim листы As Variant
    Dim i As Long
    Dim строк As Long
    Dim всего As Long
    Dim текст As String

    If Not ПроверитьКодРабочегоМеста() Then Exit Sub
    If Данные() Is Nothing Then Exit Sub
    If Not ОбщаяПапкаДоступна() Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo Сбой

    ПроставитьНедостающиеID
    путьСнимка = ПапкаКопий() & ФАЙЛ_СНИМКА
    Set снимок = ОткрытьСнимок(путьСнимка)

    MkPath ПапкаОбмена() & "Входящие"
    файлДельты = ПапкаОбмена() & "Входящие" & Application.PathSeparator & _
                 Format$(Now, "yyyymmdd-hhnnss") & "-" & КодРабочегоМеста() & ".csv"

    f = FreeFile
    Open файлДельты For Output As #f
    Print #f, "@ДЕЛЬТА;1;" & КодРабочегоМеста() & ";" & Кодировать(ТекущийСотрудник()) & ";" & МеткаВремени()

    листы = Split(ЛИСТЫ_ОБМЕНА, ",")
    For i = LBound(листы) To UBound(листы)
        строк = ВыгрузитьЛист(f, CStr(листы(i)), снимок)
        всего = всего + строк
        If строк > 0 Then текст = текст & vbCrLf & "  " & листы(i) & ": " & строк
    Next i
    Close #f

    If всего = 0 Then
        Kill файлДельты
        Application.ScreenUpdating = True
        ЗакрытьСнимок снимок
        MsgBox "С прошлой отправки ничего не менялось — отправлять нечего.", _
               vbInformation, "Обмен"
        Exit Sub
    End If

    ЗакрытьСнимок снимок
    СохранитьСнимок путьСнимка
    Application.ScreenUpdating = True

    MsgBox "Отправлено изменений: " & всего & текст & vbCrLf & vbCrLf & _
           "Файл: " & файлДельты & vbCrLf & vbCrLf & _
           "Данные появятся у остальных после того, как мастер соберёт свод.", _
           vbInformation, "Обмен"
    Exit Sub

Сбой:
    On Error Resume Next
    Close #f
    ЗакрытьСнимок снимок
    Application.ScreenUpdating = True
    MsgBox "Не удалось отправить изменения." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Обмен"
End Sub


' Сравнивает лист со снимком и пишет в дельту изменившиеся строки.
Private Function ВыгрузитьЛист(f As Integer, имяЛиста As String, снимок As Workbook) As Long
    Dim ws As Worksheet
    Dim r As Long, c As Long, last As Long, колонок As Long
    Dim ид As String
    Dim строка As String
    Dim былаСтрока As String
    Dim было As Object
    Dim шапка As String
    Dim n As Long

    Set ws = Лист(имяЛиста)
    If ws Is Nothing Then Exit Function
    колонок = ws.Cells(СТРОКА_ЗАГОЛОВКОВ, ws.Columns.Count).End(xlToLeft).Column
    last = ПоследняяСтрока(ws, 1)
    If last < ПЕРВАЯ_СТРОКА Then Exit Function

    Set было = СнимокЛиста(снимок, имяЛиста, колонок)

    For c = 1 To колонок
        шапка = шапка & ";" & Кодировать(CStr(ws.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value))
    Next c

    For r = ПЕРВАЯ_СТРОКА To last
        ид = Trim$(CStr(ws.Cells(r, 1).Value))
        If ид <> "" Then
            строка = СобратьСтроку(ws, r, колонок)
            былаСтрока = ""
            If было.Exists(ид) Then былаСтрока = было(ид)
            If строка <> былаСтрока Then
                If n = 0 Then
                    Print #f, "@ЛИСТ;" & имяЛиста
                    Print #f, "@КОЛОНКИ" & шапка
                End If
                n = n + 1
                Print #f, IIf(былаСтрока = "", "НОВ", "ИЗМ") & ";" & Кодировать(ид) & ";" & _
                          МеткаИзменения(ws, r) & строка
            End If
            If было.Exists(ид) Then было.Remove ид
        End If
    Next r

    ' Всё, что осталось в снимке, из книги пропало.
    Dim ключ As Variant
    For Each ключ In было.Keys
        If n = 0 Then
            Print #f, "@ЛИСТ;" & имяЛиста
            Print #f, "@КОЛОНКИ" & шапка
        End If
        n = n + 1
        Print #f, "УДЛ;" & Кодировать(CStr(ключ)) & ";" & МеткаВремени()
    Next ключ

    ВыгрузитьЛист = n
End Function


Private Function СобратьСтроку(ws As Worksheet, r As Long, колонок As Long) As String
    Dim c As Long
    Dim s As String
    For c = 1 To колонок
        s = s & ";" & Кодировать(ЗначениеЯчейки(ws.Cells(r, c)))
    Next c
    СобратьСтроку = s
End Function


' Расчётные ячейки в дельту не идут: у получателя там своя формула.
Private Function ЗначениеЯчейки(ячейка As Range) As String
    If Left$(CStr(ячейка.Formula), 1) = "=" Then
        ЗначениеЯчейки = Chr$(1)          ' метка «здесь формула, не трогать»
    ElseIf IsDate(ячейка.Value) And ячейка.Value <> "" Then
        ЗначениеЯчейки = Format$(ячейка.Value, "yyyy-mm-dd hh:nn:ss")
    Else
        ЗначениеЯчейки = CStr(ячейка.Value)
    End If
End Function


' Время правки строки. Берём колонку «Изменено», если она заполнена, —
' по ней разрешается спор при одновременной правке. Если её нет,
' подставляем время отправки: хуже, но лучше, чем ничего.
Private Function МеткаИзменения(ws As Worksheet, r As Long) As String
    Dim c As Long
    c = ColByHeader(ws, "Изменено")
    If c > 0 Then
        If IsDate(ws.Cells(r, c).Value) Then
            МеткаИзменения = Format$(ws.Cells(r, c).Value, "yyyy-mm-dd\Thh:nn:ss")
            Exit Function
        End If
    End If
    МеткаИзменения = МеткаВремени()
End Function


' Строка без ID в обмен не попадёт: слияние идёт построчно и опознаёт
' строки только по ID. Поэтому перед отправкой всё, что заполнено руками
' и осталось без ID, его получает.
Public Sub ПроставитьНедостающиеID()
    Dim листы As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim r As Long, last As Long, c As Long, колонок As Long
    Dim префикс As String
    Dim счёт As Long
    Dim добавлено As Long

    листы = Split(ЛИСТЫ_ОБМЕНА, ",")
    For i = LBound(листы) To UBound(листы)
        Set ws = Лист(CStr(листы(i)))
        If Not ws Is Nothing Then
            префикс = ПрефиксID(CStr(листы(i)))
            колонок = ws.Cells(СТРОКА_ЗАГОЛОВКОВ, ws.Columns.Count).End(xlToLeft).Column
            last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            счёт = МаксимальныйНомерID(ws, префикс)
            For r = ПЕРВАЯ_СТРОКА To last
                If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then
                    If Not СтрокаПустая(ws, r, колонок) Then
                        счёт = счёт + 1
                        ws.Cells(r, 1).Value = префикс & Format$(счёт, "0000")
                        добавлено = добавлено + 1
                    End If
                End If
            Next r
        End If
    Next i
    If добавлено > 0 Then
        MsgBox "Строкам без ID он проставлен: " & добавлено & ". " & _
               "Без ID строка не участвует в обмене — слияние опознаёт строки только по нему.", _
               vbInformation, "Обмен"
    End If
End Sub


Private Function СтрокаПустая(ws As Worksheet, r As Long, колонок As Long) As Boolean
    Dim c As Long
    For c = 2 To колонок
        If Left$(CStr(ws.Cells(r, c).Formula), 1) <> "=" Then
            If Trim$(CStr(ws.Cells(r, c).Value)) <> "" Then Exit Function
        End If
    Next c
    СтрокаПустая = True
End Function


Private Function МаксимальныйНомерID(ws As Worksheet, префикс As String) As Long
    Dim r As Long, last As Long, n As Long
    Dim s As String
    last = ПоследняяСтрока(ws, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        s = Trim$(CStr(ws.Cells(r, 1).Value))
        If InStr(1, s, префикс, vbTextCompare) = 1 Then
            n = CLng(Val(Mid$(s, Len(префикс) + 1)))
            If n > МаксимальныйНомерID Then МаксимальныйНомерID = n
        End If
    Next r
End Function


Private Function ПрефиксID(имяЛиста As String) As String
    Select Case имяЛиста
        Case "Журнал": ПрефиксID = КодРабочегоМеста() & "-" & Format$(Date, "yyyy") & "-"
        Case "Объекты": ПрефиксID = "ОБ-"
        Case "Специалисты": ПрефиксID = "СП-"
        Case "Аттестации": ПрефиксID = "АТ-"
        Case "СИ": ПрефиксID = "СИ-"
        Case "СИ_Обслуживание": ПрефиксID = "ТО-строка-"
        Case "НД": ПрефиксID = "НД-"
        Case "Кто_где": ПрефиксID = "КГ-"
        Case "График_отпусков": ПрефиксID = "ГО-"
        Case "Обслуживание_факт": ПрефиксID = "ОБСЛ-"
        Case "Выдача_оборудования": ПрефиксID = "ВЫД-"
        Case "Рекламации": ПрефиксID = "РЕК-"
        Case "План_график": ПрефиксID = "ПГ-"
        Case "Задачи": ПрефиксID = "ЗД-"
        Case Else: ПрефиксID = "ЗАП-"
    End Select
End Function


' -- Снимок ------------------------------------------------------------

Private Function ОткрытьСнимок(путь As String) As Workbook
    If Dir(путь) = "" Then Exit Function      ' первая отправка — сравнивать не с чем
    On Error Resume Next
    Set ОткрытьСнимок = Application.Workbooks.Open(fileName:=путь, ReadOnly:=True, UpdateLinks:=0)
End Function


Private Sub ЗакрытьСнимок(снимок As Workbook)
    On Error Resume Next
    If Not снимок Is Nothing Then снимок.Close SaveChanges:=False
End Sub


Private Sub СохранитьСнимок(путь As String)
    On Error Resume Next
    MkPath ПапкаКопий()
    Данные().Save
    If Dir(путь) <> "" Then Kill путь
    FileCopy Данные().FullName, путь
End Sub


Private Function СнимокЛиста(снимок As Workbook, имяЛиста As String, колонок As Long) As Object
    Dim d As Object
    Dim ws As Worksheet
    Dim r As Long, last As Long
    Dim ид As String

    Set d = CreateObject("Scripting.Dictionary")
    Set СнимокЛиста = d
    If снимок Is Nothing Then Exit Function

    On Error Resume Next
    Set ws = снимок.Worksheets(имяЛиста)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function        ' лист появился в новой версии

    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ПЕРВАЯ_СТРОКА To last
        ид = Trim$(CStr(ws.Cells(r, 1).Value))
        If ид <> "" Then d(ид) = СобратьСтроку(ws, r, колонок)
    Next r
End Function


' -- Сборка свода у мастера --------------------------------------------

Public Sub СобратьСвод()
    Dim папка As String
    Dim файл As String
    Dim обработано As Long
    Dim применено As Long
    Dim конфликты As Collection
    Dim список As Collection
    Dim i As Long
    Dim итог As String

    If Not ЭтоМастерИлиПодтверждено() Then Exit Sub
    If Данные() Is Nothing Then Exit Sub
    If Not ОбщаяПапкаДоступна() Then Exit Sub

    папка = ПапкаОбмена() & "Входящие" & Application.PathSeparator
    If Dir(папка, vbDirectory) = "" Then
        MsgBox "Папка «Входящие» пуста — собирать нечего.", vbInformation, "Свод"
        Exit Sub
    End If

    Set конфликты = New Collection
    Set список = New Collection
    файл = Dir(папка & "*.csv")
    Do While файл <> ""
        список.Add папка & файл
        файл = Dir
    Loop
    If список.Count = 0 Then
        MsgBox "Новых дельт нет — все обработаны.", vbInformation, "Свод"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Сбой

    СохранитьСостояниеСвода

    For i = 1 To список.Count
        применено = применено + ПрименитьДельту(CStr(список(i)), конфликты)
        обработано = обработано + 1
    Next i

    ПроверитьДублиНомеров конфликты
    ЗаписатьКонфликты конфликты

    For i = 1 To список.Count
        УбратьВОбработанные CStr(список(i))
    Next i

    ОпубликоватьСвод
    Application.ScreenUpdating = True

    итог = "Обработано дельт: " & обработано & vbCrLf & "Применено строк: " & применено
    If конфликты.Count > 0 Then
        итог = итог & vbCrLf & vbCrLf & "Спорных правок: " & конфликты.Count & _
               " — они в папке «Конфликты», разберите их вручную. " & _
               "Данные при этом не потеряны: обе версии записаны в отчёт."
    End If
    MsgBox итог, vbInformation, "Свод"
    Exit Sub

Сбой:
    Application.ScreenUpdating = True
    MsgBox "Сборка свода прервана." & vbCrLf & "Ошибка: " & Err.Description & vbCrLf & vbCrLf & _
           "Состояние свода до сборки сохранено в папке «Свод\История» — " & _
           "при необходимости к нему можно вернуться.", vbCritical, "Свод"
End Sub


Private Function ПрименитьДельту(путь As String, конфликты As Collection) As Long
    Dim f As Integer
    Dim строка As String
    Dim части() As String
    Dim ws As Worksheet
    Dim карта() As Long
    Dim автор As String
    Dim n As Long
    Dim естьКарта As Boolean

    f = FreeFile
    Open путь For Input As #f
    Do While Not EOF(f)
        Line Input #f, строка
        If строка <> "" Then
            части = Split(строка, ";")
            Select Case части(0)
                Case "@ДЕЛЬТА"
                    If UBound(части) >= 3 Then автор = Раскодировать(части(3))
                Case "@ЛИСТ"
                    Set ws = ЛистТихо(Раскодировать(части(1)))
                    естьКарта = False
                Case "@КОЛОНКИ"
                    If Not ws Is Nothing Then
                        карта = ПостроитьКарту(ws, части)
                        естьКарта = True
                    End If
                Case "НОВ", "ИЗМ", "УДЛ"
                    If Not ws Is Nothing Then
                        n = n + ПрименитьСтроку(ws, части, карта, естьКарта, автор, путь, конфликты)
                    End If
            End Select
        End If
    Loop
    Close #f
    ПрименитьДельту = n
End Function


' Колонки сопоставляются по заголовку, а не по позиции: у отправителя
' порядок колонок может отличаться, если у него другая версия книги.
Private Function ПостроитьКарту(ws As Worksheet, части() As String) As Long()
    Dim карта() As Long
    Dim i As Long
    ReDim карта(1 To UBound(части))
    For i = 1 To UBound(части)
        карта(i) = ColByHeader(ws, Раскодировать(части(i)))
    Next i
    ПостроитьКарту = карта
End Function


Private Function ПрименитьСтроку(ws As Worksheet, части() As String, карта() As Long, _
                                 естьКарта As Boolean, автор As String, файл As String, _
                                 конфликты As Collection) As Long
    Dim ид As String
    Dim метка As String
    Dim r As Long
    Dim i As Long
    Dim цель As Long
    Dim значение As String
    Dim нашаМетка As String

    ид = Раскодировать(части(1))
    If ид = "" Then Exit Function
    метка = части(2)
    r = НайтиПоID(ws, ид)

    If части(0) = "УДЛ" Then
        If r > 0 Then
            ws.Rows(r).Delete
            ПрименитьСтроку = 1
        End If
        Exit Function
    End If

    If r = 0 Then
        r = ПоследняяСтрока(ws, 1) + 1
        ws.Cells(r, 1).Value = ид
    Else
        ' Строку правили и у нас, и у отправителя — берём более позднюю
        ' по времени правки, обе версии кладём в отчёт о конфликтах.
        нашаМетка = МеткаИзменения(ws, r)
        If нашаМетка > метка Then
            конфликты.Add Array(ws.Name, ид, автор, метка, нашаМетка, файл, "оставлена наша версия")
            Exit Function
        ElseIf нашаМетка < метка And нашаМетка <> "" Then
            конфликты.Add Array(ws.Name, ид, автор, метка, нашаМетка, файл, "принята версия отправителя")
        End If
    End If

    ' Строка дельты: Действие;ID;Метка;значение1;значение2;…
    ' Значения начинаются с индекса 3, а карта колонок — с единицы.
    For i = 3 To UBound(части)
        цель = 0
        If естьКарта Then
            If (i - 2) <= UBound(карта) Then цель = карта(i - 2)
        End If
        If цель > 0 Then
            значение = Раскодировать(части(i))
            If значение <> Chr$(1) Then                      ' у отправителя тут была формула
                If Left$(CStr(ws.Cells(r, цель).Formula), 1) <> "=" Then
                    ЗаписатьЗначение ws.Cells(r, цель), значение
                End If
            End If
        End If
    Next i
    ПрименитьСтроку = 1
End Function


Private Sub ЗаписатьЗначение(ячейка As Range, значение As String)
    If значение = "" Then
        ячейка.ClearContents
    ElseIf IsDate(значение) And InStr(значение, "-") > 0 Then
        ячейка.Value = CDate(значение)
    ElseIf IsNumeric(значение) Then
        ячейка.Value = CDbl(значение)
    Else
        ячейка.Value = значение
    End If
End Sub


Private Function НайтиПоID(ws As Worksheet, ид As String) As Long
    Dim r As Long
    Dim last As Long
    last = ПоследняяСтрока(ws, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ид, vbBinaryCompare) = 0 Then
            НайтиПоID = r
            Exit Function
        End If
    Next r
End Function


' -- Конфликты и дубли -------------------------------------------------

' Даже теоретический сбой блокировки счётчика не должен пройти
' незамеченным: после слияния проверяем номера на повторы.
Private Sub ПроверитьДублиНомеров(конфликты As Collection)
    Dim ws As Worksheet
    Dim методы As Variant
    Dim i As Long, r As Long, last As Long
    Dim ключ As String
    Dim видели As Object
    Dim v As Variant
    Dim дата As Variant

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    Set видели = CreateObject("Scripting.Dictionary")
    методы = Split(МЕТОДЫ, ",")
    last = ПоследняяСтрока(ws, ColReq(ws, "Дата контроля"))

    For r = ПЕРВАЯ_СТРОКА To last
        дата = ws.Cells(r, ColReq(ws, "Дата контроля")).Value
        If IsDate(дата) Then
            For i = LBound(методы) To UBound(методы)
                v = ws.Cells(r, ColReq(ws, методы(i) & " №")).Value
                If IsNumeric(v) And Trim$(CStr(v)) <> "" Then
                    ключ = методы(i) & "-" & Year(CDate(дата)) & "-" & CStr(CLng(v))
                    If видели.Exists(ключ) Then
                        конфликты.Add Array(Л_ЖУРНАЛ, CStr(ws.Cells(r, 1).Value), "", "", "", "", _
                                            "ДУБЛЬ НОМЕРА " & ФорматНомера(CLng(v), Year(CDate(дата)), CStr(методы(i))) & _
                                            ", строки " & видели(ключ) & " и " & r)
                    Else
                        видели(ключ) = r
                    End If
                End If
            Next i
        End If
    Next r
End Sub


Private Sub ЗаписатьКонфликты(конфликты As Collection)
    Dim f As Integer
    Dim путь As String
    Dim i As Long
    Dim з As Variant

    If конфликты.Count = 0 Then Exit Sub
    MkPath ПапкаОбмена() & "Конфликты"
    путь = ПапкаОбмена() & "Конфликты" & Application.PathSeparator & _
           "Конфликты_" & Format$(Now, "yyyymmdd-hhnn") & ".csv"
    f = FreeFile
    Open путь For Output As #f
    Print #f, "Лист;ID записи;Автор дельты;Время у отправителя;Время у нас;Файл дельты;Решение"
    For i = 1 To конфликты.Count
        з = конфликты(i)
        Print #f, з(0) & ";" & з(1) & ";" & з(2) & ";" & з(3) & ";" & з(4) & ";" & з(5) & ";" & з(6)
    Next i
    Close #f
End Sub


' -- Публикация и получение свода --------------------------------------

Private Sub СохранитьСостояниеСвода()
    Dim папка As String
    Dim цель As String
    On Error Resume Next
    папка = ПапкаОбмена() & "Свод" & Application.PathSeparator & "История" & Application.PathSeparator
    MkPath папка
    If Dir(ПапкаОбмена() & "Свод" & Application.PathSeparator & "ЛНК_Свод.xlsx") = "" Then Exit Sub
    цель = папка & "ЛНК_Свод_" & Format$(Now, "yyyymmdd-hhnn") & ".xlsx"
    FileCopy ПапкаОбмена() & "Свод" & Application.PathSeparator & "ЛНК_Свод.xlsx", цель
    ПочиститьИсторию папка
End Sub


Private Sub ПочиститьИсторию(папка As String)
    Dim файлы As Collection
    Dim файл As String
    Dim держать As Long
    Dim i As Long

    держать = CLng(Val(Настройка("Хранить состояний свода")))
    If держать <= 0 Then держать = 60
    Set файлы = New Collection
    файл = Dir(папка & "ЛНК_Свод_*.xlsx")
    Do While файл <> ""
        файлы.Add файл
        файл = Dir
    Loop
    ' Порядок, в котором Dir отдаёт имена, не оговорён, поэтому сортируем
    ' сами. В именах дата и время, так что алфавитный порядок = хронологический.
    Dim j As Long
    Dim врем As String
    For i = 1 To файлы.Count - 1
        For j = i + 1 To файлы.Count
            If файлы(i) > файлы(j) Then
                врем = файлы(i)
                файлы.Remove i
                файлы.Add врем, , , j - 1
            End If
        Next j
    Next i
    On Error Resume Next
    For i = 1 To файлы.Count - держать
        Kill папка & файлы(i)
    Next i
End Sub


Private Sub ОпубликоватьСвод()
    Dim папка As String
    On Error GoTo Сбой
    папка = ПапкаОбмена() & "Свод" & Application.PathSeparator
    MkPath папка
    Данные().Save
    If Dir(папка & "ЛНК_Свод.xlsx") <> "" Then Kill папка & "ЛНК_Свод.xlsx"
    FileCopy Данные().FullName, папка & "ЛНК_Свод.xlsx"
    ЗеркалоОбщейПапки
    Exit Sub
Сбой:
    MsgBox "Свод собран, но опубликовать его не удалось." & vbCrLf & Err.Description, _
           vbExclamation, "Свод"
End Sub


' Сервер централизованно не резервируется, поэтому мастер держит у себя
' зеркало общей папки: отказ сервера тогда не авария, а смена пути.
Private Sub ЗеркалоОбщейПапки()
    Dim откуда As String
    Dim куда As String
    On Error Resume Next
    If Not ЭтоМастер() Then Exit Sub
    откуда = ПапкаОбмена()
    куда = Путь("Папка программы", "") & "Обмен_зеркало" & Application.PathSeparator
    MkPath куда
    КопироватьДерево откуда, куда
    ЗаписатьВЖурналКопирования "Зеркало общей папки", куда, 0, "Успешно"
End Sub


Private Sub КопироватьДерево(откуда As String, куда As String)
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder Left$(откуда, Len(откуда) - 1), Left$(куда, Len(куда) - 1), True
End Sub


Public Sub ПолучитьСвод()
    Dim источник As String
    Dim ответ As VbMsgBoxResult

    If Данные() Is Nothing Then Exit Sub
    If Not ОбщаяПапкаДоступна() Then Exit Sub
    источник = ПапкаОбмена() & "Свод" & Application.PathSeparator & "ЛНК_Свод.xlsx"
    If Dir(источник) = "" Then
        MsgBox "Свод ещё не опубликован — мастер его не собирал.", vbInformation, "Обмен"
        Exit Sub
    End If

    If ЕстьНеотправленное() Then
        ответ = MsgBox("У вас есть изменения, которые ещё не отправлены." & vbCrLf & vbCrLf & _
                       "Если получить свод сейчас, они потеряются. Отправить сначала?", _
                       vbYesNoCancel + vbExclamation, "Обмен")
        If ответ = vbCancel Then Exit Sub
        If ответ = vbYes Then
            ОтправитьМоиИзменения
            Exit Sub
        End If
    End If

    MsgBox "Получение свода заменит локальные данные версией мастера." & vbCrLf & vbCrLf & _
           "Порядок: закройте книгу данных, скопируйте файл" & vbCrLf & источник & vbCrLf & _
           "поверх" & vbCrLf & Данные().FullName & vbCrLf & _
           "и откройте программу заново." & vbCrLf & vbCrLf & _
           "Автоматическая замена открытого файла делается не здесь: перезапись " & _
           "книги, которая сейчас открыта, ломает больше, чем чинит.", _
           vbInformation, "Получение свода"
End Sub


Private Function ЕстьНеотправленное() As Boolean
    Dim снимок As Workbook
    Dim путь As String
    Dim листы As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim old As Object
    Dim r As Long, last As Long, колонок As Long
    Dim ид As String

    путь = ПапкаКопий() & ФАЙЛ_СНИМКА
    If Dir(путь) = "" Then
        ЕстьНеотправленное = True         ' ни разу не отправляли
        Exit Function
    End If

    On Error GoTo Выход
    Set снимок = ОткрытьСнимок(путь)
    листы = Split(ЛИСТЫ_ОБМЕНА, ",")
    For i = LBound(листы) To UBound(листы)
        Set ws = Лист(CStr(листы(i)))
        If Not ws Is Nothing Then
            колонок = ws.Cells(СТРОКА_ЗАГОЛОВКОВ, ws.Columns.Count).End(xlToLeft).Column
            Set old = СнимокЛиста(снимок, CStr(листы(i)), колонок)
            last = ПоследняяСтрока(ws, 1)
            For r = ПЕРВАЯ_СТРОКА To last
                ид = Trim$(CStr(ws.Cells(r, 1).Value))
                If ид <> "" Then
                    If Not old.Exists(ид) Then GoTo Есть
                    If old(ид) <> СобратьСтроку(ws, r, колонок) Then GoTo Есть
                End If
            Next r
        End If
    Next i
Выход:
    ЗакрытьСнимок снимок
    Exit Function
Есть:
    ЗакрытьСнимок снимок
    ЕстьНеотправленное = True
End Function


' -- Роль мастера ------------------------------------------------------

Public Sub ВзятьРольМастера()
    Dim f As Integer
    Dim файл As String

    If Not ПроверитьКодРабочегоМеста() Then Exit Sub
    If Not ОбщаяПапкаДоступна() Then Exit Sub
    If MsgBox("Взять роль мастера на себя?" & vbCrLf & vbCrLf & _
              "Сейчас мастер: " & IIf(КтоМастер() = "", "не назначен", КтоМастер()) & vbCrLf & _
              "Роль снимется с прежнего рабочего места — двух мастеров одновременно не бывает.", _
              vbYesNo + vbQuestion, "Роль мастера") = vbNo Then Exit Sub

    On Error GoTo Сбой
    MkPath ПапкаОбмена() & "Свод"
    файл = ПапкаОбмена() & "Свод" & Application.PathSeparator & "мастер.txt"
    f = FreeFile
    Open файл For Output As #f
    Print #f, ТекущийСотрудник() & " (" & КодРабочегоМеста() & "), с " & Format$(Now, "dd.mm.yyyy hh:nn")
    Close #f

    ЗаписатьНастройку "Это мастер-ПК", "Да"
    MsgBox "Роль мастера у вас. Не забудьте собирать свод — пока он не собран, " & _
           "остальные видят данные на момент последней сборки.", vbInformation, "Роль мастера"
    Exit Sub
Сбой:
    On Error Resume Next
    Close #f
    MsgBox "Не удалось записать роль мастера." & vbCrLf & Err.Description, vbCritical, "Роль мастера"
End Sub


Public Function КтоМастер() As String
    Dim файл As String
    Dim f As Integer
    Dim s As String

    файл = ПапкаОбмена() & "Свод" & Application.PathSeparator & "мастер.txt"
    If Dir(файл) = "" Then Exit Function
    On Error Resume Next
    f = FreeFile
    Open файл For Input As #f
    If Not EOF(f) Then Line Input #f, s
    Close #f
    КтоМастер = Trim$(s)
End Function


Private Function ЭтоМастерИлиПодтверждено() As Boolean
    If ЭтоМастер() Then
        ЭтоМастерИлиПодтверждено = True
        Exit Function
    End If
    ЭтоМастерИлиПодтверждено = (MsgBox( _
        "Это рабочее место не отмечено как мастер." & vbCrLf & vbCrLf & _
        "Сейчас мастер: " & IIf(КтоМастер() = "", "не назначен", КтоМастер()) & vbCrLf & vbCrLf & _
        "Собирать свод одновременно на двух машинах нельзя — правки разойдутся. " & _
        "Собрать всё равно?", vbYesNo + vbExclamation, "Свод") = vbYes)
End Function


' -- Вспомогательное ---------------------------------------------------

Private Function ОбщаяПапкаДоступна() As Boolean
    Dim п As String
    п = ПапкаОбмена()
    If Dir(п, vbDirectory) <> "" Then
        ОбщаяПапкаДоступна = True
        Exit Function
    End If
    MsgBox "Общая папка недоступна:" & vbCrLf & п & vbCrLf & vbCrLf & _
           "Если работаете через флешку — вставьте её и укажите путь в настройках. " & _
           "Без общей папки обмен не работает; всё остальное работает.", _
           vbExclamation, "Обмен"
End Function


Private Sub УбратьВОбработанные(путь As String)
    Dim папка As String
    Dim имя As String
    On Error Resume Next
    имя = Mid$(путь, InStrRev(путь, Application.PathSeparator) + 1)
    папка = ПапкаОбмена() & "Обработанные" & Application.PathSeparator & _
            Format$(Date, "yyyy-mm") & Application.PathSeparator
    MkPath папка
    FileCopy путь, папка & имя
    Kill путь
End Sub


' Дельты — обычный CSV с разделителем «;», поэтому сами разделители,
' переносы строк и амперсанд внутри значений заменяются на подстановки.
Private Function Кодировать(s As String) As String
    Dim t As String
    t = Replace(s, "&", "&amp;")
    t = Replace(t, ";", "&sc;")
    t = Replace(t, vbCr, "&cr;")
    t = Replace(t, vbLf, "&lf;")
    Кодировать = t
End Function


Private Function Раскодировать(s As String) As String
    Dim t As String
    t = Replace(s, "&lf;", vbLf)
    t = Replace(t, "&cr;", vbCr)
    t = Replace(t, "&sc;", ";")
    t = Replace(t, "&amp;", "&")
    Раскодировать = t
End Function
