Attribute VB_Name = "modPassport"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Сборка Паспорта ЛНК из реестров.
'
'  Из одиннадцати разделов шесть собираются целиком из данных: разделы 2
'  и 5-9 — это реестры специалистов, аттестаций, СИ и НД. Остальные пять
'  берутся с листа «Паспорт_постоянная»: они меняются раз в несколько лет
'  и в реестрах им места нет.
'
'  Собирается в Excel, а не в Word. Причина простая: Паспорт — это почти
'  сплошные таблицы переменной длины, и в Excel они получаются точнее.
'  Если нужно попасть в оформление действующей редакции слово в слово,
'  пришлите её файлом — переделаю на подстановку в шаблон Word.
' ═══════════════════════════════════════════════════════════════════════

Private Const Л_ПАСПОРТ As String = "Паспорт_печать"
Private Const Л_ПОСТОЯННАЯ As String = "Паспорт_постоянная"


Public Sub СобратьПаспорт()
    Dim ws As Worksheet
    Dim r As Long
    Dim пробелы As String

    If Данные() Is Nothing Then Exit Sub

    пробелы = ПроверитьПередВыпуском()
    If пробелы <> "" Then
        MsgBox "Паспорт не собран — не заполнено то, без чего он выйдет с пробелами:" & _
               vbCrLf & пробелы & vbCrLf & vbCrLf & _
               "Заполните перечисленное и повторите. Документ с пустыми местами " & _
               "хуже, чем несобранный: пустоту в нём замечают не сразу.", _
               vbExclamation, "Паспорт ЛНК"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Сбой
    Set ws = ПересоздатьЛист()

    r = Титул(ws, 1)
    r = Раздел1(ws, r)
    r = Раздел2(ws, r)
    r = РазделИзПостоянной(ws, r, 3, "Объекты неразрушающего контроля", _
                           "Область аттестации: объекты контроля")
    r = РазделИзПостоянной(ws, r, 4, "Виды (методы) неразрушающего контроля", _
                           "Область аттестации: методы")
    r = РазделСИ(ws, r, 5, "Оснащённость средствами неразрушающего контроля", "Средства НК", True)
    r = РазделСИ(ws, r, 6, "Меры, контрольные и настроечные образцы", "Меры и образцы", False)
    r = РазделСИ(ws, r, 7, "Дозиметрические и радиометрические средства измерений", "Дозиметрия", False)
    r = РазделСИ(ws, r, 8, "Вспомогательное оборудование", "Вспомогательное", False)
    r = Раздел9(ws, r)
    r = РазделИзПостоянной(ws, r, 10, "Помещения", "Помещения ЛНК: номер, площадь, назначение")
    r = РазделИзПостоянной(ws, r, 11, "Передвижные лаборатории", "Передвижные лаборатории")

    НастроитьПечать ws
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True

    MsgBox "Паспорт собран на листе «" & Л_ПАСПОРТ & "»." & vbCrLf & vbCrLf & _
           "Разделы 2 и 5-9 взяты из реестров и всегда соответствуют им. " & _
           "Разделы 1, 3, 4, 10, 11 — с листа «Паспорт_постоянная»." & vbCrLf & vbCrLf & _
           "Первую собранную редакцию стоит один раз сверить с действующей глазами: " & _
           "опечатка в постоянной части иначе уедет во все следующие редакции.", _
           vbInformation, "Паспорт ЛНК"
    Exit Sub

Сбой:
    Application.ScreenUpdating = True
    MsgBox "Не удалось собрать Паспорт." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Паспорт ЛНК"
End Sub


' ── Проверка перед выпуском ───────────────────────────────────────────

' Паспорт с пустыми местами хуже несобранного: пустоту в нём замечают
' не сразу, а он идёт в комплект для аттестации лаборатории.
Public Function ПроверитьПередВыпуском() As String
    Dim ws As Worksheet
    Dim r As Long, last As Long
    Dim текст As String
    Dim n As Long

    Set ws = ЛистТихо(Л_ПОСТОЯННАЯ)
    If ws Is Nothing Then
        ПроверитьПередВыпуском = vbCrLf & "• нет листа «" & Л_ПОСТОЯННАЯ & "»"
        Exit Function
    End If

    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, 4).Value)), "да", vbTextCompare) = 0 Then
            If Trim$(CStr(ws.Cells(r, 2).Value)) = "" Then
                текст = текст & vbCrLf & "• «" & CStr(ws.Cells(r, 1).Value) & _
                        "» — лист «" & Л_ПОСТОЯННАЯ & "», строка " & r
            End If
        End If
    Next r

    n = ЧислоСтрок(Л_СПЕЦИАЛИСТЫ)
    If n = 0 Then текст = текст & vbCrLf & "• реестр специалистов пуст — раздел 2 будет пустым"
    n = ЧислоСтрок(Л_АТТЕСТАЦИИ)
    If n = 0 Then текст = текст & vbCrLf & "• реестр аттестаций пуст — раздел 2 будет без удостоверений"
    n = ЧислоСтрок(Л_СИ)
    If n = 0 Then текст = текст & vbCrLf & "• реестр СИ пуст — разделы 5-8 будут пустыми"
    n = ЧислоСтрок("НД")
    If n = 0 Then текст = текст & vbCrLf & "• реестр НД пуст — раздел 9 будет пустым"

    ПроверитьПередВыпуском = текст
End Function


Private Function ЧислоСтрок(имяЛиста As String) As Long
    Dim ws As Worksheet
    Set ws = ЛистТихо(имяЛиста)
    If ws Is Nothing Then Exit Function
    ЧислоСтрок = ПоследняяСтрока(ws, 1) - ПЕРВАЯ_СТРОКА + 1
    If ЧислоСтрок < 0 Then ЧислоСтрок = 0
End Function


' ── Разделы ───────────────────────────────────────────────────────────

Private Function Титул(ws As Worksheet, r As Long) As Long
    ws.Cells(r, 1).Value = "ПАСПОРТ"
    ws.Cells(r + 1, 1).Value = "лаборатории неразрушающего контроля"
    ws.Cells(r + 2, 1).Value = ЗначениеПостоянной("Полное наименование ЛНК")
    ws.Cells(r + 4, 1).Value = "Редакция № " & ЗначениеПостоянной("Номер редакции Паспорта") & _
                               "  от " & ЗначениеПостоянной("Дата редакции")
    ws.Cells(r + 5, 1).Value = "Собран " & Format$(Now, "dd.mm.yyyy hh:nn") & _
                               " из реестров ЛНК"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + 2, 1))
        .Font.Size = 16
        .Font.Bold = True
    End With
    ws.Cells(r + 5, 1).Font.Italic = True
    ws.Cells(r + 5, 1).Font.Color = RGB(120, 120, 120)
    ws.Range(ws.Cells(r, 1), ws.Cells(r + 5, 1)).HorizontalAlignment = xlCenter
    Титул = r + 7
End Function


Private Function Раздел1(ws As Worksheet, r As Long) As Long
    Dim поля As Variant
    Dim i As Long
    Dim значение As String

    r = Заголовок(ws, r, 1, "Общие данные о лаборатории")
    поля = Array("Полное наименование ЛНК", "Сокращённое наименование", "Почтовый адрес", _
                 "Телефон, электронная почта", "Наименование организации, в составе которой ЛНК", _
                 "№ свидетельства об аттестации ЛНК", "Дата выдачи свидетельства", _
                 "Свидетельство действительно до", "Орган по аттестации", _
                 "Начальник ЛНК: должность, ФИО, телефон", _
                 "Ответственный за функционирование СМ: ФИО", _
                 "Ответственный за метрологическое обеспечение: ФИО", _
                 "Ответственный за фонд НД: ФИО")
    For i = LBound(поля) To UBound(поля)
        значение = ЗначениеПостоянной(CStr(поля(i)))
        If значение <> "" Then
            ws.Cells(r, 1).Value = поля(i)
            ws.Cells(r, 2).Value = значение
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Merge
            ОформитьСтроку ws, r, 6
            r = r + 1
        End If
    Next i
    Раздел1 = r + 1
End Function


Private Function Раздел2(ws As Worksheet, r As Long) As Long
    Dim sp As Worksheet
    Dim at As Worksheet
    Dim i As Long, j As Long, lastSp As Long, lastAt As Long
    Dim n As Long
    Dim фио As String
    Dim удост As String

    r = Заголовок(ws, r, 2, "Профессиональная квалификация сотрудников")
    r = Шапка(ws, r, Array("№ п/п", "Фамилия, имя, отчество", "Основание привлечения", _
                           "Должность, функции в ЛНК", _
                           "Образование, повышение квалификации, стаж", _
                           "Метод, уровень, № удостоверения, срок"), _
              Array(6, 30, 20, 32, 40, 46))

    Set sp = ЛистТихо(Л_СПЕЦИАЛИСТЫ)
    Set at = ЛистТихо(Л_АТТЕСТАЦИИ)
    If sp Is Nothing Then
        Раздел2 = r + 1
        Exit Function
    End If
    lastSp = ПоследняяСтрока(sp, 1)

    For i = ПЕРВАЯ_СТРОКА To lastSp
        фио = Trim$(CStr(sp.Cells(i, ColReq(sp, "ФИО полностью")).Value))
        If фио <> "" Then
            n = n + 1
            удост = ""
            If Not at Is Nothing Then
                lastAt = ПоследняяСтрока(at, 1)
                For j = ПЕРВАЯ_СТРОКА To lastAt
                    If StrComp(Trim$(CStr(at.Cells(j, ColReq(at, "ID специалиста")).Value)), _
                               Trim$(CStr(sp.Cells(i, 1).Value)), vbTextCompare) = 0 Then
                        удост = удост & IIf(удост = "", "", vbLf) & _
                                CStr(at.Cells(j, ColReq(at, "Метод контроля")).Value) & ", " & _
                                CStr(at.Cells(j, ColReq(at, "Уровень")).Value) & " ур., уд. № " & _
                                CStr(at.Cells(j, ColReq(at, "№ квалификационного удостоверения")).Value) & _
                                ", до " & Format$(at.Cells(j, ColReq(at, "Действительно до")).Value, "dd.mm.yyyy")
                    End If
                Next j
            End If
            ws.Cells(r, 1).Value = n
            ws.Cells(r, 2).Value = фио
            ws.Cells(r, 3).Value = sp.Cells(i, ColReq(sp, "Основание привлечения")).Value
            ws.Cells(r, 4).Value = Trim$(CStr(sp.Cells(i, ColReq(sp, "Должность")).Value) & vbLf & _
                                         CStr(sp.Cells(i, ColReq(sp, "Функции в ЛНК")).Value))
            ws.Cells(r, 5).Value = sp.Cells(i, ColReq(sp, "Образование, повышение квалификации")).Value
            ws.Cells(r, 6).Value = IIf(удост = "", "удостоверения не заведены в реестр", удост)
            ОформитьСтроку ws, r, 6
            r = r + 1
        End If
    Next i
    Раздел2 = r + 1
End Function


Private Function РазделСИ(ws As Worksheet, r As Long, номер As Long, название As String, _
                          разделПаспорта As String, поМетодам As Boolean) As Long
    Dim si As Worksheet
    Dim i As Long, last As Long, n As Long
    Dim методы As Variant
    Dim m As Long
    Dim текущийМетод As String

    r = Заголовок(ws, r, номер, название)
    r = Шапка(ws, r, Array("№ п/п", "Наименование, тип, заводской №", "Назначение", _
                           "Год выпуска, ввод в эксплуатацию", _
                           "Технические и метрологические характеристики", _
                           "Владелец, основание владения", "Поверка / калибровка до"), _
              Array(6, 40, 30, 18, 44, 26, 18))

    Set si = ЛистТихо(Л_СИ)
    If si Is Nothing Then
        РазделСИ = r + 1
        Exit Function
    End If
    last = ПоследняяСтрока(si, 1)

    If поМетодам Then
        методы = Split(МЕТОДЫ, ",")
        For m = LBound(методы) To UBound(методы)
            текущийМетод = CStr(методы(m))
            If ЕстьПоМетоду(si, last, разделПаспорта, текущийМетод) Then
                ws.Cells(r, 1).Value = "Метод " & текущийМетод
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Merge
                ws.Cells(r, 1).Font.Bold = True
                ws.Cells(r, 1).Interior.Color = RGB(226, 234, 246)
                ОформитьСтроку ws, r, 7
                r = r + 1
                For i = ПЕРВАЯ_СТРОКА To last
                    If ПодходитРаздел(si, i, разделПаспорта) Then
                        If StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "Метод контроля")).Value)), _
                                   текущийМетод, vbTextCompare) = 0 Then
                            n = n + 1
                            СтрокаСИ ws, r, si, i, n
                            r = r + 1
                        End If
                    End If
                Next i
            End If
        Next m
    Else
        For i = ПЕРВАЯ_СТРОКА To last
            If ПодходитРаздел(si, i, разделПаспорта) Then
                n = n + 1
                СтрокаСИ ws, r, si, i, n
                r = r + 1
            End If
        Next i
    End If

    If n = 0 Then
        ws.Cells(r, 1).Value = "по этому разделу оборудование в реестре не отмечено"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Merge
        ws.Cells(r, 1).Font.Italic = True
        ОформитьСтроку ws, r, 7
        r = r + 1
    End If
    РазделСИ = r + 1
End Function


Private Function ПодходитРаздел(si As Worksheet, i As Long, разделПаспорта As String) As Boolean
    If StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "Состояние")).Value)), "Списан", vbTextCompare) = 0 Then
        Exit Function
    End If
    ПодходитРаздел = (StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "Раздел Паспорта")).Value)), _
                              разделПаспорта, vbTextCompare) = 0)
End Function


Private Function ЕстьПоМетоду(si As Worksheet, last As Long, разделПаспорта As String, _
                              метод As String) As Boolean
    Dim i As Long
    For i = ПЕРВАЯ_СТРОКА To last
        If ПодходитРаздел(si, i, разделПаспорта) Then
            If StrComp(Trim$(CStr(si.Cells(i, ColReq(si, "Метод контроля")).Value)), _
                       метод, vbTextCompare) = 0 Then
                ЕстьПоМетоду = True
                Exit Function
            End If
        End If
    Next i
End Function


Private Sub СтрокаСИ(ws As Worksheet, r As Long, si As Worksheet, i As Long, n As Long)
    Dim наим As String
    Dim годы As String
    Dim поверка As Variant

    наим = Trim$(CStr(si.Cells(i, ColReq(si, "Наименование")).Value))
    If Trim$(CStr(si.Cells(i, ColReq(si, "Тип / обозначение")).Value)) <> "" Then
        наим = наим & ", " & CStr(si.Cells(i, ColReq(si, "Тип / обозначение")).Value)
    End If
    If Trim$(CStr(si.Cells(i, ColReq(si, "Заводской №")).Value)) <> "" Then
        наим = наим & ", зав. № " & CStr(si.Cells(i, ColReq(si, "Заводской №")).Value)
    End If

    годы = Trim$(CStr(si.Cells(i, ColReq(si, "Год выпуска")).Value))
    If IsDate(si.Cells(i, ColReq(si, "Дата ввода в эксплуатацию")).Value) Then
        годы = Trim$(годы & vbLf & "ввод " & _
               Format$(si.Cells(i, ColReq(si, "Дата ввода в эксплуатацию")).Value, "dd.mm.yyyy"))
    End If

    поверка = si.Cells(i, ColReq(si, "Действительно до")).Value

    ws.Cells(r, 1).Value = n
    ws.Cells(r, 2).Value = наим
    ws.Cells(r, 3).Value = si.Cells(i, ColReq(si, "Назначение")).Value
    ws.Cells(r, 4).Value = годы
    ws.Cells(r, 5).Value = si.Cells(i, ColReq(si, "Технические и метрологические характеристики")).Value
    ws.Cells(r, 6).Value = Trim$(CStr(si.Cells(i, ColReq(si, "Владелец оборудования")).Value) & " " & _
                                 CStr(si.Cells(i, ColReq(si, "Основание владения")).Value))
    If IsDate(поверка) Then
        ws.Cells(r, 7).Value = Format$(поверка, "dd.mm.yyyy")
        ' Просроченное не прячем: Паспорт с истёкшей поверкой всё равно
        ' вскроется при проверке, лучше увидеть это заранее.
        If CDate(поверка) < Date Then
            ws.Cells(r, 7).Font.Color = RGB(156, 0, 6)
            ws.Cells(r, 7).Font.Bold = True
        End If
    Else
        ws.Cells(r, 7).Value = "не заполнено"
        ws.Cells(r, 7).Font.Color = RGB(156, 0, 6)
    End If
    ОформитьСтроку ws, r, 7
End Sub


Private Function Раздел9(ws As Worksheet, r As Long) As Long
    Dim нд As Worksheet
    Dim i As Long, last As Long, n As Long

    r = Заголовок(ws, r, 9, "Перечень нормативных документов")
    r = Шапка(ws, r, Array("№ п/п", "Обозначение", "Наименование документа", _
                           "Издательство / разработчик", "Год издания", _
                           "Проверен на актуальность"), _
              Array(6, 24, 56, 28, 12, 20))

    Set нд = ЛистТихо("НД")
    If нд Is Nothing Then
        Раздел9 = r + 1
        Exit Function
    End If
    last = ПоследняяСтрока(нд, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If Trim$(CStr(нд.Cells(i, ColReq(нд, "Наименование документа")).Value)) <> "" Then
            n = n + 1
            ws.Cells(r, 1).Value = n
            ws.Cells(r, 2).Value = нд.Cells(i, ColReq(нд, "Обозначение")).Value
            ws.Cells(r, 3).Value = нд.Cells(i, ColReq(нд, "Наименование документа")).Value
            ws.Cells(r, 4).Value = нд.Cells(i, ColReq(нд, "Издательство / разработчик")).Value
            ws.Cells(r, 5).Value = нд.Cells(i, ColReq(нд, "Год издания")).Value
            If IsDate(нд.Cells(i, ColReq(нд, "Дата последней проверки на актуальность")).Value) Then
                ws.Cells(r, 6).Value = Format$( _
                    нд.Cells(i, ColReq(нд, "Дата последней проверки на актуальность")).Value, "dd.mm.yyyy")
            End If
            ОформитьСтроку ws, r, 6
            r = r + 1
        End If
    Next i
    Раздел9 = r + 1
End Function


' Разделы 3, 4, 10, 11: перечень с листа постоянной части. Значение
' может занимать несколько строк подряд — читаем до первой пустой.
Private Function РазделИзПостоянной(ws As Worksheet, r As Long, номер As Long, _
                                    название As String, поле As String) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long, начало As Long
    Dim n As Long

    r = Заголовок(ws, r, номер, название)
    Set src = ЛистТихо(Л_ПОСТОЯННАЯ)
    If src Is Nothing Then
        РазделИзПостоянной = r + 1
        Exit Function
    End If

    last = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(src.Cells(i, 1).Value)), поле, vbTextCompare) = 0 Then
            начало = i
            Exit For
        End If
    Next i
    If начало = 0 Then
        РазделИзПостоянной = r + 1
        Exit Function
    End If

    For i = начало To начало + 40
        If i > начало Then
            If Trim$(CStr(src.Cells(i, 1).Value)) <> "" Then Exit For
        End If
        If Trim$(CStr(src.Cells(i, 2).Value)) = "" Then
            If i > начало Then Exit For
        Else
            n = n + 1
            ws.Cells(r, 1).Value = n
            ws.Cells(r, 2).Value = src.Cells(i, 2).Value
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Merge
            ОформитьСтроку ws, r, 7
            r = r + 1
        End If
    Next i
    РазделИзПостоянной = r + 1
End Function


' ── Оформление ────────────────────────────────────────────────────────

Private Function Заголовок(ws As Worksheet, r As Long, номер As Long, текст As String) As Long
    ws.Cells(r, 1).Value = CStr(номер) & ". " & текст
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Merge
    With ws.Cells(r, 1)
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 22
    Заголовок = r + 1
End Function


Private Function Шапка(ws As Worksheet, r As Long, заголовки As Variant, ширины As Variant) As Long
    Dim i As Long
    For i = LBound(заголовки) To UBound(заголовки)
        ws.Cells(r, i + 1).Value = заголовки(i)
        If ws.Columns(i + 1).ColumnWidth < ширины(i) Then
            ws.Columns(i + 1).ColumnWidth = ширины(i)
        End If
    Next i
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(заголовки) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(226, 234, 246)
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


Private Function ПересоздатьЛист() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Л_ПАСПОРТ).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = Л_ПАСПОРТ
    ws.Tab.Color = RGB(31, 56, 100)
    ws.Cells.Font.Name = "Times New Roman"
    ws.Cells.Font.Size = 10
    Set ПересоздатьЛист = ws
End Function


Private Sub НастроитьПечать(ws As Worksheet)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .CenterFooter = "&P"
    End With
End Sub


Private Function ЗначениеПостоянной(поле As String) As String
    Dim ws As Worksheet
    Dim r As Long, last As Long
    Set ws = ЛистТихо(Л_ПОСТОЯННАЯ)
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), поле, vbTextCompare) = 0 Then
            ЗначениеПостоянной = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function
