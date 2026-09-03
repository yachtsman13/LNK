Attribute VB_Name = "modGenerate"
Option Explicit

' =======================================================================
'  Формирование документов Word из шаблонов.
'
'  В шаблонах поля помечены как #ИмяПоля — так же, как в ваших рабочих
'  шаблонах. Что система знает, она подставляет; чего не знает —
'  оставляет плейсхолдером и честно перечисляет в конце, чтобы
'  незаполненное не уехало в документ незамеченным.
' =======================================================================

Private Const wdReplaceAll As Long = 2
Private Const wdFindContinue As Long = 1


' Формирует документы по всем методам текущей строки журнала.
Public Sub СформироватьДокументы()
    Dim ws As Worksheet
    Dim r As Long
    Dim методы As Variant
    Dim i As Long
    Dim метод As String
    Dim сделано As String
    Dim путь As String
    Dim предупреждения As String
    Dim безШаблона As String

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ws)
    If r = 0 Then Exit Sub

    предупреждения = ПроверитьСтрокуЖурнала(r)
    If предупреждения <> "" Then
        If MsgBox("Перед формированием стоит посмотреть:" & vbCrLf & предупреждения & vbCrLf & vbCrLf & _
                  "Формировать всё равно?", vbYesNo + vbExclamation, "Проверка") = vbNo Then Exit Sub
    End If

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        If МетодПрименён(ws, r, метод) Then
            If ФайлШаблона(метод) = "" Then
                безШаблона = безШаблона & IIf(безШаблона = "", "", ", ") & метод
            Else
                путь = СформироватьОдин(ws, r, метод)
                If путь <> "" Then сделано = сделано & vbCrLf & "  " & путь
            End If
        End If
    Next i

    If сделано = "" Then
        MsgBox "Ни одного документа не сформировано." & vbCrLf & _
               IIf(безШаблона = "", "В строке не отмечен ни один метод контроля.", _
                   "Для методов " & безШаблона & " шаблонов пока нет — эти документы " & _
                   "оформляются вручную."), vbExclamation, "Формирование документов"
    Else
        ws.Cells(r, ColReq(ws, "Документ сформирован")).Value = Date
        MsgBox "Сформировано:" & сделано & _
               IIf(безШаблона = "", "", vbCrLf & vbCrLf & "Без шаблона остались методы: " & _
                   безШаблона & " — эти документы оформляются вручную.") & vbCrLf & vbCrLf & _
               "Незаполненные поля остались в виде #Плейсхолдеров — допишите их в документе.", _
               vbInformation, "Формирование документов"
    End If
End Sub


Private Function СформироватьОдин(ws As Worksheet, r As Long, метод As String) As String
    Dim шаблон As String
    Dim путьШаблона As String
    Dim папка As String
    Dim файл As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim дата As Variant
    Dim год As Long
    Dim лишних As Long

    шаблон = ФайлШаблона(метод)
    If шаблон = "" Then Exit Function
    путьШаблона = ПапкаШаблонов() & шаблон
    If Not ФайлСуществует(путьШаблона) Then
        MsgBox "Не найден шаблон для метода " & метод & ":" & vbCrLf & путьШаблона, _
               vbExclamation, "Формирование документов"
        Exit Function
    End If

    дата = ws.Cells(r, ColReq(ws, "Дата контроля")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))

    On Error GoTo Сбой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(fileName:=путьШаблона, ReadOnly:=False, AddToRecentFiles:=False)

    Подставить wdDoc, "Номер", РегНомер(ws, r, метод)
    Подставить wdDoc, "Год", CStr(год)
    Подставить wdDoc, "Вид_контроля", метод
    Подставить wdDoc, "Дата_и_место_проведения_контроля", ДатаИМесто(ws, r)
    Подставить wdDoc, "Объект_контроля", ОписаниеОбъекта(ws, r)
    Подставить wdDoc, "Наименование_объекта", ОписаниеОбъекта(ws, r)
    Подставить wdDoc, "Заказывающее_подразделение", _
               CStr(ws.Cells(r, ColReq(ws, "Заказывающее подразделение")).Value)
    Подставить wdDoc, "Уровень_квалификации_№_удостоверения_ФИО", ПодписьИсполнителя(ws, r, метод)
    Подставить wdDoc, "Уровень_№_ФИО", ПодписьИсполнителя(ws, r, метод)
    Подставить wdDoc, "Наименование_ЛНК", Настройка("Полное наименование ЛНК")
    Подставить wdDoc, "Дата_проведения_контроля", _
               IIf(IsDate(дата), Format$(дата, "dd.mm.yyyy"), "")
    Подставить wdDoc, "Заводской_номер", _
               CStr(ws.Cells(r, ColReq(ws, "Заводской № объекта")).Value)
    Подставить wdDoc, "Место_расположения", МестоРасположения(ws, r)
    Подставить wdDoc, "Тип_объекта", ПолеОбъекта(ws, r, "Тип / группа")

    ' Паспорт объекта — из справочника, один раз заполненный на объект.
    ПодставитьПаспортОбъекта wdDoc, ws, r
    ' Поля этой конкретной работы: заключение, дефекты, НД, средства НК.
    ' Они же могут переопределить отдельные паспортные значения.
    лишних = ПодставитьДанныеПротокола(wdDoc, ws, r, метод)
    ' Настройки аппаратуры АЭ — свой лист: двенадцать полей, и в других
    ' методах они пустуют, поэтому в общий лист их класть незачем.
    If StrComp(метод, "АЭ", vbTextCompare) = 0 Then
        ПодставитьИзЛиста wdDoc, Л_ДАННЫЕ_АЭ, _
                          CStr(ws.Cells(r, ColReq(ws, "ID записи")).Value), 3
    End If
    ' Точки замеров — пока по УЗТ и ТВ, у остальных методов таблиц в
    ' шаблонах нет, там результат пишется текстом.
    лишних = лишних + ПодставитьЗамеры(wdDoc, _
                 CStr(ws.Cells(r, ColReq(ws, "ID записи")).Value), метод)
    If лишних > 0 Then
        MsgBox "По методу " & метод & " заведено на " & лишних & " точек больше, " & _
               "чем мест в шаблоне." & vbCrLf & vbCrLf & _
               "В документ попали первые — остальные остались в системе. " & _
               "Допишите их в документе руками или добавьте строки в таблицу шаблона.", _
               vbExclamation, "Формирование документов"
    End If

    папка = ПапкаГотовых() & CStr(год) & Application.PathSeparator & метод & Application.PathSeparator
    MkPath папка
    файл = папка & ОчиститьДляИмениФайла(РегНомер(ws, r, метод)) & " " & _
           ОчиститьДляИмениФайла(Left$(ОписаниеОбъекта(ws, r), 60)) & ".docx"

    wdDoc.SaveAs2 fileName:=файл
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    СформироватьОдин = файл
    Exit Function

Сбой:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Не удалось сформировать документ по методу " & метод & "." & vbCrLf & _
           "Ошибка: " & Err.Description, vbCritical, "Формирование документов"
End Function


' Публичная: её зовёт и modMeasure, подставляя точки замеров.
Public Sub ЗаменитьВДокументе(doc As Object, поле As String, значение As String)
    Подставить doc, поле, значение
End Sub


Private Sub Подставить(doc As Object, поле As String, значение As String)
    Dim rng As Object
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = "#" & поле
        .Replacement.ClearFormatting
        .Replacement.Text = значение
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub


' -- Паспорт объекта и данные работы -----------------------------------

' Все паспортные поля справочника объектов идут в документ под теми же
' именами, что стоят в заголовках колонок: «Марка материала» ->
' #Марка_материала. Отдельной таблицы соответствия нет намеренно —
' её пришлось бы править при каждом новом поле.
Private Sub ПодставитьПаспортОбъекта(doc As Object, ws As Worksheet, r As Long)
    Dim об As Worksheet
    Dim идОбъекта As String
    Dim строка As Long
    Dim c As Long, колонок As Long
    Dim имя As String

    идОбъекта = Trim$(CStr(ws.Cells(r, ColReq(ws, "ID объекта")).Value))
    If идОбъекта = "" Then Exit Sub
    Set об = ЛистТихо(Л_ОБЪЕКТЫ)
    If об Is Nothing Then Exit Sub

    строка = НайтиПоПервойКолонке(об, идОбъекта)
    If строка = 0 Then Exit Sub

    колонок = об.Cells(СТРОКА_ЗАГОЛОВКОВ, об.Columns.Count).End(xlToLeft).Column
    For c = 2 To колонок
        имя = ИмяПоля(CStr(об.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value))
        If имя <> "" Then
            Подставить doc, имя, ЗначениеТекстом(об.Cells(строка, c))
        End If
    Next c
End Sub


' Подставляет все колонки строки листа, найденной по ID записи журнала,
' начиная с указанной. Имя плейсхолдера берётся из заголовка колонки.
Public Sub ПодставитьИзЛиста(doc As Object, имяЛиста As String, _
                             идЗаписи As String, перваяКолонка As Long)
    Dim ws As Worksheet
    Dim i As Long, last As Long
    Dim c As Long, колонок As Long
    Dim строка As Long
    Dim имя As String

    If Trim$(идЗаписи) = "" Then Exit Sub
    Set ws = ЛистТихо(имяЛиста)
    If ws Is Nothing Then Exit Sub

    last = ПоследняяСтрока(ws, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(i, ColReq(ws, "ID записи журнала")).Value)), _
                   Trim$(идЗаписи), vbTextCompare) = 0 Then
            строка = i
            Exit For
        End If
    Next i
    If строка = 0 Then Exit Sub

    колонок = ws.Cells(СТРОКА_ЗАГОЛОВКОВ, ws.Columns.Count).End(xlToLeft).Column
    For c = перваяКолонка To колонок
        имя = ИмяПоля(CStr(ws.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value))
        If имя <> "" Then Подставить doc, имя, ЗначениеТекстом(ws.Cells(строка, c))
    Next c
End Sub


' Поля конкретной работы с листа «Данные_протоколов». Возвращает 0 —
' он ничего не переполняет, счёт нужен вызывающему для единообразия.
Private Function ПодставитьДанныеПротокола(doc As Object, ws As Worksheet, _
                                           r As Long, метод As String) As Long
    Dim дп As Worksheet
    Dim идЗаписи As String
    Dim i As Long, last As Long
    Dim c As Long, колонок As Long
    Dim имя As String
    Dim строка As Long

    идЗаписи = Trim$(CStr(ws.Cells(r, ColReq(ws, "ID записи")).Value))
    Set дп = ЛистТихо(Л_ДАННЫЕ_ПРОТОКОЛОВ)
    If дп Is Nothing Or идЗаписи = "" Then Exit Function

    last = ПоследняяСтрока(дп, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(дп.Cells(i, ColReq(дп, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(дп.Cells(i, ColReq(дп, "Метод")).Value)), _
                       метод, vbTextCompare) = 0 Then
                строка = i
                Exit For
            End If
        End If
    Next i
    If строка = 0 Then Exit Function

    колонок = дп.Cells(СТРОКА_ЗАГОЛОВКОВ, дп.Columns.Count).End(xlToLeft).Column
    For c = 4 To колонок
        имя = ИмяПоля(CStr(дп.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value))
        If имя <> "" Then
            Подставить doc, имя, ЗначениеТекстом(дп.Cells(строка, c))
        End If
    Next c

    ' Руководитель работ идёт в документ одной строкой: должность и ФИО.
    Подставить doc, "Должность_ФИО", Trim$( _
        ЗначениеТекстом(дп.Cells(строка, ColReq(дп, "Должность руководителя"))) & " " & _
        ЗначениеТекстом(дп.Cells(строка, ColReq(дп, "Руководитель работ"))))

    ПодставитьСредстваНК doc, дп, строка
    ПрименитьПереопределения doc, CStr(дп.Cells(строка, ColReq(дп, "Переопределения паспорта")).Value)
End Function


' Средства НК подставляются целиком: наименование, тип, заводской номер
' и срок поверки — чтобы их не переписывали в каждый протокол руками,
' как было в вашем шаблоне акта ВИК.
Private Sub ПодставитьСредстваНК(doc As Object, дп As Worksheet, строка As Long)
    Dim k As Long
    Dim ид As String
    Dim si As Worksheet
    Dim r As Long

    Set si = ЛистТихо(Л_СИ)
    If si Is Nothing Then Exit Sub

    For k = 1 To 2
        ид = Trim$(CStr(дп.Cells(строка, ColReq(дп, "ID средства НК " & k)).Value))
        r = 0
        If ид <> "" Then r = НайтиПоПервойКолонке(si, ид)
        If r > 0 Then
            Подставить doc, "СИ" & k & "_Наименование", ЗначениеТекстом(si.Cells(r, ColReq(si, "Наименование")))
            Подставить doc, "СИ" & k & "_Тип", ЗначениеТекстом(si.Cells(r, ColReq(si, "Тип / обозначение")))
            Подставить doc, "СИ" & k & "_Зав_номер", ЗначениеТекстом(si.Cells(r, ColReq(si, "Заводской №")))
            Подставить doc, "СИ" & k & "_Поверка_до", ЗначениеТекстом(si.Cells(r, ColReq(si, "Действительно до")))
        End If
    Next k
End Sub


' «Поле=Значение; Поле=Значение» — редкий случай, когда по этой работе
' паспортное значение другое. Правится только документ, справочник
' остаётся историей объекта.
Private Sub ПрименитьПереопределения(doc As Object, текст As String)
    Dim части As Variant
    Dim i As Long
    Dim пара As Variant

    If Trim$(текст) = "" Then Exit Sub
    части = Split(текст, ";")
    For i = LBound(части) To UBound(части)
        If InStr(части(i), "=") > 0 Then
            пара = Split(части(i), "=")
            Подставить doc, ИмяПоля(CStr(пара(0))), Trim$(CStr(пара(1)))
        End If
    Next i
End Sub


' Заголовок колонки -> имя плейсхолдера: пробелы и переносы в
' подчёркивания. Служебные колонки пропускаются.
Private Function ИмяПоля(заголовок As String) As String
    Dim s As String
    s = Trim$(Replace(Replace(заголовок, vbLf, " "), vbCr, " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Select Case True
        Case s = "": Exit Function
        Case InStr(1, s, "ID", vbTextCompare) = 1: Exit Function
        Case InStr(1, s, "Кем внесено", vbTextCompare) > 0: Exit Function
        Case InStr(1, s, "Изменено", vbTextCompare) > 0: Exit Function
        Case InStr(1, s, "Работ в журнале", vbTextCompare) > 0: Exit Function
        Case InStr(1, s, "Переопределения", vbTextCompare) > 0: Exit Function
        Case InStr(1, s, "Прежние наименования", vbTextCompare) > 0: Exit Function
    End Select
    ИмяПоля = Replace(Replace(s, " ", "_"), "-", "_")
End Function


Private Function ЗначениеТекстом(ячейка As Range) As String
    If IsError(ячейка.Value) Then Exit Function
    If IsDate(ячейка.Value) Then
        ЗначениеТекстом = Format$(ячейка.Value, "dd.mm.yyyy")
    Else
        ЗначениеТекстом = Trim$(CStr(ячейка.Value))
    End If
End Function


' Одно поле объекта по строке журнала — для случаев, когда имя
' плейсхолдера не совпадает с заголовком колонки.
Private Function ПолеОбъекта(ws As Worksheet, r As Long, заголовок As String) As String
    Dim об As Worksheet
    Dim строка As Long
    Dim c As Long

    Set об = ЛистТихо(Л_ОБЪЕКТЫ)
    If об Is Nothing Then Exit Function
    строка = НайтиПоПервойКолонке(об, Trim$(CStr(ws.Cells(r, ColReq(ws, "ID объекта")).Value)))
    If строка = 0 Then Exit Function
    c = ColByHeader(об, заголовок)
    If c > 0 Then ПолеОбъекта = ЗначениеТекстом(об.Cells(строка, c))
End Function


Private Function МестоРасположения(ws As Worksheet, r As Long) As String
    МестоРасположения = ПолеОбъекта(ws, r, "Место расположения")
    If МестоРасположения = "" Then
        МестоРасположения = CStr(ws.Cells(r, ColReq(ws, "Заказывающее подразделение")).Value)
    End If
End Function


Private Function НайтиПоПервойКолонке(ws As Worksheet, ид As String) As Long
    Dim r As Long, last As Long
    last = ПоследняяСтрока(ws, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), ид, vbTextCompare) = 0 Then
            НайтиПоПервойКолонке = r
            Exit Function
        End If
    Next r
End Function


Public Function ФайлШаблона(метод As String) As String
    Select Case UCase$(Trim$(метод))
        Case "ВИК": ФайлШаблона = "Акт_ВИК.docx"
        Case "УЗД": ФайлШаблона = "Заключение_УЗД.docx"
        Case "УЗТ": ФайлШаблона = "Протокол_УЗТ.docx"
        Case "АЭ": ФайлШаблона = "Протокол_АЭ_полный.docx"
        Case "ТВ": ФайлШаблона = "Протокол_ТВ.docx"
        Case "ТК": ФайлШаблона = "Протокол_ТК.docx"
        Case "ВД": ФайлШаблона = "Протокол_ВД.docx"
        Case Else: ФайлШаблона = ""      ' МК — шаблона пока нет
    End Select
End Function


Private Function ОписаниеОбъекта(ws As Worksheet, r As Long) As String
    Dim имя As String
    Dim зав As String
    имя = Trim$(CStr(ws.Cells(r, ColReq(ws, "Наименование объекта в документе")).Value))
    зав = Trim$(CStr(ws.Cells(r, ColReq(ws, "Заводской № объекта")).Value))
    If зав <> "" Then имя = имя & ", зав. № " & зав
    ОписаниеОбъекта = имя
End Function


Private Function ДатаИМесто(ws As Worksheet, r As Long) As String
    Dim s As String
    Dim место As String
    If IsDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value) Then
        s = Format$(ws.Cells(r, ColReq(ws, "Дата контроля")).Value, "dd.mm.yyyy")
    End If
    место = Trim$(CStr(ws.Cells(r, ColReq(ws, "Заказывающее подразделение")).Value))
    If место <> "" Then
        If s <> "" Then s = s & ", "
        s = s & место
    End If
    ДатаИМесто = s
End Function


' Подпись исполнителя: уровень и номер удостоверения берутся из реестра
' аттестаций — переписывать их в каждый документ руками незачем.
Private Function ПодписьИсполнителя(ws As Worksheet, r As Long, метод As String) As String
    Dim фио As String
    Dim att As Worksheet
    Dim i As Long, last As Long
    Dim уровень As String
    Dim удост As String

    фио = Trim$(CStr(ws.Cells(r, ColReq(ws, метод & " исполнитель")).Value))
    If фио = "" Then Exit Function

    Set att = Лист(Л_АТТЕСТАЦИИ)
    If att Is Nothing Then
        ПодписьИсполнителя = фио
        Exit Function
    End If
    last = ПоследняяСтрока(att, ColReq(att, "ID"))
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(att.Cells(i, ColReq(att, "ФИО")).Value)), фио, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(att.Cells(i, ColReq(att, "Метод контроля")).Value)), метод, vbTextCompare) = 0 Then
            уровень = Trim$(CStr(att.Cells(i, ColReq(att, "Уровень")).Value))
            удост = Trim$(CStr(att.Cells(i, ColReq(att, "№ квалификационного удостоверения")).Value))
            Exit For
        End If
    Next i

    ПодписьИсполнителя = Trim$(IIf(уровень = "", "", уровень & " уровень, ") & _
                               IIf(удост = "", "", "уд. № " & удост & ", ") & фио)
End Function
