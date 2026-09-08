Attribute VB_Name = "modEntry"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Ввод замеров и настроек АЭ на временном листе.
'
'  Зачем. Лист «Замеры» — хранилище: двадцать колонок, из которых по УЗТ
'  не нужны твёрдость и НВ, а по ТВ не нужны замеры и минимум. Точки всех
'  записей лежат вперемешку, схемы контроля перед глазами нет. Вводить
'  туда напрямую можно, но неудобно, а неудобство при вводе — это ошибки
'  в протоколе.
'
'  Как. Кнопка открывает временный лист: шапка работы, схема контроля
'  картинкой справа, слева — таблица только из нужных колонок, с
'  сечениями и точками. Минимум и НВ считаются на глазах, значение ниже
'  нормы краснеет сразу. По «Записать» проверенное уходит в «Замеры», а
'  лист удаляется.
'
'  Чего здесь нарочно НЕ делается: лист не удаляется при закрытии книги.
'  До записи данных нет больше нигде, и потерять полсотни точек из-за
'  случайно нажатой «Отмены» или сбоя Excel — цена, которой эта затея не
'  стоит. Незакрытый лист переживает закрытие книги, а при следующем
'  открытии система спрашивает: продолжить или выбросить.
' ═══════════════════════════════════════════════════════════════════════

Public Const Л_ВВОД As String = "Ввод_замеров"
Public Const Л_ВВОД_АЭ As String = "Ввод_АЭ"

' Служебные ячейки: в них лист помнит, к чему он относится. Колонка
' далеко справа и скрыта — человек её не видит, а макрос читает.
Private Const К_СЛУЖЕБНАЯ As Long = 30       ' AD
Private Const СТР_ЗАПИСЬ As Long = 1
Private Const СТР_МЕТОД As Long = 2
Private Const СТР_СТРОКА_Ж As Long = 3

Private Const СТР_ТАБЛИЦЫ As Long = 9        ' строка заголовков таблицы
Private Const К_СХЕМЫ As String = "R2"       ' куда прикалывается картинка


' ── Открыть ввод замеров ──────────────────────────────────────────────

Public Sub ВвестиЗамеры()
    Dim ж As Worksheet
    Dim r As Long
    Dim идЗаписи As String
    Dim метод As String

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub

    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))
    If идЗаписи = "" Then
        MsgBox "У этой строки нет ID записи — замеры не к чему привязать.", _
               vbExclamation, "Ввод замеров"
        Exit Sub
    End If

    If Not МожноЗанятьЛист(Л_ВВОД) Then Exit Sub

    метод = modMeasure.СпроситьМетод(ж, r)
    If метод = "" Then Exit Sub

    ПостроитьЛистЗамеров ж, r, идЗаписи, метод
End Sub


Private Sub ПостроитьЛистЗамеров(ж As Worksheet, r As Long, идЗаписи As String, метод As String)
    Dim ws As Worksheet
    Dim сечений As Long, точек As Long
    Dim строка As Long
    Dim с As Long, т As Long
    Dim норма As String
    Dim прибор As String
    Dim уже As Long

    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        сечений = СпроситьЧисло("Сколько сечений на объекте?", 5, 1, 30)
        If сечений = 0 Then Exit Sub
        точек = СпроситьЧисло("Сколько точек в каждом сечении?", 6, 1, 30)
        If точек = 0 Then Exit Sub
    Else
        сечений = 1
        точек = СпроситьЧисло("Сколько точек замера твёрдости?", 3, 1, 30)
        If точек = 0 Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Set ws = ПересоздатьЛист(Л_ВВОД)

    ws.Cells(СТР_ЗАПИСЬ, К_СЛУЖЕБНАЯ).Value = идЗаписи
    ws.Cells(СТР_МЕТОД, К_СЛУЖЕБНАЯ).Value = метод
    ws.Cells(СТР_СТРОКА_Ж, К_СЛУЖЕБНАЯ).Value = r
    ws.Columns(К_СЛУЖЕБНАЯ).Hidden = True

    норма = НормаОбъекта(ж, r, метод)
    прибор = ПриборРаботы(идЗаписи, метод)

    ШапкаЛиста ws, ж, r, метод, норма
    ЗаголовкиТаблицы ws, метод
    ВставитьСхему ws, ж, r, метод

    строка = СТР_ТАБЛИЦЫ + 1
    For с = 1 To сечений
        For т = 1 To точек
            ЗаполнитьСтрокуВвода ws, строка, метод, с, т, норма, прибор
            строка = строка + 1
        Next т
    Next с

    ' Уже введённое по этой записи и методу подтягивается сюда же:
    ' лист служит и вводом, и правкой — в «Замеры» руками лазить незачем.
    уже = ПодтянутьВведённое(ws, идЗаписи, метод)

    ' Оформляем по фактической последней строке, а не по счётчику выше:
    ' подтянутые точки могли выйти за пределы построенной сетки, и без
    ' этого у них не было бы ни рамок, ни подсветки ниже нормы.
    ОформитьТаблицу ws, метод, ПоследняяСтрокаВвода(ws)
    КнопкиЛиста ws
    ws.Activate
    ws.Cells(СТР_ТАБЛИЦЫ + 1, ПерваяКолонкаВвода(метод)).Select
    Application.ScreenUpdating = True

    MsgBox "Лист ввода готов." & vbCrLf & vbCrLf & _
           IIf(уже > 0, "Уже введённых точек подтянуто: " & уже & "." & vbCrLf & vbCrLf, "") & _
           IIf(StrComp(метод, "УЗТ", vbTextCompare) = 0, _
               "Заполняйте «Замер 1…5» — минимум посчитается сам.", _
               "Заполняйте «Твёрдость, МПа» — НВ пересчитается сам.") & vbCrLf & _
           "Значение ниже нормы краснеет сразу." & vbCrLf & vbCrLf & _
           "Закончите — «Записать в журнал замеров». Лист сам не удалится: " & _
           "закрытие книги ничего не потеряет.", _
           vbInformation, "Ввод замеров"
End Sub


' ── Шапка, таблица, схема ─────────────────────────────────────────────

Private Sub ШапкаЛиста(ws As Worksheet, ж As Worksheet, r As Long, _
                       метод As String, норма As String)
    Dim объект As String

    объект = Trim$(CStr(ж.Cells(r, ColReq(ж, "Наименование объекта в документе")).Value))
    If Trim$(CStr(ж.Cells(r, ColReq(ж, "Заводской № объекта")).Value)) <> "" Then
        объект = объект & ", зав. № " & CStr(ж.Cells(r, ColReq(ж, "Заводской № объекта")).Value)
    End If

    ws.Range("A2").Value = "ВВОД ЗАМЕРОВ — " & метод
    ws.Range("A2").Font.Size = 16
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Color = RGB(31, 56, 100)

    ws.Range("A3").Value = "Объект: " & объект
    ws.Range("A4").Value = "Запись: " & CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value) & _
                           "     Акт: " & CStr(ж.Cells(r, ColReq(ж, "№ акта")).Value) & _
                           "     Дата: " & Формат(ж.Cells(r, ColReq(ж, "Дата контроля")).Value)
    ws.Range("A5").Value = "Норма" & IIf(StrComp(метод, "УЗТ", vbTextCompare) = 0, _
                           " (мин. толщина, мм): ", " (твёрдость, НВ): ") & _
                           IIf(норма = "", "не заполнена в справочнике объектов", норма)
    ws.Range("A3:A5").Font.Color = RGB(89, 89, 89)
    If норма = "" Then ws.Range("A5").Font.Color = RGB(192, 0, 0)
End Sub


Private Sub ЗаголовкиТаблицы(ws As Worksheet, метод As String)
    Dim имена As Variant
    Dim i As Long

    имена = ИменаКолонок(метод)
    For i = LBound(имена) To UBound(имена)
        ws.Cells(СТР_ТАБЛИЦЫ, i + 1).Value = CStr(имена(i))
    Next i
    With ws.Range(ws.Cells(СТР_ТАБЛИЦЫ, 1), ws.Cells(СТР_ТАБЛИЦЫ, UBound(имена) + 1))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(СТР_ТАБЛИЦЫ).RowHeight = 30
End Sub


' Колонки листа ввода — только те, что нужны этому методу. В «Замерах»
' их двадцать на оба метода сразу; здесь лишних нет.
Private Function ИменаКолонок(метод As String) As Variant
    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        ИменаКолонок = Array("№ сеч.", "№ точки", "Замер 1", "Замер 2", "Замер 3", _
                             "Замер 4", "Замер 5", "Минимум", "Норма", _
                             "Зав. № прибора", "Заключение по точке", "Примечание")
    Else
        ИменаКолонок = Array("№ точки", "Твёрдость, МПа", "НВ", "Норма", _
                             "Зав. № прибора", "Заключение по точке", "Примечание")
    End If
End Function


Private Function ПерваяКолонкаВвода(метод As String) As Long
    ' УЗТ: сечение и точка проставлены, курсор ставится на «Замер 1».
    ' ТВ: точка проставлена, курсор — на «Твёрдость».
    ПерваяКолонкаВвода = IIf(StrComp(метод, "УЗТ", vbTextCompare) = 0, 3, 2)
End Function


Private Sub ЗаполнитьСтрокуВвода(ws As Worksheet, строка As Long, метод As String, _
                                 сечение As Long, точка As Long, _
                                 норма As String, прибор As String)
    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        ws.Cells(строка, 1).Value = сечение
        ws.Cells(строка, 2).Value = точка
        ws.Cells(строка, 8).Formula = _
            "=IF(COUNT(C" & строка & ":G" & строка & ")=0,"""",MIN(C" & строка & ":G" & строка & "))"
        If норма <> "" Then ws.Cells(строка, 9).Value = CDbl(Val(Replace(норма, ",", ".")))
        If прибор <> "" Then ws.Cells(строка, 10).Value = прибор
    Else
        ws.Cells(строка, 1).Value = точка
        ws.Cells(строка, 3).Formula = _
            "=IF(B" & строка & "="""","""",IFERROR(INDEX(Твёрдость_НВ!$A:$A," & _
            "MATCH(B" & строка & ",Твёрдость_НВ!$B:$B,1)),""вне таблицы""))"
        If норма <> "" Then ws.Cells(строка, 4).Value = CDbl(Val(Replace(норма, ",", ".")))
        If прибор <> "" Then ws.Cells(строка, 5).Value = прибор
    End If
End Sub


Private Sub ОформитьТаблицу(ws As Worksheet, метод As String, последняя As Long)
    Dim колонок As Long
    Dim расч As Long, норм As Long
    Dim обл As Range

    колонок = UBound(ИменаКолонок(метод)) + 1
    расч = IIf(StrComp(метод, "УЗТ", vbTextCompare) = 0, 8, 3)
    норм = расч + 1

    Set обл = ws.Range(ws.Cells(СТР_ТАБЛИЦЫ, 1), ws.Cells(последняя, колонок))
    обл.Borders.LineStyle = xlContinuous
    обл.Borders.Color = RGB(150, 150, 150)

    ' Расчётные колонки серым — их не заполняют руками.
    With ws.Range(ws.Cells(СТР_ТАБЛИЦЫ + 1, расч), ws.Cells(последняя, расч))
        .Interior.Color = RGB(235, 235, 235)
        .Font.Bold = True
    End With

    ' Ниже нормы — красным, сразу при вводе. Ради этого всё и затевалось:
    ' ошибку видно, пока прибор ещё в руках, а не в готовом протоколе.
    With ws.Range(ws.Cells(СТР_ТАБЛИЦЫ + 1, расч), ws.Cells(последняя, расч))
        .FormatConditions.Delete
        With .FormatConditions.Add(Type:=xlExpression, _
             Formula1:="=AND(" & Буква(ws, расч) & (СТР_ТАБЛИЦЫ + 1) & "<>""""," & _
                       Буква(ws, норм) & (СТР_ТАБЛИЦЫ + 1) & "<>""""," & _
                       Буква(ws, расч) & (СТР_ТАБЛИЦЫ + 1) & "<" & Буква(ws, норм) & (СТР_ТАБЛИЦЫ + 1) & ")")
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
    End With

    ws.Columns(1).ColumnWidth = 8
    ws.Columns(2).ColumnWidth = 9
    ws.Columns(3).Resize(, колонок - 2).ColumnWidth = 11
    ws.Columns(колонок - 1).ColumnWidth = 26
    ws.Columns(колонок).ColumnWidth = 20
    ws.Cells.Font.Name = "Calibri"
End Sub


' Схема ставится с колонки R, а не сразу за таблицей: кнопки шириной
' 150 пунктов занимают N-P, и картинка легла бы поверх них.
' Схема контроля — та же картинка, что уйдёт в протокол: берётся тем же
' поиском по методу и типу объекта. Ставится справа от таблицы, чтобы не
' налезать на строки и не мешать их добавлять.
Private Sub ВставитьСхему(ws As Worksheet, ж As Worksheet, r As Long, метод As String)
    Dim файл As String
    Dim подпись As String
    Dim ширина As Single, высота As Single
    Dim рис As Object
    Dim якорь As Range

    файл = modGenerate.НайтиСхему(метод, modGenerate.ПолеОбъекта(ж, r, "Тип / группа"), _
                                  подпись, ширина, высота)
    Set якорь = ws.Range(К_СХЕМЫ)

    If файл = "" Then
        якорь.Value = "Схема контроля не заведена: реестр «Схемы_контроля» или папка схем пусты."
        якорь.Font.Italic = True
        якорь.Font.Color = RGB(150, 150, 150)
        Exit Sub
    End If

    On Error Resume Next
    Set рис = ws.Shapes.AddPicture(fileName:=файл, LinkToFile:=False, _
                                   SaveWithDocument:=True, _
                                   Left:=якорь.Left, Top:=якорь.Top, Width:=-1, Height:=-1)
    On Error GoTo 0
    If рис Is Nothing Then Exit Sub

    ' Вписываем в рамку тех же размеров, что и в протоколе, сохраняя
    ' пропорции: то, что видно при замере, и то, что уйдёт в документ, —
    ' должно быть одной картинкой в одном масштабе.
    рис.LockAspectRatio = True
    If рис.Width > ширина Then рис.Width = ширина
    If рис.Height > высота Then рис.Height = высота
    If подпись <> "" Then
        ws.Range(К_СХЕМЫ).Offset(1, 0).Value = подпись
    End If
End Sub


' ── Запись в «Замеры» ─────────────────────────────────────────────────

Public Sub ЗаписатьЗамеры()
    Dim ws As Worksheet
    Dim зам As Worksheet
    Dim идЗаписи As String, метод As String
    Dim строка As Long, последняя As Long
    Dim записано As Long, обновлено As Long
    Dim беда As String

    Set ws = ЛистТихо(Л_ВВОД)
    If ws Is Nothing Then
        MsgBox "Лист ввода не открыт.", vbExclamation, "Ввод замеров"
        Exit Sub
    End If
    идЗаписи = Trim$(CStr(ws.Cells(СТР_ЗАПИСЬ, К_СЛУЖЕБНАЯ).Value))
    метод = Trim$(CStr(ws.Cells(СТР_МЕТОД, К_СЛУЖЕБНАЯ).Value))
    If идЗаписи = "" Or метод = "" Then
        MsgBox "Лист ввода не помнит, к какой записи относится. " & _
               "Закройте его и откройте ввод заново.", vbExclamation, "Ввод замеров"
        Exit Sub
    End If

    последняя = ПоследняяСтрокаВвода(ws)
    беда = ПроверитьВвод(ws, метод, последняя)
    If беда <> "" Then
        MsgBox "Записывать пока рано:" & vbCrLf & vbCrLf & беда, _
               vbExclamation, "Ввод замеров"
        Exit Sub
    End If

    Set зам = Лист(Л_ЗАМЕРЫ)
    If зам Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    For строка = СТР_ТАБЛИЦЫ + 1 To последняя
        If Не_пустая(ws, строка, метод) Then
            If ПеренестиСтроку(ws, строка, зам, идЗаписи, метод) Then
                обновлено = обновлено + 1
            Else
                записано = записано + 1
            End If
        End If
    Next строка
    Application.ScreenUpdating = True

    If записано + обновлено = 0 Then
        MsgBox "Ни одной заполненной точки — записывать нечего.", _
               vbInformation, "Ввод замеров"
        Exit Sub
    End If

    ПерейтиКЖурналу
    УдалитьЛист Л_ВВОД
    MsgBox "Записано в журнал замеров." & vbCrLf & vbCrLf & _
           "Новых точек: " & записано & vbCrLf & _
           "Обновлено ранее введённых: " & обновлено & vbCrLf & vbCrLf & _
           "Лист ввода закрыт.", vbInformation, "Ввод замеров"
End Sub


' True — строка обновила существующую точку, False — завела новую.
Private Function ПеренестиСтроку(ws As Worksheet, строка As Long, зам As Worksheet, _
                                 идЗаписи As String, метод As String) As Boolean
    Dim r As Long
    Dim сечение As Variant, точка As Variant
    Dim c As Long

    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        сечение = ws.Cells(строка, 1).Value
        точка = ws.Cells(строка, 2).Value
    Else
        сечение = ""
        точка = ws.Cells(строка, 1).Value
    End If

    r = НайтиТочку(зам, идЗаписи, метод, сечение, точка)
    If r = 0 Then
        r = ПоследняяСтрока(зам, 1) + 1
        зам.Cells(r, 1).Value = НовыйIDЗамера(зам)
        зам.Cells(r, ColReq(зам, "ID записи журнала")).Value = идЗаписи
        зам.Cells(r, ColReq(зам, "Метод")).Value = метод
        зам.Cells(r, ColReq(зам, "№ сечения")).Value = сечение
        зам.Cells(r, ColReq(зам, "№ точки")).Value = точка
        ПеренестиСтроку = False
    Else
        ПеренестиСтроку = True
    End If

    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        For c = 1 To 5
            зам.Cells(r, ColReq(зам, "Замер " & c)).Value = ws.Cells(строка, 2 + c).Value
        Next c
        зам.Cells(r, ColReq(зам, "Нормируемое значение")).Value = ws.Cells(строка, 9).Value
        зам.Cells(r, ColReq(зам, "Зав. № прибора")).Value = ws.Cells(строка, 10).Value
        зам.Cells(r, ColReq(зам, "Заключение по точке")).Value = ws.Cells(строка, 11).Value
        зам.Cells(r, ColReq(зам, "Примечание")).Value = ws.Cells(строка, 12).Value
    Else
        зам.Cells(r, ColReq(зам, "Твёрдость, МПа")).Value = ws.Cells(строка, 2).Value
        зам.Cells(r, ColReq(зам, "Нормируемое значение")).Value = ws.Cells(строка, 4).Value
        зам.Cells(r, ColReq(зам, "Зав. № прибора")).Value = ws.Cells(строка, 5).Value
        зам.Cells(r, ColReq(зам, "Заключение по точке")).Value = ws.Cells(строка, 6).Value
        зам.Cells(r, ColReq(зам, "Примечание")).Value = ws.Cells(строка, 7).Value
    End If

    ' «Минимум» и «НВ» не переносятся: в «Замерах» на них свои формулы,
    ' и записать туда число значило бы затереть формулу значением.
    зам.Cells(r, ColReq(зам, "Кем внесено")).Value = ТекущийСотрудник()
    зам.Cells(r, ColReq(зам, "Изменено")).Value = Now
End Function


Private Function НайтиТочку(зам As Worksheet, идЗаписи As String, метод As String, _
                            сечение As Variant, точка As Variant) As Long
    Dim r As Long, last As Long
    last = ПоследняяСтрока(зам, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(зам.Cells(r, ColReq(зам, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(зам.Cells(r, ColReq(зам, "Метод")).Value)), _
                       метод, vbTextCompare) = 0 Then
                If CStr(зам.Cells(r, ColReq(зам, "№ сечения")).Value) = CStr(сечение) Then
                    If CStr(зам.Cells(r, ColReq(зам, "№ точки")).Value) = CStr(точка) Then
                        НайтиТочку = r
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r
End Function


' ── Проверки перед записью ────────────────────────────────────────────

Private Function ПроверитьВвод(ws As Worksheet, метод As String, последняя As Long) As String
    Dim строка As Long
    Dim c As Long, откуда As Long, докуда As Long
    Dim ключ As String
    Dim видели As Object
    Dim нечисло As String
    Dim дубли As String
    Dim v As Variant

    Set видели = CreateObject("Scripting.Dictionary")
    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        откуда = 3: докуда = 7
    Else
        откуда = 2: докуда = 2
    End If

    For строка = СТР_ТАБЛИЦЫ + 1 To последняя
        If Не_пустая(ws, строка, метод) Then
            For c = откуда To докуда
                v = ws.Cells(строка, c).Value
                If Trim$(CStr(v)) <> "" And Not IsNumeric(v) Then
                    нечисло = нечисло & vbCrLf & "  строка " & строка & ", «" & _
                              CStr(ws.Cells(СТР_ТАБЛИЦЫ, c).Value) & "»: " & CStr(v)
                End If
            Next c

            If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
                ключ = CStr(ws.Cells(строка, 1).Value) & "/" & CStr(ws.Cells(строка, 2).Value)
            Else
                ключ = CStr(ws.Cells(строка, 1).Value)
            End If
            If видели.Exists(ключ) Then
                дубли = дубли & vbCrLf & "  " & ключ & " — строки " & видели(ключ) & " и " & строка
            Else
                видели(ключ) = строка
            End If
        End If
    Next строка

    If нечисло <> "" Then
        ПроверитьВвод = "В замерах есть не числа:" & нечисло & vbCrLf & vbCrLf & _
                        "Десятичный разделитель — запятая, как везде в книге."
        Exit Function
    End If
    If дубли <> "" Then
        ПроверитьВвод = "Одна и та же точка встречается дважды" & _
                        IIf(StrComp(метод, "УЗТ", vbTextCompare) = 0, " (сечение/точка)", "") & _
                        ":" & дубли
    End If
End Function


Private Function Не_пустая(ws As Worksheet, строка As Long, метод As String) As Boolean
    Dim c As Long, откуда As Long, докуда As Long
    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        откуда = 3: докуда = 7
    Else
        откуда = 2: докуда = 2
    End If
    For c = откуда To докуда
        If Trim$(CStr(ws.Cells(строка, c).Value)) <> "" Then
            Не_пустая = True
            Exit Function
        End If
    Next c
End Function


' ── Кнопки листа ──────────────────────────────────────────────────────

Public Sub ОтменитьВвод()
    Dim ws As Worksheet
    Dim есть As Boolean
    Dim имяЛиста As String

    Set ws = ЛистТихо(Л_ВВОД)
    If ws Is Nothing Then Set ws = ЛистТихо(Л_ВВОД_АЭ)
    If ws Is Nothing Then Exit Sub

    есть = ЕстьЗаполненное(ws)
    If есть Then
        If MsgBox("На листе есть заполненное, и оно нигде больше не сохранено." & _
                  vbCrLf & vbCrLf & "Точно выбросить?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Ввод замеров") = vbNo Then Exit Sub
    End If

    имяЛиста = ws.Name
    ПерейтиКЖурналу
    УдалитьЛист имяЛиста
End Sub


Public Sub ДобавитьСтрокиВвода()
    Dim ws As Worksheet
    Dim метод As String
    Dim последняя As Long, сколько As Long, i As Long
    Dim сечение As Long, точка As Long
    Dim норма As String

    Set ws = ЛистТихо(Л_ВВОД)
    If ws Is Nothing Then Exit Sub
    метод = Trim$(CStr(ws.Cells(СТР_МЕТОД, К_СЛУЖЕБНАЯ).Value))
    последняя = ПоследняяСтрокаВвода(ws)

    сколько = СпроситьЧисло("Сколько строк добавить?", 6, 1, 60)
    If сколько = 0 Then Exit Sub

    If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
        сечение = CLng(Val(ws.Cells(последняя, 1).Value))
        точка = CLng(Val(ws.Cells(последняя, 2).Value))
        норма = CStr(ws.Cells(последняя, 9).Value)
    Else
        сечение = 0
        точка = CLng(Val(ws.Cells(последняя, 1).Value))
        норма = CStr(ws.Cells(последняя, 4).Value)
    End If

    For i = 1 To сколько
        точка = точка + 1
        ЗаполнитьСтрокуВвода ws, последняя + i, метод, сечение, точка, норма, ""
    Next i
    ОформитьТаблицу ws, метод, последняя + сколько
    ws.Cells(последняя + 1, ПерваяКолонкаВвода(метод)).Select
End Sub


Private Sub КнопкиЛиста(ws As Worksheet)
    ПоложитьКнопку ws, "N2", "Записать в журнал замеров", "ЗаписатьЗамеры"
    ПоложитьКнопку ws, "N4", "Добавить строки", "ДобавитьСтрокиВвода"
    ПоложитьКнопку ws, "N6", "Отмена", "ОтменитьВвод"
End Sub


Private Sub ПоложитьКнопку(ws As Worksheet, адрес As String, надпись As String, макрос As String)
    Dim кн As Object
    Dim ячейка As Range
    Set ячейка = ws.Range(адрес)
    Set кн = ws.Buttons.Add(ячейка.Left, ячейка.Top, 150, 26)
    кн.Caption = надпись
    ' С именем книги: кнопка на листе книги данных, макрос — в программе.
    кн.OnAction = "'" & ThisWorkbook.Name & "'!" & макрос
    кн.Font.Size = 9
End Sub


' ── Настройки АЭ ──────────────────────────────────────────────────────

' Двенадцать полей аппаратуры в общей таблице — это горизонтальная
' простыня, в которой не видно, что чему соответствует. Здесь они стоят
' столбиком: слева название, справа значение.
Public Sub ВвестиНастройкиАЭ()
    Dim ж As Worksheet
    Dim ws As Worksheet
    Dim аэ As Worksheet
    Dim r As Long, i As Long, строка As Long
    Dim идЗаписи As String
    Dim источник As Long
    Dim колонок As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub
    идЗаписи = Trim$(CStr(ж.Cells(r, ColReq(ж, "ID записи")).Value))
    If идЗаписи = "" Then Exit Sub

    If Not МетодПрименён(ж, r, "АЭ") Then
        MsgBox "В этой строке не отмечен метод АЭ.", vbExclamation, "Настройки АЭ"
        Exit Sub
    End If
    If Not МожноЗанятьЛист(Л_ВВОД_АЭ) Then Exit Sub

    Set аэ = Лист(Л_ДАННЫЕ_АЭ)
    If аэ Is Nothing Then Exit Sub
    колонок = аэ.Cells(СТРОКА_ЗАГОЛОВКОВ, аэ.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False
    Set ws = ПересоздатьЛист(Л_ВВОД_АЭ)
    ws.Cells(СТР_ЗАПИСЬ, К_СЛУЖЕБНАЯ).Value = идЗаписи
    ws.Cells(СТР_МЕТОД, К_СЛУЖЕБНАЯ).Value = "АЭ"
    ws.Columns(К_СЛУЖЕБНАЯ).Hidden = True

    ws.Range("A2").Value = "НАСТРОЙКИ АППАРАТУРЫ АЭ"
    ws.Range("A2").Font.Size = 16
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Color = RGB(31, 56, 100)
    ws.Range("A3").Value = "Запись: " & идЗаписи & "     Объект: " & _
        CStr(ж.Cells(r, ColReq(ж, "Наименование объекта в документе")).Value)
    ws.Range("A3").Font.Color = RGB(89, 89, 89)

    ' Уже введённое по этой записи; если его нет — берём прошлую работу
    ' по АЭ: приборы и полоса частот от работы к работе меняются мало.
    источник = СтрокаАЭ(аэ, идЗаписи)
    If источник = 0 Then источник = ПоследняяЗаполненнаяАЭ(аэ)

    строка = 5
    For i = 4 To колонок - 3          ' поля между «Метод» и «Кем внесено»
        ws.Cells(строка, 1).Value = CStr(аэ.Cells(СТРОКА_ЗАГОЛОВКОВ, i).Value)
        ws.Cells(строка, 1).Font.Bold = True
        If источник > 0 Then ws.Cells(строка, 2).Value = аэ.Cells(источник, i).Value
        ws.Cells(строка, 2).Interior.Color = RGB(255, 253, 231)
        ws.Cells(строка, 2).Borders.LineStyle = xlContinuous
        строка = строка + 1
    Next i

    ws.Columns(1).ColumnWidth = 38
    ws.Columns(2).ColumnWidth = 46
    ws.Cells.Font.Name = "Calibri"
    ПоложитьКнопку ws, "E2", "Записать настройки", "ЗаписатьНастройкиАЭ"
    ПоложитьКнопку ws, "E4", "Отмена", "ОтменитьВвод"
    ws.Activate
    ws.Range("B5").Select
    Application.ScreenUpdating = True

    MsgBox IIf(источник = 0, "Прошлых настроек АЭ в системе нет — заполняйте с нуля.", _
               "Подставлены настройки " & IIf(СтрокаАЭ(аэ, идЗаписи) > 0, _
               "этой же работы.", "прошлой работы по АЭ — поправьте, что изменилось.")), _
           vbInformation, "Настройки АЭ"
End Sub


Public Sub ЗаписатьНастройкиАЭ()
    Dim ws As Worksheet
    Dim аэ As Worksheet
    Dim идЗаписи As String
    Dim r As Long, i As Long, строка As Long
    Dim колонок As Long

    Set ws = ЛистТихо(Л_ВВОД_АЭ)
    If ws Is Nothing Then Exit Sub
    идЗаписи = Trim$(CStr(ws.Cells(СТР_ЗАПИСЬ, К_СЛУЖЕБНАЯ).Value))
    If идЗаписи = "" Then Exit Sub

    Set аэ = Лист(Л_ДАННЫЕ_АЭ)
    If аэ Is Nothing Then Exit Sub
    колонок = аэ.Cells(СТРОКА_ЗАГОЛОВКОВ, аэ.Columns.Count).End(xlToLeft).Column

    r = СтрокаАЭ(аэ, идЗаписи)
    If r = 0 Then
        r = ПоследняяСтрока(аэ, 1) + 1
        аэ.Cells(r, 1).Value = НовыйIDАЭ(аэ)
        аэ.Cells(r, ColReq(аэ, "ID записи журнала")).Value = идЗаписи
    End If

    строка = 5
    For i = 4 To колонок - 3
        аэ.Cells(r, i).Value = ws.Cells(строка, 2).Value
        строка = строка + 1
    Next i
    аэ.Cells(r, ColReq(аэ, "Кем внесено")).Value = ТекущийСотрудник()
    аэ.Cells(r, ColReq(аэ, "Изменено")).Value = Now

    ПерейтиКЖурналу
    УдалитьЛист Л_ВВОД_АЭ
    MsgBox "Настройки АЭ записаны.", vbInformation, "Настройки АЭ"
End Sub


' ── Незакрытый ввод ───────────────────────────────────────────────────

' Вызывается при открытии книги. Лист ввода мог остаться с прошлого раза
' — значит, кто-то не дописал. Молча удалять нельзя: там единственный
' экземпляр данных.
Public Sub ПродолжитьНезакрытыйВвод()
    Dim имя As Variant
    Dim ws As Worksheet

    For Each имя In Array(Л_ВВОД, Л_ВВОД_АЭ)
        Set ws = ЛистТихо(CStr(имя))
        If Not ws Is Nothing Then
            If ЕстьЗаполненное(ws) Then
                ws.Activate
                MsgBox "Остался незакрытый лист ввода «" & CStr(имя) & "» — " & _
                       "в прошлый раз его не дописали." & vbCrLf & vbCrLf & _
                       "Данные на месте. Допишите и нажмите «Записать», " & _
                       "либо «Отмена», если они не нужны.", _
                       vbInformation, "Ввод замеров"
                Exit Sub
            Else
                УдалитьЛист CStr(имя)      ' пустой — уносим молча
            End If
        End If
    Next имя
End Sub


' ── Мелочи ────────────────────────────────────────────────────────────

Private Function МожноЗанятьЛист(имя As String) As Boolean
    Dim ws As Worksheet
    Set ws = ЛистТихо(имя)
    If ws Is Nothing Then
        МожноЗанятьЛист = True
        Exit Function
    End If
    If Not ЕстьЗаполненное(ws) Then
        МожноЗанятьЛист = True
        Exit Function
    End If
    ws.Activate
    MsgBox "Уже открыт лист ввода, и в нём есть заполненное." & vbCrLf & vbCrLf & _
           "Допишите и нажмите «Записать» — или «Отмена», если не нужно. " & _
           "Двух листов ввода сразу не бывает: перепутать, куда что пишется, " & _
           "слишком легко.", vbExclamation, "Ввод замеров"
End Function


' Есть ли на листе то, что ввёл человек. Считаются ТОЛЬКО те колонки,
' которые он и заполняет: номера сечений и точек, норма и прибор
' проставлены программой, расчётные посчитаны формулой — сами по себе
' они не означают, что кто-то что-то ввёл.
'
' Функция важнее, чем кажется: по ней решается, можно ли молча удалить
' оставшийся лист. Ошибётся в одну сторону — сотрёт чужую работу,
' ошибётся в другую — не даст открыть новый ввод.
Private Function ЕстьЗаполненное(ws As Worksheet) As Boolean
    Dim r As Long, c As Long
    Dim последняя As Long
    Dim колонки As Variant
    Dim i As Long

    последняя = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    If StrComp(ws.Name, Л_ВВОД_АЭ, vbTextCompare) = 0 Then
        ' Настройки АЭ: значения стоят в колонке B, начиная с пятой строки.
        For r = 5 To последняя
            If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then
                ЕстьЗаполненное = True
                Exit Function
            End If
        Next r
        Exit Function
    End If

    ' Замеры: колонки, которые заполняет человек, — свои у каждого метода.
    If StrComp(Trim$(CStr(ws.Cells(СТР_МЕТОД, К_СЛУЖЕБНАЯ).Value)), "УЗТ", vbTextCompare) = 0 Then
        колонки = Array(3, 4, 5, 6, 7, 11, 12)      ' замеры, заключение, примечание
    Else
        колонки = Array(2, 6, 7)                     ' твёрдость, заключение, примечание
    End If

    For r = СТР_ТАБЛИЦЫ + 1 To последняя
        For i = LBound(колонки) To UBound(колонки)
            c = CLng(колонки(i))
            If Trim$(CStr(ws.Cells(r, c).Value)) <> "" Then
                ЕстьЗаполненное = True
                Exit Function
            End If
        Next i
    Next r
End Function


Private Function ПоследняяСтрокаВвода(ws As Worksheet) As Long
    ПоследняяСтрокаВвода = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ПоследняяСтрокаВвода < СТР_ТАБЛИЦЫ + 1 Then ПоследняяСтрокаВвода = СТР_ТАБЛИЦЫ + 1
End Function


Private Function ПодтянутьВведённое(ws As Worksheet, идЗаписи As String, метод As String) As Long
    Dim зам As Worksheet
    Dim r As Long, last As Long, строка As Long
    Dim сечение As Variant, точка As Variant
    Dim c As Long
    Dim найдено As Long

    Set зам = ЛистТихо(Л_ЗАМЕРЫ)
    If зам Is Nothing Then Exit Function
    last = ПоследняяСтрока(зам, 1)

    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(зам.Cells(r, ColReq(зам, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(зам.Cells(r, ColReq(зам, "Метод")).Value)), _
                       метод, vbTextCompare) = 0 Then
                сечение = зам.Cells(r, ColReq(зам, "№ сечения")).Value
                точка = зам.Cells(r, ColReq(зам, "№ точки")).Value
                строка = НайтиСтрокуВвода(ws, метод, сечение, точка)
                If строка = 0 Then
                    строка = ПоследняяСтрокаВвода(ws) + 1
                    ЗаполнитьСтрокуВвода ws, строка, метод, CLng(Val(сечение)), _
                                         CLng(Val(точка)), "", ""
                End If
                If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
                    For c = 1 To 5
                        ws.Cells(строка, 2 + c).Value = зам.Cells(r, ColReq(зам, "Замер " & c)).Value
                    Next c
                    ws.Cells(строка, 9).Value = зам.Cells(r, ColReq(зам, "Нормируемое значение")).Value
                    ws.Cells(строка, 10).Value = зам.Cells(r, ColReq(зам, "Зав. № прибора")).Value
                    ws.Cells(строка, 11).Value = зам.Cells(r, ColReq(зам, "Заключение по точке")).Value
                    ws.Cells(строка, 12).Value = зам.Cells(r, ColReq(зам, "Примечание")).Value
                Else
                    ws.Cells(строка, 2).Value = зам.Cells(r, ColReq(зам, "Твёрдость, МПа")).Value
                    ws.Cells(строка, 4).Value = зам.Cells(r, ColReq(зам, "Нормируемое значение")).Value
                    ws.Cells(строка, 5).Value = зам.Cells(r, ColReq(зам, "Зав. № прибора")).Value
                    ws.Cells(строка, 6).Value = зам.Cells(r, ColReq(зам, "Заключение по точке")).Value
                    ws.Cells(строка, 7).Value = зам.Cells(r, ColReq(зам, "Примечание")).Value
                End If
                найдено = найдено + 1
            End If
        End If
    Next r
    ПодтянутьВведённое = найдено
End Function


Private Function НайтиСтрокуВвода(ws As Worksheet, метод As String, _
                                  сечение As Variant, точка As Variant) As Long
    Dim r As Long, последняя As Long
    последняя = ПоследняяСтрокаВвода(ws)
    For r = СТР_ТАБЛИЦЫ + 1 To последняя
        If StrComp(метод, "УЗТ", vbTextCompare) = 0 Then
            If CStr(ws.Cells(r, 1).Value) = CStr(сечение) And _
               CStr(ws.Cells(r, 2).Value) = CStr(точка) Then
                НайтиСтрокуВвода = r
                Exit Function
            End If
        Else
            If CStr(ws.Cells(r, 1).Value) = CStr(точка) Then
                НайтиСтрокуВвода = r
                Exit Function
            End If
        End If
    Next r
End Function


Private Function НормаОбъекта(ж As Worksheet, r As Long, метод As String) As String
    Dim заголовок As String
    заголовок = IIf(StrComp(метод, "УЗТ", vbTextCompare) = 0, _
                    "Мин. допустимая толщина, мм", "Допустимая твёрдость, НВ")
    НормаОбъекта = modGenerate.ПолеОбъекта(ж, r, заголовок)
End Function


Private Function ПриборРаботы(идЗаписи As String, метод As String) As String
    Dim пс As Worksheet
    Dim r As Long, last As Long
    Dim си As Worksheet
    Dim идСИ As String
    Dim i As Long

    Set пс = ЛистТихо("Применённые_СИ")
    If пс Is Nothing Then Exit Function
    last = ПоследняяСтрока(пс, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(пс.Cells(r, ColReq(пс, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(пс.Cells(r, ColReq(пс, "Метод")).Value)), _
                       метод, vbTextCompare) = 0 Then
                идСИ = Trim$(CStr(пс.Cells(r, ColReq(пс, "ID средства НК")).Value))
                Exit For
            End If
        End If
    Next r
    If идСИ = "" Then Exit Function

    Set си = ЛистТихо(Л_СИ)
    If си Is Nothing Then Exit Function
    last = ПоследняяСтрока(си, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(си.Cells(i, 1).Value)), идСИ, vbTextCompare) = 0 Then
            ПриборРаботы = Trim$(CStr(си.Cells(i, ColReq(си, "Заводской №")).Value))
            Exit Function
        End If
    Next i
End Function


Private Function НовыйIDЗамера(зам As Worksheet) As String
    НовыйIDЗамера = НомерСПрефиксом(зам, modExchange.ПрефиксID(Л_ЗАМЕРЫ))
End Function


Private Function НовыйIDАЭ(аэ As Worksheet) As String
    НовыйIDАЭ = НомерСПрефиксом(аэ, modExchange.ПрефиксID(Л_ДАННЫЕ_АЭ))
End Function


Private Function НомерСПрефиксом(ws As Worksheet, префикс As String) As String
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
    НомерСПрефиксом = префикс & Format$(n + 1, "0000")
End Function


Private Function СтрокаАЭ(аэ As Worksheet, идЗаписи As String) As Long
    Dim r As Long, last As Long
    last = ПоследняяСтрока(аэ, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(аэ.Cells(r, ColReq(аэ, "ID записи журнала")).Value)), _
                   идЗаписи, vbTextCompare) = 0 Then
            СтрокаАЭ = r
            Exit Function
        End If
    Next r
End Function


Private Function ПоследняяЗаполненнаяАЭ(аэ As Worksheet) As Long
    Dim r As Long, last As Long
    last = ПоследняяСтрока(аэ, 1)
    For r = last To ПЕРВАЯ_СТРОКА Step -1
        If Trim$(CStr(аэ.Cells(r, 4).Value)) <> "" Then
            ПоследняяЗаполненнаяАЭ = r
            Exit Function
        End If
    Next r
End Function


Private Function ПересоздатьЛист(имя As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Данные()
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(имя).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(1))
    On Error Resume Next
    ws.Name = имя
    On Error GoTo 0
    ws.Tab.Color = RGB(0, 176, 80)
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Set ПересоздатьЛист = ws
End Function


Private Sub УдалитьЛист(имя As String)
    Dim wb As Workbook
    Set wb = Данные()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(имя).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub


Private Sub ПерейтиКЖурналу()
    Dim ж As Worksheet
    Set ж = ЛистТихо(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    ж.Activate
End Sub


Private Function СпроситьЧисло(вопрос As String, поумолчанию As Long, _
                               наименьшее As Long, наибольшее As Long) As Long
    Dim s As String
    Dim n As Long
    s = InputBox(вопрос, "Ввод замеров", CStr(поумолчанию))
    If StrPtr(s) = 0 Then Exit Function        ' нажали «Отмена»
    s = Trim$(s)
    If Not IsNumeric(s) Then Exit Function
    n = CLng(Val(s))
    If n < наименьшее Or n > наибольшее Then
        MsgBox "Ожидается число от " & наименьшее & " до " & наибольшее & ".", _
               vbExclamation, "Ввод замеров"
        Exit Function
    End If
    СпроситьЧисло = n
End Function


' Буква колонки. Лист передаётся явно: без него Cells берёт активный,
' а он в этот момент может быть уже не тот.
Private Function Буква(ws As Worksheet, колонка As Long) As String
    Буква = Split(ws.Cells(1, колонка).Address(True, False), "$")(0)
End Function


Private Function Формат(значение As Variant) As String
    If IsDate(значение) Then
        Формат = Format$(значение, "dd.mm.yyyy")
    Else
        Формат = Trim$(CStr(значение))
    End If
End Function
