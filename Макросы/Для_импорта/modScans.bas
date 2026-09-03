Attribute VB_Name = "modScans"
Option Explicit

' =======================================================================
'  Сканы свидетельств и удостоверений — в пакет, а не мимо него.
'
'  Раньше это была единственная неустранимая ручная операция: сканы
'  печатались отдельно и вкладывались руками, поэтому в сквозную
'  нумерацию пакета они не попадали.
'
'  Устранимая её часть — картинки. Скан в JPG, PNG или TIFF Word
'  вставляет сам, а значит, страницы со сканами становятся обычными
'  страницами свода: и нумерация сквозная, и на бирке всё сходится.
'  Неустранимая часть остаётся ровно одна — сканы в PDF: объединить их
'  без Acrobat нельзя, и про них программа честно говорит отдельно.
'
'  Картинки кладутся в таблицу с жёстко заданной высотой строк, а не
'  абзацами подряд: тогда четвёртый скан не съезжает на следующую
'  страницу из-за того, что второй оказался на сантиметр выше.
' =======================================================================

Private Const wdPageBreakS As Long = 7
Private Const wdRowHeightExactly As Long = 2
Private Const wdPreferredWidthPoints As Long = 3
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdCellAlignVerticalCenter As Long = 1
Private Const wdLineStyleNone As Long = 0

Private Const РАСШИРЕНИЯ As String = "*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.bmp;*.emf;*.wmf"


' Папка сканов по разделу: Поверка, Удостоверения, ЛНК.
Public Function ПапкаСканов(раздел As String) As String
    Dim база As String
    база = Настройка("Папка сканов")
    If база = "" Then Exit Function
    If Right$(база, 1) <> Application.PathSeparator Then база = база & Application.PathSeparator
    ПапкаСканов = база & раздел & Application.PathSeparator
End Function


' Сканы, в имени которых встречается признак: заводской номер прибора,
' фамилия специалиста. Точное имя файла требовать нельзя — сканируют
' люди и называют файлы по-своему.
Public Function НайтиСканы(раздел As String, признак As String) As Collection
    Dim найдено As New Collection
    Dim папка As String
    Dim маски As Variant
    Dim i As Long
    Dim имя As String

    Set НайтиСканы = найдено
    папка = ПапкаСканов(раздел)
    If папка = "" Then Exit Function
    If Not ПапкаСуществует(папка) Then Exit Function

    маски = Split(РАСШИРЕНИЯ, ";")
    For i = LBound(маски) To UBound(маски)
        имя = ПервыйФайл(папка & CStr(маски(i)))
        Do While имя <> ""
            If признак = "" Then
                найдено.Add папка & имя
            ElseIf InStr(1, имя, признак, vbTextCompare) > 0 Then
                найдено.Add папка & имя
            End If
            имя = СледующийФайл()
        Loop
    Next i
End Function


' -- Сборка документа со сканами ---------------------------------------

' Один раздел — со своей страницы. Возвращает число добавленных
' страниц, чтобы вызывающий знал, на сколько вырос пакет.
Public Function ДобавитьРазделСканов(док As Object, заголовок As String, _
                                     файлы As Collection, наСтранице As Long) As Long
    Dim i As Long
    Dim таблица As Object
    Dim колонок As Long
    Dim вРяду As Long
    Dim добавлено As Long

    If файлы Is Nothing Then Exit Function
    If файлы.Count = 0 Then Exit Function

    If наСтранице >= 4 Then
        колонок = 2
    Else
        колонок = 1
    End If
    вРяду = наСтранице \ колонок
    If вРяду < 1 Then вРяду = 1

    i = 1
    Do While i <= файлы.Count
        Если_НеПервая док, добавлено
        ЗаголовокРаздела док, заголовок, добавлено + 1, _
                          ((файлы.Count - 1) \ наСтранице) + 1
        Set таблица = СеткаНаСтранице(док, вРяду, колонок)
        ЗаполнитьСетку таблица, файлы, i, наСтранице
        i = i + наСтранице
        добавлено = добавлено + 1
    Loop
    ДобавитьРазделСканов = добавлено
End Function


Private Sub Если_НеПервая(док As Object, добавлено As Long)
    If добавлено = 0 And док.Content.End <= 2 Then Exit Sub
    док.Content.InsertParagraphAfter
    док.Paragraphs(док.Paragraphs.Count).Range.InsertBreak Type:=wdPageBreakS
End Sub


Private Sub ЗаголовокРаздела(док As Object, текст As String, _
                             лист As Long, всего As Long)
    Dim p As Object
    док.Content.InsertParagraphAfter
    Set p = док.Paragraphs(док.Paragraphs.Count)
    p.Range.Text = текст & IIf(всего > 1, "  (" & лист & " из " & всего & ")", "")
    p.Range.Font.Bold = True
    p.Range.Font.Size = 11
    p.Alignment = wdAlignParagraphCenter
End Sub


' Сетка фиксированного размера: высота строк задана точно, поэтому
' содержимое не может её растянуть и утащить последнюю картинку на
' следующую страницу.
Private Function СеткаНаСтранице(док As Object, строк As Long, колонок As Long) As Object
    Dim таблица As Object
    Dim rng As Object
    Dim ширинаСтр As Single, высотаСтр As Single

    док.Content.InsertParagraphAfter
    Set rng = док.Paragraphs(док.Paragraphs.Count).Range
    Set таблица = док.Tables.Add(Range:=rng, NumRows:=строк, NumColumns:=колонок)

    With док.Sections(док.Sections.Count).PageSetup
        ширинаСтр = .PageWidth - .LeftMargin - .RightMargin
        высотаСтр = .PageHeight - .TopMargin - .BottomMargin - 34   ' минус заголовок
    End With

    таблица.PreferredWidthType = wdPreferredWidthPoints
    таблица.PreferredWidth = ширинаСтр
    таблица.Rows.HeightRule = wdRowHeightExactly
    таблица.Rows.Height = высотаСтр / строк
    таблица.Borders.InsideLineStyle = wdLineStyleNone
    таблица.Borders.OutsideLineStyle = wdLineStyleNone
    таблица.Rows.Alignment = 1                              ' по центру страницы
    Set СеткаНаСтранице = таблица
End Function


Private Sub ЗаполнитьСетку(таблица As Object, файлы As Collection, _
                           первый As Long, сколько As Long)
    Dim n As Long
    Dim индекс As Long
    Dim ячейка As Object
    Dim картинка As Object
    Dim рамкаШ As Single, рамкаВ As Single

    For n = 0 To сколько - 1
        индекс = первый + n
        If индекс > файлы.Count Then Exit For
        Set ячейка = ЯчейкаПоНомеру(таблица, n)
        If ячейка Is Nothing Then Exit For

        ячейка.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        ячейка.VerticalAlignment = wdCellAlignVerticalCenter
        рамкаШ = ячейка.Width - 8
        рамкаВ = таблица.Rows.Height - 8

        Set картинка = ячейка.Range.InlineShapes.AddPicture( _
                          fileName:=CStr(файлы(индекс)), _
                          LinkToFile:=False, SaveWithDocument:=True)
        ВписатьВРамку картинка, рамкаШ, рамкаВ
    Next n
End Sub


Private Function ЯчейкаПоНомеру(таблица As Object, n As Long) As Object
    Dim r As Long, c As Long
    Dim колонок As Long
    колонок = таблица.Columns.Count
    r = n \ колонок + 1
    c = n Mod колонок + 1
    On Error Resume Next
    Set ЯчейкаПоНомеру = таблица.Cell(r, c)
End Function


' Вписать картинку в рамку, сохранив пропорции. Растянуть по рамке
' было бы проще, но скан свидетельства с искажёнными пропорциями
' выглядит подделкой — и читается хуже.
Public Sub ВписатьВРамку(картинка As Object, рамкаШ As Single, рамкаВ As Single)
    Dim к As Single
    Dim кШ As Single, кВ As Single

    If картинка.Width <= 0 Or картинка.Height <= 0 Then Exit Sub
    картинка.LockAspectRatio = 0                 ' msoFalse: ставим оба размера сами
    кШ = рамкаШ / картинка.Width
    кВ = рамкаВ / картинка.Height
    к = кШ
    If кВ < к Then к = кВ
    картинка.Width = картинка.Width * к
    картинка.Height = картинка.Height * к
End Sub


' -- Что печатать по конкретному акту ----------------------------------

' Заводские номера приборов, применённых по всем записям акта.
Public Function ПризнакиПриборов(строки As Collection) As Collection
    Dim ж As Worksheet
    Dim пр As Worksheet
    Dim si As Worksheet
    Dim результат As New Collection
    Dim i As Long, k As Long, last As Long
    Dim идЗаписи As String, идСИ As String, зав As String
    Dim строкаСИ As Long

    Set ПризнакиПриборов = результат
    Set ж = ЛистТихо(Л_ЖУРНАЛ)
    Set пр = ЛистТихо("Применённые_СИ")
    Set si = ЛистТихо(Л_СИ)
    If ж Is Nothing Or пр Is Nothing Or si Is Nothing Then Exit Function

    last = ПоследняяСтрока(пр, 1)
    For k = 1 To строки.Count
        идЗаписи = Trim$(CStr(ж.Cells(CLng(строки(k)), ColReq(ж, "ID записи")).Value))
        If идЗаписи <> "" Then
            For i = ПЕРВАЯ_СТРОКА To last
                If StrComp(Trim$(CStr(пр.Cells(i, ColReq(пр, "ID записи журнала")).Value)), _
                           идЗаписи, vbTextCompare) = 0 Then
                    идСИ = Trim$(CStr(пр.Cells(i, ColReq(пр, "ID средства НК")).Value))
                    строкаСИ = СтрокаСИ(si, идСИ)
                    If строкаСИ > 0 Then
                        зав = Trim$(CStr(si.Cells(строкаСИ, ColReq(si, "Заводской №")).Value))
                        If зав <> "" Then ДобавитьРаз результат, зав
                    End If
                End If
            Next i
        End If
    Next k
End Function


' Фамилии исполнителей по всем записям акта.
Public Function ПризнакиИсполнителей(строки As Collection) As Collection
    Dim ж As Worksheet
    Dim результат As New Collection
    Dim методы As Variant
    Dim k As Long, i As Long
    Dim r As Long
    Dim фио As String
    Dim c As Long

    Set ПризнакиИсполнителей = результат
    Set ж = ЛистТихо(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Function

    методы = Split(МЕТОДЫ, ",")
    For k = 1 To строки.Count
        r = CLng(строки(k))
        For i = LBound(методы) To UBound(методы)
            c = ColByHeader(ж, CStr(методы(i)) & " исполнитель")
            If c > 0 Then
                фио = Trim$(CStr(ж.Cells(r, c).Value))
                If фио <> "" Then ДобавитьРаз результат, Фамилия(фио)
            End If
        Next i
    Next k
End Function


Private Function Фамилия(фио As String) As String
    Dim p As Long
    p = InStr(Trim$(фио), " ")
    If p > 0 Then
        Фамилия = Left$(Trim$(фио), p - 1)
    Else
        Фамилия = Trim$(фио)
    End If
End Function


Private Sub ДобавитьРаз(куда As Collection, что As String)
    Dim i As Long
    For i = 1 To куда.Count
        If StrComp(CStr(куда(i)), что, vbTextCompare) = 0 Then Exit Sub
    Next i
    куда.Add что
End Sub


Private Function СтрокаСИ(si As Worksheet, идСИ As String) As Long
    Dim r As Long, last As Long
    If идСИ = "" Then Exit Function
    last = ПоследняяСтрока(si, 1)
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(si.Cells(r, 1).Value)), идСИ, vbTextCompare) = 0 Then
            СтрокаСИ = r
            Exit Function
        End If
    Next r
End Function


' Все сканы пакета в требуемом порядке: свидетельства лаборатории,
' удостоверения исполнителей, свидетельства о поверке приборов.
Public Function СобратьСканыПакета(wdApp As Object, строки As Collection, _
                                   ByRef страниц As Long) As String
    Dim док As Object
    Dim наСтранице As Long
    Dim признаки As Collection
    Dim файлы As Collection
    Dim путь As String

    страниц = 0
    наСтранице = CLng(Val(Настройка("Сканов на страницу")))
    If наСтранице <= 0 Then наСтранице = 4

    Set док = wdApp.Documents.Add
    страниц = страниц + ДобавитьРазделСканов(док, "Свидетельства лаборатории", _
                            НайтиСканы("ЛНК", ""), наСтранице)

    Set признаки = ПризнакиИсполнителей(строки)
    Set файлы = СканыПоПризнакам("Удостоверения", признаки)
    страниц = страниц + ДобавитьРазделСканов(док, _
                            "Квалификационные удостоверения исполнителей", файлы, наСтранице)

    Set признаки = ПризнакиПриборов(строки)
    Set файлы = СканыПоПризнакам("Поверка", признаки)
    страниц = страниц + ДобавитьРазделСканов(док, _
                            "Свидетельства о поверке применённых средств измерений", _
                            файлы, наСтранице)

    If страниц = 0 Then
        док.Close SaveChanges:=False
        Exit Function
    End If

    путь = ПапкаГотовых() & "Сканы_пакета_" & Format$(Now, "yyyymmdd_hhnnss") & ".docx"
    док.SaveAs2 fileName:=путь
    док.Close SaveChanges:=False
    СобратьСканыПакета = путь
End Function


Private Function СканыПоПризнакам(раздел As String, признаки As Collection) As Collection
    Dim результат As New Collection
    Dim i As Long, k As Long
    Dim найденные As Collection

    Set СканыПоПризнакам = результат
    If признаки Is Nothing Then Exit Function
    For i = 1 To признаки.Count
        Set найденные = НайтиСканы(раздел, CStr(признаки(i)))
        For k = 1 To найденные.Count
            ДобавитьРаз результат, CStr(найденные(k))
        Next k
    Next i
End Function


' Отдельная печать: свидетельства о поверке по акту, без пакета.
Public Sub НапечататьСвидетельстваПоверки()
    ПечатьРаздела "Поверка", "Свидетельства о поверке применённых средств измерений"
End Sub


Public Sub НапечататьУдостоверения()
    ПечатьРаздела "Удостоверения", "Квалификационные удостоверения исполнителей"
End Sub


Private Sub ПечатьРаздела(раздел As String, заголовок As String)
    Dim ж As Worksheet
    Dim r As Long
    Dim строки As Collection
    Dim признаки As Collection
    Dim файлы As Collection
    Dim wdApp As Object
    Dim док As Object
    Dim наСтранице As Long
    Dim страниц As Long

    Set ж = Лист(Л_ЖУРНАЛ)
    If ж Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ж)
    If r = 0 Then Exit Sub

    If ПапкаСканов(раздел) = "" Then
        MsgBox "Не задана папка сканов в настройках.", vbExclamation, заголовок
        Exit Sub
    End If
    If Not ПапкаСуществует(ПапкаСканов(раздел)) Then
        MsgBox "Нет папки " & ПапкаСканов(раздел) & "." & vbCrLf & vbCrLf & _
               "Создайте её и положите туда сканы. В имени файла должен быть " & _
               "заводской номер прибора или фамилия специалиста — по ним и ищется.", _
               vbExclamation, заголовок
        Exit Sub
    End If

    Set строки = modPackage.СтрокиАкта(ж, r, modJournal.НомерАкта(ж, r))
    If StrComp(раздел, "Поверка", vbTextCompare) = 0 Then
        Set признаки = ПризнакиПриборов(строки)
    Else
        Set признаки = ПризнакиИсполнителей(строки)
    End If
    Set файлы = СканыПоПризнакам(раздел, признаки)

    If файлы.Count = 0 Then
        MsgBox "Ни одного скана не нашлось." & vbCrLf & vbCrLf & _
               "Искали в " & ПапкаСканов(раздел) & " по " & признаки.Count & _
               " признакам. Проверьте, что в именах файлов есть заводской номер " & _
               "прибора или фамилия специалиста, и что это картинки, а не PDF: " & _
               "PDF Word вставить не может.", vbExclamation, заголовок
        Exit Sub
    End If

    наСтранице = CLng(Val(Настройка("Сканов на страницу")))
    If наСтранице <= 0 Then наСтранице = 4

    On Error GoTo Сбой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set док = wdApp.Documents.Add
    страниц = ДобавитьРазделСканов(док, заголовок, файлы, наСтранице)
    wdApp.Visible = True
    док.Activate
    MsgBox "Собрано сканов: " & файлы.Count & " на " & страниц & " стр." & vbCrLf & vbCrLf & _
           "Документ открыт в Word — проверьте и печатайте. В пакет по акту эти же " & _
           "сканы попадают сами, отдельно печатать их не нужно.", _
           vbInformation, заголовок
    Exit Sub

Сбой:
    On Error Resume Next
    If Not wdApp Is Nothing Then wdApp.Visible = True
    MsgBox "Не удалось собрать сканы." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, заголовок
End Sub
