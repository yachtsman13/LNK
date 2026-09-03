Attribute VB_Name = "modScans"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Сканы свидетельств и удостоверений — в пакет, а не мимо него.
'
'  Раньше это была единственная неустранимая ручная операция: сканы
'  печатались отдельно и вкладывались руками, поэтому в сквозную
'  нумерацию пакета они не попадали.
'
'  Часть решена полностью: скан в JPG, PNG или TIFF Word вставляет сам,
'  и такие страницы становятся обычными страницами свода — нумерация
'  сквозная, на бирке всё сходится.
'
'  Часть сканов приходит в PDF — так сложилось у разных исполнителей,
'  переучивать никого не будем. Раз в лаборатории есть Acrobat (не
'  Reader), эти файлы можно склеить программно: PDDoc.InsertPages —
'  ровно то, для чего этот вызов существует, и самая обкатанная часть
'  автоматизации Acrobat. Результат — один файл вместо десятка. Но
'  вписать в этот уже готовый PDF номер по правилам «пакета» (лист, а
'  не страница, без номера на первом листе) программно нельзя без риска
'  сломать что-то на машине пользователя: это делается через JavaScript
'  Acrobat, а такие вызовы через автоматизацию Acrobat не всегда
'  разрешены его настройками безопасности и версией. Поэтому объединённый
'  PDF печатается и подшивается отдельно, без сквозного номера, — как
'  прежде подшивались сканы по одному, только теперь одним файлом,
'  а не россыпью.
'
'  Картинки кладутся в таблицу с жёстко заданной высотой строк, а не
'  абзацами подряд: тогда четвёртый скан не съезжает на следующую
'  страницу из-за того, что второй оказался на сантиметр выше.
' ═══════════════════════════════════════════════════════════════════════

Private Const wdPageBreakS As Long = 7
Private Const wdRowHeightExactly As Long = 2
Private Const wdPreferredWidthPoints As Long = 3
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdCellAlignVerticalCenter As Long = 1
Private Const wdLineStyleNone As Long = 0

Private Const РАСШИРЕНИЯ As String = "*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.bmp;*.emf;*.wmf"
Private Const МАСКА_PDF As String = "*.pdf"


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


' ── Сборка документа со сканами ───────────────────────────────────────

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


' ── Что печатать по конкретному акту ──────────────────────────────────

' Заводские номера приборов, применённых по всем записям акта.
Public Function ПризнакиПриборов(строки As Collection) As Collection
    Dim ж As Worksheet
    Dim пр As Worksheet
    Dim si As Worksheet
    Dim результат As New Collection
    Dim i As Long, k As Long, last As Long
    Dim идЗаписи As String, идСИ As String, зав As String
    Dim rСИ As Long

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
                    rСИ = СтрокаСИ(si, идСИ)
                    If rСИ > 0 Then
                        зав = Trim$(CStr(si.Cells(rСИ, ColReq(si, "Заводской №")).Value))
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


' Все сканы-картинки пакета в требуемом порядке: свидетельства
' лаборатории, удостоверения исполнителей, свидетельства о поверке
' приборов. PDF-сканы сюда не входят — Word не умеет вставлять их как
' картинку; для них отдельная сборка через Acrobat, см. ниже.
'
' путьPDF, страницPDF (ByRef) — файл, в который объединились PDF-сканы по
' тому же акту, тем же трём разделам и в том же порядке, и число страниц
' в нём. путьPDF пуст, если PDF-сканов не нашлось ИЛИ Acrobat не смог их
' объединить — эти два случая различает ошибкаPDF: пусто в первом,
' текст причины во втором.
Public Function СобратьСканыПакета(wdApp As Object, строки As Collection, _
                                   ByRef страниц As Long, ByRef путьPDF As String, _
                                   ByRef страницPDF As Long, ByRef ошибкаPDF As String) As String
    Dim док As Object
    Dim наСтранице As Long
    Dim признаки As Collection
    Dim файлы As Collection
    Dim pdfВсе As New Collection
    Dim путь As String

    страниц = 0
    путьPDF = ""
    наСтранице = CLng(Val(Настройка("Сканов на страницу")))
    If наСтранице <= 0 Then наСтранице = 4

    Set док = wdApp.Documents.Add
    страниц = страниц + ДобавитьРазделСканов(док, "Свидетельства лаборатории", _
                            НайтиСканы("ЛНК", ""), наСтранице)
    ДобавитьВсе pdfВсе, НайтиPDF("ЛНК", "")

    Set признаки = ПризнакиИсполнителей(строки)
    Set файлы = СканыПоПризнакам("Удостоверения", признаки)
    страниц = страниц + ДобавитьРазделСканов(док, _
                            "Квалификационные удостоверения исполнителей", файлы, наСтранице)
    ДобавитьВсе pdfВсе, PDFПоПризнакам("Удостоверения", признаки)

    Set признаки = ПризнакиПриборов(строки)
    Set файлы = СканыПоПризнакам("Поверка", признаки)
    страниц = страниц + ДобавитьРазделСканов(док, _
                            "Свидетельства о поверке применённых средств измерений", _
                            файлы, наСтранице)
    ДобавитьВсе pdfВсе, PDFПоПризнакам("Поверка", признаки)

    If страниц = 0 Then
        док.Close SaveChanges:=False
    Else
        путь = ПапкаГотовых() & "Сканы_пакета_" & Format$(Now, "yyyymmdd_hhnnss") & ".docx"
        док.SaveAs2 fileName:=путь
        док.Close SaveChanges:=False
        СобратьСканыПакета = путь
    End If

    страницPDF = 0
    путьPDF = ""
    ошибкаPDF = ""
    If pdfВсе.Count > 0 Then
        If Not ОбъединитьPDF(pdfВсе, _
                    ПапкаГотовых() & "Сканы_PDF_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf", _
                    путьPDF, страницPDF, ошибкаPDF) Then
            путьPDF = ""
        End If
    End If
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


' ── PDF-сканы: поиск и объединение через Acrobat ──────────────────────

' То же самое, что НайтиСканы, только маска одна — *.pdf. Отдельная
' функция, а не общий список масок: Word и Acrobat читают файлы
' по-разному, и ниже по коду важно не перепутать, какой список идёт
' куда — в картинки или в объединение PDF.
Public Function НайтиPDF(раздел As String, признак As String) As Collection
    Dim найдено As New Collection
    Dim папка As String
    Dim имя As String

    Set НайтиPDF = найдено
    папка = ПапкаСканов(раздел)
    If папка = "" Then Exit Function
    If Not ПапкаСуществует(папка) Then Exit Function

    имя = ПервыйФайл(папка & МАСКА_PDF)
    Do While имя <> ""
        If признак = "" Then
            найдено.Add папка & имя
        ElseIf InStr(1, имя, признак, vbTextCompare) > 0 Then
            найдено.Add папка & имя
        End If
        имя = СледующийФайл()
    Loop
End Function


Private Function PDFПоПризнакам(раздел As String, признаки As Collection) As Collection
    Dim результат As New Collection
    Dim i As Long, k As Long
    Dim найденные As Collection

    Set PDFПоПризнакам = результат
    If признаки Is Nothing Then Exit Function
    For i = 1 To признаки.Count
        Set найденные = НайтиPDF(раздел, CStr(признаки(i)))
        For k = 1 To найденные.Count
            ДобавитьРаз результат, CStr(найденные(k))
        Next k
    Next i
End Function


Private Sub ДобавитьВсе(куда As Collection, откуда As Collection)
    Dim i As Long
    For i = 1 To откуда.Count
        ДобавитьРаз куда, CStr(откуда(i))
    Next i
End Sub


' Склеивает несколько PDF в один через Acrobat (не Reader — тот
' автоматизацию не даёт). PDDoc.InsertPages — старейшая и самая
' обкатанная часть автоматизации Acrobat, десятилетиями используемая
' именно для этой задачи, поэтому здесь ей и место.
'
' Файлы, которые Acrobat не смог открыть (повреждён скан, не PDF,
' несмотря на расширение) — пропускаются, а не обрывают всю сборку:
' один плохой файл не должен стоить всех остальных; пропуски видны в
' ошибка (не пусто) даже при True, чтобы вызывающий мог показать это
' человеку, не считая сборку полным провалом.
'
' Возвращает True при успехе (путь и страницы заполнены), False при
' полном провале (Acrobat не найден или ни один файл не открылся) —
' тогда заполнена только ошибка, а путь и страницы остаются пустыми.
Public Function ОбъединитьPDF(файлы As Collection, выходной As String, _
                              ByRef путь As String, ByRef страниц As Long, _
                              ByRef ошибка As String) As Boolean
    Const PDSaveFull As Long = 1
    Dim базовый As Object
    Dim вставка As Object
    Dim i As Long
    Dim открыт As Boolean
    Dim пропущено As Long

    путь = "": страниц = 0: ошибка = ""
    If файлы Is Nothing Then Exit Function
    If файлы.Count = 0 Then Exit Function

    On Error GoTo НетAcrobat
    Set базовый = CreateObject("AcroExch.PDDoc")
    On Error GoTo Сбой

    ' Первый файл, который реально открылся, становится основой —
    ' не обязательно самый первый в списке, если он вдруг повреждён.
    For i = 1 To файлы.Count
        If базовый.Open(CStr(файлы(i))) Then
            открыт = True
            Exit For
        Else
            пропущено = пропущено + 1
        End If
    Next i
    If Not открыт Then
        ошибка = "ни один из " & файлы.Count & " файлов не открылся"
        Exit Function
    End If

    For i = i + 1 To файлы.Count
        Set вставка = CreateObject("AcroExch.PDDoc")
        If вставка.Open(CStr(файлы(i))) Then
            базовый.InsertPages базовый.GetNumPages() - 1, вставка, 0, _
                                вставка.GetNumPages(), 0
            вставка.Close
        Else
            пропущено = пропущено + 1
        End If
        Set вставка = Nothing
    Next i

    страниц = базовый.GetNumPages()
    базовый.Save PDSaveFull, выходной
    базовый.Close
    Set базовый = Nothing

    путь = выходной
    If пропущено > 0 Then ошибка = "не открылось файлов: " & пропущено & " — пропущены"
    ОбъединитьPDF = True
    Exit Function

НетAcrobat:
    ошибка = "не нашёл Acrobat (AcroExch.PDDoc) — он должен быть установлен, " & _
             "не только Reader"
    Exit Function

Сбой:
    On Error Resume Next
    If Not базовый Is Nothing Then базовый.Close
    If Not вставка Is Nothing Then вставка.Close
    ошибка = "Acrobat не справился: " & Err.Description
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
    Dim pdfФайлы As Collection
    Dim wdApp As Object
    Dim док As Object
    Dim наСтранице As Long
    Dim страниц As Long
    Dim путьPDF As String
    Dim страницPDF As Long
    Dim ошибкаPDF As String

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
    Set pdfФайлы = PDFПоПризнакам(раздел, признаки)

    If файлы.Count = 0 And pdfФайлы.Count = 0 Then
        MsgBox "Ни одного скана не нашлось." & vbCrLf & vbCrLf & _
               "Искали в " & ПапкаСканов(раздел) & " по " & признаки.Count & _
               " признакам (JPG/PNG/TIFF и PDF). Проверьте, что в имени файла есть " & _
               "заводской номер прибора или фамилия специалиста.", vbExclamation, заголовок
        Exit Sub
    End If

    наСтранице = CLng(Val(Настройка("Сканов на страницу")))
    If наСтранице <= 0 Then наСтранице = 4

    If pdfФайлы.Count > 0 Then
        путьPDF = ПапкаГотовых() & "Сканы_PDF_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
        If Not ОбъединитьPDF(pdfФайлы, путьPDF, путьPDF, страницPDF, ошибкаPDF) Then
            путьPDF = ""
        End If
    End If

    If файлы.Count = 0 Then
        ' Только PDF — Word открывать незачем, объединённого файла хватит.
        ПоказатьРезультатПечати заголовок, 0, 0, путьPDF, страницPDF, ошибкаPDF
        Exit Sub
    End If

    On Error GoTo Сбой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set док = wdApp.Documents.Add
    страниц = ДобавитьРазделСканов(док, заголовок, файлы, наСтранице)
    wdApp.Visible = True
    док.Activate
    ПоказатьРезультатПечати заголовок, файлы.Count, страниц, путьPDF, страницPDF, ошибкаPDF
    Exit Sub

Сбой:
    On Error Resume Next
    If Not wdApp Is Nothing Then wdApp.Visible = True
    MsgBox "Не удалось собрать сканы." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, заголовок
End Sub


Private Sub ПоказатьРезультатПечати(заголовок As String, картинок As Long, страницКартинок As Long, _
                                    путьPDF As String, страницPDF As Long, ошибкаPDF As String)
    Dim текст As String

    If картинок > 0 Then
        текст = "Картинок: " & картинок & ", на " & страницКартинок & " стр. " & _
                "Документ открыт в Word — проверьте и печатайте."
    End If
    If путьPDF <> "" Then
        текст = текст & IIf(текст = "", "", vbCrLf & vbCrLf) & _
                "PDF объединены Acrobat, " & страницPDF & " стр.: " & путьPDF
    ElseIf ошибкаPDF <> "" Then
        текст = текст & IIf(текст = "", "", vbCrLf & vbCrLf) & _
                "PDF-сканы найдены, но не объединились: " & ошибкаPDF & _
                ". Печатайте их по одному."
    End If
    текст = текст & vbCrLf & vbCrLf & _
            "В пакет по акту эти же сканы попадают сами, отдельно печатать их не нужно " & _
            "— эта кнопка для случая, когда пакет ещё рано собирать целиком."
    MsgBox текст, vbInformation, заголовок
End Sub
