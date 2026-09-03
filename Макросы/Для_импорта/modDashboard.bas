Attribute VB_Name = "modDashboard"
Option Explicit

' =======================================================================
'  Плитки пульта: что требует внимания сегодня.
'
'  Правило одно — плитка показывается, только если по ней есть что
'  делать. Дашборд, на котором всегда всё зелёное, перестают читать
'  через неделю.
'
'  Своего листа у плиток больше нет: они живут в верхней части пульта,
'  над кнопками, — чтобы увиденное и действие по нему были на одном
'  экране.
' =======================================================================

Private Const ЦВ_ОК As Long = 12379352        ' спокойный зелёный
Private Const ЦВ_ВНИМАНИЕ As Long = 10284031  ' жёлтый
Private Const ЦВ_СРОЧНО As Long = 13551615    ' красный
Private Const ЦВ_ТЕКСТ_ОК As Long = 25600
Private Const ЦВ_ТЕКСТ_ВН As Long = 26012
Private Const ЦВ_ТЕКСТ_СР As Long = 2500


' Плитки в столбцах B, C, D начиная со строки r. Возвращает первую
' свободную строку — под ней пульт рисует кнопки.
'
' Дашборд больше не отдельный лист: смотреть на одном экране, что
' горит, и тут же нажимать нужную кнопку — это одно действие, а не
' два, и переключаться между листами ради него незачем.
Public Function ПлиткиНа(ws As Worksheet, ByVal r As Long) As Long
    Dim начало As Long

    начало = r
    ' Каждая плитка со своей защитой: не собралась одна — остальные
    ' всё равно показываются. Наполовину собранный пульт полезнее,
    ' чем сообщение об ошибке вместо него.
    On Error Resume Next
    r = ПлиткаАттестацияЛНК(ws, r)
    r = ПлиткаСрокиПоЛисту(ws, r, Л_АТТЕСТАЦИИ, "Действительно до", "Удостоверения специалистов", _
                           "ФИО", "Метод контроля", 90)
    r = ПлиткаСрокиПоЛисту(ws, r, Л_СИ, "Действительно до", "Поверка средств измерений", _
                           "Наименование", "Заводской №", 90)
    r = ПлиткаСрокиПоЛисту(ws, r, "НД", "Дата следующей проверки", "Проверка актуальности НД", _
                           "Наименование документа", "Обозначение", 60)
    r = ПлиткаПлана(ws, r)
    r = ПлиткаЧерновиков(ws, r)
    r = ПлиткаВыданного(ws, r)
    r = ПлиткаРекламаций(ws, r)
    r = ПлиткаОбмена(ws, r)
    r = ПлиткаНосителя(ws, r)
    r = ПлиткаЭлектроинструмента(ws, r)
    If Err.Number <> 0 Then
        r = Плитка(ws, r, "Часть плиток не собралась", "сбой", _
                   "Ошибка " & Err.Number & ": " & Err.Description & _
                   ". Остальное на пульте верно.", 1)
        Err.Clear
    End If
    On Error GoTo 0

    If r = начало Then
        ws.Cells(r, 2).Value = "Просроченного и ближайшего нет — всё в сроках."
        ws.Cells(r, 2).Font.Color = RGB(0, 97, 0)
        ws.Cells(r, 2).Font.Size = 12
        r = r + 1
    End If
    ПлиткиНа = r
End Function


' Выполнение плана-графика работ в текущем месяце и работы сверх плана.
Private Function ПлиткаПлана(ws As Worksheet, r As Long) As Long
    Dim пл As Worksheet
    Dim ж As Worksheet
    Dim i As Long, last As Long
    Dim cГод As Long, cМес As Long, cСтатус As Long
    Dim всего As Long, выполнено As Long, вне As Long
    Dim cДата As Long, cВид As Long
    Dim д As Date

    ПлиткаПлана = r
    Set пл = ЛистТихо("План_график")
    Set ж = ЛистТихо(Л_ЖУРНАЛ)
    If пл Is Nothing Or ж Is Nothing Then Exit Function

    cГод = ColByHeader(пл, "Год")
    cМес = ColByHeader(пл, "Месяц")
    cСтатус = ColByHeader(пл, "Статус")
    If cГод = 0 Or cМес = 0 Or cСтатус = 0 Then Exit Function
    last = ПоследняяСтрока(пл, 1)
    For i = ПЕРВАЯ_СТРОКА To last
        If CLng(Val(пл.Cells(i, cГод).Value)) = Year(Date) Then
            If МесяцСовпал(CStr(пл.Cells(i, cМес).Value), Month(Date)) Then
                всего = всего + 1
                If StrComp(CStr(пл.Cells(i, cСтатус).Value), "Выполнено", vbTextCompare) = 0 Then
                    выполнено = выполнено + 1
                End If
            End If
        End If
    Next i

    cДата = ColByHeader(ж, "Дата контроля")
    cВид = ColByHeader(ж, "Вид работ")
    If cДата > 0 And cВид > 0 Then
        last = ПоследняяСтрока(ж, cДата)
        For i = ПЕРВАЯ_СТРОКА To last
            If IsDate(ж.Cells(i, cДата).Value) Then
                д = CDate(ж.Cells(i, cДата).Value)
                If Year(д) = Year(Date) And Month(д) = Month(Date) Then
                    If StrComp(Trim$(CStr(ж.Cells(i, cВид).Value)), "По плану", vbTextCompare) <> 0 Then
                        вне = вне + 1
                    End If
                End If
            End If
        Next i
    End If

    If всего = 0 And вне = 0 Then Exit Function
    ПлиткаПлана = Плитка(ws, r, "План работ на " & Format$(Date, "mmmm"), _
        выполнено & " из " & всего, _
        "По плану выполнено " & выполнено & " из " & всего & _
        ", осталось " & (всего - выполнено) & ". Внеплановых работ за месяц: " & вне & ". " & _
        "Плановой работа становится, когда её отметили в плане-графике.", _
        IIf(всего > 0 And выполнено < всего, 1, 0))
End Function


Private Function МесяцСовпал(текст As String, месяц As Long) As Boolean
    Dim s As String
    s = Trim$(текст)
    If s = "" Then Exit Function
    If IsNumeric(s) Then
        МесяцСовпал = (CLng(Val(s)) = месяц)
        Exit Function
    End If
    МесяцСовпал = (StrComp(Left$(s, 3), Left$(Format$(DateSerial(2000, месяц, 1), "mmmm"), 3), _
                           vbTextCompare) = 0)
End Function


' -- Плитки ------------------------------------------------------------

' Общая плитка «сроки по колонке с датой»: считает просроченное и
' ближайшее, перечисляет до пяти позиций поимённо.
Private Function ПлиткаСрокиПоЛисту(ws As Worksheet, r As Long, имяЛиста As String, _
                                    колонкаДаты As String, заголовок As String, _
                                    поле1 As String, поле2 As String, порог As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long
    Dim cДата As Long, c1 As Long, c2 As Long, cСост As Long
    Dim просрочено As Long, скоро As Long
    Dim список As String
    Dim д As Date
    Dim осталось As Long

    ПлиткаСрокиПоЛисту = r
    Set src = ЛистТихо(имяЛиста)
    If src Is Nothing Then Exit Function
    cДата = ColByHeader(src, колонкаДаты)
    If cДата = 0 Then Exit Function
    c1 = ColByHeader(src, поле1)
    c2 = ColByHeader(src, поле2)
    cСост = ColByHeader(src, "Состояние")
    last = ПоследняяСтрока(src, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If cСост = 0 Or StrComp(CStr(src.Cells(i, cСост).Value), "Списан", vbTextCompare) <> 0 Then
            If IsDate(src.Cells(i, cДата).Value) Then
                д = CDate(src.Cells(i, cДата).Value)
                осталось = DateDiff("d", Date, д)
                If осталось < 0 Then
                    просрочено = просрочено + 1
                ElseIf осталось <= порог Then
                    скоро = скоро + 1
                End If
                If осталось <= порог Then
                    If Len(список) < 400 Then
                        список = список & IIf(список = "", "", "; ") & _
                                 CStr(src.Cells(i, c1).Value) & _
                                 IIf(c2 = 0, "", " " & CStr(src.Cells(i, c2).Value)) & _
                                 " — " & IIf(осталось < 0, "просрочено " & -осталось & " дн.", _
                                             "осталось " & осталось & " дн.")
                    End If
                End If
            End If
        End If
    Next i

    If просрочено = 0 And скоро = 0 Then Exit Function
    ПлиткаСрокиПоЛисту = Плитка(ws, r, заголовок, _
        IIf(просрочено > 0, просрочено & " просрочено", скоро & " истекает"), _
        список, IIf(просрочено > 0, 2, 1))
End Function


' Срок аттестации самой лаборатории. Пороги здесь другие — переаттестация
' готовится месяцами, поэтому предупреждаем за полгода.
Private Function ПлиткаАттестацияЛНК(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long
    Dim значение As String
    Dim д As Date
    Dim осталось As Long

    ПлиткаАттестацияЛНК = r
    Set src = ЛистТихо("Паспорт_постоянная")
    If src Is Nothing Then Exit Function
    last = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    For i = ПЕРВАЯ_СТРОКА To last
        If InStr(1, CStr(src.Cells(i, 1).Value), "Свидетельство действительно до", vbTextCompare) > 0 Then
            значение = Trim$(CStr(src.Cells(i, 2).Value))
            Exit For
        End If
    Next i

    If значение = "" Then
        ПлиткаАттестацияЛНК = Плитка(ws, r, "Аттестация лаборатории", "срок не заполнен", _
            "Впишите срок свидетельства на листе «Паспорт_постоянная» — это самый дорогой " & _
            "по последствиям срок во всей системе: с его истечением лаборатория " & _
            "не может работать вообще.", 1)
        Exit Function
    End If
    If Not IsDate(значение) Then Exit Function

    д = CDate(значение)
    осталось = DateDiff("d", Date, д)
    If осталось > 180 Then Exit Function
    ПлиткаАттестацияЛНК = Плитка(ws, r, "Аттестация лаборатории", _
        IIf(осталось < 0, "ПРОСРОЧЕНА", осталось & " дн."), _
        "Свидетельство действительно до " & Format$(д, "dd.mm.yyyy") & _
        ". Переаттестация готовится месяцами — начинать пора.", _
        IIf(осталось <= 90, 2, 1))
End Function


Private Function ПлиткаЧерновиков(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long, n As Long
    Dim cСтатус As Long, cДата As Long
    Dim старейший As Date

    ПлиткаЧерновиков = r
    Set src = ЛистТихо(Л_ЖУРНАЛ)
    If src Is Nothing Then Exit Function
    cСтатус = ColByHeader(src, "Статус")
    cДата = ColByHeader(src, "Дата контроля")
    If cСтатус = 0 Then Exit Function
    last = ПоследняяСтрока(src, cДата)

    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(CStr(src.Cells(i, cСтатус).Value), "Черновик", vbTextCompare) = 0 Then
            n = n + 1
            If IsDate(src.Cells(i, cДата).Value) Then
                If старейший = 0 Or CDate(src.Cells(i, cДата).Value) < старейший Then
                    старейший = CDate(src.Cells(i, cДата).Value)
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    ПлиткаЧерновиков = Плитка(ws, r, "Черновики в журнале", n & " шт.", _
        "Незарегистрированные записи: номера им не выданы, в бумажный журнал они не идут." & _
        IIf(старейший = 0, "", " Самая ранняя от " & Format$(старейший, "dd.mm.yyyy") & "."), _
        IIf(старейший > 0 And DateDiff("d", старейший, Date) > 14, 2, 1))
End Function


Private Function ПлиткаВыданного(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long, n As Long
    Dim cВыдачи As Long, cВозврат As Long, cКому As Long, cНаим As Long
    Dim список As String
    Dim дней As Long

    ПлиткаВыданного = r
    Set src = ЛистТихо("Выдача_оборудования")
    If src Is Nothing Then Exit Function
    cВыдачи = ColByHeader(src, "Дата выдачи")
    cВозврат = ColByHeader(src, "Дата возврата")
    cКому = ColByHeader(src, "Кому выдано")
    cНаим = ColByHeader(src, "Наименование")
    If cВыдачи = 0 Then Exit Function
    last = ПоследняяСтрока(src, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If IsDate(src.Cells(i, cВыдачи).Value) And Not IsDate(src.Cells(i, cВозврат).Value) Then
            n = n + 1
            дней = DateDiff("d", CDate(src.Cells(i, cВыдачи).Value), Date)
            If Len(список) < 300 Then
                список = список & IIf(список = "", "", "; ") & _
                         CStr(src.Cells(i, cНаим).Value) & " — " & _
                         CStr(src.Cells(i, cКому).Value) & ", " & дней & " дн."
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    ПлиткаВыданного = Плитка(ws, r, "Оборудование на руках", n & " шт.", список, 0)
End Function


Private Function ПлиткаРекламаций(ws As Worksheet, r As Long) As Long
    Dim src As Worksheet
    Dim i As Long, last As Long
    Dim cСрок As Long, cОтметка As Long, cРекв As Long
    Dim просрочено As Long, вРаботе As Long
    Dim список As String

    ПлиткаРекламаций = r
    Set src = ЛистТихо("Рекламации")
    If src Is Nothing Then Exit Function
    cСрок = ColByHeader(src, "Срок выполнения")
    cОтметка = ColByHeader(src, "Отметка об выполненных мероприятиях")
    cРекв = ColByHeader(src, "Реквизиты документа")
    If cСрок = 0 Then Exit Function
    last = ПоследняяСтрока(src, 1)

    For i = ПЕРВАЯ_СТРОКА To last
        If Trim$(CStr(src.Cells(i, cОтметка).Value)) = "" Then
            If IsDate(src.Cells(i, cСрок).Value) Then
                вРаботе = вРаботе + 1
                If CDate(src.Cells(i, cСрок).Value) < Date Then
                    просрочено = просрочено + 1
                    If Len(список) < 300 Then
                        список = список & IIf(список = "", "", "; ") & _
                                 CStr(src.Cells(i, cРекв).Value) & " — срок был " & _
                                 Format$(src.Cells(i, cСрок).Value, "dd.mm.yyyy")
                    End If
                End If
            End If
        End If
    Next i

    If вРаботе = 0 Then Exit Function
    ПлиткаРекламаций = Плитка(ws, r, "Рекламации", _
        IIf(просрочено > 0, просрочено & " просрочено", вРаботе & " в работе"), _
        IIf(список = "", "Мероприятия ещё не отмечены выполненными.", список), _
        IIf(просрочено > 0, 2, 1))
End Function


Private Function ПлиткаОбмена(ws As Worksheet, r As Long) As Long
    Dim папка As String
    Dim файл As String
    Dim n As Long

    ПлиткаОбмена = r
    папка = ПапкаОбмена() & "Входящие" & Application.PathSeparator
    If Not ПапкаСуществует(папка) Then Exit Function
    файл = ПервыйФайл(папка & "*.csv")
    Do While файл <> ""
        n = n + 1
        файл = Dir
    Loop
    If n = 0 Then Exit Function
    If Not ЭтоМастер() Then Exit Function
    ПлиткаОбмена = Плитка(ws, r, "Дельты ждут сборки", n & " шт.", _
        "Пока свод не собран, остальные видят данные на момент прошлой сборки.", 1)
End Function


' Электроинструмент проверяется раз в полгода, и просроченным им
' работать нельзя. Пороги короче, чем у поверки приборов: месяц запаса
' при полугодовой периодичности — это уже много.
Private Function ПлиткаЭлектроинструмента(ws As Worksheet, r As Long) As Long
    Dim список As String
    Dim просрочено As Long

    ПлиткаЭлектроинструмента = r
    список = ПроверитьСрокиЭлектроинструмента()
    If список = "" Then Exit Function
    просрочено = СколькоПросроченоЭлектроинструмента()
    ПлиткаЭлектроинструмента = Плитка(ws, r, "Электроинструмент", _
        IIf(просрочено > 0, просрочено & " просрочено", "истекает"), _
        список, IIf(просрочено > 0, 2, 1))
End Function


' Пока сети нет, номера выдаёт флешка. Если её нет под рукой, об этом
' лучше узнать сразу, а не в момент регистрации.
Private Function ПлиткаНосителя(ws As Worksheet, r As Long) As Long
    ПлиткаНосителя = r
    If ПапкаОбмена() <> "" Then Exit Function
    ПлиткаНосителя = Плитка(ws, r, "Общая папка недоступна", "нет", _
        "Ни сетевой путь, ни штатная флешка сейчас не доступны. Работать можно: " & _
        "записи заводятся черновиками, документы формируются. Нельзя только выдавать " & _
        "номера — подключите носитель и нажмите «Зарегистрировать все черновики».", 1)
End Function


' -- Отрисовка ---------------------------------------------------------

' Пояснения бывают от пары слов до семисот символов (список из полусотни
' записей). Раньше строка была фиксированной высоты 34 пункта, и длинный
' текст обрезался визуально — WrapText сам по себе высоту не меняет.
' Здесь высота считается автоматически (Rows.AutoFit), после того как
' очень длинные пояснения обрезаны до разумного предела: без этого
' предела одна плитка с полусотней записей растянулась бы на пол-листа.
Private Const ПРЕДЕЛ_ПОЯСНЕНИЯ As Long = 280
Private Const ВЫСОТА_МИН As Single = 20
Private Const ВЫСОТА_МАКС As Single = 80

Private Function Плитка(ws As Worksheet, r As Long, заголовок As String, _
                        число As String, пояснение As String, уровень As Long) As Long
    Dim фон As Long
    Dim цвет As Long

    Select Case уровень
        Case 2: фон = ЦВ_СРОЧНО: цвет = ЦВ_ТЕКСТ_СР
        Case 1: фон = ЦВ_ВНИМАНИЕ: цвет = ЦВ_ТЕКСТ_ВН
        Case Else: фон = ЦВ_ОК: цвет = ЦВ_ТЕКСТ_ОК
    End Select

    ws.Cells(r, 2).Value = заголовок
    ws.Cells(r, 3).Value = число
    ws.Cells(r, 4).Value = Обрезано(пояснение, ПРЕДЕЛ_ПОЯСНЕНИЯ)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Interior.Color = фон
        .Font.Color = цвет
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
        .Borders.Weight = xlMedium
    End With
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Font.Bold = True
    ws.Cells(r, 3).HorizontalAlignment = xlCenter
    ws.Cells(r, 4).Font.Size = 9

    ws.Rows(r).RowHeight = ВЫСОТА_МИН        ' с этого AutoFit посчитает верно
    ws.Rows(r).AutoFit
    If ws.Rows(r).RowHeight < ВЫСОТА_МИН Then ws.Rows(r).RowHeight = ВЫСОТА_МИН
    If ws.Rows(r).RowHeight > ВЫСОТА_МАКС Then ws.Rows(r).RowHeight = ВЫСОТА_МАКС

    Плитка = r + 1
End Function


Private Function Обрезано(текст As String, предел As Long) As String
    If Len(текст) <= предел Then
        Обрезано = текст
    Else
        Обрезано = Left$(текст, предел) & "… (список длиннее — подробности в соответствующем реестре)"
    End If
End Function
