Attribute VB_Name = "modJournal"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Работа с журналом: новая запись, регистрация, аннулирование.
'
'  Порядок такой же, как на бумаге: сначала запись живёт черновиком —
'  её можно править и по ней можно сформировать документ на проверку;
'  настоящие номера выдаются одним действием в момент регистрации.
' ═══════════════════════════════════════════════════════════════════════

' Сколько строк добавлено в журнал выдачи за пакетную регистрацию.
' Объявление стоит здесь, а не рядом с процедурой, которая его считает:
' переменные уровня модуля Excel читает только в секции объявлений, до
' первой процедуры, а ниже по файлу просто не видит.
Private ВыдачаОтмечена As Long


' Новая строка журнала: заводится черновиком, курсор встаёт на дату.
Public Sub НоваяЗапись()
    Dim ws As Worksheet
    Dim r As Long
    Dim методы As Variant
    Dim i As Long

    If Not ПроверитьКодРабочегоМеста() Then Exit Sub
    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub

    r = ПоследняяСтрока(ws, ColReq(ws, "Дата контроля")) + 1
    ws.Cells(r, ColReq(ws, "ID записи")).Value = НовыйID()
    ws.Cells(r, ColReq(ws, "Статус")).Value = "Черновик"
    ws.Cells(r, ColReq(ws, "Дата контроля")).Value = Date
    ws.Cells(r, ColReq(ws, "Кем внесено")).Value = ТекущийСотрудник()
    ws.Cells(r, ColReq(ws, "Изменено")).Value = Now

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        ws.Cells(r, ColReq(ws, методы(i) & " №")).ClearContents
    Next i

    ws.Activate
    ws.Cells(r, ColReq(ws, "Дата контроля")).Select
    MsgBox "Заведена черновая запись " & ws.Cells(r, ColReq(ws, "ID записи")).Value & "." & vbCrLf & vbCrLf & _
           "Заполните дату, объект и по каждому применённому методу — ИСПОЛНИТЕЛЯ. " & _
           "Именно исполнитель отмечает, что метод применялся; колонку «№» не трогайте, " & _
           "номер туда впишет система." & vbCrLf & vbCrLf & _
           "Потом «Зарегистрировать» — тогда будут выданы настоящие номера. " & _
           "До регистрации запись в бумажный журнал не идёт.", _
           vbInformation, "Новая запись"
End Sub


' Внутренний идентификатор записи: код рабочего места + год + счётчик.
' Локальный счётчик живёт в настройках этой машины, поэтому записи,
' заведённые на разных ПК без сети, не столкнутся при слиянии.
Public Function НовыйID() As String
    Dim n As Long
    n = CLng(Val(Настройка("Счётчик записей рабочего места"))) + 1
    ЗаписатьНастройку "Счётчик записей рабочего места", n
    НовыйID = КодРабочегоМеста() & "-" & Format$(Date, "yyyy") & "-" & Format$(n, "0000")
End Function


' Регистрация: выдаёт номера сразу по всем заполненным методам строки.
Public Sub ЗарегистрироватьЗапись()
    Dim ws As Worksheet
    Dim r As Long
    Dim методы As Variant
    Dim i As Long
    Dim метод As String
    Dim год As Long
    Dim предупреждения As String
    Dim заполнено As Long
    Dim выдано As String
    Dim номер As Long
    Dim cНомер As Long
    Dim выдач As Long
    Dim расхождения As String

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    If Not ПроверитьКодРабочегоМеста() Then Exit Sub

    r = ТекущаяСтрока(ws)
    If r = 0 Then Exit Sub

    If StrComp(CStr(ws.Cells(r, ColReq(ws, "Статус")).Value), "Зарегистрирован", vbTextCompare) = 0 Then
        MsgBox "Эта запись уже зарегистрирована. Повторная регистрация выдала бы " & _
               "второй номер на ту же работу.", vbExclamation, "Регистрация"
        Exit Sub
    End If

    If Not IsDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value) Then
        MsgBox "Не заполнена дата контроля — по ней определяется год нумерации.", _
               vbExclamation, "Регистрация"
        Exit Sub
    End If
    год = Year(CDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value))

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        If МетодПрименён(ws, r, CStr(методы(i))) Then
            заполнено = заполнено + 1
        End If
    Next i
    If заполнено = 0 Then
        MsgBox "В строке не отмечен ни один метод контроля." & vbCrLf & vbCrLf & _
               "Поставьте любой знак в колонке «№» того метода, который применялся, " & _
               "и укажите исполнителя — номер подставится при регистрации.", _
               vbExclamation, "Регистрация"
        Exit Sub
    End If

    предупреждения = ПроверитьСтрокуЖурнала(r)
    If предупреждения <> "" Then
        If MsgBox("Перед регистрацией стоит посмотреть:" & vbCrLf & предупреждения & vbCrLf & vbCrLf & _
                  "Зарегистрировать всё равно?", vbYesNo + vbExclamation, "Проверка") = vbNo Then Exit Sub
    End If

    ' Номера выдаются по одному на метод. Если на середине оборвётся сеть,
    ' уже выданные останутся в строке — потерянных номеров не будет.
    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        cНомер = ColReq(ws, метод & " №")
        ' Номер нужен там, где метод применялся, а числа в клетке ещё нет.
        If МетодПрименён(ws, r, метод) And Not IsNumeric(ws.Cells(r, cНомер).Value) Then
            номер = ВыдатьНомер(метод, год)
            If номер = 0 Then
                MsgBox "Регистрация прервана на методе " & метод & "." & vbCrLf & _
                       "Выданные до него номера сохранены в строке — повторите " & _
                       "регистрацию, недостающие номера доберутся.", _
                       vbExclamation, "Регистрация"
                Exit Sub
            End If
            ws.Cells(r, cНомер).Value = номер
            выдано = выдано & vbCrLf & "  " & ФорматНомера(номер, год, метод)
        End If
    Next i

    ws.Cells(r, ColReq(ws, "Статус")).Value = "Зарегистрирован"
    ws.Cells(r, ColReq(ws, "Изменено")).Value = Now
    If Trim$(CStr(ws.Cells(r, ColReq(ws, "Кем внесено")).Value)) = "" Then
        ws.Cells(r, ColReq(ws, "Кем внесено")).Value = ТекущийСотрудник()
    End If

    ' Приборы, отмеченные применёнными, тем самым были выданы —
    ' журнал выдачи заполняется отсюда, а не второй раз руками.
    расхождения = modIssue.ВыдачаПоПрименённым(ws, r, выдач)

    MsgBox "Запись зарегистрирована. Выданы номера:" & выдано & _
           IIf(выдач = 0, "", vbCrLf & vbCrLf & "В журнал выдачи добавлено строк: " & выдач & _
                              " — по применённым приборам.") & _
           IIf(расхождения = "", "", vbCrLf & vbCrLf & расхождения), _
           IIf(расхождения = "", vbInformation, vbExclamation), "Регистрация"
End Sub


' Регистрация всех черновиков разом.
'
' Пока сети нет, носитель ходит по рукам: он оказывается у вас на
' несколько минут, и за это время надо зарегистрировать всё
' накопившееся. По одной записи это делать неудобно, а неудобное
' перестают делать — и записи копятся черновиками.
Public Sub ЗарегистрироватьВсеЧерновики()
    Dim ws As Worksheet
    Dim r As Long, last As Long
    Dim cСтатус As Long, cДата As Long
    Dim черновиков As Long
    Dim готово As Long
    Dim пропущено As Long
    Dim перечень As String
    Dim предупреждения As String

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    If Not ПроверитьКодРабочегоМеста() Then Exit Sub
    If ПапкаОбмена() = "" Then
        MsgBox "Общая папка недоступна — номера выдавать неоткуда." & vbCrLf & vbCrLf & _
               "Подключите штатную флешку и повторите. Черновики никуда не денутся.", _
               vbExclamation, "Регистрация черновиков"
        Exit Sub
    End If

    cСтатус = ColReq(ws, "Статус")
    cДата = ColReq(ws, "Дата контроля")
    last = ПоследняяСтрока(ws, cДата)

    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(CStr(ws.Cells(r, cСтатус).Value), "Черновик", vbTextCompare) = 0 Then
            черновиков = черновиков + 1
            If Len(перечень) < 500 Then
                перечень = перечень & vbCrLf & "  строка " & r & ": " & _
                           Format$(ws.Cells(r, cДата).Value, "dd.mm.yyyy") & ", " & _
                           Left$(CStr(ws.Cells(r, ColReq(ws, "Наименование объекта в документе")).Value), 44)
            End If
        End If
    Next r

    If черновиков = 0 Then
        MsgBox "Черновиков нет — регистрировать нечего.", vbInformation, "Регистрация черновиков"
        Exit Sub
    End If

    If MsgBox("Черновиков к регистрации: " & черновиков & перечень & _
              IIf(черновиков > 8, vbCrLf & "  …", "") & vbCrLf & vbCrLf & _
              "Номера будут выданы через " & ОписаниеОбмена() & "." & vbCrLf & _
              "Записи с замечаниями будут пропущены и названы в конце." & vbCrLf & vbCrLf & _
              "Зарегистрировать всё?", vbYesNo + vbQuestion, "Регистрация черновиков") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    перечень = ""
    ВыдачаОтмечена = 0
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(CStr(ws.Cells(r, cСтатус).Value), "Черновик", vbTextCompare) = 0 Then
            предупреждения = ПричинаНеРегистрировать(ws, r)
            If предупреждения <> "" Then
                пропущено = пропущено + 1
                If Len(перечень) < 700 Then
                    перечень = перечень & vbCrLf & "  строка " & r & ": " & предупреждения
                End If
            ElseIf ВыдатьНомераСтроки(ws, r) Then
                готово = готово + 1
            Else
                пропущено = пропущено + 1
                перечень = перечень & vbCrLf & "  строка " & r & ": не удалось выдать номер"
                Exit For                  ' носитель отвалился — дальше бессмысленно
            End If
        End If
    Next r
    Application.ScreenUpdating = True

    MsgBox "Зарегистрировано: " & готово & " из " & черновиков & "." & _
           IIf(ВыдачаОтмечена = 0, "", vbCrLf & "В журнал выдачи добавлено строк: " & _
               ВыдачаОтмечена & " — по применённым приборам.") & _
           IIf(пропущено = 0, "", vbCrLf & vbCrLf & "Пропущено " & пропущено & ":" & перечень & _
               vbCrLf & vbCrLf & "Эти записи остались черновиками — поправьте и повторите.") , _
           vbInformation, "Регистрация черновиков"
End Sub


' Что мешает зарегистрировать строку. Пусто — можно регистрировать.
Private Function ПричинаНеРегистрировать(ws As Worksheet, r As Long) As String
    Dim методы As Variant
    Dim i As Long
    Dim заполнено As Long

    If Not IsDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value) Then
        ПричинаНеРегистрировать = "не заполнена дата контроля"
        Exit Function
    End If
    If Trim$(CStr(ws.Cells(r, ColReq(ws, "ID объекта")).Value)) = "" Then
        ПричинаНеРегистрировать = "не выбран объект контроля"
        Exit Function
    End If

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        If МетодПрименён(ws, r, CStr(методы(i))) Then
            заполнено = заполнено + 1
            If Trim$(CStr(ws.Cells(r, ColReq(ws, методы(i) & " исполнитель")).Value)) = "" Then
                ПричинаНеРегистрировать = "метод " & методы(i) & " без исполнителя"
                Exit Function
            End If
        End If
    Next i
    If заполнено = 0 Then ПричинаНеРегистрировать = "не отмечен ни один метод"
End Function


' Выдаёт номера по всем отмеченным методам строки. False — если
' носитель пропал на середине; уже выданные номера при этом
' остаются в строке и вторично не запрашиваются.
Private Function ВыдатьНомераСтроки(ws As Worksheet, r As Long) As Boolean
    Dim методы As Variant
    Dim i As Long
    Dim год As Long
    Dim номер As Long
    Dim cНомер As Long
    Dim выдач As Long

    год = Year(CDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value))
    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        cНомер = ColReq(ws, методы(i) & " №")
        If МетодПрименён(ws, r, CStr(методы(i))) And Not IsNumeric(ws.Cells(r, cНомер).Value) Then
            номер = ВыдатьНомер(CStr(методы(i)), год)
            If номер = 0 Then Exit Function
            ws.Cells(r, cНомер).Value = номер
        End If
    Next i

    ws.Cells(r, ColReq(ws, "Статус")).Value = "Зарегистрирован"
    ws.Cells(r, ColReq(ws, "Изменено")).Value = Now
    If Trim$(CStr(ws.Cells(r, ColReq(ws, "Кем внесено")).Value)) = "" Then
        ws.Cells(r, ColReq(ws, "Кем внесено")).Value = ТекущийСотрудник()
    End If

    ' То же, что при поштучной регистрации: применённый прибор был выдан.
    ' Расхождения здесь не показываются по одному — их видно на строке
    ' кнопкой «Отметить выдачу по применённым».
    modIssue.ВыдачаПоПрименённым ws, r, выдач
    ВыдачаОтмечена = ВыдачаОтмечена + выдач
    ВыдатьНомераСтроки = True
End Function


' Открыть новый акт и присвоить его выделенным строкам журнала.
'
' Один акт — одна работа, а объектов в ней обычно несколько: выделяете
' строки этих объектов и жмёте кнопку. Раньше номер акта вписывался
' руками, и именно на этом акты дублировались.
Public Sub ОткрытьНовыйАкт()
    Dim ws As Worksheet
    Dim строки As Range
    Dim ячейка As Range
    Dim номер As String
    Dim год As Long
    Dim дата As Variant
    Dim сколько As Long
    Dim cАкт As Long

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    If Not ПроверитьКодРабочегоМеста() Then Exit Sub
    If ActiveSheet.Name <> ws.Name Then
        MsgBox "Выделите на листе «" & ws.Name & "» строки объектов, которые войдут в акт.", _
               vbInformation, "Новый акт"
        Exit Sub
    End If

    Set строки = Selection
    cАкт = ColReq(ws, "№ акта")
    дата = ws.Cells(строки.Row, ColReq(ws, "Дата контроля")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))

    For Each ячейка In строки.Columns(1).Cells
        If ячейка.Row >= ПЕРВАЯ_СТРОКА Then
            If IsDate(ws.Cells(ячейка.Row, ColReq(ws, "Дата контроля")).Value) Then
                сколько = сколько + 1
            End If
        End If
    Next ячейка

    If сколько = 0 Then
        MsgBox "Не выделено ни одной строки с датой контроля.", vbExclamation, "Новый акт"
        Exit Sub
    End If

    If MsgBox("Открыть новый акт за " & год & " год и присвоить его " & сколько & _
              " строкам?" & vbCrLf & vbCrLf & _
              "Номер выдаст счётчик — он один на всех, поэтому два акта с одним " & _
              "номером появиться не могут.", vbYesNo + vbQuestion, "Новый акт") = vbNo Then Exit Sub

    номер = ВыдатьНомерАкта(год)
    If номер = "" Then Exit Sub

    For Each ячейка In строки.Columns(1).Cells
        If ячейка.Row >= ПЕРВАЯ_СТРОКА Then
            If IsDate(ws.Cells(ячейка.Row, ColReq(ws, "Дата контроля")).Value) Then
                ws.Cells(ячейка.Row, cАкт).Value = номер
                ws.Cells(ячейка.Row, ColReq(ws, "Изменено")).Value = Now
            End If
        End If
    Next ячейка

    MsgBox "Открыт " & номер & ", присвоен " & сколько & " строкам." & vbCrLf & vbCrLf & _
           "Печать пакета собирает все строки одного акта: акт, затем протоколы " & _
           "по всем его объектам.", vbInformation, "Новый акт"
End Sub


' Номер акта текущей строки — по нему собирается пакет.
Public Function НомерАкта(ws As Worksheet, строка As Long) As String
    НомерАкта = Trim$(CStr(ws.Cells(строка, ColReq(ws, "№ акта")).Value))
End Function


' Аннулирование по итогам рекламации: строка остаётся на месте,
' иначе в журнале появится дыра, а по Правилам аннулированное
' заключение должно быть в журнале видно.
Public Sub АннулироватьЗапись()
    Dim ws As Worksheet
    Dim r As Long
    Dim основание As String
    Dim замена As String

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ws)
    If r = 0 Then Exit Sub

    основание = InputBox("Основание аннулирования (номер рекламации, решение комиссии):", _
                         "Аннулирование заключения")
    If Trim$(основание) = "" Then Exit Sub
    замена = InputBox("Номер нового заключения, выданного взамен (можно оставить пустым):", _
                      "Аннулирование заключения")

    ws.Cells(r, ColReq(ws, "Статус")).Value = "Аннулирован"
    ws.Cells(r, ColReq(ws, "Примечание")).Value = _
        Trim$(CStr(ws.Cells(r, ColReq(ws, "Примечание")).Value) & _
              " Аннулировано " & Format$(Date, "dd.mm.yyyy") & ". Основание: " & основание & _
              IIf(Trim$(замена) = "", "", ". Взамен выдано: " & замена) & ".")
    ws.Cells(r, ColReq(ws, "Изменено")).Value = Now

    MsgBox "Запись помечена аннулированной. Номер за ней сохранён — он не будет " & _
           "выдан повторно, и в печатном журнале строка останется с пометкой.", _
           vbInformation, "Аннулирование"
End Sub


' Строка, на которой стоит курсор, с понятной подсказкой при промахе.
Public Function ТекущаяСтрока(ws As Worksheet) As Long
    Dim r As Long
    If ActiveSheet.Name <> ws.Name Then
        MsgBox "Встаньте на нужную строку листа «" & ws.Name & "» и повторите.", _
               vbInformation, "Выбор строки"
        Exit Function
    End If
    r = ActiveCell.Row
    If r < ПЕРВАЯ_СТРОКА Then
        MsgBox "Курсор стоит в шапке таблицы. Встаньте на строку с записью.", _
               vbInformation, "Выбор строки"
        Exit Function
    End If
    ТекущаяСтрока = r
End Function


' Регистрационный номер по методу для строки — для документов и печати.
Public Function РегНомер(ws As Worksheet, строка As Long, метод As String) As String
    Dim v As Variant
    Dim дата As Variant
    v = ws.Cells(строка, ColReq(ws, метод & " №")).Value
    If Trim$(CStr(v)) = "" Then
        ' Метод применялся, но номер ещё не выдан: в печатном журнале
        ' такая строка должна быть видна, а не пропасть.
        If МетодПрименён(ws, строка, метод) Then РегНомер = "не зарегистрирован"
        Exit Function
    End If
    дата = ws.Cells(строка, ColReq(ws, "Дата контроля")).Value
    If Not IsNumeric(v) Then
        РегНомер = CStr(v)          ' черновик или «?» из миграции
    ElseIf IsDate(дата) Then
        РегНомер = ФорматНомера(CLng(v), Year(CDate(дата)), метод)
    Else
        РегНомер = CStr(v)
    End If
End Function
