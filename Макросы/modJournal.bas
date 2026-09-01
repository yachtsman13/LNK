Attribute VB_Name = "modJournal"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Работа с журналом: новая запись, регистрация, аннулирование.
'
'  Порядок такой же, как на бумаге: сначала запись живёт черновиком —
'  её можно править и по ней можно сформировать документ на проверку;
'  настоящие номера выдаются одним действием в момент регистрации.
' ═══════════════════════════════════════════════════════════════════════


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
           "Заполните объект и методы, затем нажмите «Зарегистрировать» — тогда будут " & _
           "выданы настоящие номера. До регистрации запись в бумажный журнал не идёт.", _
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
        If Trim$(CStr(ws.Cells(r, ColReq(ws, методы(i) & " №")).Value)) <> "" Then
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
        If Trim$(CStr(ws.Cells(r, cНомер).Value)) <> "" Then
            If Not IsNumeric(ws.Cells(r, cНомер).Value) Then
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
        End If
    Next i

    ws.Cells(r, ColReq(ws, "Статус")).Value = "Зарегистрирован"
    ws.Cells(r, ColReq(ws, "Изменено")).Value = Now
    If Trim$(CStr(ws.Cells(r, ColReq(ws, "Кем внесено")).Value)) = "" Then
        ws.Cells(r, ColReq(ws, "Кем внесено")).Value = ТекущийСотрудник()
    End If

    MsgBox "Запись зарегистрирована. Выданы номера:" & выдано, vbInformation, "Регистрация"
End Sub


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
    If Trim$(CStr(v)) = "" Then Exit Function
    дата = ws.Cells(строка, ColReq(ws, "Дата контроля")).Value
    If Not IsNumeric(v) Then
        РегНомер = CStr(v)          ' черновик или «?» из миграции
    ElseIf IsDate(дата) Then
        РегНомер = ФорматНомера(CLng(v), Year(CDate(дата)), метод)
    Else
        РегНомер = CStr(v)
    End If
End Function
