Attribute VB_Name = "modConfig"
Option Explicit

' =======================================================================
'  Настройки, пути и поиск колонок по заголовку.
'
'  Колонки нигде не заданы номерами: их ищет ColByHeader по тексту
'  заголовка в строке 3. Вставили колонку в середину листа — код
'  продолжает работать, ничего править не надо.
' =======================================================================

Public Const ИМЯ_ФАЙЛА_ДАННЫХ As String = "ЛНК_Данные.xlsx"
Public Const СТРОКА_ЗАГОЛОВКОВ As Long = 3
Public Const ПЕРВАЯ_СТРОКА As Long = 4

Public Const Л_ЖУРНАЛ As String = "Журнал"
Public Const Л_ОБЪЕКТЫ As String = "Объекты"
Public Const Л_СПЕЦИАЛИСТЫ As String = "Специалисты"
Public Const Л_АТТЕСТАЦИИ As String = "Аттестации"
Public Const Л_СИ As String = "СИ"
Public Const Л_КТО_ГДЕ As String = "Кто_где"
Public Const Л_НАСТРОЙКИ As String = "Настройки"
Public Const Л_КОПИРОВАНИЕ As String = "Журнал_копирования"

Public Const МЕТОДЫ As String = "ВИК,УЗД,УЗТ,АЭ,ТВ,ТК,ВД,МК"

Private mДанные As Workbook


' -- Книга данных ------------------------------------------------------

' Возвращает книгу данных, при необходимости открывая её.
' Ищем рядом с книгой программы; если там нет — спрашиваем у пользователя.
Public Function Данные() As Workbook
    Dim путь As String
    Dim wb As Workbook

    If Not mДанные Is Nothing Then
        On Error Resume Next
        путь = mДанные.Name          ' обращение к закрытой книге даст ошибку
        On Error GoTo 0
        If путь <> "" Then
            Set Данные = mДанные
            Exit Function
        End If
        Set mДанные = Nothing
    End If

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, ИМЯ_ФАЙЛА_ДАННЫХ, vbTextCompare) = 0 Then
            Set mДанные = wb
            Set Данные = wb
            Exit Function
        End If
    Next wb

    путь = ThisWorkbook.Path & Application.PathSeparator & ИМЯ_ФАЙЛА_ДАННЫХ
    If Dir(путь) = "" Then
        MsgBox "Не найден файл данных:" & vbCrLf & путь & vbCrLf & vbCrLf & _
               "Он должен лежать в той же папке, что и программа. " & _
               "Если файл перенесли — верните его на место или укажите путь в настройках.", _
               vbCritical, "Файл данных"
        Exit Function
    End If

    Set mДанные = Application.Workbooks.Open(fileName:=путь, UpdateLinks:=0)
    Set Данные = mДанные
End Function


Public Function Лист(имяЛиста As String) As Worksheet
    Dim wb As Workbook
    Set wb = Данные()
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set Лист = wb.Worksheets(имяЛиста)
    On Error GoTo 0
    If Лист Is Nothing Then
        MsgBox "В файле данных нет листа «" & имяЛиста & "».", vbCritical, "Файл данных"
    End If
End Function


' -- Колонки по заголовку ----------------------------------------------

' Номер колонки по тексту заголовка. Переносы строк в заголовке
' приравниваются к пробелу, регистр и лишние пробелы не важны.
Public Function ColByHeader(ws As Worksheet, заголовок As String) As Long
    Dim c As Long
    Dim last As Long
    Dim v As String
    Dim искомое As String

    искомое = НормШапка(заголовок)
    last = ws.Cells(СТРОКА_ЗАГОЛОВКОВ, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To last
        v = НормШапка(CStr(ws.Cells(СТРОКА_ЗАГОЛОВКОВ, c).Value))
        If v = искомое Then
            ColByHeader = c
            Exit Function
        End If
    Next c
    ColByHeader = 0
End Function


' То же, но без колонки работать нельзя — сообщаем и останавливаемся.
Public Function ColReq(ws As Worksheet, заголовок As String) As Long
    ColReq = ColByHeader(ws, заголовок)
    If ColReq = 0 Then
        Err.Raise vbObjectError + 513, "ColReq", _
                  "На листе «" & ws.Name & "» нет колонки «" & заголовок & "». " & _
                  "Вероятно, заголовок переименовали — верните прежний текст."
    End If
End Function


Private Function НормШапка(s As String) As String
    Dim t As String
    t = Replace(Replace(s, vbLf, " "), vbCr, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    НормШапка = LCase$(Trim$(t))
End Function


' Последняя заполненная строка таблицы: считаем по ключевой колонке,
' а не по UsedRange — формулы протянуты далеко вниз и портят счёт.
Public Function ПоследняяСтрока(ws As Worksheet, колонкаКлюча As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, колонкаКлюча).End(xlUp).Row
    If r < ПЕРВАЯ_СТРОКА Then r = ПЕРВАЯ_СТРОКА - 1
    ПоследняяСтрока = r
End Function


' -- Настройки ---------------------------------------------------------

Public Function Настройка(имя As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Dim last As Long

    Set ws = Лист(Л_НАСТРОЙКИ)
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), имя, vbTextCompare) = 0 Then
            Настройка = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function


Public Sub ЗаписатьНастройку(имя As String, значение As Variant)
    Dim ws As Worksheet
    Dim r As Long
    Dim last As Long

    Set ws = Лист(Л_НАСТРОЙКИ)
    If ws Is Nothing Then Exit Sub
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), имя, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value = значение
            Exit Sub
        End If
    Next r
End Sub


Public Function КодРабочегоМеста() As String
    КодРабочегоМеста = Настройка("Код рабочего места")
End Function


Public Function ТекущийСотрудник() As String
    ТекущийСотрудник = Настройка("ФИО сотрудника")
End Function


Public Function ЭтоМастер() As Boolean
    ЭтоМастер = (StrComp(Настройка("Это мастер-ПК"), "Да", vbTextCompare) = 0)
End Function


' Рабочее место обязано быть подписано: без кода ID записей на разных
' машинах совпадут, и при слиянии одна затрёт другую.
Public Function ПроверитьКодРабочегоМеста() As Boolean
    If КодРабочегоМеста() = "" Then
        MsgBox "Не заполнен код рабочего места." & vbCrLf & vbCrLf & _
               "Откройте лист «Настройки» и впишите код (ПК1 … ПК6). " & _
               "Он входит во внутренний номер записи и в документах не печатается. " & _
               "Без него две машины заведут записи с одинаковыми номерами.", _
               vbExclamation, "Настройки"
        Exit Function
    End If
    ПроверитьКодРабочегоМеста = True
End Function


' -- Пути --------------------------------------------------------------

Public Function Путь(имяНастройки As String, поУмолчанию As String) As String
    Dim p As String
    p = Настройка(имяНастройки)
    If p = "" Then p = ThisWorkbook.Path & Application.PathSeparator & поУмолчанию
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    Путь = p
End Function

Public Function ПапкаШаблонов() As String
    ПапкаШаблонов = Путь("Папка шаблонов", "Шаблоны")
End Function

Public Function ПапкаГотовых() As String
    ПапкаГотовых = Путь("Папка готовых документов", "Готовые")
End Function

Public Function ПапкаКопий() As String
    ПапкаКопий = Путь("Папка локальных копий", "Копии")
End Function

Public Function ПапкаОбмена() As String
    ПапкаОбмена = Путь("Общая папка обмена", "Обмен")
End Function

Public Function ПапкаСчётчиков() As String
    ПапкаСчётчиков = ПапкаОбмена() & "Счётчики" & Application.PathSeparator
End Function


' Создаёт путь целиком, включая промежуточные папки.
Public Sub MkPath(ByVal п As String)
    Dim части() As String
    Dim i As Long
    Dim текущий As String

    п = Replace(п, "/", "\")
    If Right$(п, 1) = "\" Then п = Left$(п, Len(п) - 1)
    части = Split(п, "\")
    If Left$(п, 2) = "\\" Then
        текущий = "\\" & части(2) & "\" & части(3)
        i = 4
    Else
        текущий = части(0)
        i = 1
    End If
    For i = i To UBound(части)
        текущий = текущий & "\" & части(i)
        If Dir(текущий, vbDirectory) = "" Then MkDir текущий
    Next i
End Sub


Public Function ОчиститьДляИмениФайла(ByVal s As String) As String
    Dim плохие As Variant
    Dim i As Long
    плохие = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbLf, vbCr)
    For i = LBound(плохие) To UBound(плохие)
        s = Replace(s, плохие(i), " ")
    Next i
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ОчиститьДляИмениФайла = Trim$(s)
End Function


Public Function МеткаВремени() As String
    МеткаВремени = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
End Function


' Пауза без Application.Wait: тот умеет только целые секунды.
Public Sub Подождать(мс As Long)
    Dim t As Double
    t = Timer
    Do While (Timer - t) * 1000 < мс
        DoEvents
        If Timer < t Then Exit Do    ' полночь
    Loop
End Sub
