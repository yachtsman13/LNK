Attribute VB_Name = "ПробаФормы"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  ПРОБА. Это не часть системы: один модуль, один запуск, после себя
'  ничего не оставляет — построенная форма удаляется, книга не меняется.
'
'  Проверяет три вещи, которые нельзя выяснить, не запустив Excel:
'    1. строится ли форма поиска прямо из кода на этой машине — от этого
'       зависит, каким способом делать поиск во всех списках;
'    2. крутится ли список колесом мыши;
'    3. поспевает ли отбор за набором на трёх сотнях позиций.
'
'  Запуск: Alt+F11 → File → Import File… → выбрать этот файл →
'          F5 → ПроверитьФормуПоиска → Run.
'
'  Нужна та же настройка Excel, что и для установщика: Файл → Параметры →
'  Центр управления безопасностью → Параметры центра управления
'  безопасностью → Параметры макросов → «Доверять доступ к объектной
'  модели проектов VBA».
'
'  Запускать можно в любой книге, хоть в пустой. Если рядом открыта книга
'  данных, проба возьмёт настоящие объекты; если нет — выдуманные, той же
'  формы и длины.
' ═══════════════════════════════════════════════════════════════════════

Private Const ИМЯ_ФОРМЫ As String = "фПробаПоиска"
Private Const ТИП_ФОРМЫ As Long = 3          ' vbext_ct_MSForm

' Код формы обращается к этим напрямую, поэтому они Public.
Public ПробаПозиции As Variant               ' весь список
Public ПробаВыбрано As String                ' что выбрали
Public ПробаМиллисекунд As Long              ' сколько занял последний отбор
Public ПробаОтборов As Long                  ' сколько раз отбирали


Public Sub ПроверитьФормуПоиска()
    Dim vbc As Object
    Dim источник As String
    Dim отчёт As String

    If Not ЕстьДоступКVBA() Then
        MsgBox "Нет доступа к объектной модели VBA — проба не может " & _
               "построить форму." & vbCrLf & vbCrLf & _
               "Файл → Параметры → Центр управления безопасностью → " & _
               "Параметры центра управления безопасностью → Параметры " & _
               "макросов → включить «Доверять доступ к объектной модели " & _
               "проектов VBA», затем запустить пробу заново." & vbCrLf & vbCrLf & _
               "Если эта настройка недоступна — так и напишите: значит, " & _
               "поиск нужно делать без формы.", _
               vbExclamation, "Проба формы поиска"
        Exit Sub
    End If

    ПробаПозиции = СобратьПозиции(источник)
    ПробаВыбрано = ""
    ПробаМиллисекунд = 0
    ПробаОтборов = 0

    On Error GoTo Сбой
    УдалитьФорму                     ' вдруг осталась от прошлого запуска
    Set vbc = ПостроитьФорму()

    ' Форма показывается модально: выполнение стоит здесь, пока её не
    ' закроют, — и только потом собирается отчёт.
    VBA.UserForms.Add(vbc.Name).Show

    отчёт = "Форма построена: ДА" & vbCrLf & _
            "Позиций в списке: " & (UBound(ПробаПозиции) + 1) & _
            " (" & источник & ")" & vbCrLf & _
            "Отборов за сеанс: " & ПробаОтборов & vbCrLf & _
            "Последний отбор занял, мс: " & ПробаМиллисекунд & vbCrLf & _
            "Выбрано: " & IIf(ПробаВыбрано = "", "— ничего не выбирали", ПробаВыбрано)

    УдалитьФорму
    MsgBox отчёт & vbCrLf & vbCrLf & _
           "Форма удалена, книга не изменена — сохранять её не нужно." & vbCrLf & vbCrLf & _
           "Пришлите этот текст и скажите отдельно, самое важное: " & _
           "КРУТИЛСЯ ЛИ СПИСОК КОЛЕСОМ МЫШИ.", _
           vbInformation, "Проба формы поиска"
    Exit Sub

Сбой:
    УдалитьФорму
    MsgBox "Форму построить не удалось." & vbCrLf & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Пришлите этот текст — значит, этот способ не годится и поиск " & _
           "нужно делать без формы.", _
           vbCritical, "Проба формы поиска"
End Sub


' ── Построение формы ──────────────────────────────────────────────────

' Форма собирается из кода целиком: сама форма, шесть элементов на ней и
' её собственный код. Именно это и проверяется — прислать готовую форму
' файлом нельзя, её двоичную часть делает только сам Excel.
Private Function ПостроитьФорму() As Object
    Dim vbc As Object
    Dim эл As Object

    Set vbc = ThisWorkbook.VBProject.VBComponents.Add(ТИП_ФОРМЫ)
    vbc.Name = ИМЯ_ФОРМЫ
    vbc.Properties("Caption") = "Поиск позиции — проба"
    vbc.Properties("Width") = 430
    vbc.Properties("Height") = 345

    Set эл = vbc.Designer.Controls.Add("Forms.Label.1", "Пояснение", True)
    эл.Left = 8: эл.Top = 6: эл.Width = 408: эл.Height = 28
    эл.Caption = "Наберите часть названия или номера — можно несколько " & _
                 "слов в любом порядке. ПОПРОБУЙТЕ ПРОКРУТИТЬ СПИСОК КОЛЕСОМ МЫШИ."

    Set эл = vbc.Designer.Controls.Add("Forms.TextBox.1", "Запрос", True)
    эл.Left = 8: эл.Top = 38: эл.Width = 408: эл.Height = 20

    Set эл = vbc.Designer.Controls.Add("Forms.Label.1", "Счётчик", True)
    эл.Left = 8: эл.Top = 62: эл.Width = 408: эл.Height = 14

    Set эл = vbc.Designer.Controls.Add("Forms.ListBox.1", "Список", True)
    эл.Left = 8: эл.Top = 80: эл.Width = 408: эл.Height = 186

    Set эл = vbc.Designer.Controls.Add("Forms.CommandButton.1", "Выбрать", True)
    эл.Left = 8: эл.Top = 274: эл.Width = 120: эл.Height = 24
    эл.Caption = "Выбрать"

    Set эл = vbc.Designer.Controls.Add("Forms.CommandButton.1", "Закрыть", True)
    эл.Left = 296: эл.Top = 274: эл.Width = 120: эл.Height = 24
    эл.Caption = "Закрыть"

    vbc.CodeModule.AddFromString КодФормы()
    Set ПостроитьФорму = vbc
End Function


' Код самой формы. Он держится нарочно коротким: вся работа — в этом
' модуле, форма только передаёт себя в него. Чем меньше кода собирается
' строками, тем меньше в нём может сломаться.
Private Function КодФормы() As String
    Dim s As String
    s = "Option Explicit" & vbCrLf & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    ПробаФормы.ОтобратьВСписок Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Запрос_Change()" & vbCrLf
    s = s & "    ПробаФормы.ОтобратьВСписок Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Список_DblClick(ByVal Отмена As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    ПробаФормы.ЗабратьВыбор Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Выбрать_Click()" & vbCrLf
    s = s & "    ПробаФормы.ЗабратьВыбор Me" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Закрыть_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    КодФормы = s
End Function


Private Sub УдалитьФорму()
    Dim vbc As Object
    On Error Resume Next
    Set vbc = ThisWorkbook.VBProject.VBComponents(ИМЯ_ФОРМЫ)
    If Not vbc Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove vbc
    On Error GoTo 0
End Sub


Private Function ЕстьДоступКVBA() As Boolean
    Dim n As Long
    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.Count
    ЕстьДоступКVBA = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' ── Отбор ─────────────────────────────────────────────────────────────

' Список заполняется одним присвоением массива, а не построчно: на
' трёхстах позициях разница между «мгновенно» и «подтормаживает».
Public Sub ОтобратьВСписок(ф As Object)
    Dim начало As Single
    Dim найдено As Variant
    Dim сколько As Long

    начало = Timer
    найдено = ОтобратьПозиции(CStr(ф.Запрос.Text), сколько)
    ф.Список.Clear
    If сколько > 0 Then ф.Список.List = найдено
    ПробаМиллисекунд = CLng((Timer - начало) * 1000)
    ПробаОтборов = ПробаОтборов + 1
    ф.Счётчик.Caption = "Найдено: " & сколько & " из " & (UBound(ПробаПозиции) + 1) & _
                        "     ·     отбор занял " & ПробаМиллисекунд & " мс"
End Sub


Public Sub ЗабратьВыбор(ф As Object)
    If ф.Список.ListIndex >= 0 Then ПробаВыбрано = CStr(ф.Список.Value)
    ф.Hide
End Sub


' Слова ищутся все и в любом порядке: «бак 500» находит «Бак
' водоподъёмной установки, 500 л». Регистр не важен.
Private Function ОтобратьПозиции(запрос As String, ByRef сколько As Long) As Variant
    Dim рез() As String
    Dim слова As Variant
    Dim i As Long, k As Long, w As Long
    Dim подходит As Boolean
    Dim текст As String

    ReDim рез(0 To UBound(ПробаПозиции))
    слова = Split(Trim$(запрос), " ")
    k = 0
    For i = 0 To UBound(ПробаПозиции)
        текст = CStr(ПробаПозиции(i))
        подходит = True
        For w = LBound(слова) To UBound(слова)
            If Trim$(CStr(слова(w))) <> "" Then
                If InStr(1, текст, Trim$(CStr(слова(w))), vbTextCompare) = 0 Then
                    подходит = False
                    Exit For
                End If
            End If
        Next w
        If подходит Then
            рез(k) = текст
            k = k + 1
        End If
    Next i

    сколько = k
    If k > 0 Then ReDim Preserve рез(0 To k - 1)
    ОтобратьПозиции = рез
End Function


' ── Откуда берутся позиции ────────────────────────────────────────────

Private Function СобратьПозиции(ByRef источник As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim список As Collection
    Dim мас() As String
    Dim r As Long, last As Long, i As Long

    Set список = New Collection

    On Error Resume Next
    For Each wb In Application.Workbooks
        Set ws = Nothing
        Set ws = wb.Worksheets("Объекты")
        If Not ws Is Nothing Then Exit For
    Next wb
    On Error GoTo 0

    If Not ws Is Nothing Then
        ' Колонка 12 — то самое читаемое обозначение, которое сейчас стоит
        ' в выпадающем списке журнала.
        last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        For r = 4 To last
            If Trim$(CStr(ws.Cells(r, 12).Value)) <> "" Then
                список.Add CStr(ws.Cells(r, 12).Value)
            End If
        Next r
        источник = "настоящие объекты из книги данных"
    End If

    If список.Count = 0 Then
        For i = 1 To 300
            список.Add Образец(i)
        Next i
        источник = "выдуманные: книга данных не открыта"
    End If

    ReDim мас(0 To список.Count - 1)
    For i = 1 To список.Count
        мас(i - 1) = список(i)
    Next i
    СобратьПозиции = мас
End Function


' Выдуманная позиция той же формы, что и настоящие: наименование,
' заводской и регистрационный номера — чтобы поиск было на чём щупать.
Private Function Образец(i As Long) As String
    Dim виды As Variant
    виды = Array("Баллон (воздух)", "Бак водоподъёмной установки", "Ресивер", _
                 "Испаритель", "Конденсатор", "Сосуд под давлением", "Трубопровод")
    Образец = CStr(виды(i Mod 7)) & " V=" & (10 + (i Mod 9) * 5) & " л, зав. № " & _
              Format$(190000 + i, "000000") & ", рег. № " & (i Mod 97)
End Function
