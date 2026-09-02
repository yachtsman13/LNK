Attribute VB_Name = "modNumbers"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Счётчики регистрационных номеров.
'
'  Номер выдаётся только через файл-счётчик в общей папке — он один на
'  всех, поэтому пропусков и повторов не возникает. Пока номера нет,
'  запись живёт с черновым номером и в бумажный журнал не попадает.
'
'  Формат номера: NN-ГГГГ-Метод. Ведущий ноль до двух знаков
'  (07-2026-ВИК), дальше двух знаков дополнения нет (204-2026-ВИК).
' ═══════════════════════════════════════════════════════════════════════

Private Const ЖДАТЬ_БЛОКИРОВКУ_МС As Long = 5000
Private Const УСТАРЕВШАЯ_БЛОКИРОВКА_МИН As Long = 2


Public Function ФорматНомера(номер As Long, год As Long, метод As String) As String
    Dim s As String
    If номер < 10 Then
        s = "0" & CStr(номер)
    Else
        s = CStr(номер)
    End If
    ФорматНомера = s & "-" & CStr(год) & "-" & метод
End Function


Public Function ЧерновойНомер(порядковый As Long) As String
    ЧерновойНомер = "ЧЕРНОВИК-" & КодРабочегоМеста() & "-" & Format$(порядковый, "000")
End Function


Public Function ЭтоЧерновик(значение As Variant) As Boolean
    ЭтоЧерновик = (InStr(1, CStr(значение), "ЧЕРНОВИК", vbTextCompare) = 1)
End Function


' Акт — свой уровень учёта и своя серия номеров: он один на работу и
' охватывает несколько объектов, у каждого из которых внутри акта свои
' номера протоколов по методам. Поэтому счётчик у акта отдельный, а не
' общий с методами.
Public Function ФорматНомераАкта(номер As Long, год As Long) As String
    Dim s As String
    If номер < 10 Then s = "0" & CStr(номер) Else s = CStr(номер)
    ФорматНомераАкта = "АКТ " & s & "-" & CStr(год)
End Function


Public Function ВыдатьНомерАкта(год As Long) As String
    Dim n As Long
    n = ВыдатьНомер("АКТ", год)
    If n = 0 Then Exit Function
    ВыдатьНомерАкта = ФорматНомераАкта(n, год)
End Function


' ── Выдача номера ─────────────────────────────────────────────────────

' Забирает следующий номер по методу и году. Возвращает 0, если не вышло,
' и объясняет причину — молча номер не пропадает.
Public Function ВыдатьНомер(метод As String, год As Long) As Long
    Dim папка As String
    Dim файлСчётчика As String
    Dim файлБлокировки As String
    Dim мояМетка As String
    Dim текущий As Long
    Dim попытка As Long

    папка = ПапкаСчётчиков()
    If Not ПапкаСуществует(ПапкаОбмена()) Then
        MsgBox "Общая папка недоступна:" & vbCrLf & ПапкаОбмена() & vbCrLf & vbCrLf & _
               "Номер можно выдать только через неё — иначе два человека получат один и тот же. " & _
               "Работайте с черновым номером, зарегистрируете, когда появится сеть.", _
               vbExclamation, "Выдача номера"
        Exit Function
    End If

    On Error GoTo Сбой
    MkPath папка
    файлСчётчика = папка & метод & "-" & CStr(год) & ".txt"
    файлБлокировки = папка & метод & "-" & CStr(год) & ".lock"
    Randomize
    мояМетка = КодРабочегоМеста() & "|" & МеткаВремени() & "|" & CStr(Int(Rnd() * 100000))

    For попытка = 1 To 3
        If ЗахватитьБлокировку(файлБлокировки, мояМетка) Then
            текущий = ПрочитатьСчётчик(файлСчётчика)
            текущий = текущий + 1
            ЗаписатьСчётчик файлСчётчика, текущий
            ' проверяем, что записалось именно наше — файловая система могла подвести
            If ПрочитатьСчётчик(файлСчётчика) <> текущий Then
                СнятьБлокировку файлБлокировки
                Err.Raise vbObjectError + 520, "ВыдатьНомер", _
                          "Счётчик не сохранился: записали " & текущий & ", прочитали другое."
            End If
            СнятьБлокировку файлБлокировки
            ЗаписатьВРеестрВыданных метод, год, текущий
            СкопироватьСчётчики          ' резервная копия, уровень 1
            ВыдатьНомер = текущий
            Exit Function
        End If
        Подождать 400
    Next попытка

    MsgBox "Счётчик по методу " & метод & " за " & год & " год занят другим рабочим местом " & _
           "дольше обычного." & vbCrLf & vbCrLf & _
           "Подождите несколько секунд и повторите. Если повторяется — проверьте, " & _
           "не остался ли файл " & метод & "-" & год & ".lock после аварийного закрытия.", _
           vbExclamation, "Выдача номера"
    Exit Function

Сбой:
    On Error Resume Next
    СнятьБлокировку файлБлокировки
    MsgBox "Не удалось выдать номер." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Выдача номера"
End Function


' Текущее значение счётчика, не забирая номер, — для сверки при переходе.
Public Function ПосмотретьСчётчик(метод As String, год As Long) As Long
    ПосмотретьСчётчик = ПрочитатьСчётчик(ПапкаСчётчиков() & метод & "-" & CStr(год) & ".txt")
End Function


Public Sub УстановитьСчётчик(метод As String, год As Long, значение As Long)
    MkPath ПапкаСчётчиков()
    ЗаписатьСчётчик ПапкаСчётчиков() & метод & "-" & CStr(год) & ".txt", значение
End Sub


' Каждый выданный номер дописывается в реестр на носителе: номер, кто
' взял, когда. Файл счётчика хранит только последнее число, а реестр —
' всю историю. Пока сети нет и книги живут порознь, это единственное
' место, где видна вся выданная нумерация сразу; заодно по нему
' восстанавливается счётчик, если файл счётчика потеряется.
Private Sub ЗаписатьВРеестрВыданных(метод As String, год As Long, номер As Long)
    Dim файл As String
    Dim f As Integer

    On Error Resume Next
    файл = ПапкаСчётчиков() & "Выданные-" & метод & "-" & CStr(год) & ".csv"
    If Not ФайлСуществует(файл) Then
        f = FreeFile
        Open файл For Output As #f
        Print #f, "Номер;Рабочее место;Сотрудник;Дата и время"
        Close #f
    End If
    f = FreeFile
    Open файл For Append As #f
    Print #f, CStr(номер) & ";" & КодРабочегоМеста() & ";" & ТекущийСотрудник() & ";" & _
              Format$(Now, "dd.mm.yyyy hh:nn:ss")
    Close #f
End Sub


' ── Блокировка ────────────────────────────────────────────────────────

' Пытается занять счётчик. Настоящей атомарной блокировки в VBA нет,
' поэтому после записи перечитываем файл: если в нём чужая метка —
' значит, нас опередили, и номер мы не берём. Плюс мастер при сборке
' свода проверяет номера на дубли (раздел 5 ТЗ).
Private Function ЗахватитьБлокировку(файл As String, мояМетка As String) As Boolean
    Dim f As Integer
    Dim ждём As Double

    ждём = Timer
    Do While ФайлСуществует(файл)
        If БлокировкаУстарела(файл) Then
            On Error Resume Next
            Kill файл
            On Error GoTo 0
            Exit Do
        End If
        If (Timer - ждём) * 1000 > ЖДАТЬ_БЛОКИРОВКУ_МС Then Exit Function
        Подождать 200
    Loop

    On Error GoTo Сбой
    f = FreeFile
    Open файл For Output As #f
    Print #f, мояМетка
    Close #f

    Подождать 120
    If ПрочитатьСтроку(файл) <> мояМетка Then Exit Function   ' нас опередили

    ЗахватитьБлокировку = True
    Exit Function
Сбой:
    On Error Resume Next
    Close #f
End Function


Private Sub СнятьБлокировку(файл As String)
    On Error Resume Next
    If ФайлСуществует(файл) Then Kill файл
End Sub


' Книга могла упасть, не сняв блокировку. Через две минуты считаем её
' брошенной — иначе одна авария остановит выдачу номеров всем.
Private Function БлокировкаУстарела(файл As String) As Boolean
    On Error Resume Next
    БлокировкаУстарела = (DateDiff("n", FileDateTime(файл), Now) >= УСТАРЕВШАЯ_БЛОКИРОВКА_МИН)
End Function


' ── Чтение и запись файлов счётчиков ──────────────────────────────────

Private Function ПрочитатьСчётчик(файл As String) As Long
    Dim s As String
    s = Trim$(ПрочитатьСтроку(файл))
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then
        Err.Raise vbObjectError + 521, "ПрочитатьСчётчик", _
                  "В файле счётчика " & файл & " лежит не число, а «" & s & "»."
    End If
    ПрочитатьСчётчик = CLng(s)
End Function


Private Sub ЗаписатьСчётчик(файл As String, значение As Long)
    Dim f As Integer
    f = FreeFile
    Open файл For Output As #f
    Print #f, CStr(значение)
    Close #f
End Sub


Private Function ПрочитатьСтроку(файл As String) As String
    Dim f As Integer
    Dim s As String
    If Not ФайлСуществует(файл) Then Exit Function
    On Error GoTo Сбой
    f = FreeFile
    Open файл For Input As #f
    If Not EOF(f) Then Line Input #f, s
    Close #f
    ПрочитатьСтроку = s
    Exit Function
Сбой:
    On Error Resume Next
    Close #f
End Function
