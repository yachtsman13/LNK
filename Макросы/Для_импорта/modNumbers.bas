Attribute VB_Name = "modNumbers"
Option Explicit

' =======================================================================
'  —чЄтчики регистрационных номеров.
'
'  Ќомер выдаЄтс€ только через файл-счЄтчик в общей папке Ч он один на
'  всех, поэтому пропусков и повторов не возникает. ѕока номера нет,
'  запись живЄт с черновым номером и в бумажный журнал не попадает.
'
'  ‘ормат номера: NN-√√√√-ћетод. ¬едущий ноль до двух знаков
'  (07-2026-¬» ), дальше двух знаков дополнени€ нет (204-2026-¬» ).
' =======================================================================

Private Const ∆ƒј“№_ЅЋќ »–ќ¬ ”_ћ— As Long = 5000
Private Const ”—“ј–≈¬Ўјя_ЅЋќ »–ќ¬ ј_ћ»Ќ As Long = 2


Public Function ‘орматЌомера(номер As Long, год As Long, метод As String) As String
    Dim s As String
    If номер < 10 Then
        s = "0" & CStr(номер)
    Else
        s = CStr(номер)
    End If
    ‘орматЌомера = s & "-" & CStr(год) & "-" & метод
End Function


Public Function „ерновойЌомер(пор€дковый As Long) As String
    „ерновойЌомер = "„≈–Ќќ¬» -" &  од–абочегоћеста() & "-" & Format$(пор€дковый, "000")
End Function


Public Function Ёто„ерновик(значение As Variant) As Boolean
    Ёто„ерновик = (InStr(1, CStr(значение), "„≈–Ќќ¬» ", vbTextCompare) = 1)
End Function


' јкт Ч свой уровень учЄта и сво€ сери€ номеров: он один на работу и
' охватывает несколько объектов, у каждого из которых внутри акта свои
' номера протоколов по методам. ѕоэтому счЄтчик у акта отдельный, а не
' общий с методами.
Public Function ‘орматЌомерајкта(номер As Long, год As Long) As String
    Dim s As String
    If номер < 10 Then s = "0" & CStr(номер) Else s = CStr(номер)
    ‘орматЌомерајкта = "ј “ " & s & "-" & CStr(год)
End Function


Public Function ¬ыдатьЌомерјкта(год As Long) As String
    Dim n As Long
    n = ¬ыдатьЌомер("ј “", год)
    If n = 0 Then Exit Function
    ¬ыдатьЌомерјкта = ‘орматЌомерајкта(n, год)
End Function


' -- ¬ыдача номера -----------------------------------------------------

' «абирает следующий номер по методу и году. ¬озвращает 0, если не вышло,
' и объ€сн€ет причину Ч молча номер не пропадает.
Public Function ¬ыдатьЌомер(метод As String, год As Long) As Long
    Dim папка As String
    Dim файл—чЄтчика As String
    Dim файлЅлокировки As String
    Dim мо€ћетка As String
    Dim текущий As Long
    Dim попытка As Long

    папка = ѕапка—чЄтчиков()
    If Not ѕапка—уществует(ѕапкаќбмена()) Then
        MsgBox "ќбща€ папка недоступна:" & vbCrLf & ѕапкаќбмена() & vbCrLf & vbCrLf & _
               "Ќомер можно выдать только через неЄ Ч иначе два человека получат один и тот же. " & _
               "–аботайте с черновым номером, зарегистрируете, когда по€витс€ сеть.", _
               vbExclamation, "¬ыдача номера"
        Exit Function
    End If

    On Error GoTo —бой
    MkPath папка
    файл—чЄтчика = папка & метод & "-" & CStr(год) & ".txt"
    файлЅлокировки = папка & метод & "-" & CStr(год) & ".lock"
    Randomize
    мо€ћетка =  од–абочегоћеста() & "|" & ћетка¬ремени() & "|" & CStr(Int(Rnd() * 100000))

    For попытка = 1 To 3
        If «ахватитьЅлокировку(файлЅлокировки, мо€ћетка) Then
            текущий = ѕрочитать—чЄтчик(файл—чЄтчика)
            текущий = текущий + 1
            «аписать—чЄтчик файл—чЄтчика, текущий
            ' провер€ем, что записалось именно наше Ч файлова€ система могла подвести
            If ѕрочитать—чЄтчик(файл—чЄтчика) <> текущий Then
                —н€тьЅлокировку файлЅлокировки
                Err.Raise vbObjectError + 520, "¬ыдатьЌомер", _
                          "—чЄтчик не сохранилс€: записали " & текущий & ", прочитали другое."
            End If
            —н€тьЅлокировку файлЅлокировки
            «аписать¬–еестр¬ыданных метод, год, текущий
            —копировать—чЄтчики          ' резервна€ копи€, уровень 1
            ¬ыдатьЌомер = текущий
            Exit Function
        End If
        ѕодождать 400
    Next попытка

    MsgBox "—чЄтчик по методу " & метод & " за " & год & " год зан€т другим рабочим местом " & _
           "дольше обычного." & vbCrLf & vbCrLf & _
           "ѕодождите несколько секунд и повторите. ≈сли повтор€етс€ Ч проверьте, " & _
           "не осталс€ ли файл " & метод & "-" & год & ".lock после аварийного закрыти€.", _
           vbExclamation, "¬ыдача номера"
    Exit Function

—бой:
    On Error Resume Next
    —н€тьЅлокировку файлЅлокировки
    MsgBox "Ќе удалось выдать номер." & vbCrLf & "ќшибка: " & Err.Description, _
           vbCritical, "¬ыдача номера"
End Function


' “екущее значение счЄтчика, не забира€ номер, Ч дл€ сверки при переходе.
Public Function ѕосмотреть—чЄтчик(метод As String, год As Long) As Long
    ѕосмотреть—чЄтчик = ѕрочитать—чЄтчик(ѕапка—чЄтчиков() & метод & "-" & CStr(год) & ".txt")
End Function


Public Sub ”становить—чЄтчик(метод As String, год As Long, значение As Long)
    MkPath ѕапка—чЄтчиков()
    «аписать—чЄтчик ѕапка—чЄтчиков() & метод & "-" & CStr(год) & ".txt", значение
End Sub


'  аждый выданный номер дописываетс€ в реестр на носителе: номер, кто
' вз€л, когда. ‘айл счЄтчика хранит только последнее число, а реестр Ч
' всю историю. ѕока сети нет и книги живут порознь, это единственное
' место, где видна вс€ выданна€ нумераци€ сразу; заодно по нему
' восстанавливаетс€ счЄтчик, если файл счЄтчика потер€етс€.
Private Sub «аписать¬–еестр¬ыданных(метод As String, год As Long, номер As Long)
    Dim файл As String
    Dim f As Integer

    On Error Resume Next
    файл = ѕапка—чЄтчиков() & "¬ыданные-" & метод & "-" & CStr(год) & ".csv"
    If Not ‘айл—уществует(файл) Then
        f = FreeFile
        Open файл For Output As #f
        Print #f, "Ќомер;–абочее место;—отрудник;ƒата и врем€"
        Close #f
    End If
    f = FreeFile
    Open файл For Append As #f
    Print #f, CStr(номер) & ";" &  од–абочегоћеста() & ";" & “екущий—отрудник() & ";" & _
              Format$(Now, "dd.mm.yyyy hh:nn:ss")
    Close #f
End Sub


' -- Ѕлокировка --------------------------------------------------------

' ѕытаетс€ зан€ть счЄтчик. Ќасто€щей атомарной блокировки в VBA нет,
' поэтому после записи перечитываем файл: если в нЄм чужа€ метка Ч
' значит, нас опередили, и номер мы не берЄм. ѕлюс мастер при сборке
' свода провер€ет номера на дубли (раздел 5 “«).
Private Function «ахватитьЅлокировку(файл As String, мо€ћетка As String) As Boolean
    Dim f As Integer
    Dim ждЄм As Double

    ждЄм = Timer
    Do While ‘айл—уществует(файл)
        If Ѕлокировка”старела(файл) Then
            On Error Resume Next
            Kill файл
            On Error GoTo 0
            Exit Do
        End If
        If (Timer - ждЄм) * 1000 > ∆ƒј“№_ЅЋќ »–ќ¬ ”_ћ— Then Exit Function
        ѕодождать 200
    Loop

    On Error GoTo —бой
    f = FreeFile
    Open файл For Output As #f
    Print #f, мо€ћетка
    Close #f

    ѕодождать 120
    If ѕрочитать—троку(файл) <> мо€ћетка Then Exit Function   ' нас опередили

    «ахватитьЅлокировку = True
    Exit Function
—бой:
    On Error Resume Next
    Close #f
End Function


Private Sub —н€тьЅлокировку(файл As String)
    On Error Resume Next
    If ‘айл—уществует(файл) Then Kill файл
End Sub


'  нига могла упасть, не сн€в блокировку. „ерез две минуты считаем еЄ
' брошенной Ч иначе одна авари€ остановит выдачу номеров всем.
Private Function Ѕлокировка”старела(файл As String) As Boolean
    On Error Resume Next
    Ѕлокировка”старела = (DateDiff("n", FileDateTime(файл), Now) >= ”—“ј–≈¬Ўјя_ЅЋќ »–ќ¬ ј_ћ»Ќ)
End Function


' -- „тение и запись файлов счЄтчиков ----------------------------------

Private Function ѕрочитать—чЄтчик(файл As String) As Long
    Dim s As String
    s = Trim$(ѕрочитать—троку(файл))
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then
        Err.Raise vbObjectError + 521, "ѕрочитать—чЄтчик", _
                  "¬ файле счЄтчика " & файл & " лежит не число, а Ђ" & s & "ї."
    End If
    ѕрочитать—чЄтчик = CLng(s)
End Function


Private Sub «аписать—чЄтчик(файл As String, значение As Long)
    Dim f As Integer
    f = FreeFile
    Open файл For Output As #f
    Print #f, CStr(значение)
    Close #f
End Sub


Private Function ѕрочитать—троку(файл As String) As String
    Dim f As Integer
    Dim s As String
    If Not ‘айл—уществует(файл) Then Exit Function
    On Error GoTo —бой
    f = FreeFile
    Open файл For Input As #f
    If Not EOF(f) Then Line Input #f, s
    Close #f
    ѕрочитать—троку = s
    Exit Function
—бой:
    On Error Resume Next
    Close #f
End Function
