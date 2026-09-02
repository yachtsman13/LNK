Attribute VB_Name = "modMedia"
Option Explicit

' =======================================================================
'  Назначенный носитель для обмена.
'
'  Флешка опознаётся по серийному номеру тома — он присваивается при
'  форматировании и в отличие от буквы диска и имени папки не переносится
'  вместе с файлами. Скопировать папку на другую флешку можно, серийный
'  номер — нет.
'
'  Честно про границы: без прав администратора запретить подключение
'  других флешек нельзя, это делается политиками Windows. Проверка
'  работает на уровне программы — она защищает от ошибки («я думал, эта
'  тоже подойдёт»), а не от намерения. От подмены файла защищает
'  контрольная сумма, а не привязка к носителю.
' =======================================================================

Private Const DRIVE_REMOVABLE As Long = 1


' Путь к штатной флешке, если она подключена; иначе "" и внятное объяснение.
Public Function НайтиШтатныйНоситель(Optional молча As Boolean = False) As String
    Dim fso As Object
    Dim д As Object
    Dim разрешённые As String
    Dim серийный As String
    Dim найденные As String
    Dim сколькоСъёмных As Long

    разрешённые = СписокРазрешённых()
    If разрешённые = "" Then
        If Not молча Then
            MsgBox "Не задан ни один разрешённый носитель." & vbCrLf & vbCrLf & _
                   "Мастер должен подключить штатную флешку и выполнить " & _
                   "«Зарегистрировать носитель». До этого обмен через флешку не работает — " & _
                   "иначе данные уедут на случайный носитель.", _
                   vbExclamation, "Носитель обмена"
        End If
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each д In fso.Drives
        If д.DriveType = DRIVE_REMOVABLE And д.IsReady Then
            сколькоСъёмных = сколькоСъёмных + 1
            серийный = ФорматСерийного(д.SerialNumber)
            найденные = найденные & IIf(найденные = "", "", ", ") & д.DriveLetter & ": " & серийный
            If InStr(1, разрешённые, серийный, vbTextCompare) > 0 Then
                НайтиШтатныйНоситель = д.DriveLetter & ":\"
                Exit Function
            End If
        End If
    Next д

    If Not молча Then
        If сколькоСъёмных = 0 Then
            MsgBox "Съёмный носитель не подключён." & vbCrLf & vbCrLf & _
                   "Вставьте штатную флешку ЛНК-Обмен и повторите.", _
                   vbExclamation, "Носитель обмена"
        Else
            MsgBox "Обмен разрешён только со штатным носителем." & vbCrLf & vbCrLf & _
                   "Разрешены серийные номера: " & разрешённые & vbCrLf & _
                   "Подключено: " & найденные & vbCrLf & vbCrLf & _
                   "Возьмите штатную флешку у мастера. Если эта флешка и есть штатная, " & _
                   "её нужно зарегистрировать — её могли переформатировать, тогда " & _
                   "серийный номер сменился.", _
                   vbExclamation, "Носитель обмена"
        End If
    End If
End Function


' Регистрация носителя — действие мастера. Прежний номер не затирается,
' а уходит в примечание с датой: замену носителя потом видно.
Public Sub ЗарегистрироватьНоситель()
    Dim fso As Object
    Dim д As Object
    Dim список As String
    Dim выбор As String
    Dim буква As String
    Dim серийный As String

    If Not ЭтоМастер() Then
        If MsgBox("Это рабочее место не отмечено как мастер." & vbCrLf & _
                  "Регистрация носителя — действие мастера: список разрешённых " & _
                  "флешек общий для всех." & vbCrLf & vbCrLf & "Всё равно продолжить?", _
                  vbYesNo + vbExclamation, "Носитель обмена") = vbNo Then Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each д In fso.Drives
        If д.DriveType = DRIVE_REMOVABLE And д.IsReady Then
            список = список & vbCrLf & "  " & д.DriveLetter & ":  " & _
                     ФорматСерийного(д.SerialNumber) & _
                     IIf(д.VolumeName = "", "", "  «" & д.VolumeName & "»")
        End If
    Next д

    If список = "" Then
        MsgBox "Съёмный носитель не подключён.", vbExclamation, "Носитель обмена"
        Exit Sub
    End If

    выбор = InputBox("Подключённые съёмные носители:" & vbCrLf & список & vbCrLf & vbCrLf & _
                     "Введите букву диска, который делаем штатным:", "Носитель обмена")
    буква = UCase$(Left$(Trim$(выбор), 1))
    If буква = "" Then Exit Sub

    On Error GoTo Сбой
    Set д = fso.GetDrive(буква & ":")
    If д.DriveType <> DRIVE_REMOVABLE Then
        MsgBox "Диск " & буква & ": не съёмный. Регистрировать под обмен нужно флешку.", _
               vbExclamation, "Носитель обмена"
        Exit Sub
    End If
    серийный = ФорматСерийного(д.SerialNumber)

    If InStr(1, СписокРазрешённых(), серийный, vbTextCompare) > 0 Then
        MsgBox "Этот носитель уже зарегистрирован (" & серийный & ").", _
               vbInformation, "Носитель обмена"
        Exit Sub
    End If

    ДобавитьРазрешённый серийный
    ЗаписатьФайлНосителей
    ПодписатьНоситель буква & ":\", серийный

    MsgBox "Носитель зарегистрирован: " & буква & ": " & серийный & vbCrLf & vbCrLf & _
           "Всего разрешённых носителей: " & (UBound(Split(СписокРазрешённых(), ",")) + 1) & _
           ". Их должно быть два — основной и резервный у мастера: " & _
           "потерянная или умершая флешка иначе останавливает обмен целиком.", _
           vbInformation, "Носитель обмена"
    Exit Sub
Сбой:
    MsgBox "Не удалось зарегистрировать носитель." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Носитель обмена"
End Sub


' Серийный номер тома в привычном виде 1A2B-3C4D.
Private Function ФорматСерийного(значение As Long) As String
    Dim h As String
    h = Right$("00000000" & Hex$(значение), 8)
    ФорматСерийного = Left$(h, 4) & "-" & Right$(h, 4)
End Function


Private Function СписокРазрешённых() As String
    Dim s As String
    s = Настройка("Серийные номера разрешённых флешек")
    If s = "" Then s = ПрочитатьФайлНосителей()
    СписокРазрешённых = s
End Function


Private Sub ДобавитьРазрешённый(серийный As String)
    Dim s As String
    s = Настройка("Серийные номера разрешённых флешек")
    If s = "" Then
        s = серийный
    Else
        s = s & ", " & серийный
    End If
    ЗаписатьНастройку "Серийные номера разрешённых флешек", s
End Sub


' Список дублируется в общей папке: рабочие места берут его оттуда,
' а не выспрашивают у мастера.
Private Sub ЗаписатьФайлНосителей()
    Dim файл As String
    Dim f As Integer

    On Error Resume Next
    MkPath ПапкаОбмена() & "Версия"
    файл = ПапкаОбмена() & "Версия" & Application.PathSeparator & "носители.txt"
    f = FreeFile
    Open файл For Output As #f
    Print #f, Настройка("Серийные номера разрешённых флешек")
    Print #f, "' записано " & МеткаВремени() & " с " & КодРабочегоМеста()
    Close #f
End Sub


Private Function ПрочитатьФайлНосителей() As String
    Dim файл As String
    Dim f As Integer
    Dim s As String

    файл = ПапкаОбмена() & "Версия" & Application.PathSeparator & "носители.txt"
    If Not ФайлСуществует(файл) Then Exit Function
    On Error Resume Next
    f = FreeFile
    Open файл For Input As #f
    If Not EOF(f) Then Line Input #f, s
    Close #f
    ПрочитатьФайлНосителей = Trim$(s)
End Function


' Записка в корне флешки: чтобы найденную флешку вернули, а не
' отформатировали под музыку.
Private Sub ПодписатьНоситель(корень As String, серийный As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open корень & "ПРОЧТИ.txt" For Output As #f
    Print #f, "Это служебный носитель лаборатории неразрушающего контроля."
    Print #f, "Серийный номер тома: " & серийный
    Print #f, "Не форматировать. Нашедшего просим вернуть в ЛНК."
    Print #f, "Зарегистрирован " & Format$(Date, "dd.mm.yyyy") & ", " & ТекущийСотрудник()
    Close #f
End Sub
