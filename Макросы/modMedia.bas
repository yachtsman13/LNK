Attribute VB_Name = "modMedia"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
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
' ═══════════════════════════════════════════════════════════════════════

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
    If разрешённые = "" Then разрешённые = ПодхватитьСписокСНосителя()
    If разрешённые = "" Then
        If Not молча Then
            MsgBox "Не задан ни один разрешённый носитель." & vbCrLf & vbCrLf & _
                   "Подключите штатную флешку и выполните «Зарегистрировать носитель», " & _
                   "а если её ещё не готовили — «Подготовить флешку для обмена». " & _
                   "До этого обмен через флешку не работает: иначе данные уедут " & _
                   "на случайный носитель.", _
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


' Свежее рабочее место списка ещё не знает. Обходим съёмные диски: если
' на каком-то уже лежит ЛНК_Обмен со списком, и этот диск в списке есть —
' значит, флешка та самая, список забираем себе. Обращений к ПапкаОбмена
' здесь нет, поэтому петли не возникает.
Private Function ПодхватитьСписокСНосителя() As String
    Dim fso As Object
    Dim д As Object
    Dim список As String
    Dim серийный As String

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each д In fso.Drives
        If д.DriveType = DRIVE_REMOVABLE And д.IsReady Then
            список = ПрочитатьФайлНосителей(д.DriveLetter & ":\ЛНК_Обмен\")
            If список <> "" Then
                серийный = ФорматСерийного(д.SerialNumber)
                If InStr(1, список, серийный, vbTextCompare) > 0 Then
                    ЗаписатьНастройку "Серийные номера разрешённых флешек", список
                    ПодхватитьСписокСНосителя = список
                    Exit Function
                End If
            End If
        End If
    Next д
End Function


' Подготовка флешки к работе: структура папок, регистрация серийного
' номера, записка в корне. Одна кнопка — и обмен работает без сети.
Public Sub ПодготовитьФлешку()
    Dim fso As Object
    Dim д As Object
    Dim список As String
    Dim выбор As String
    Dim буква As String
    Dim серийный As String
    Dim корень As String
    Dim обмен As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each д In fso.Drives
        If д.DriveType = DRIVE_REMOVABLE And д.IsReady Then
            список = список & vbCrLf & "  " & д.DriveLetter & ":  " & _
                     ФорматСерийного(д.SerialNumber) & _
                     IIf(д.VolumeName = "", "", "  «" & д.VolumeName & "»") & _
                     "  свободно " & Format$(д.FreeSpace / 1073741824, "0.0") & " ГБ"
        End If
    Next д

    If список = "" Then
        MsgBox "Съёмный носитель не подключён." & vbCrLf & vbCrLf & _
               "Вставьте флешку, которая станет штатной для обмена, и повторите.", _
               vbExclamation, "Подготовка флешки"
        Exit Sub
    End If

    выбор = InputBox("Подключённые съёмные носители:" & vbCrLf & список & vbCrLf & vbCrLf & _
                     "Введите букву диска, который готовим под обмен." & vbCrLf & _
                     "Данные на флешке не пострадают — будет создана папка ЛНК_Обмен.", _
                     "Подготовка флешки")
    буква = UCase$(Left$(Trim$(выбор), 1))
    If буква = "" Then Exit Sub

    On Error GoTo Сбой
    Set д = fso.GetDrive(буква & ":")
    If д.DriveType <> DRIVE_REMOVABLE Then
        MsgBox "Диск " & буква & ": не съёмный. Под обмен готовится флешка.", _
               vbExclamation, "Подготовка флешки"
        Exit Sub
    End If

    корень = буква & ":\"
    обмен = корень & "ЛНК_Обмен\"
    MkPath обмен & "Счётчики"
    MkPath обмен & "Входящие"
    MkPath обмен & "Обработанные"
    MkPath обмен & "Свод\История"
    MkPath обмен & "Конфликты"
    MkPath обмен & "Версия"

    серийный = ФорматСерийного(д.SerialNumber)
    If InStr(1, СписокРазрешённых(), серийный, vbTextCompare) = 0 Then
        ДобавитьРазрешённый серийный
    End If
    ЗаписатьФайлНосителей обмен
    ПодписатьНоситель корень, серийный

    ЗаписатьНастройку "Общая папка обмена", ""
    СброситьПутьОбмена

    MsgBox "Флешка готова: " & буква & ": " & серийный & vbCrLf & _
           "Создана структура " & обмен & vbCrLf & vbCrLf & _
           "Теперь программа находит её сама, по серийному номеру тома — " & _
           "буква диска может меняться, это не помешает." & vbCrLf & vbCrLf & _
           "Дальше: на каждом рабочем месте один раз выполните " & _
           "«Зарегистрировать носитель», подключив эту же флешку. " & _
           "Появится сеть — впишете путь в настройках, и она станет " & _
           "использоваться вместо флешки автоматически.", _
           vbInformation, "Подготовка флешки"
    Exit Sub
Сбой:
    MsgBox "Не удалось подготовить флешку." & vbCrLf & "Ошибка: " & Err.Description, _
           vbCritical, "Подготовка флешки"
End Sub


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
    ЗаписатьФайлНосителей буква & ":\ЛНК_Обмен\"
    ПодписатьНоситель буква & ":\", серийный
    СброситьПутьОбмена

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


' Только из настроек. Раньше отсюда читался ещё и файл в общей папке —
' и получалась петля: чтобы найти общую папку, надо знать разрешённые
' носители, а чтобы прочитать список носителей, надо знать общую папку.
' Список с флешки подхватывается иначе — см. ПодхватитьСписокСНосителя.
Private Function СписокРазрешённых() As String
    СписокРазрешённых = Настройка("Серийные номера разрешённых флешек")
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
' Путь передаётся явно: обращаться отсюда к ПапкаОбмена нельзя, она
' сама зовёт поиск носителя.
Private Sub ЗаписатьФайлНосителей(папкаОбмена As String)
    Dim файл As String
    Dim f As Integer

    On Error Resume Next
    If папкаОбмена = "" Then Exit Sub
    MkPath папкаОбмена & "Версия"
    файл = папкаОбмена & "Версия" & Application.PathSeparator & "носители.txt"
    f = FreeFile
    Open файл For Output As #f
    Print #f, Настройка("Серийные номера разрешённых флешек")
    Print #f, "' записано " & МеткаВремени() & " с " & КодРабочегоМеста()
    Close #f
End Sub


Private Function ПрочитатьФайлНосителей(папкаОбмена As String) As String
    Dim файл As String
    Dim f As Integer
    Dim s As String

    If папкаОбмена = "" Then Exit Function
    файл = папкаОбмена & "Версия" & Application.PathSeparator & "носители.txt"
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
