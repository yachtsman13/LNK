Attribute VB_Name = "modAccessConfig"
Option Explicit
Option Compare Database

' Пути — рядом с базой данных, чтобы работало и на общей сетевой папке.
Public Function TemplatesFolder() As String
    TemplatesFolder = CurrentProject.Path & "\Шаблоны\"
End Function

Public Function OutputFolder() As String
    Dim p As String
    p = CurrentProject.Path & "\Готовые\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    OutputFolder = p
End Function

Public Function TemplateFileFor(methodName As String) As String
    Select Case methodName
        Case "ВИК": TemplateFileFor = "Акт_ВИК.docx"
        Case "УЗТ": TemplateFileFor = "Протокол_УЗТ.docx"
        Case "УЗД": TemplateFileFor = "Заключение_УЗД.docx"
        Case "ТВ":  TemplateFileFor = "Протокол_ТВ.docx"
        Case "АЭ":  TemplateFileFor = "Заключение_АЭ.docx"
        Case "ТК":  TemplateFileFor = "Протокол_ТК.docx"
        Case "ВД":  TemplateFileFor = "Протокол_ВД.docx"
        Case Else:  TemplateFileFor = ""
    End Select
End Function

Public Sub MkPath(p As String)
    Dim parts() As String, cur As String, i As Long
    parts = Split(p, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        If parts(i) = "" Then Exit For
        cur = cur & "\" & parts(i)
        If Dir(cur, vbDirectory) = "" Then MkDir cur
    Next i
End Sub

Public Function ОчиститьДляИмениФайла(s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|", Chr(10), Chr(13))
    ОчиститьДляИмениФайла = Nz(s, "")
    For i = LBound(bad) To UBound(bad)
        ОчиститьДляИмениФайла = Replace(ОчиститьДляИмениФайла, bad(i), " ")
    Next i
    Do While InStr(ОчиститьДляИмениФайла, "  ") > 0
        ОчиститьДляИмениФайла = Replace(ОчиститьДляИмениФайла, "  ", " ")
    Loop
    ОчиститьДляИмениФайла = Trim(ОчиститьДляИмениФайла)
    If Len(ОчиститьДляИмениФайла) > 60 Then ОчиститьДляИмениФайла = Left(ОчиститьДляИмениФайла, 60)
End Function

' Безопасный (устойчивый к параллельной работе нескольких пользователей) следующий номер
' в серии метод+год — читает максимум прямо из таблицы через DMax, а не по кэшированной
' формуле, как это было в Excel. Полной защиты от гонки два человека кликнули одновременно
' это не даёт (для этого нужна транзакция уровня SQL Server), но на практике при 3
' пользователях и ручном нажатии кнопки коллизия практически исключена, в отличие от Excel,
' где COUNTIFS вообще не видит несохранённые чужие правки до пересчёта книги.
Public Function СледующийНомерВСерии(methodName As String, yearNum As Integer) As Long
    Dim maxNum As Variant
    maxNum = DMax("НомерВСерии", "tblЖурнал", "Метод='" & methodName & "' AND Год=" & yearNum)
    If IsNull(maxNum) Then
        СледующийНомерВСерии = 1
    Else
        СледующийНомерВСерии = maxNum + 1
    End If
End Function
