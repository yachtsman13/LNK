Attribute VB_Name = "modChecks"
Option Explicit

' Проверяет по листу "Специалисты", действительна ли аттестация исполнителя
' по нужному методу на дату контроля. Возвращает "" если всё в порядке,
' иначе — текст предупреждения для показа пользователю перед формированием документа.
Public Function ПроверитьАттестацию(fio As String, methodName As String, onDate As Variant) As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim checkDate As Date
    Dim found As Boolean

    If IsDate(onDate) Then
        checkDate = CDate(onDate)
    Else
        checkDate = Date
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Специалисты")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    found = False

    For r = 4 To lastRow
        If Len(Trim(ws.Cells(r, 2).Value)) > 0 Then
            If InStr(1, ws.Cells(r, 2).Value, fio, vbTextCompare) > 0 _
               And Trim(ws.Cells(r, 4).Value) = methodName Then
                found = True
                If IsDate(ws.Cells(r, 9).Value) Then
                    If CDate(ws.Cells(r, 9).Value) < checkDate Then
                        ПроверитьАттестацию = "Удостоверение «" & fio & "» по методу " & methodName & _
                            " просрочено (действительно до " & Format(ws.Cells(r, 9).Value, "dd.mm.yyyy") & ")."
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r

    If Not found Then
        ПроверитьАттестацию = "В реестре «Специалисты» не нашлась запись «" & fio & "» по методу " & _
            methodName & " — проверьте квалификацию вручную перед выдачей документа."
    End If
End Function
