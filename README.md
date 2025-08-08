# VBA
Sort for rep
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim c As Long
    Dim rowRange As Range
    Dim isEmpty As Boolean
    Dim checkRange As Range
    
    Set ws = Me
    
    ' Отслеживаем изменения в диапазоне B3:N последней строки
    Set checkRange = Intersect(Target, ws.Range("B3:N" & ws.Rows.Count))
    If checkRange Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Находим последнюю заполненную строку по столбцу B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then lastRow = 3
    
    For r = 3 To lastRow
        Set rowRange = ws.Range("B" & r & ":N" & r)
        
        isEmpty = True
        ' Проверяем все ячейки строки в диапазоне B:N
        For c = 2 To 14 ' B=2, N=14
            If Trim(ws.Cells(r, c).Value) <> "" Then
                isEmpty = False
                Exit For
            End If
        Next c
        
        If isEmpty Then
            rowRange.Borders.LineStyle = xlNone
        Else
            With rowRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1 ' черные границы
            End With
        End If
    Next r
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
