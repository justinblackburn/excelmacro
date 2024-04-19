Sub HighlightRows()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim dict As Object, lastRow1 As Long, lastRow2 As Long, i As Long
    Dim colIndex1 As Long, colIndex2 As Long

    ' Assign worksheets
    Set ws1 = ThisWorkbook.Sheets("Spreadsheet1")
    Set ws2 = ThisWorkbook.Sheets("Spreadsheet2")

    ' Set the column index for comparison (17 for Q, 16 for P, 26 for Z, etc.)
    colIndex1 = 17 ' Column Q in Spreadsheet 1
    colIndex2 = 17 ' Column Q in Spreadsheet 2, change as necessary

    ' Create a dictionary
    Set dict = CreateObject("Scripting.Dictionary")

    ' Determine the last row in each sheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, colIndex1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, colIndex2).End(xlUp).Row

    ' Loop through Spreadsheet 1 and add items to dictionary
    For i = 1 To lastRow1
        If Not dict.Exists(ws1.Cells(i, colIndex1).Value) Then
            dict.Add ws1.Cells(i, colIndex1).Value, Array(False, i)
        End If
    Next i

    ' Loop through Spreadsheet 2 and update dictionary / color Spreadsheet 2 rows
    For i = 1 To lastRow2
        If dict.Exists(ws2.Cells(i, colIndex2).Value) Then
            ' Both have the item: color row in Spreadsheet 2 yellow
            ws2.Rows(i).Interior.Color = vbYellow
            ' Update the dictionary to indicate the item is found in both sheets
            dict(ws2.Cells(i, colIndex2).Value)(0) = True
        Else
            ' Only Spreadsheet 2 has the item: color row blue
            ws2.Rows(i).Interior.Color = vbBlue
        End If
    Next i

    ' Loop through the dictionary to color Spreadsheet 1 rows green if only in Spreadsheet 1
    For Each key In dict.Keys
        If Not dict(key)(0) Then ' item was not found in Spreadsheet 2
            ws1.Rows(dict(key)(1)).Interior.Color = vbGreen
        End If
    Next key

    ' Call the date checking function after running main comparison
    Call CheckDates(ws2, 20) ' Change 20 to the actual column index for the date column
End Sub

Sub CheckDates(ws As Worksheet, dateCol As Long)
    Dim lastRow As Long, i As Long
    Dim currentDate As Date

    currentDate = Date
    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row

    For i = 1 To lastRow
        If ws.Cells(i, dateCol).Value <> "" Then
            If ws.Cells(i, dateCol).Value <= (currentDate + 14) And ws.Cells(i, dateCol).Value >= currentDate Then
                ws.Cells(i, dateCol).Interior.Color = vbRed
            End If
        End If
    Next i
End Sub
