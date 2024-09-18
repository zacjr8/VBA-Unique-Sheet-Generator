' VBA Code to seperate unique values into different sheets

Sub SeparateDuplicatesIntoSheets()
    Dim ws As Worksheet
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim newWs As Worksheet
    Dim dataRange As Range
    Dim value As Variant
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet and range
    Set ws = ThisWorkbook.Sheets("All Sections") ' Change to your sheet name
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set dataRange = ws.Range("A1:A" & lastRow)

    ' Create a collection to hold unique values
    Set uniqueValues = New Collection

    ' Collect unique values
    On Error Resume Next
    For Each cell In dataRange
        uniqueValues.Add cell.value, CStr(cell.value)
    Next cell
    On Error GoTo 0

    ' Loop through unique values and create new sheets
    For i = 1 To uniqueValues.Count
        value = uniqueValues(i)
        
        ' Skip empty values
        If Trim(value) <> "" Then
            ' Create a new sheet for each unique value
            Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newWs.Name = value
            
            ' Copy the header if needed (adjust as per your header row)
            ws.Rows(1).Copy Destination:=newWs.Rows(1)

            ' Copy rows with the corresponding duplicate values
            Dim rowCount As Long
            rowCount = 2 ' Start pasting from the second row in the new sheet
            
            For Each cell In dataRange
                If cell.value = value Then
                    ws.Rows(cell.Row).Copy Destination:=newWs.Rows(rowCount)
                    rowCount = rowCount + 1
                End If
            Next cell
            
            ' Match the column widths
            Dim col As Integer
            For col = 1 To ws.Columns.Count
                If ws.Cells(1, col).value <> "" Then
                    newWs.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth
                End If
            Next col
        End If
    Next i

    MsgBox "Sheets created for unique values!", vbInformation
End Sub

