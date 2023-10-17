Sub ImportAndProcessCSV()
    Dim ws_source As Worksheet
    Dim ws_result As Worksheet
    Dim filePath As String
    Dim lastRow As Long
    Dim newWorkbook As Workbook
    Dim saveFilePath As Variant
    
    ' Open the file dialog to select a CSV file
    Dim fd2 As FileDialog
    Set fd2 = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd2
        .AllowMultiSelect = False
        .Filters.Add "Text Files", "*.csv", 1
        .FilterIndex = 1
        .Filters.Add "text files| ", "*.csv"
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Open the CSV file and set ws_source as the active worksheet
    Workbooks.Open filePath
    Set ws_source = ActiveWorkbook.ActiveSheet
    
    ' Get the number of rows in ws_source
    lastRow = ws_source.Cells(ws_source.Rows.Count, "B").End(xlUp).Row

    ' Create a new workbook for ws_result
    Set newWorkbook = Workbooks.Add
    Set ws_result = newWorkbook.Sheets(1)
    
     'Copy row 1 from ThisWorkbook's sheet("copy") to row 1 in ws_result
    ThisWorkbook.Sheets("copy").Rows(1).Copy Destination:=ws_result.Rows(1)
    
    ' Copy B2 from ws_source to C2 in ws_result
    'ws_source.Range("B2").Copy Destination:=ws_result.Range("C2")
    
    For i = 2 To lastRow
    ws_result.Range("F" & i) = ws_source.Range("B" & i)
    ws_result.Range("L" & i) = ws_source.Range("E" & i)
    ws_result.Range("Q" & i) = ws_source.Range("D" & i)
    ws_result.Range("R" & i) = ws_source.Range("E" & i) & ", " & ws_source.Range("F" & i) & ", " & ws_source.Range("G" & i)
    ws_result.Range("S" & i) = ws_source.Range("H" & i)
    ws_result.Range("T" & i) = ws_source.Range("I" & i)
    ws_result.Range("U" & i) = ws_source.Range("J" & i)
    ws_result.Range("V" & i) = ws_source.Range("K" & i)
    ws_result.Range("W" & i) = ws_source.Range("L" & i)
    ws_result.Range("X" & i) = ws_source.Range("M" & i)
    ws_result.Range("Y" & i) = ws_source.Range("N" & i)
    ws_result.Range("Z" & i) = ws_source.Range("O" & i)
    ws_result.Range("AA" & i) = ws_source.Range("P" & i)
    ws_result.Range("AB" & i) = ws_source.Range("C" & i)
    ws_result.Range("AC" & i) = ws_source.Range("Q" & i)
    ws_result.Range("AD" & i) = ws_source.Range("R" & i)
    ws_result.Range("AE" & i) = ws_source.Range("S" & i)
    ws_result.Range("AF" & i) = ws_source.Range("T" & i)
    ws_result.Range("AG" & i) = ws_source.Range("U" & i)
    ws_result.Range("AH" & i) = ws_source.Range("V" & i)
    ws_result.Range("AH" & i).NumberFormat = "dd/mm/yyyy"
    ws_result.Range("AI" & i) = ws_source.Range("W" & i)
    ws_result.Range("AJ" & i) = ws_source.Range("X" & i)
    ws_result.Range("AO" & i) = ws_source.Range("Y" & i)
    ws_result.Range("AP" & i) = ws_source.Range("AC" & i)
    ws_result.Range("AQ" & i) = ws_source.Range("AD" & i)
    
    'ws_result.Range("BA" & i) = ws_source.Range("" & i)
    'ws_result.Range("BB" & i) = ws_source.Range("" & i)
    'ws_result.Range("BC" & i) = ws_source.Range("" & i)
    'ws_result.Range("BD" & i) = ws_source.Range("E" & i)
    ws_result.Range("BE" & i) = ws_source.Range("E" & i)
    ws_result.Range("BF" & i) = ws_source.Range("F" & i)
    ws_result.Range("BG" & i) = ws_source.Range("G" & i)
    ws_result.Range("BH" & i) = ws_source.Range("U" & i)
    ws_result.Range("BI" & i) = ws_source.Range("V" & i)
    ws_result.Range("BI" & i).NumberFormat = "dd/mm/yyyy"
    ws_result.Range("BJ" & i) = ws_source.Range("X" & i)
    ws_result.Range("BK" & i) = ws_source.Range("Z" & i)
    ws_result.Range("BL" & i) = ws_source.Range("AA" & i)
    ws_result.Range("BM" & i) = ws_source.Range("AB" & i)
    ws_result.Range("BN" & i) = ws_source.Range("AE" & i)
    ws_result.Range("BO" & i) = ws_source.Range("AF" & i)
    ws_result.Range("BP" & i) = ws_source.Range("AG" & i)
    
    Next i
    
    ' Save ws_result as a new Excel file
    saveFilePath = Application.GetSaveAsFilename(InitialFileName:="result", FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    ' Exit if no save file path is provided
    If saveFilePath = False Then
        newWorkbook.Close SaveChanges:=False
        ws_source.Parent.Close SaveChanges:=False
        Exit Sub
    End If
    
    ws_result.Parent.SaveAs saveFilePath
    
    ' Close the workbooks
    newWorkbook.Close SaveChanges:=False
    ws_source.Parent.Close SaveChanges:=False
    
    MsgBox "CSV file imported and processed successfully!"
End Sub
