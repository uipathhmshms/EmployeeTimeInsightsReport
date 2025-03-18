Sub AddStyleToSheet()
    Dim ws As Worksheet
    Dim firstRowRange As Range
    Dim tableStart As Range
    Dim intTableWidth As Integer

    ' Use the active sheet as the target
    Set ws = ActiveSheet
    ' Assume table starts at cell A1
    Set tableStart = ws.Range("A1")
    ' Determine table width from the used range
    intTableWidth = ws.UsedRange.Columns.Count
    ' Set the first row range based on table start and width
    Set firstRowRange = tableStart.Resize(1, intTableWidth)
    
    ' Apply right-to-left setting on all sheets, freeze first row, auto-fit columns
    SetAllSheetsRTL
    FreezeFirstRow
    AutoFitColumns firstRowRange
    ' Color the first column (except header)
    ColorFirstColumn ws
	
	ColorLowPercentagesInColumnI
	
	ColorHighVacationHoursInColumnO
	
	ColorAllNumbersBlue
	
	ColorHeaderRow firstRowRange

	Align_A_C_E_ColumnsText

	CenterTextInFirstRow firstRowRange

	FormatTotalRows
	
	AddGrandTotalRow
	
	FormatBigNumbersWithCommas
End Sub

Sub SetAllSheetsRTL()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.DisplayRightToLeft = True
    Next ws
End Sub

Sub FreezeFirstRow()
    ' Freeze the first row of the active window
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Sub AutoFitColumns(firstRowRange As Range)
    ' Auto-fit columns based on the content of the first row
    firstRowRange.EntireColumn.AutoFit
End Sub

Sub ColorFirstColumn(ws As Worksheet)
    Dim lastRow As Long
    ' Determine the last used row on the sheet
    lastRow = ws.UsedRange.Rows.Count
    ' Color the first column (excluding header in row 1) with a light blue background
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).Interior.Color = RGB(68,179,225)
End Sub

Sub ColorHeaderRow(firstRowRange As Range)
    firstRowRange.Interior.Color = RGB(192, 230, 245)
End Sub

Sub Align_A_C_E_ColumnsText()
    ' Get the used range in the active sheet
    Dim usedRange As Range
    Set usedRange = ActiveSheet.UsedRange
 
	' Center align columns A and C (except the header row)
    Dim i As Long
    For i = 2 To usedRange.Rows.Count
        Cells(i, 1).HorizontalAlignment = xlCenter  ' Column A
        Cells(i, 3).HorizontalAlignment = xlCenter  ' Column C
		Cells(i, 5).HorizontalAlignment = xlCenter  ' Column E
    Next i
End Sub

Sub CenterTextInFirstRow(firstRowRange As Range)
    firstRowRange.HorizontalAlignment = xlCenter
End Sub

Sub FormatTotalRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
	Dim lastCol As Long

    ' Set the active sheet
    Set ws = ActiveSheet

    ' Get the last used row in column C
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
	lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Find last used column
	
    ' Loop through all rows in column G
    For i = 2 To lastRow
        ' Check if the month=0 - meaning thats a total row
        If ws.Cells(i, 5).Value="" Then
            ' Color the entire row grey
			ws.Range(ws.Cells(i, 2), ws.Cells(i, lastCol)).Interior.Color = RGB(192, 230, 245)

            ' Write "total" in Column C
            ws.Cells(i, 3).Value = "Total"
            ' Delete the month number (Column G)
            ws.Cells(i, 5).Value = ""
        End If
    Next i
End Sub

Sub AddGrandTotalRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim i As Long
    Dim hasNumericData As Boolean
    
    ' Set the active sheet
    Set ws = ActiveSheet
    
    ' Find last used row and column
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Add a row for the grand total
    lastRow = lastRow + 1
    
    ' Label for Grand Total
    ws.Cells(lastRow, 3).Value = "Grand Total"
    
    ' Format the grand total row
    ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, lastCol)).Interior.Color = RGB(68,179,225)
    ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, lastCol)).Font.Bold = True
    
    ' Calculate and add grand totals for each numeric column
    For col = 6 To lastCol
        ' Check if column contains numeric data
        hasNumericData = False
        For i = 2 To lastRow - 1
            If IsNumeric(ws.Cells(i, col).Value) Then
                hasNumericData = True
                Exit For
            End If
        Next i
        
        ' If column has numeric data, calculate grand total
        If hasNumericData Then
            ' Create a SUM formula that excludes header row
            ws.Cells(lastRow, col).Formula = "=SUM(" & ws.Cells(2, col).Address & ":" & ws.Cells(lastRow - 1, col).Address & ")"
        End If
    Next col
End Sub

' Colors all numeric values in the table blue
Sub ColorAllNumbersBlue()
    Dim ws As Worksheet
    Dim usedRange As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Set the active sheet
    Set ws = ActiveSheet
    
    ' Get the used range of the sheet
    Set usedRange = ws.UsedRange
    
    ' Determine the last row and column
    lastRow = usedRange.Rows.Count
    lastCol = usedRange.Columns.Count
    
    ' Loop through each cell in the used range, starting from row 2 (skip header)
    For Each cell In usedRange.Cells
        If cell.Row > 1 Then ' Skip the header row
            ' Skip column B (2nd column)
            If cell.Column <> 2 Then
                ' Check if the cell contains a numeric value
                If IsNumeric(cell.Value) Or (IsNumeric(Replace(cell.Text, "%", "")) And InStr(cell.Text, "%") > 0) Then
                    ' Set font color to blue
                    cell.Font.Color = vbBlue
                    Debug.Print "Row " & cell.Row & ", Column " & cell.Column & " colored blue (Value = '" & cell.Text & "')"
                End If
            Else
                Debug.Print "Row " & cell.Row & ", Column " & cell.Column & " skipped (Column B)"
            End If
        End If
    Next cell
End Sub

' Colors values less than 100% in column I red
Sub ColorLowPercentagesInColumnI()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim numericValue As Double
    Dim isNumericValue As Boolean
    
    ' Set the active sheet
    Set ws = ActiveSheet
    
    ' Find the last used row in column I (9th column)
    lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    
    ' Loop through all rows in column I, starting from row 2 (skip header)
    For i = 2 To lastRow
        ' Get the cell value as a string and remove leading/trailing spaces
        cellValue = Trim(ws.Cells(i, 9).Text) ' Use .Text to get the displayed value
        
        ' Reset flags and values
        isNumericValue = False
        numericValue = 0
        
        ' Check if the cell is not empty and contains a percentage
        If cellValue <> "" And InStr(cellValue, "%") > 0 Then
            ' Remove the % sign
            cellValue = Replace(cellValue, "%", "")
            cellValue = Trim(cellValue) ' Trim again after removing %
            
            ' Replace any potential comma decimal separator with a dot
            cellValue = Replace(cellValue, ",", ".")
            
            ' Try to convert the string to a number
            On Error Resume Next
            numericValue = CDbl(cellValue)
            If Err.Number = 0 Then
                isNumericValue = True
            Else
                Debug.Print "Conversion Error at Row " & i & ": Value = '" & ws.Cells(i, 9).Text & "', Cleaned = '" & cellValue & "'"
            End If
            On Error GoTo 0 ' Reset error handling
            
            ' Check if the value is a valid number and less than 100
            If isNumericValue And numericValue > 0 And numericValue < 100 Then
                ' Set the background color to red
                ws.Cells(i, 9).Interior.Color = RGB(255, 192, 203)
                Debug.Print "Row " & i & " colored red (Value < 100)"
            End If
        Else
            ' Handle non-percentage cells (e.g., empty or total rows)
            ws.Cells(i, 9).Font.Color = vbBlack
            Debug.Print "Row " & i & " skipped (No % or empty): '" & cellValue & "'"
        End If
    Next i
End Sub

Sub FormatBigNumbersWithCommas()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    Set ws = ActiveSheet
    Set rng = ws.UsedRange
    
    For Each cell In rng
        ' Skip column B (2nd column)
        If cell.Column <> 2 Then
            ' Check if the cell contains a numeric value and is not empty
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                ' Include percentage values; no need to skip them
                ' Check if the absolute value is >= 1000
                If Abs(cell.Value) >= 1000 Then
                    ' Check if the cell has a percentage format
                    If InStr(cell.NumberFormat, "%") > 0 Then
                        ' Preserve the percentage format but add commas
                        cell.NumberFormat = "#,##0\%"
                    Else
                        ' Apply comma formatting for non-percentage numbers
                        cell.NumberFormat = "#,##0"
                    End If
                    Debug.Print "Row " & cell.Row & ", Column " & cell.Column & " formatted with commas (Value = '" & cell.Text & "')"
                End If
            End If
        Else
            Debug.Print "Row " & cell.Row & ", Column " & cell.Column & " skipped (Column B)"
        End If
    Next cell
End Sub

' colors values greater than 90 in column N red
Sub ColorHighVacationHoursInColumnO()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim numericValue As Double
    Dim isNumericValue As Boolean
    
    ' Set the active sheet
    Set ws = ActiveSheet
    
    ' Find the last used row in column N (14th column)
    lastRow = ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
    
    ' Loop through all rows in column L, starting from row 2 (skip header)
    For i = 2 To lastRow
        ' Skip the grand total row by checking if column C (3rd column) contains "Grand Total"
        If ws.Cells(i, 3).Value = "Grand Total" Then
            Debug.Print "Row " & i & " skipped (Grand Total row)"
            GoTo NextRow
        End If
        
        ' Get the cell value as a string and remove leading/trailing spaces
        cellValue = Trim(ws.Cells(i, 15).Text) ' Use .Text to get the displayed value
        
        ' Reset flags and values
        isNumericValue = False
        numericValue = 0
        
        ' Check if the cell is not empty
        If cellValue <> "" Then
            ' Try to convert the string to a number
            On Error Resume Next
            numericValue = CDbl(cellValue)
            If Err.Number = 0 Then
                isNumericValue = True
            Else
                Debug.Print "Conversion Error at Row " & i & ": Value = '" & ws.Cells(i, 12).Text & "', Cleaned = '" & cellValue & "'"
            End If
            On Error GoTo 0 ' Reset error handling
            
            ' Check if the value is a valid number and greater than 90
            If isNumericValue And numericValue > 90 Then
                ' Set the background color to red
				ws.Cells(i, 15).Interior.Color = RGB(255, 192, 203)
                Debug.Print "Row " & i & " colored red (Value > 90)"
            Else
                ' Reset to black if not greater than 90 or invalid
                Debug.Print "Row " & i & " not colored (Value <= 90 or invalid)"
            End If
        Else
            ' Handle empty cells
            Debug.Print "Row " & i & " skipped (Empty): '" & cellValue & "'"
        End If
NextRow:
    Next i
End Sub