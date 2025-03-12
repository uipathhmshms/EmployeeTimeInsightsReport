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
	' Color the header
	ColorHeaderRow firstRowRange
	' Aligns column A,C,E to the center
	Align_A_C_E_ColumnsText
	' Center the text in the first row
	CenterTextInFirstRow firstRowRange
	' Add the text 'total' and color to total row 
	FormatTotalRows
	
	AddGrandTotalRow
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
    ws.Range(ws.Cells(lastRow, 2), ws.Cells(lastRow, lastCol)).Interior.Color = RGB(128, 200, 225)
    ws.Range(ws.Cells(lastRow, 2), ws.Cells(lastRow, lastCol)).Font.Bold = True
    
    ' Calculate and add grand totals for each numeric column
    For col = 2 To lastCol
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
