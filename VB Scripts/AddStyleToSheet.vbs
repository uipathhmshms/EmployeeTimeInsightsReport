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
	' Aligns column A and C to the center
	Align_A_And_C_ColumnsText
	' Center the text in the first row
	CenterTextInFirstRow firstRowRange
	' Add the text 'total' and color to total row 
	FormatTotalRows
	
	' Merge so the department name will be written once
	'MergeFirstColumnRows
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

Sub Align_A_And_C_ColumnsText()
    ' Get the used range in the active sheet
    Dim usedRange As Range
    Set usedRange = ActiveSheet.UsedRange
 
	' Center align columns A and C (except the header row)
    Dim i As Long
    For i = 2 To usedRange.Rows.Count
        Cells(i, 1).HorizontalAlignment = xlCenter  ' Column A
        Cells(i, 3).HorizontalAlignment = xlCenter  ' Column C
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

    ' Get the last used row in column G
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
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


' Sub MergeFirstColumnRows()
    ' Dim ws As Worksheet
    ' Dim lastRow As Long
    ' Dim i As Long, startRow As Long
    
    ' Set ws = ActiveSheet
    ' lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' ' Start grouping from row 2 (assuming header is in row 1)
    ' startRow = 2
    
    ' ' Loop from row 3 to lastRow (fixing off-by-one issue)
    ' For i = 3 To lastRow
        ' ' If value changes or end of sheet reached
        ' If ws.Cells(i, 1).Value <> ws.Cells(startRow, 1).Value Then
            ' ' If there is more than one row in the group, merge them
            ' If i - startRow > 1 Then
                ' ws.Range(ws.Cells(startRow, 1), ws.Cells(i - 1, 1)).Merge
                ' With ws.Range(ws.Cells(startRow, 1), ws.Cells(i - 1, 1))
                    ' .HorizontalAlignment = xlCenter
                    ' .VerticalAlignment = xlCenter
                ' End With
            ' End If
            ' ' Update startRow to the new group start
            ' startRow = i
        ' End If
    ' Next i
    
    ' ' Handle merging for the last group (fixes last department issue)
    ' If lastRow - startRow > 0 Then
        ' ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, 1)).Merge
        ' With ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, 1))
            ' .HorizontalAlignment = xlCenter
            ' .VerticalAlignment = xlCenter
        ' End With
    ' End If
' End Sub

