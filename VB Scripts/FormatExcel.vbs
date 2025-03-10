Option Explicit

Sub FormatExcel()
    ' Main routine to reformat the sheet
    Dim wsInput As Worksheet, wsOutput As Worksheet
    
    Set wsInput = ThisWorkbook.Sheets(1)
    Set wsOutput = CreateOutputSheet(wsInput)
    
    WriteHeaders wsInput, wsOutput
    ProcessRows wsInput, wsOutput
End Sub

Function CreateOutputSheet(wsInput As Worksheet) As Worksheet
    ' Deletes any existing sheet named "EmployeesInsight" and creates a new one
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("EmployeesInsight").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "EmployeesInsight"
    Set CreateOutputSheet = ws
End Function

Sub WriteHeaders(wsInput As Worksheet, wsOutput As Worksheet)
    ' Copies headers dynamically to preserve Hebrew encoding
    Dim lastCol As Integer
    lastCol = wsInput.Cells(1, Columns.Count).End(xlToLeft).Column
    wsInput.Range(wsInput.Cells(1, 1), wsInput.Cells(1, lastCol)).Copy
    wsOutput.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
End Sub

' Loop through the input sheet rows and copy data into the output sheet.
' The grouping is by מחלקה (input column 1) and קוד עובד (input column 2).
Sub ProcessRows(wsInput As Worksheet, wsOutput As Worksheet)
    Dim lastRow As Long, i As Long, outRow As Long
    Dim currDept As String, currAgent As String
    Dim dept As String, agentCode As String, agentName As String, contractType As String, monthNum As Variant
    
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row
    outRow = 2
    currDept = ""
    currAgent = ""
    
    For i = 2 To lastRow
        ' Read grouping fields from input based on new structure:
        dept = wsInput.Cells(i, 1).Value          ' מחלקה
        agentCode = wsInput.Cells(i, 2).Value       ' קוד עובד
        agentName = wsInput.Cells(i, 3).Value       ' שם עובד
        contractType = wsInput.Cells(i, 4).Value      ' סוג חוזה
        monthNum = wsInput.Cells(i, 5).Value        ' Month
        
        ' Write the grouping columns only if they change
        If dept <> currDept Then
            wsOutput.Cells(outRow, 1).Value = dept
            currDept = dept
            currAgent = ""  ' Reset agent grouping when department changes
        End If
        
        If agentCode <> currAgent Then
            wsOutput.Cells(outRow, 2).Value = agentCode
            wsOutput.Cells(outRow, 3).Value = agentName
            wsOutput.Cells(outRow, 4).Value = contractType
            currAgent = agentCode
        End If
        
        ' Copy the remaining columns from the input row.
        wsOutput.Cells(outRow, 5).Value = monthNum
        
        ' Copy all numerical columns from input to output (columns 6 to 17)
        wsOutput.Range(wsOutput.Cells(outRow, 6), wsOutput.Cells(outRow, 17)).Value = _
            wsInput.Range(wsInput.Cells(i, 6), wsInput.Cells(i, 17)).Value
        
        outRow = outRow + 1
    Next i
End Sub




