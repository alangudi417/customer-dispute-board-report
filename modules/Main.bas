Sub Report_Africa()

    Dim wbMain As Workbook, wbDispute As Workbook, wbNotes As Workbook
    Dim wsOpen As Worksheet, wsPivot As Worksheet, wsManager As Worksheet, wsData As Worksheet
    Dim srcRng As Range, pCache As PivotCache, pTable As PivotTable
    Dim lastRow As Long
    Dim filePathRaw As String, filePathOutput As String
    Dim todayStr As String
    Dim ans As VbMsgBoxResult, ansn As VbMsgBoxResult

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wbMain = ThisWorkbook

        ' Ask user before proceeding
    ans = MsgBox("Is the UDM_File Macro-Enabled?" & vbCrLf & vbCrLf & _
                 "Select Yes to continue.", vbYesNo + vbQuestion, "Confirmation")
    If ans = vbNo Then GoTo Cleanup
    ansn = MsgBox("Is the data from Column 'Case ID' changed from text to column?" & vbCrLf & vbCrLf & _
                 "Select Yes to continue.", vbYesNo + vbQuestion, "Confirmation")
    If ansn = vbNo Then GoTo Cleanup

    todayStr = Format(Date, "mm-dd-yy")

        ' Locate and open the required files
    filePathRaw = "/sample_data/Board _ Africa/UDM_Dispute " & Format(Date, "mm-dd-yy") & " _ Africa.xlsm"
    filePathOutput = "/sample_data/UDM Comments/UDM_Notes " & Format(Date, "mm-dd-yy") & ".xlsx"

    Set wbDispute = Workbooks.Open(filePathRaw)
    Set wbNotes = Workbooks.Open(filePathOutput)
    Set wsOpen = wbDispute.Sheets("Open")

        ' Change text to columns
    wsOpen.Columns("I:I").TextToColumns Destination:=Range("I1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
        FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

        ' Add Category, UDM Notes and Manager Name2 columns
    lastRow = wsOpen.Cells(wsOpen.Rows.Count, "A").End(xlUp).Row
    
    wsOpen.Range("AE1").Value = "Manager Name2"
    wsOpen.Range("AE2").Formula = "=AB2"
    wsOpen.Range("AE2").AutoFill wsOpen.Range("AE2:AE" & lastRow)
    wsOpen.Columns("AE").EntireColumn.Hidden = True

    wsOpen.Range("AF1").Value = "UDM Notes"
    wsOpen.Range("AF2").Formula = "=IFERROR(INDEX('[" & wbNotes.Name & "]Sheet1'!AW:AW," & "MATCH(I2,'[" & wbNotes.Name & "]Sheet1'!C:C,0)),"""")"
    wsOpen.Range("AF2").AutoFill wsOpen.Range("AF2:AF" & lastRow)
    
        ' Create "Pivot" sheet
    DeleteSheetIfExists wbDispute, "Pivot"
    Set srcRng = wsOpen.UsedRange
    Set pCache = wbDispute.PivotCaches.Create(xlDatabase, srcRng)
    Set wsPivot = wbDispute.Sheets.Add
    wsPivot.Name = "Pivot"

    Set pTable = pCache.CreatePivotTable(wsPivot.Range("A3"), "PivotData")
    With pTable
        .RowAxisLayout xlTabularRow
        .PivotFields("Customer Name ").Orientation = xlRowField
        .AddDataField .PivotFields("Customer Number"), "Disputes", xlCount
        .AddDataField .PivotFields("Disputed Amount"), "Disp. Amount", xlSum
        .PivotFields("Customer Name ").AutoSort xlDescending, "Disp. Amount"
        RemoveSubtotals .RowFields
    End With
    wsPivot.Columns("A:C").AutoFit

        ' Create "Data" sheet
    DeleteSheetIfExists wbDispute, "Data"
    Set wsData = wbDispute.Sheets.Add
    wsData.Name = "Data"

    With wsOpen
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        .Range("A1").AutoFilter Field:=20, _
            Criteria1:=Array("Sales Team", "New_case"), _
            Operator:=xlFilterValues
        
        .Range("Z2:Z" & lastRow).SpecialCells(xlCellTypeVisible).Copy wsData.Range("A2")
    End With

    wsData.Range("A1").Value = "Processor Name"
    wsData.Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    wsData.Range("B1").Value = "ID"
    wsData.Range("B2").Formula = "=IFERROR(INDEX(Open!Y:Y,MATCH(A2,Open!Z:Z,0)),"""")"
    wsData.Range("B2").AutoFill wsData.Range("B2:B" & lastRow)

        ' Adjust the Width of columns A:C
    wsData.Columns("A:C").AutoFit
    If wsOpen.AutoFilterMode Then wsOpen.AutoFilter.ShowAllData

        ' Create "Manager" sheet
    DeleteSheetIfExists wbDispute, "Manager"
    Set pCache = wbDispute.PivotCaches.Create(xlDatabase, wsOpen.UsedRange)
    Set wsManager = wbDispute.Sheets.Add
    wsManager.Name = "Manager"

    Set pTable = pCache.CreatePivotTable(wsManager.Range("A3"), "ManagerData")
    With pTable
        .RowAxisLayout xlTabularRow
        .PivotFields("Manager Name").Orientation = xlRowField
        .PivotFields("Manager Name").Position = 1
        .AddDataField .PivotFields("Manager Name2"), "Disputes", xlCount
        .AddDataField .PivotFields("Disputed Amount"), "Disp. Amount", xlSum
        .PivotFields("Manager name").AutoSort xlDescending, "Disp. Amount"
        RemoveSubtotals .RowFields
    End With
    
        ' Adjust the Width of columns A:C
    wsManager.Columns("A:C").AutoFit

        ' Final Output
    MsgBox "The warboard report has been completed."

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub