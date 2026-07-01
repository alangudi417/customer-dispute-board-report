    ' Remove subtotals for row fields
Private Sub RemoveSubtotals(rFields As PivotFields)
    Dim pf As PivotField
    For Each pf In rFields
        pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next pf
End Sub


'================================================================================================================


' Delete sheet if exists
Private Sub DeleteSheetIfExists(wb As Workbook, sName As String)
    On Error Resume Next
    wb.Sheets(sName).Delete
    On Error GoTo 0
End Sub