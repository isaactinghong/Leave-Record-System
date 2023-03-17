
Public Function generateBlankWorksheet(worksheetName As String)

    On Error Resume Next
    'Check if the worksheet already exists and delete if it does
    If Not Worksheets(worksheetName) Is Nothing Then
        Application.DisplayAlerts = False
        Worksheets(worksheetName).Delete
        Application.DisplayAlerts = True
    End If
    'Create the new worksheet
    Worksheets.Add.Name = worksheetName
End Function


Sub generateEmployeeLeaveReport()

    'Set Active Sheet to "Employee Record Generator"
    ActiveWorkbook.Worksheets("Employee Record Generator").Activate

    'Read from named cells
    Dim EmployeeName As Variant
    EmployeeName = Range("EmployeeName").Value
    Debug.Print "EmployeeName: " & EmployeeName

    Dim EmployeeId As Variant
    EmployeeId = Range("EmployeeId").Value
    Debug.Print "EmployeeId: " & EmployeeId

    'Read from range of values starting from particular cell
    Dim startCellAddress As String: startCellAddress = "LeavesToApply"
    Dim currentCellAddress As String: currentCellAddress = startCellAddress
    Dim currentCellValue As Variant


    Debug.Print "LeavesToApply: "
    Do While Not IsEmpty(Range(currentCellAddress))

        currentCellValue = Range(currentCellAddress).Value
        Debug.Print currentCellValue
        currentCellAddress = Range(currentCellAddress).Offset(1, 0).Address
    Loop

    'Enter the name of the worksheet to be created
    Dim worksheetName As String: worksheetName = "John"
    generateBlankWorksheet (worksheetName)

End Sub
