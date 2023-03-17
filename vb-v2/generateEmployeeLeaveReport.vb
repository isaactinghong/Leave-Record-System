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
    Dim employeeName As String
    employeeName = Range("EmployeeName").Value
    Debug.Print "EmployeeName: " & employeeName

    Dim EmployeeId As String
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
    Dim worksheetName As String: worksheetName = employeeName
    generateBlankWorksheet (worksheetName)

    'Get the Employee Profile
    Dim aEmployeeProfile As Object
    Set aEmployeeProfile = EmployeeProfile(EmployeeId)

    Dim employmentStartDate As Date, initialCarryOver As Integer, yearsWorkedBeforeExcelCreation As Integer
    employmentStartDate = aEmployeeProfile("EmploymentStartDate")
    initialCarryOver = aEmployeeProfile("InitialCarryOver")
    yearsWorkedBeforeExcelCreation = aEmployeeProfile("yearsWorkedBeforeExcelCreation")

    'Log the employee profile. dict.
    Debug.Print "---------------------------------"
    Debug.Print "Employee Profile: "
    Debug.Print "EmployeeId: " & EmployeeId
    Debug.Print "EmployeeName: " & employeeName
    Debug.Print "EmploymentStartDate: " & employmentStartDate
    Debug.Print "InitialCarryOver: " & initialCarryOver
    Debug.Print "yearsWorkedBeforeExcelCreation: " & yearsWorkedBeforeExcelCreation

    'Get the Meta Data
    Dim excelCreateDate As Variant
    excelCreateDate = MetaData("Excel Create Date")
    Debug.Print "---------------------------------"
    Debug.Print "ExcelCreateDate: " & excelCreateDate


    'Get the Leave Entitlement Policy
    'Dim leaves As Variant
    'leaves = LeaveEntitlementPolicy(8)
    'Debug.Print "---------------------------------"
    'Debug.Print "LeaveEntitlementPolicy 8 years: " & leaves


    'Set column width for the new worksheet
    Worksheets(worksheetName).Columns("A").ColumnWidth = 25
    Worksheets(worksheetName).Columns("B").ColumnWidth = 15

    'Write to the new worksheet
    Worksheets(worksheetName).Range("A1").Value = "Employee Name: "
    Worksheets(worksheetName).Range("B1").Value = employeeName

    Worksheets(worksheetName).Range("A2").Value = "Employee ID: "
    Worksheets(worksheetName).Range("B2").Value = EmployeeId

    Worksheets(worksheetName).Range("A3").Value = "Employment Start Date: "
    Worksheets(worksheetName).Range("B3").Value = employmentStartDate

    Worksheets(worksheetName).Range("A4").Value = "Initial Carry Over: "
    Worksheets(worksheetName).Range("B4").Value = initialCarryOver

    Worksheets(worksheetName).Range("A5").Value = "Years Worked Before Excel Creation: "
    Worksheets(worksheetName).Range("B5").Value = yearsWorkedBeforeExcelCreation


End Sub
