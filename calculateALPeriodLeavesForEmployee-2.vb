' LeaveRecords, Function to fetch leave records of an employee
Public Function LeaveRecords(employeeId As String) As Variant
    Dim leaveRecords As Variant
    leaveRecords = EmployeeProfiles(employeeId)("LeaveRecords")
    Dim leaveRecordsByDate As Variant
    leaveRecordsByDate = Application.Transpose(leaveRecords)
    Dim leaveRecordsByDateDict As Variant
    leaveRecordsByDateDict = Application.WorksheetFunction.CreateLookupArray(leaveRecordsByDate, 1, 2)
    LeaveRecords = leaveRecordsByDateDict
End Function

' EmployeeProfiles, Function to fetch employee profile of an employee
Public Function EmployeeProfiles(employeeId As String) As Variant
    Dim employeeProfiles As Variant
    employeeProfiles = Sheets("Employee Profiles").Range("A:E").Value
    Dim employeeProfilesDict As Variant
    employeeProfilesDict = Application.WorksheetFunction.CreateLookupArray(employeeProfiles, 1, 2)
    EmployeeProfiles = employeeProfilesDict(employeeId)
End Function

' MetaData, Function to fetch meta data of the excel file
Public Function MetaData(key As String) As Variant
    Dim metaData As Variant
    metaData = Sheets("Meta Data").Range("A:B").Value
    Dim metaDataDict As Variant
    metaDataDict = Application.WorksheetFunction.CreateLookupArray(metaData, 1, 2)
    MetaData = metaDataDict(key)
End Function

' HRPolicy, Function to fetch HR policy of the company
Public Function HRPolicy(key As String) As Variant
    Dim hrPolicy As Variant
    hrPolicy = Sheets("HR Policy").Range("A:B").Value
    Dim hrPolicyDict As Variant
    hrPolicyDict = Application.WorksheetFunction.CreateLookupArray(hrPolicy, 1, 2)
    HRPolicy = hrPolicyDict(key)
End Function


Public Function calculateALPeriodLeavesForEmployee(employeeId As String, leavesToApply As Variant) As Variant
    Dim employeeProfile As Variant
    employeeProfile = EmployeeProfiles(employeeId)
    Dim initialCarryOver As Integer
    initialCarryOver = employeeProfile("InitialCarryOver")

    ' Determine current AL period
    Dim currentDate As Date
    currentDate = MetaData("ExcelCreateDate")
    Dim yearsWorked As Integer
    yearsWorked = employeeProfile("YearsWorked")
    Dim currentPeriodStart As Date
    currentPeriodStart = WorksheetFunction.Max(employeeProfile("EmploymentStartDate"), DateAdd("m", -11, currentDate))
    Dim currentPeriodEnd As Date
    currentPeriodEnd = DateAdd("d", -1, DateAdd("yyyy", 1, currentPeriodStart))

    ' Calculate prorated and carry over leaves
    Dim carryOverEnd As Date
    carryOverEnd = DateAdd("m", -HRPolicy("carryOverMonths"), currentPeriodEnd)
    Dim carryOverLeaves As Integer
    carryOverLeaves = carryOverLeavesOfTargetDate(employeeId, carryOverEnd)
    Dim proratedLeaves As Integer
    proratedLeaves = prorataLeavesOfTargetDate(employeeId, currentDate)

    Dim ALPeriodLeavesTaken As Variant
    ALPeriodLeavesTaken = LeaveRecords(employeeId)(Array(currentPeriodStart, currentPeriodEnd))
    Dim ALPeriodCombinedLeaves As Variant

    ' concatenate leavesToApply to ALPeriodLeavesTaken if leavesToApply is not empty
    If IsEmpty(leavesToApply) Then
        ALPeriodCombinedLeaves = ALPeriodLeavesTaken
    Else
        ' concat leavesToApply to ALPeriodLeavesTaken in excel vba
        ALPeriodCombinedLeaves = Application.Union(ALPeriodLeavesTaken, leavesToApply)
    End If


    Dim carryOverLeavesApplied As Integer
    carryOverLeavesApplied = 0
    Dim prorataLeavesApplied As Integer
    prorataLeavesApplied = 0

    Dim i As Long
    For i = LBound(ALPeriodCombinedLeaves) To UBound(ALPeriodCombinedLeaves)
        Dim currLeave As Date
        currLeave = ALPeriodCombinedLeaves(i)

        Dim carryOverLeavesAvailable As Integer
        carryOverLeavesAvailable = WorksheetFunction.Min(initialCarryOver + carryOverLeaves - carryOverLeavesApplied, HRPolicy("carryOverLimit"))
        Dim prorataLeavesAvailable As Integer
        prorataLeavesAvailable = HRPolicy("baseLeaves") + WorksheetFunction.Min(yearsWorked, HRPolicy("maxYears")) - 2 - prorataLeavesApplied

        If currLeave >= currentPeriodStart And currLeave <= currentPeriodEnd Then
            If carryOverLeavesApplied < carryOverLeavesAvailable Then
                carryOverLeavesApplied = carryOverLeavesApplied + 1
            ElseIf prorataLeavesApplied < prorataLeavesAvailable Then
                prorataLeavesApplied = prorataLeavesApplied + 1
            Else
                calculateALPeriodLeavesForEmployee = "No more available leaves"
                Exit Function
            End If
        End If
    Next i

    Dim remainingLeavesInCarryOverPeriod As Integer
    remainingLeavesInCarryOverPeriod = carryOverLeavesAvailable - carryOverLeavesApplied
    Dim remainingLeavesAfterCarryOverPeriod As Integer
    remainingLeavesAfterCarryOverPeriod = prorataLeavesAvailable - prorataLeavesApplied

    calculateALPeriodLeavesForEmployee = Array("remainingLeavesInCarryOverPeriod", remainingLeavesInCarryOverPeriod, _
        "remainingLeavesAfterCarryOverPeriod", remainingLeavesAfterCarryOverPeriod)
End Function