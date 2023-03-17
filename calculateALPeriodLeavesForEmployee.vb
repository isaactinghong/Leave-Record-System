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
    ALPeriodCombinedLeaves = concatArrays(ALPeriodLeavesTaken, leavesToApply)

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