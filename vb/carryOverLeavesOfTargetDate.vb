Function carryOverLeavesOfTargetDate(employeeId As Integer, targetDate As Date) As Integer
    ' Get the initial carry-over leaves of the employee profile
    Dim initialCarryOver As Integer
    initialCarryOver = WorksheetFunction.VLookup(employeeId, Sheets("Employee Profiles").Range("A:E"), 4, False)

    ' Check if the target date is after the ALPeriod's carryover period
    Dim carryOverEndDate As Date
    carryOverEndDate = WorksheetFunction.VLookup("Excel Create Date", Sheets("Meta Data").Range("A:B"), 2, False) + (WorksheetFunction.VLookup(employeeId, Sheets("HR Policy").Range("A:B"), 2, False) - 1) * 365
    If targetDate > carryOverEndDate Then
        carryOverLeavesOfTargetDate = 0
    Else
        ' Check if the target date is in the first ever ALPeriod of the employee
        Dim employmentStartDate As Date
        employmentStartDate = WorksheetFunction.VLookup(employeeId, Sheets("Employee Profiles").Range("A:E"), 3, False)
        Dim firstALPeriodStartDate As Date
        firstALPeriodStartDate = employmentStartDate + WorksheetFunction.VLookup(employeeId, Sheets("HR Policy").Range("A:B"), 2, False)
        If targetDate < firstALPeriodStartDate Then
            carryOverLeavesOfTargetDate = initialCarryOver
        Else
            ' Get the remaining leaves after carry-over period of the last ALPeriod
            Dim lastALPeriodEndDate As Date
            lastALPeriodEndDate = carryOverEndDate - WorksheetFunction.VLookup(employeeId, Sheets("HR Policy").Range("A:B"), 2, False) * 365
            Dim lastALPeriodRemaining As Variant
            lastALPeriodRemaining = calculateALPeriodLeavesForEmployee(employeeId, Array(lastALPeriodEndDate))(1)

            ' Return the remaining leaves after carry-over period of the last ALPeriod as the carry-over leaves
            carryOverLeavesOfTargetDate = lastALPeriodRemaining
        End If
    End If
End Function