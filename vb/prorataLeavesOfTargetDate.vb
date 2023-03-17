Function prorataLeavesOfTargetDate(employeeId As Integer, targetDate As Date) As Integer
    Debug.Print "----------------------------------------"
    Debug.Print "prorataLeavesOfTargetDate start..."

    Dim employeeProfileSheet As Worksheet, policySheet As Worksheet, employeeProfileRange As Range, policyRange As Range
    Dim employeeProfileArray() As Variant, policyArray() As Variant
    Dim currentDate As Date, activePeriodStart As Date, activePeriodEnd As Date
    Dim totalAllowedLeaves As Integer, totalYearsWorked As Long, proratedLeavesDouble As Double, proratedLeaves As Integer
    Dim i As Long

    Set employeeProfileSheet = ThisWorkbook.Sheets("Employee Profiles")
    Set policySheet = ThisWorkbook.Sheets("HR Policy")

    ' Fetch Employee Record from Employee Profiles using the provided employeeId
    Set employeeProfileRange = employeeProfileSheet.Range("A2:E" & employeeProfileSheet.Cells(Rows.Count, 1).End(xlUp).Row)
    employeeProfileArray = employeeProfileRange.Value

    ' Find the record of the employee with the provided employeeId
    Dim employeeProfile As Variant
    For i = LBound(employeeProfileArray) To UBound(employeeProfileArray)
        If employeeProfileArray(i, 1) = employeeId Then
            employeeProfile = employeeProfileArray(i, 1)
            ' log the values for debugging
            Debug.Print "Employee Profile: " & employeeProfile
            Exit For
        End If
    Next i

    ' set all the values from the array to the variables
    ' in the order of Employee ID, Employee Name,   Employment Start Date,  Initial Carry-Over Leaves,  Years Worked
    Dim employeeName As String, employmentStartDate As Date, initialCarryOverLeaves As Integer, yearsWorked As Integer
    employeeName = employeeProfileArray(i, 2)
    employmentStartDate = employeeProfileArray(i, 3)
    initialCarryOverLeaves = employeeProfileArray(i, 4)
    yearsWorked = employeeProfileArray(i, 5)

    ' log the values for debugging
    Debug.Print "Employee ID: " & employeeId
    Debug.Print "Employee Name: " & employeeName
    Debug.Print "Employment Start Date: " & employmentStartDate
    Debug.Print "Initial Carry-Over Leaves: " & initialCarryOverLeaves
    Debug.Print "Years Worked: " & yearsWorked

    ' Get Active Leave Period of the Employee that includes the targetDate
    currentDate = targetDate
    Debug.Print "Current Date: " & currentDate

    activePeriodStart = employmentStartDate
    Do Until activePeriodStart > currentDate
        activePeriodStart = DateAdd("yyyy", 1, activePeriodStart)
    Loop
    activePeriodStart = DateAdd("yyyy", -1, activePeriodStart)
    activePeriodEnd = DateAdd("yyyy", 1, activePeriodStart)
    activePeriodEnd = DateAdd("d", -1, activePeriodEnd)

    Debug.Print "Active Period Start Date: " & activePeriodStart
    Debug.Print "Active Period End Date: " & activePeriodEnd

    ' Get the total leaves allowed in the current leave period using HR Policy Sheet
    Set policyRange = policySheet.Range("A2:B11")
    policyArray = policyRange.Value

    ' Find the total allowed leaves for the current leave period based on the years worked

    totalYearsWorked = yearsWorked + Year(activePeriodStart) - Year(employmentStartDate)
    Debug.Print "Years Worked before excel created: " & yearsWorked
    Debug.Print "Total Years Worked: " & totalYearsWorked

    totalAllowedLeaves = 18
    For i = LBound(policyArray) To UBound(policyArray)
        ' policyArray(i, 1) is the years worked
        ' policyArray(i, 2) is the total leaves entitled

        Dim years As Integer, totalLeavesEntitled As Integer
        years = policyArray(i, 1)
        totalLeavesEntitled = policyArray(i, 2)

        If years = totalYearsWorked Then
            totalAllowedLeaves = totalLeavesEntitled
            Exit For
        End If
    Next i
    Debug.Print "Total Allowed Leaves: " & totalAllowedLeaves

    ' Calculate the day difference between the current date and the start of the active leave period
    Dim dayDiff As Long
    dayDiff = DateDiff("d", activePeriodStart, currentDate)

    ' Calculate Prorated Leaves for the current leave period for Employee
    proratedLeavesDouble = totalAllowedLeaves * dayDiff / 365

    Debug.Print "Prorated Leaves: " & Round(proratedLeavesDouble, 3)

    ' Round the prorated leaves to the nearest integer
    proratedLeaves = Round(proratedLeavesDouble, 0)

    Debug.Print "Prorated Leaves (Rounded): " & proratedLeaves

    ' Return the prorata leaves
    prorataLeavesOfTargetDate = proratedLeaves


End Function



Sub testingLeaveSys()
    Ans = prorataLeavesOfTargetDate(10001, "2019-03-16")
    MsgBox Ans
End Sub
