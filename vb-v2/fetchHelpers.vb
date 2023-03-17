
' EmployeeProfiles, Function to fetch employee profile of an employee
Public Function EmployeeProfile(EmployeeId As String) As Dictionary
    Dim employeeProfiles As Variant
    employeeProfiles = Sheets("Employee Profiles").Range("A:E").Value

    Dim employeeProfilesDict As Object
    Set employeeProfilesDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To UBound(employeeProfiles, 1)

        'if key is empty, break
        If CStr(employeeProfiles(i, 1)) = "" Then
            Exit For
        End If

        'Debug.Print "CStr(employeeProfiles(i, 1)):" & CStr(employeeProfiles(i, 1))


        Dim employeeProfileDict As Object
        Set employeeProfileDict = CreateObject("Scripting.Dictionary")

        employeeProfileDict.Add "EmployeeId", employeeProfiles(i, 1)
        employeeProfileDict.Add "EmployeeName", employeeProfiles(i, 2)
        employeeProfileDict.Add "EmploymentStartDate", employeeProfiles(i, 3)
        employeeProfileDict.Add "InitialCarryOver", employeeProfiles(i, 4)
        employeeProfileDict.Add "yearsWorkedBeforeExcelCreation", employeeProfiles(i, 5)


        'add key value pair to dictionary
        ' the value is a variant array of 5 elements
        ' Dim aEmployeeProfile As Variant
        ' aEmployeeProfile = Array(employeeProfiles(i, 1), employeeProfiles(i, 2), employeeProfiles(i, 3), employeeProfiles(i, 4), employeeProfiles(i, 5))
        employeeProfilesDict.Add CStr(employeeProfiles(i, 1)), employeeProfileDict

    Next i

    Dim result As Object
    Set result = employeeProfilesDict(EmployeeId)

    'Return null if employeeId is not found
    If result Is Nothing Then
        Set EmployeeProfile = Nothing
    Else
        Set EmployeeProfile = result
    End If

End Function

' MetaData, Function to fetch meta data of the excel file
' metaData = Sheets("Meta Data").Range("A:B").Value
' reference the method of EmployeeProfile Function
Public Function MetaData(key As String) As Variant
    Dim metaDatas As Variant
    metaDatas = Sheets("Meta Data").Range("A:B").Value
    Dim metaDataDict As Object
    Set metaDataDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To UBound(metaDatas, 1)

        'if key is empty, break
        If CStr(metaDatas(i, 1)) = "" Then
            Exit For
        End If

        'Debug.Print "CStr(metaData(i, 1)):" & CStr(metaDatas(i, 1))
        'Debug.Print "metaDatas(i, 2):" & metaDatas(i, 2)

        'add key value pair to dictionary
        ' the value is a variant array of 5 elements
        Dim aMetaData As Variant
        aMetaData = metaDatas(i, 2)
        metaDataDict.Add CStr(metaDatas(i, 1)), aMetaData

    Next i

    MetaData = metaDataDict(key)
End Function

' LeaveEntitlementPolicy, Function to fetch Leave Entitlement Policy of the company
' LeaveEntitlementPolicy = Sheets("Leave Entitlement Policy").Range("A:B").Value
' reference the method of EmployeeProfile Function
Public Function LeaveEntitlementPolicy(years As String) As Integer
    Dim leaveEntitlementPolicies As Variant
    leaveEntitlementPolicies = Sheets("Leave Entitlement Policy").Range("A:B").Value
    Dim LeaveEntitlementPolicyDict As Object
    Set LeaveEntitlementPolicyDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To UBound(leaveEntitlementPolicies, 1)

        'if key is empty, break
        If CStr(leaveEntitlementPolicies(i, 1)) = "" Then
            Exit For
        End If

        'Debug.Print "CStr(LeaveEntitlementPolicy(i, 1)):" & CStr(leaveEntitlementPolicies(i, 1))
        'Debug.Print "leaveEntitlementPolicies(i, 2):" & leaveEntitlementPolicies(i, 2)

        'add key value pair to dictionary
        ' the value is a variant array of 5 elements
        Dim aLeaveEntitlementPolicy As Integer
        aLeaveEntitlementPolicy = leaveEntitlementPolicies(i, 2)
        LeaveEntitlementPolicyDict.Add CStr(leaveEntitlementPolicies(i, 1)), aLeaveEntitlementPolicy

    Next i


    ' If LeaveEntitlementPolicyDict contains the years, LeaveEntitlementPolicy = result
    ' Else, LeaveEntitlementPolicy = 18
    if LeaveEntitlementPolicyDict.Exists(years) Then
        LeaveEntitlementPolicy = LeaveEntitlementPolicyDict(years)
    Else
        LeaveEntitlementPolicy = 18
    End If

End Function