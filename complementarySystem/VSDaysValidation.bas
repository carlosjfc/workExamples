Attribute VB_Name = "VSDaysValidation"
Option Compare Database

Function FVSDaysValidation(ByVal EmpID As String, ByVal VType As String, ByVal HrsRequested As Long)
    EmployeeID = EmpID
    Dim mensaje As String
    mensaje = ""
    If (VType = "S") Then
        dblAccuredSickHours = Round(AccrualSickHours(EmployeeID), 2)
        dblSickHoursTaken = Round(SickHoursTaken(EmployeeID), 2)
        dblSickHoursCanUse = Round(SickHoursCanUse(EmployeeID), 2)
        If (HrsRequested > dblAccuredSickHours) Then
                mensaje = "We can't assign the " & HrsRequested & " Sick Hours Requested." & vbCrLf & " The employee only have " & dblAccuredSickHours & " Accured Sick Hours. " _
                          & vbCrLf & "This Employee have taked " & dblSickHoursTaken & " Sick hours and as a TOTAL for this current year the employee could take " & dblSickHoursCanUse _
                          & vbCrLf & "The Maximun that the supervisor coul approve is: " & dblSickHoursCanUse - dblSickHoursTaken & " Hours"
        End If
    ElseIf (VType = "V") Then
        dblAccuredVacationHours = Round(AccrualVacationHours(EmployeeID), 2)
        dblVacationHoursTaken = Round(VacationHoursTaken(EmployeeID), 2)
        dblVacationHoursCanUse = Round(VacationHoursCanUse(EmployeeID), 2)
        If (HrsRequested > dblAccuredVacationHours) Then
                mensaje = "We can't assign the " & HrsRequested & " Vacation Hours Requested." & vbCrLf & " The employee only have " & dblAccuredVacationHours & " Accured Vacations Hours. " _
                          & vbCrLf & "This Employee have taked " & dblVacationHoursTaken & " Vacation hours and as a TOTAL for this current year the employee could take " & dblVacationHoursCanUse _
                          & vbCrLf & "The Maximun that the supervisor coul approve is: " & dblVacationHoursCanUse - dblVacationHoursTaken & " Hours"
        End If
    ElseIf (VType = "E") Then
        dblAccuredBereavementHours = Round(AccrualBereavementHours(EmployeeID), 2)
        dblBereavementHoursTaken = Round(BereavementHoursTaken(EmployeeID), 2)
        dblBereavementHoursCanUse = Round(BereavementHoursCanUse(EmployeeID), 2)
        If (HrsRequested > dblAccuredBereavementHours) Then
                mensaje = "We can't assign the " & HrsRequested & " Bereavement Hours Requested." & vbCrLf & " The employee only have " & dblAccuredBereavementHours & " Accured Bereavement Hours. " _
                          & vbCrLf & "This Employee have taked " & dblBereavementHoursTaken & " Bereavement hours and as a TOTAL for this current year the employee could take " & dblBereavementHoursCanUse _
                          & vbCrLf & "The Maximun that the supervisor coul approve is: " & dblBereavementHoursCanUse - dblBereavementHoursTaken & " Hours"
        End If
    End If
 FVSDaysValidation = mensaje
End Function
