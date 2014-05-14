Attribute VB_Name = "workingDates"
Option Compare Database
Public Function WorkingDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
WorkingDays = WorkDays(dtBegin, dtEnd) - HollidayDays(dtBegin, dtEnd)
End Function

' WorkDays
' returns the number of working days between two dates
Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long

   Dim dtFirstSunday As Date
   Dim dtLastSaturday As Date
   Dim lngWorkDays As Long

   ' get first sunday in range
   dtFirstSunday = dtBegin + ((8 - Weekday(dtBegin)) Mod 7)

   ' get last saturday in range
   dtLastSaturday = dtEnd - (Weekday(dtEnd) Mod 7)

   ' get work days between first sunday and last saturday
   lngWorkDays = (((dtLastSaturday - dtFirstSunday) + 1) / 7) * 5

   ' if first sunday is not begin date
   If dtFirstSunday <> dtBegin Then

      ' assume first sunday is after begin date
      ' add workdays from begin date to first sunday
      lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

   End If

   ' if last saturday is not end date
   If dtLastSaturday <> dtEnd Then

      ' assume last saturday is before end date
      ' add workdays from last saturday to end date
      lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

   End If

   ' return working days
   WorkDays = lngWorkDays

End Function

Public Function DateWithinRange(ByVal theDate As Date, ByVal theBeginDate As Date, ByVal theEndDate As Date) As Boolean
On Error GoTo DateWithinRange_err

' PURPOSE: To determine whether-or-not a date is within a given range of dates
' ACCEPTS: - The date we want to check
'          - First date in the range
'          - Last date in range
' RETURNS: True if the date is within the range, Else False
Dim x  As Long
Dim Y  As Long

   x = DateDiff("d", theBeginDate, theDate)
   If x > -1 Then
      Y = DateDiff("d", theDate, theEndDate)
      If Y > -1 Then
         DateWithinRange = True
      End If
   End If

DateWithinRange_xit:
On Error Resume Next
Exit Function

DateWithinRange_err:
Resume DateWithinRange_xit
End Function


Public Function HollidayDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    'Obtener la lista de dias festivos
    Dim hollidayDaysList As DAO.Recordset
    Dim someDay As Date
    Dim belongPeriod As Boolean
    Dim countHollidays As Integer
    countHollidays = 0
    
    Set hollidayDaysList = db.OpenRecordset("qrHollidayDates", dbOpenDynaset, dbSeeChanges)
    If Not (hollidayDaysList.BOF And hollidayDaysList.EOF) Then
        hollidayDaysList.MoveNext
        Do While Not hollidayDaysList.EOF
            someDay = hollidayDaysList![HolidayDate]
            belongPeriod = False
            belongPeriod = DateWithinRange(someDay, dtBegin, dtEnd)
            If (belongPeriod) Then countHollidays = countHollidays + 1
            hollidayDaysList.MoveNext
        Loop
    End If
    HollidayDays = countHollidays
End Function
Sub dateFunctions()
   Dim strDateString As String

   strDateString = "The days between 3/15/2000 and today is: " & _
                    DateDiff("d", "3/15/2000", Now) & vbCrLf & _
                   "The months between 3/15/2000 and today is: " & _
                    DateDiff("m", "3/15/2000", Now)

   MsgBox strDateString
End Sub

Public Function dateDiffInDays(ByVal prmDate As Date) As Integer
Debug.Print "Entro a la diferencias dias"
dateDiffInDays = DateDiff("d", prmDate, Now())
End Function

Public Function dateDiffInMonths(ByVal prmDate As Date) As Integer
Debug.Print "Entro a la diferencias meses"
dateDiffInDays = DateDiff("m", prmDate, Now())
End Function

Public Function EmpDateHired(ByVal EmpID As Integer) As Date

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
    End If
    EmpDateHired = hireDate
End Function

Public Function AccrualSickHours(ByVal EmpID As Integer) As Double

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workDaysCurrentYear As Long
    Dim sickUnitValueHoursForTheYear As Double ' Cuanto vale la unidad de hora de enfermedad por dia de trabajo
    Dim facultyHours As Byte ' Total Hours for Faculty
    Dim staffHours As Byte ' Total Hours for Staff
    Dim totalHours As Byte ' Total Hours for everyone
    facultyHours = 25
    staffHours = 40
    
    workDaysCurrentYear = WorkingDays(Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy"), Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
    'Debug.Print "workDaysCurrentYear: " & workDaysCurrentYear
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    
    If (EmployeeType = "F") Then
        sickUnitValueHoursForTheYear = facultyHours / workDaysCurrentYear
        totalHours = facultyHours
    Else
        sickUnitValueHoursForTheYear = staffHours / workDaysCurrentYear
        totalHours = staffHours
    End If
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        'Horas tomadas por el empleado
        Dim SickHoursTaken As DAO.Recordset
        Dim dblSickHoursTaken As Double
        dblSickHoursTaken = 0
        
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS SickHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""S"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set SickHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (SickHoursTaken.BOF And SickHoursTaken.EOF) Then
            SickHoursTaken.MoveFirst
            dblSickHoursTaken = dblSickHoursTaken + SickHoursTaken![SickHours]
        End If
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular las horas a usar reales (restantes horas)
                workDaysCurrentYear = WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                'Debug.Print "WorkingDays: " & WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                hoursCanUse = workDaysCurrentYear * sickUnitValueHoursForTheYear
                'Debug.Print workDaysCurrentYear & " * " & sickUnitValueHoursForTheYear
            Else
                hoursCanUse = totalHours
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       AccrualSickHours = hoursCanUse - dblSickHoursTaken
       If (AccrualSickHours < 0) Then AccrualSickHours = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function AccrualBereavementHours(ByVal EmpID As Integer) As Double

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workDaysCurrentYear As Long
    Dim sickUnitValueHoursForTheYear As Double ' Cuanto vale la unidad de hora de enfermedad por dia de trabajo
    Dim facultyHours As Byte ' Total Hours for Faculty
    Dim staffHours As Byte ' Total Hours for Staff
    Dim totalHours As Byte ' Total Hours for everyone
    facultyHours = 15
    staffHours = 24
    
    workDaysCurrentYear = WorkingDays(Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy"), Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    
    If (EmployeeType = "F") Then
        'sickUnitValueHoursForTheYear = facultyHours / workDaysCurrentYear
        sickUnitValueHoursForTheYear = facultyHours
        totalHours = facultyHours
    Else
        'sickUnitValueHoursForTheYear = staffHours / workDaysCurrentYear
        sickUnitValueHoursForTheYear = staffHours
        totalHours = staffHours
    End If
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        'Horas tomadas por el empleado
        Dim SickHoursTaken As DAO.Recordset
        Dim dblSickHoursTaken As Double
        dblSickHoursTaken = 0
        
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS SickHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)='E') AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set SickHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (SickHoursTaken.BOF And SickHoursTaken.EOF) Then
            SickHoursTaken.MoveFirst
            dblSickHoursTaken = dblSickHoursTaken + SickHoursTaken![SickHours]
        End If
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = totalHours
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular las horas a usar reales (restantes horas)
                'workDaysCurrentYear = WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                'hoursCanUse = workDaysCurrentYear * sickUnitValueHoursForTheYear
                hoursCanUse = sickUnitValueHoursForTheYear
            Else
                hoursCanUse = totalHours
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       AccrualBereavementHours = hoursCanUse - dblSickHoursTaken
       If (AccrualBereavementHours < 0) Then AccrualBereavementHours = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function SickHoursTaken(ByVal EmpID As Integer) As Double

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workDaysCurrentYear As Long
    Dim sickUnitValueHoursForTheYear As Double ' Cuanto vale la unidad de hora de enfermedad por dia de trabajo
    Dim facultyHours As Byte ' Total Hours for Faculty
    Dim staffHours As Byte ' Total Hours for Staff
    Dim totalHours As Byte ' Total Hours for everyone
    facultyHours = 25
    staffHours = 40
    
    workDaysCurrentYear = WorkingDays(Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy"), Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    
    If (EmployeeType = "F") Then
        sickUnitValueHoursForTheYear = facultyHours / workDaysCurrentYear
        totalHours = facultyHours
    Else
        sickUnitValueHoursForTheYear = staffHours / workDaysCurrentYear
        totalHours = staffHours
    End If
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        'Horas tomadas por el empleado
        Dim rSickHoursTaken As DAO.Recordset
        Dim dblSickHoursTaken As Double
        dblSickHoursTaken = 0
        
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS SickHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""S"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set rSickHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (rSickHoursTaken.BOF And rSickHoursTaken.EOF) Then
            rSickHoursTaken.MoveFirst
            dblSickHoursTaken = dblSickHoursTaken + rSickHoursTaken![SickHours]
        End If
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular las horas a usar reales (restantes horas)
                workDaysCurrentYear = WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                hoursCanUse = workDaysCurrentYear * sickUnitValueHoursForTheYear
            Else
                hoursCanUse = totalHours
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       SickHoursTaken = dblSickHoursTaken
       If (SickHoursTaken < 0) Then SickHoursTaken = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function BereavementHoursTaken(ByVal EmpID As Integer) As Double

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workDaysCurrentYear As Long
    Dim sickUnitValueHoursForTheYear As Double ' Cuanto vale la unidad de hora de enfermedad por dia de trabajo
    Dim facultyHours As Byte ' Total Hours for Faculty
    Dim staffHours As Byte ' Total Hours for Staff
    Dim totalHours As Byte ' Total Hours for everyone
    facultyHours = 15
    staffHours = 24
    
    workDaysCurrentYear = WorkingDays(Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy"), Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    
    If (EmployeeType = "F") Then
        'sickUnitValueHoursForTheYear = facultyHours / workDaysCurrentYear
        sickUnitValueHoursForTheYear = facultyHours
        totalHours = facultyHours
    Else
        'sickUnitValueHoursForTheYear = staffHours / workDaysCurrentYear
        sickUnitValueHoursForTheYear = staffHours
        totalHours = staffHours
    End If
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        'Horas tomadas por el empleado
        Dim rSickHoursTaken As DAO.Recordset
        Dim dblSickHoursTaken As Double
        dblSickHoursTaken = 0
        
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS SickHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)='E') AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set rSickHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (rSickHoursTaken.BOF And rSickHoursTaken.EOF) Then
            rSickHoursTaken.MoveFirst
            dblSickHoursTaken = dblSickHoursTaken + rSickHoursTaken![SickHours]
        End If
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = totalHours
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular las horas a usar reales (restantes horas)
                'workDaysCurrentYear = WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                'hoursCanUse = workDaysCurrentYear * sickUnitValueHoursForTheYear
                hoursCanUse = sickUnitValueHoursForTheYear
            Else
                hoursCanUse = totalHours
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       BereavementHoursTaken = dblSickHoursTaken
       If (BereavementHoursTaken < 0) Then BereavementHoursTaken = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function SickHoursCanUse(ByVal EmpID As Integer) As Double

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workDaysCurrentYear As Long
    Dim sickUnitValueHoursForTheYear As Double ' Cuanto vale la unidad de hora de enfermedad por dia de trabajo
    Dim facultyHours As Byte ' Total Hours for Faculty
    Dim staffHours As Byte ' Total Hours for Staff
    Dim totalHours As Byte ' Total Hours for everyone
    facultyHours = 25
    staffHours = 40
    
    workDaysCurrentYear = WorkingDays(Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy"), Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    
    If (EmployeeType = "F") Then
        sickUnitValueHoursForTheYear = facultyHours / workDaysCurrentYear
        totalHours = facultyHours
    Else
        sickUnitValueHoursForTheYear = staffHours / workDaysCurrentYear
        totalHours = staffHours
    End If
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        'Horas tomadas por el empleado
        Dim SickHoursTaken As DAO.Recordset
        Dim dblSickHoursTaken As Double
        dblSickHoursTaken = 0
        
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS SickHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""S"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set SickHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (SickHoursTaken.BOF And SickHoursTaken.EOF) Then
            SickHoursTaken.MoveFirst
            dblSickHoursTaken = dblSickHoursTaken + SickHoursTaken![SickHours]
        End If
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular las horas a usar reales (restantes horas)
                'workDaysCurrentYear = WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                workDaysCurrentYear = 12
                hoursCanUse = workDaysCurrentYear * sickUnitValueHoursForTheYear
            Else
                hoursCanUse = totalHours
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       SickHoursCanUse = hoursCanUse
       If (SickHoursCanUse < 0) Then SickHoursCanUse = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function BereavementHoursCanUse(ByVal EmpID As Integer) As Double

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workDaysCurrentYear As Long
    Dim sickUnitValueHoursForTheYear As Double ' Cuanto vale la unidad de hora de enfermedad por dia de trabajo
    Dim facultyHours As Byte ' Total Hours for Faculty
    Dim staffHours As Byte ' Total Hours for Staff
    Dim totalHours As Byte ' Total Hours for everyone
    facultyHours = 15
    staffHours = 24
    
    workDaysCurrentYear = WorkingDays(Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy"), Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    
    If (EmployeeType = "F") Then
        'sickUnitValueHoursForTheYear = facultyHours / workDaysCurrentYear
        sickUnitValueHoursForTheYear = facultyHours
        totalHours = facultyHours
    Else
        'sickUnitValueHoursForTheYear = staffHours
        sickUnitValueHoursForTheYear = staffHours / workDaysCurrentYear
        totalHours = staffHours
    End If
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        'Horas tomadas por el empleado
        Dim SickHoursTaken As DAO.Recordset
        Dim dblSickHoursTaken As Double
        dblSickHoursTaken = 0
        
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS SickHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)='E') AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set SickHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (SickHoursTaken.BOF And SickHoursTaken.EOF) Then
            SickHoursTaken.MoveFirst
            dblSickHoursTaken = dblSickHoursTaken + SickHoursTaken![SickHours]
        End If
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = totalHours
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular las horas a usar reales (restantes horas)
                'workDaysCurrentYear = WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                'workDaysCurrentYear = 12
                'hoursCanUse = workDaysCurrentYear * sickUnitValueHoursForTheYear
                hoursCanUse = sickUnitValueHoursForTheYear
            Else
                hoursCanUse = totalHours
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       BereavementHoursCanUse = hoursCanUse
       If (BereavementHoursCanUse < 0) Then BereavementHoursCanUse = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function AccrualSickDays(ByVal EmpID As Integer) As Double

    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workDaysCurrentYear As Long
    Dim sickUnitValueDaysForTheYear As Double ' Cuanto vale la unidad de hora de enfermedad por dia de trabajo
    Dim facultyDays As Byte ' Total Days for Faculty
    Dim staffDays As Byte ' Total Days for Staff
    Dim totalDays As Byte ' Total Days for everyone
    facultyDays = 5
    staffDays = 5
    
    workDaysCurrentYear = WorkingDays(Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy"), Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    'MsgBox "EmployeeType: " & EmployeeType
    If (EmployeeType = "F") Then
        sickUnitValueDaysForTheYear = facultyDays / workDaysCurrentYear
        totalDays = facultyDays
    Else
        sickUnitValueDaysForTheYear = staffDays / workDaysCurrentYear
        totalDays = staffDays
    End If
   ' MsgBox "Total Days: " & totalDays
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        'Horas tomadas por el empleado
        Dim sickDaysTaken As DAO.Recordset
        Dim dblSickDaysTaken As Double
        dblSickDaysTaken = 0
        
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffDays) AS SickDays " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""S"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set sickDaysTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (sickDaysTaken.BOF And sickDaysTaken.EOF) Then
            sickDaysTaken.MoveFirst
            dblSickDaysTaken = dblSickDaysTaken + sickDaysTaken![SickDays]
        End If
       ' MsgBox "dblSickDaysTaken = dblSickDaysTaken + sickDaysTaken![SickDays]: " & dblSickDaysTaken
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim daysCanUse As Double 'Days that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        daysCanUse = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            daysCanUse = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular las horas a usar reales (restantes horas)
                workDaysCurrentYear = WorkingDays(finishInitialPeriod, Format(CDate(DateSerial(Year(Now()), 12, 31)), "mm/dd/yyyy"))
                daysCanUse = workDaysCurrentYear * sickUnitValueDaysForTheYear
            Else
                daysCanUse = totalDays
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       AccrualSickDays = daysCanUse - dblSickDaysTaken
      ' MsgBox "daysCanUse - dblSickDaysTaken: " & daysCanUse & " - " & dblSickDaysTaken
       If (AccrualSickDays < 0) Then AccrualSickDays = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function


Public Function AccrualVacationHours(ByVal EmpID As Integer) As Double
    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workedMonths As Long
    workedMonths = 0
    Dim newYear As Date
    newYear = Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy")
    Dim ratio As Double
    ratio = 0
    
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        Dim seniorityTime As Byte 'Tiempo de antiguedad
        seniorityTime = DateDiff("m", hireDate, Now())
        'Debug.Print "AVH"
        'Debug.Print seniorityTime
        
        
        'Saber en que periodo se encuentra en empleado, si en menos o en mas de 47 meses
        'No obstante si esta en menos de 47 aun asi se le aplicara la validacion
        'de si paso o no el periodo de prueba
        ratio = 0
        If (EmployeeType = "F") Then
            If (seniorityTime <= 41) Then
                ratio = (10 / 12) * 5
            Else
                ratio = (15 / 12) * 5
            End If
        Else
            If (seniorityTime <= 41) Then
                ratio = (10 / 12) * 8
            Else
                ratio = (15 / 12) * 8
                'Debug.Print "Aca"
                'Debug.Print ratio
            End If
        End If
        'Debug.Print ratio
        'Horas tomadas por el empleado
        Dim VacationHoursTaken As DAO.Recordset
        Dim dblVacationHoursTaken As Double
        dblVacationHoursTaken = 0
            
'        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS VacationHours " _
'            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
'            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber, Year([TimeOfPeriodStart]) " _
'            & " HAVING ((((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year(dbo_DCB_EmployeeTimeOffRoster.TimeOfPeriodStart))=Year(Now()))) OR (dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber='0101042013'));"
      sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS VacationHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber, Year([TimeOfPeriodStart])" _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([dbo_DCB_EmployeeTimeOffRoster].[TimeOfPeriodStart]))=Year(Now()))) OR (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber)='0101042013'));"
        
        'Debug.Print sql
                
        Set VacationHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        
        If Not (VacationHoursTaken.BOF And VacationHoursTaken.EOF) Then
            VacationHoursTaken.MoveFirst
            Do While Not VacationHoursTaken.EOF
                dblVacationHoursTaken = dblVacationHoursTaken + VacationHoursTaken![VacationHours]
                'MsgBox VacationHoursTaken![VacationHours]
                VacationHoursTaken.MoveNext
            Loop
        End If
        'MsgBox "dblVacationHoursTaken = dblVacationHoursTaken + VacationHoursTaken![VacationHours]: " & dblVacationHoursTaken
        
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        Dim dayStartWorking As Byte 'Dia del mes en que empezo a trabajar un empleado
        Dim dayofToday As Byte 'Dia del mes en que empezo a trabajar un empleado
        dayStartWorking = Day(hireDate)
        dayofToday = Day(Now())
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        workedMonths = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = 0
            workedMonths = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular los meses que hay de esa fecha a hoy
                workedMonths = DateDiff("m", finishInitialPeriod, Now())
                If (dayofToday < dayStartWorking) Then
                        workedMonths = workedMonths - 1
                        If (workedMonths < 0) Then workedMonths = 0
                End If
                hoursCanUse = ratio * workedMonths
            Else
                'calcular los meses transcurridos en el ano
                workedMonths = DateDiff("m", newYear, Now())
                'If (dayofToday < dayStartWorking) Then workedMonths = workedMonths - 1
                If (workedMonths > 0 And Month(Date) = 12) Then workedMonths = workedMonths + 1
                
                hoursCanUse = ratio * workedMonths
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       AccrualVacationHours = hoursCanUse - dblVacationHoursTaken
       'MsgBox "AccrualVacationHours = hoursCanUse - dblVacationHoursTaken: " & AccrualVacationHours & " = " & hoursCanUse & " - " & dblVacationHoursTaken
       If (AccrualVacationHours < 0) Then AccrualVacationHours = 0
       'AccrualVacationHours = 0 ' Esta linea es nada mas para el inicio de un nuevo anno. Despues de eso comentarla.
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function VacationHoursTaken(ByVal EmpID As Integer) As Double
    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workedMonths As Long
    workedMonths = 0
    Dim newYear As Date
    newYear = Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy")
    Dim ratio As Double
    ratio = 0
    
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        Dim seniorityTime As Byte 'Tiempo de antiguedad
        seniorityTime = DateDiff("m", hireDate, Now())
        'Debug.Print "VHT"
        'Debug.Print seniorityTime
        
        'Saber en que periodo se encuentra en empleado, si en menos o en mas de 47 meses
        'No obstante si esta en menos de 47 aun asi se le aplicara la validacion
        'de si paso o no el periodo de prueba
        If (EmployeeType = "F") Then
            If (seniorityTime <= 41) Then
                ratio = (10 / 12) * 5
            Else
                ratio = (15 / 12) * 5
            End If
        Else
            If (seniorityTime <= 41) Then
                ratio = (10 / 12) * 8
            Else
                ratio = (15 / 12) * 8
            End If
        End If
        'Debug.Print ratio
        
        'Horas tomadas por el empleado
        Dim rVacationHoursTaken As DAO.Recordset
        Dim dblVacationHoursTaken As Double
        dblVacationHoursTaken = 0
            
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS VacationHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set rVacationHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (rVacationHoursTaken.BOF And rVacationHoursTaken.EOF) Then
            rVacationHoursTaken.MoveFirst
            dblVacationHoursTaken = dblVacationHoursTaken + rVacationHoursTaken![VacationHours]
        End If
        
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        Dim dayStartWorking As Byte 'Dia del mes en que empezo a trabajar un empleado
        Dim dayofToday As Byte 'Dia del mes en que empezo a trabajar un empleado
        dayStartWorking = Day(hireDate)
        dayofToday = Day(Now())
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        workedMonths = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = 0
            workedMonths = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular los meses que hay de esa fecha a hoy
                workedMonths = DateDiff("m", finishInitialPeriod, Now())
                If (dayofToday < dayStartWorking) Then workedMonths = workedMonths - 1
                hoursCanUse = ratio * workedMonths
            Else
                'calcular los meses transcurridos en el ano
                workedMonths = DateDiff("m", newYear, Now())
                'If (dayofToday < dayStartWorking) Then workedMonths = workedMonths - 1
                hoursCanUse = ratio * workedMonths
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       VacationHoursTaken = dblVacationHoursTaken
       If (VacationHoursTaken < 0) Then VacationHoursTaken = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function VacationHoursCanUse(ByVal EmpID As Integer) As Double
    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workedMonths As Long
    workedMonths = 0
    Dim newYear As Date
    newYear = Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy")
    Dim ratio As Double
    ratio = 0
    
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        Dim seniorityTime As Byte 'Tiempo de antiguedad
        seniorityTime = DateDiff("m", hireDate, Now())
        'Debug.Print "VHCU"
        'Debug.Print seniorityTime
        
        'Saber en que periodo se encuentra en empleado, si en menos o en mas de 47 meses
        'No obstante si esta en menos de 47 aun asi se le aplicara la validacion
        'de si paso o no el periodo de prueba
        If (EmployeeType = "F") Then
            If (seniorityTime <= 41) Then
                ratio = (10 / 12) * 5
            Else
                ratio = (15 / 12) * 5
            End If
        Else
            If (seniorityTime <= 41) Then
                ratio = (10 / 12) * 8
            Else
                ratio = (15 / 12) * 8
            End If
        End If
        'Debug.Print ratio
        'Horas tomadas por el empleado
        Dim VacationHoursTaken As DAO.Recordset
        Dim dblVacationHoursTaken As Double
        dblVacationHoursTaken = 0
            
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffHrs) AS VacationHours " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, Year([TimeOfPeriodStart]) " _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([TimeOfPeriodStart]))=Year(Now())));"
        
        Set VacationHoursTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        If Not (VacationHoursTaken.BOF And VacationHoursTaken.EOF) Then
            VacationHoursTaken.MoveFirst
            dblVacationHoursTaken = dblVacationHoursTaken + VacationHoursTaken![VacationHours]
        End If
        
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim hoursCanUse As Double 'Hours that employee has for use
        initialEP = 3 ' 3 Months is the initial period
        Dim dayStartWorking As Byte 'Dia del mes en que empezo a trabajar un empleado
        Dim dayofToday As Byte 'Dia del mes en que empezo a trabajar un empleado
        dayStartWorking = Day(hireDate)
        dayofToday = Day(Now())
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        hoursCanUse = 0
        workedMonths = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            hoursCanUse = 0
            workedMonths = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular los meses que hay de esa fecha a hoy
                'workedMonths = DateDiff("m", finishInitialPeriod, Now())
                'If (dayofToday < dayStartWorking) Then workedMonths = workedMonths - 1
                workedMonths = 12
                hoursCanUse = ratio * workedMonths
            Else
                'calcular los meses transcurridos en el ano
                'workedMonths = DateDiff("m", newYear, Now())
                'If (dayofToday < dayStartWorking) Then workedMonths = workedMonths - 1
                workedMonths = 12
                hoursCanUse = ratio * workedMonths
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       VacationHoursCanUse = hoursCanUse
       If (VacationHoursCanUse < 0) Then VacationHoursCanUse = 0
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

Public Function AccrualVacationDays(ByVal EmpID As Integer) As Double
    'Variables generales
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim MyRS As DAO.Recordset
    Dim sql As String
    Dim hireDate As Date
    Dim initialEmployeePeriod As Boolean
    
    'Variables a setear desde un inicio
    Dim EmployeeType As String
    Dim workedMonths As Long
    workedMonths = 0
    Dim newYear As Date
    newYear = Format(CDate(DateSerial(Year(Now()), 1, 1)), "mm/dd/yyyy")
    Dim ratio As Double
    ratio = 0
    
    
    sql = "SELECT dbo_DCB_EmployeeExtension.EmployeeType FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        EmployeeType = MyRS![EmployeeType]
    End If
    
    sql = "SELECT dbo_Employees.HireDate, dbo_Employees.EmpID FROM dbo_Employees WHERE (((dbo_Employees.EmpID)=" & EmpID & "));"
    Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not (MyRS.BOF And MyRS.EOF) Then
        'Fecha empleado
        MyRS.MoveFirst
        hireDate = MyRS![hireDate]
        Dim seniorityTime As Byte 'Tiempo de antiguedad
        seniorityTime = DateDiff("m", hireDate, Now())
        'Debug.Print "AVD"
        'Debug.Print seniorityTime
        
        'Saber en que periodo se encuentra en empleado, si en menos o en mas de 47 meses
        'No obstante si esta en menos de 47 aun asi se le aplicara la validacion
        'de si paso o no el periodo de prueba
        If (EmployeeType = "F") Then
            If (seniorityTime <= 41) Then
                ratio = (10 / 12)
            Else
                ratio = (15 / 12)
            End If
        Else
            If (seniorityTime <= 41) Then
                ratio = (10 / 12)
            Else
                ratio = (15 / 12)
            End If
        End If
        'Debug.Print ratio
        'Horas tomadas por el empleado
        Dim vacationDaysTaken As DAO.Recordset
        Dim dblVacationDaysTaken As Double
        dblVacationDaysTaken = 0
            
'        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffDays) AS VacationDays " _
'            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
'            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber, Year([TimeOfPeriodStart]) " _
'            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year(dbo_DCB_EmployeeTimeOffRoster.TimeOfPeriodStart))=Year(Now())) OR (dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber='0101042013'));"
        sql = "SELECT Sum(dbo_DCB_EmployeeTimeOffRoster.TimeOffDays) AS VacationDays " _
            & " FROM dbo_DCB_EmployeeTimeOffRoster " _
            & " GROUP BY dbo_DCB_EmployeeTimeOffRoster.TimeOffType, dbo_DCB_EmployeeTimeOffRoster.EmpID, dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber, Year([TimeOfPeriodStart])" _
            & " HAVING (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((Year([dbo_DCB_EmployeeTimeOffRoster].[TimeOfPeriodStart]))=Year(Now()))) OR (((dbo_DCB_EmployeeTimeOffRoster.TimeOffType)=""V"") AND ((dbo_DCB_EmployeeTimeOffRoster.EmpID)=" & EmpID & ") AND ((dbo_DCB_EmployeeTimeOffRoster.PayCycleNumber)='0101042013'));"
        
        
        Set vacationDaysTaken = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
        
        If Not (vacationDaysTaken.BOF And vacationDaysTaken.EOF) Then
            vacationDaysTaken.MoveFirst
            Do While Not vacationDaysTaken.EOF
                dblVacationDaysTaken = dblVacationDaysTaken + vacationDaysTaken![VacationDays]
                vacationDaysTaken.MoveNext
            Loop
        End If
        
        'Calcular si es un empleado de este ano
        Dim newEmployee As Boolean
        Dim initialEP As Byte ' el periodo incial de los empleados
        Dim finishInitialPeriod As Date ' cuando termina el periodo de prueba
        Dim daysCanUse As Double 'Days that employee has for use
        Dim dayStartWorking As Byte 'Dia del mes en que empezo a trabajar un empleado
        Dim dayofToday As Byte 'Dia del mes en que empezo a trabajar un empleado
        dayStartWorking = Day(hireDate)
        dayofToday = Day(Now())
        initialEP = 3 ' 3 Months is the initial period
        finishInitialPeriod = Format(CDate(DateSerial(Year(hireDate), Month(hireDate) + initialEP, Day(hireDate))), "mm/dd/yyyy")
        newEmployee = False
        daysCanUse = 0
        workedMonths = 0
        Dim diff As Integer
        diff = DateDiff("m", hireDate, Now())
        If (diff < initialEP) Then
            newEmployee = True
            daysCanUse = 0
            workedMonths = 0
        Else
            newEmployee = False
            If (Year(finishInitialPeriod) = Year(Now())) Then
                'no es nuevo empleado pero debo saber si el empleado
                'termino su periodo inicial en este anno vigente
                'si ese es el caso debo calcular los meses que hay de esa fecha a hoy
                workedMonths = DateDiff("m", finishInitialPeriod, Now())
                If (dayofToday < dayStartWorking) Then workedMonths = workedMonths - 1
                daysCanUse = ratio * workedMonths
            Else
                'calcular los meses transcurridos en el ano
                workedMonths = DateDiff("m", newYear, Now(), vbMonday)
                'If (dayofToday < dayStartWorking) Then workedMonths = workedMonths - 1
                If (workedMonths > 0 And Month(Date) = 12) Then workedMonths = workedMonths + 1
                daysCanUse = ratio * workedMonths
            End If 'If (Year(finishInitialPeriod) = Year(Now()))
        End If 'If (dateDiffInMonths(hireDate) < initialEP)
       AccrualVacationDays = daysCanUse - dblVacationDaysTaken
       If (AccrualVacationDays < 0) Then AccrualVacationDays = 0
       'AccrualVacationDays = 0 'Esta linea solo es para quitar las vacaciones al inicio del nuevo anno
       Exit Function
    Else
        MsgBox "You need to set up the hire date of the Employee ID: " & EmpID & " on DiamondD"
        Exit Function
    End If
End Function

