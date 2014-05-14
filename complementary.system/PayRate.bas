Attribute VB_Name = "PayRate"
Option Compare Database

Sub PayRateCalculation(EmpID As Integer, ADPFileNumber As Integer, PayClass As String, DeltaChangePayRate As Double, _
                        NewPayRateEffectiveFrom As Date, NewPayRateApprovedBy As Integer, NewPayRateApprovalDate As Date)

'General Variables
Dim db As DAO.Database
Set db = CurrentDb()
Dim sql As String
Dim dataCycle As DAO.Recordset
Dim MyRS As DAO.Recordset

'Pay Class
sql = "SELECT dbo_DCB_EmployeeExtension.PayClass FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
Dim CurrentPayClass As String
Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
If Not (MyRS.BOF And MyRS.EOF) Then
    MyRS.MoveFirst
    CurrentPayClass = MyRS![PayClass]
End If

'Pay Rate
sql = "SELECT dbo_DCB_EmployeeExtension.PayRate FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
Dim CurrentPayRate As String
Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
If Not (MyRS.BOF And MyRS.EOF) Then
    MyRS.MoveFirst
    CurrentPayRate = MyRS![PayRate]
End If

'HourlyRate
sql = "SELECT dbo_DCB_EmployeeExtension.HourlyRate FROM dbo_DCB_EmployeeExtension WHERE (((dbo_DCB_EmployeeExtension.EmpID)= " & EmpID & " )); "
Dim CurrentHourlyRate As String
Set MyRS = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
If Not (MyRS.BOF And MyRS.EOF) Then
    MyRS.MoveFirst
    CurrentHourlyRate = MyRS![HourlyRate]
End If

End Sub

