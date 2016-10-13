Attribute VB_Name = "DatabaseObjects"
Option Compare Database

    Function dmwListAllTables() As String
    Dim tbl As AccessObject, db As Object
    Dim strMsg As String
     
    On Error GoTo Error_Handler
     
    Set db = Application.CurrentData
    For Each tbl In db.AllTables
    'Debug.Print tbl.NAME
    Next tbl
     
    strMsg = " -- Tables listing complete -- "
     
Procedure_Done:
    dmwListAllTables = strMsg
    Exit Function
     
Error_Handler:
    strMsg = Err.Number & " " & Err.Description
    Resume Procedure_Done
     
    End Function

    Function dmwListAllTablesNotMSys() As String
    Dim tbl As AccessObject, db As Object
    Dim strMsg As String
     
    On Error GoTo Error_Handler
     
    Set db = Application.CurrentData
    For Each tbl In db.AllTables
    If Not Left(tbl.NAME, 4) = "MSys" Then
    'Debug.Print tbl.NAME
    End If
    Next tbl
     
    strMsg = " -- Tables listing complete -- "
     
Procedure_Done:
    dmwListAllTablesNotMSys = strMsg
    Exit Function
     
Error_Handler:
    strMsg = Err.Number & " " & Err.Description
    Resume Procedure_Done
     
    End Function


    Function dmwListAllQueries() As String
    Dim strMsg As String
    Dim qry As AccessObject, db As Object
     
    On Error GoTo Error_Handler
     
    Set db = Application.CurrentData
    For Each qry In db.AllQueries
    'Debug.Print qry.NAME
    Next qry
     
    strMsg = " -- Queries listing complete -- "
     
Procedure_Done:
    dmwListAllQueries = strMsg
    Exit Function
     
Error_Handler:
    strMsg = Err.Number & " " & Err.Description
    Resume Procedure_Done
     
    End Function

    Function dmwListAllForms(ByRef obj As Object) As String
    Dim strMsg As String
    Dim frm As AccessObject, db As Object
     
    On Error GoTo Error_Handler
     
    Set db = Application.CurrentProject
    For Each frm In db.AllForms
    'Debug.Print frm.NAME
    obj.AddItem (frm.NAME)
    Next frm
     
    strMsg = " -- Forms listing complete -- "
     
Procedure_Done:
    dmwListAllForms = strMsg
    Exit Function
     
Error_Handler:
    strMsg = Err.Number & " " & Err.Description
    Resume Procedure_Done
     
    End Function

    Function dmwListAllReports(ByRef obj As Object) As String
    Dim strMsg As String
    Dim rpt As AccessObject, db As Object
     
    On Error GoTo Error_Handler
     
    Set db = Application.CurrentProject
    For Each rpt In db.AllReports
    'Debug.Print rpt.NAME
    obj.AddItem (rpt.NAME)
    Next rpt
     
    strMsg = " -- Reports listing complete -- "
     
Procedure_Done:
    dmwListAllReports = strMsg
    Exit Function
     
Error_Handler:
    strMsg = Err.Number & " " & Err.Description
    Resume Procedure_Done
     
    End Function

