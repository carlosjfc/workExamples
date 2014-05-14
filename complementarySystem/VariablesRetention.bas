Attribute VB_Name = "VariablesRetention"
Option Compare Database

Sub RetentionParameterQueryDAO(pRetentionDate As String, pRetentionProgram As String)

  Const cstrQueryName As String = "Basics: Parameters"
  Dim dbs As DAO.Database
  Dim qdf As DAO.QueryDef


  Set dbs = CurrentDb()
  On Error Resume Next
                                    '  is run for the first time
  With dbs              '  it would be better to check to see if the
                                    '  querydef exists and then delete it
      .QueryDefs.Delete (cstrQueryName)
                                    '  createquerydef command line follows
    Set qdf = .CreateQueryDef(cstrQueryName)
    qdf.Parameters("RetentionDate") = pRetentionDate
    qdf.Parameters("RetentionProgram") = pRetentionProgram
    .Close
End With
  'Set qdf = dbs.QueryDefs(cstrQueryName)
  'qdf.Parameters("RetentionDate") = pRetentionDate
  'qdf.Parameters("RetentionProgram") = pRetentionProgram
 ' qdf.Close
  dbs.Close
End Sub
