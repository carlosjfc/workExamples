Attribute VB_Name = "TableExists"
Option Compare Database

Public Function TableExists(tableName As String) As Boolean
'=================================================  ============================
' hlfUtils.TableExists
'-----------------------------------------------------------------------------
' Copyright by Heather L. Floyd - Floyd Innovations - www.floydinnovations.com
' Created 08-01-2005
'-----------------------------------------------------------------------------
' Purpose:  Checks to see whether the named table exists in the database
'-----------------------------------------------------------------------------
' Parameters:
' ARGUEMENT             :   DESCRIPTION
'-----------------------------------------------------------------------------
' TableName (String)    :   Name of table to check for
'-----------------------------------------------------------------------------
' Returns:  True, if table found in current db, False if not found.
'=================================================  ============================

Dim strTableNameCheck
On Error GoTo ErrorCode

'try to assign tablename value
strTableNameCheck = CurrentDb.TableDefs(tableName)

'If no error and we get to this line, true
TableExists = True

ExitCode:
    On Error Resume Next
    Exit Function

ErrorCode:
    Select Case Err.Number
        Case 3265  'Item not found in this collection
            TableExists = False
            Resume ExitCode
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "hlfUtils.TableExists"
            'Debug.Print "Error " & Err.number & ": " & Err.Description & "hlfUtils.TableExists"
            Resume ExitCode
    End Select

End Function
