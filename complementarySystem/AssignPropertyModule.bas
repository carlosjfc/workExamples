Attribute VB_Name = "AssignPropertyModule"
Option Compare Database

Function SetPropertyDAO(obj As Object, strPropertyName As String, intType As Integer, _
    varValue As Variant, Optional strErrMsg As String) As Boolean
On Error GoTo ErrHandler
    'Purpose:   Set a property for an object, creating if necessary.
    'Arguments: obj = the object whose property should be set.
    '           strPropertyName = the name of the property to set.
    '           intType = the type of property (needed for creating)
    '           varValue = the value to set this property to.
    '           strErrMsg = string to append any error message to.
    
    If HasProperty(obj, strPropertyName) Then
        obj.Properties(strPropertyName) = varValue
    Else
        obj.Properties.Append obj.CreateProperty(strPropertyName, intType, varValue)
    End If
    SetPropertyDAO = True

ExitHandler:
    Exit Function

ErrHandler:
    strErrMsg = strErrMsg & obj.NAME & "." & strPropertyName & " not set to " & varValue & _
        ". Error " & Err.Number & " - " & Err.Description & vbCrLf
    Resume ExitHandler
End Function

Public Function HasProperty(obj As Object, strPropName As String) As Boolean
    'Purpose:   Return true if the object has the property.
    Dim varDummy As Variant
    
    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function


