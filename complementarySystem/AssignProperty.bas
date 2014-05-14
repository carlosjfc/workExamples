Attribute VB_Name = "AssignProperty"
Option Compare Database

Sub SetObjProperty(pObject As Object, pProperty As String, pType As Integer, pValue As Variant)
Const PROPERTY_NOT_FOUND As Long = 3270
Dim prp As Property
'
On Error GoTo SetObjProperty_Err
'
pObject.Properties(pProperty) = pValue
pObject.Properties.Refresh

SetObjProperty_Exit:
Set prp = Nothing
Exit Sub

SetObjProperty_Err:
If Err.Number = PROPERTY_NOT_FOUND Then
With pObject
       Set prp = .CreateProperty(pProperty, pType, pValue)
      .Properties.Append prp
      .Properties.Refresh
End With
Resume SetObjProperty_Exit
Else
      MsgBox Err.Number & ": " & Err.Description, vbCritical, "SetObjProperty"
      Resume SetObjProperty_Exit
End If

End Sub
