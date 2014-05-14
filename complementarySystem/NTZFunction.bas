Attribute VB_Name = "NTZFunction"
Option Compare Database

Public Function NTZ(anyValue As String) As Double

If IsNull(anyValue) Or IsEmpty(anyValue) Or IsError(anyValue) Or IsMissing(anyValue) Then
    NTZ = 0
Else
    NTZ = anyValue
End If

End Function
