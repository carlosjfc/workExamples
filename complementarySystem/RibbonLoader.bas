Attribute VB_Name = "RibbonLoader"
Option Compare Database

Function CreateFormButtons()
  Dim xml As String
  xml = ""
  xml = _
   "<customUI xmlns=""http://schemas.microsoft.com/" & _
   "office/2006/01/customui"">" & vbCrLf & _
   "  <ribbon startFromScratch=""false"">" & vbCrLf & _
   "    <tabs>" & vbCrLf & _
   "      <tab id=""DemoTab"" label=""LoadCustomUI Demo"">" & _
     vbCrLf & _
   "        <group id=""loadFormsGroup"" label=""Load Forms"">" & _
     vbCrLf & _
   "{0}" & vbCrLf & _
   "        </group>" & vbCrLf & _
   "      </tab>" & vbCrLf & _
   "    </tabs>" & vbCrLf & _
   "  </ribbon>" & vbCrLf & _
   "</customUI>"

  Dim template As String
  template = "<button id=""load{0}Button"" " & _
   "label=""Load {0}"" onAction=""HandleOnAction"" " & _
   "tag=""{0}""/>" & vbCrLf
  
  Dim formContent As String
  Dim frm As AccessObject
  For Each frm In CurrentProject.AllForms
    formContent = formContent & _
     replace(template, "{0}", frm.NAME)
  Next frm
  
  xml = replace(xml, "{0}", formContent)
  Debug.Print xml
  Debug.Print "test test test"
  On Error Resume Next
  ' If you call this code from the AutoExec macro,
  ' the only way it can fail is if you have a
  ' customization with the same name in the
  ' USysRibbons table.
  Application.LoadCustomUI "FormNames", xml
  'Debug.Print xml
End Function
'Public Function OpenMyform(strF As String)
'
'  DoCmd.OpenForm strF
'
'End Function

Function CreateRibbon()
Dim xml As String
'Select Microsoft XML, v3.0 Reference from the list
'Dim XMLDOC As New MSXML2.DOMDocument
'Set XMLDOC = Server.CreateObject("Msxml2.DOMDocument.3.0")
'XMLDOC.async = False
'XMLDOC.validateOnParse = False
'XMLDOC.Load ("C:\Documents and Settings\AdminAst\My Documents\cacharreo\config\Ribbon.xml")
'xml = XMLDOC.Text
'Debug.Print xml
Dim FSO As Object
Dim TS As Object
Dim c As Long
Dim Data As Variant
Data = ""

Set FSO = CreateObject("Scripting.FileSystemObject")
Set TS = FSO.OpenTextFile("C:\Documents and Settings\AdminAst\My Documents\cacharreo\config\Ribbon.xml", 1, False, -2)
c = 1
    Do Until TS.AtEndOfStream
        Data = Data + TS.ReadLine
    Loop
TS.Close

  xml = Data
  
  Dim template As String
  template = "<button id=""load{0}Button"" " & _
   "label=""Load {0}"" onAction=""HandleOnAction"" " & _
   "tag=""{0}""/>" & vbCrLf
  
  Dim formContent As String
  Dim frm As AccessObject
  For Each frm In CurrentProject.AllForms
    formContent = formContent & _
     replace(template, "{0}", frm.NAME)
  Next frm
  
  xml = replace(xml, "{0}", formContent)
  'Debug.Print xml
  On Error GoTo Proximo:
  ' If you call this code from the AutoExec macro,
  ' the only way it can fail is if you have a
  ' customization with the same name in the
  ' USysRibbons table.
  Application.LoadCustomUI "MyRibbon", xml
Proximo:
Debug.Print Err.Number & " " & Err.Description
Resume Next
End Function

Public Sub HandleOnAction(control As IRibbonControl)
    ' Load the specified form, and set its
    ' RibbonName property so that it displays
    ' the custom UI.
    DoCmd.OpenForm control.Tag
    Forms(control.Tag).RibbonName = "FormNames"
End Sub

