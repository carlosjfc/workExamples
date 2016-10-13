Attribute VB_Name = "NewRibbonFunctions"
Option Compare Database
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                                    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                                    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
                                     ByVal lpsz2 As String) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
                                    (ByVal hwnd As Long, ByVal wMsg As Long, _
                                     ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                    (ByVal hwnd As Long, ByVal lpOperation As String, _
                                     ByVal lpFile As String, ByVal lpParameters As String, _
                                     ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Const WM_GETTEXTLENGTH = 14
Const WM_GETTEXT = 13
Private Function GetMyWindow()
    Dim GetWnd As Long 'hWnd

    Const WndClassA = "SciCalc"  'window name for Nt4 and XP
    Const WndClassB = "Edit"     'Edit control name for XP
    'Const WndClassB = "Static"   'Edit control name for Nt4

    GetWnd = FindWindow(WndClassA, "Calculator")
    If GetWnd Then
        GetWnd = FindWindowEx(GetWnd, ByVal 0&, WndClassB, vbNullString)
        If GetWnd > 0 Then
          GetMyWindow = GetWnd
        End If
    End If
End Function

Public Function ShoCalc()
Dim lHnd As Long
  lHnd = ShellExecute(0, "open", "Calc.exe", 0, 0, SW_SHOWNORMAL)
End Function

Private Sub WindowText(window_hwnd As Long, Ctrl As control)
Dim txtlen As Long
Dim txt As String

    'WindowText = ""
    If window_hwnd = 0 Then Exit Sub
    
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
    If txtlen = 0 Then Exit Sub
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
    ''WindowText = Left$(txt, txtlen)
    Ctrl.Value = Left$(txt, txtlen)
    
End Sub

Public Function fncLoadRibbon()
Dim rsRib As DAO.Recordset
Dim db As DAO.Database
On Error GoTo fError
'-----------------------------------------------------------------
'This function loads the ribbons stored in the table dbo_DCB_CS_tblRibbons,
'that must be called by the macro AutoExec
'
'Create the macro AutoExec, select the action RunCode
'and type the function name in the argument : fncLoadRibbon()
'------------------------------------------------------------------
Set db = CurrentDb()
Set rsRib = db.OpenRecordset("dbo_DCB_CS_tblRibbons", dbOpenDynaset, dbSeeChanges)
If (Not (rsRib.EOF And rsRib.BOF)) Then
    rsRib.MoveFirst
    Do While Not rsRib.EOF
      Application.LoadCustomUI rsRib!RibbonName, rsRib!RibbonXML
      'Debug.Print rsRib!RibbonName & " " & rsRib!RibbonXML
      'Debug.
      rsRib.MoveNext
    Loop
End If

fExit:
  Exit Function
fError:
  Select Case Err.Number
    Case 3078
      MsgBox "Table not found...", vbInformation, "Warning"
    Case Else
'      MsgBox "Error miooooo: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, _
'      vbCritical, "Warning", Err.HelpFile, Err.HelpContext
  End Select
  Resume fExit:
End Function

