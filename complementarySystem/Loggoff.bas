Attribute VB_Name = "Loggoff"
Option Compare Database

Public Function logOff()
'MsgBox userID
  Dim db As DAO.Database
  'Dim prop As DAO.Property
  Set db = CurrentDb()
  'db.Properties.Refresh
  SetObjProperty db, "userID", dbText, "a"
  SetObjProperty db, "id_profile", dbText, "a"
  SetObjProperty db, "id_ribbon", dbText, "a"
  SetObjProperty db, "RibbonName", dbText, "a"
  SetObjProperty db, "userPssw", dbText, "a"
 ' Me.txtPassword = ""
  'Me.txtuser = ""
  
  'DoCmd.Close
  CurrentDb.Properties("userID") = "a"
  CurrentDb.Properties("id_profile") = "a"
  CurrentDb.Properties("id_ribbon") = "a"
  CurrentDb.Properties("RibbonName") = "a"
  CurrentDb.Properties("userPssw") = "a"
  
  CurrentDb().Properties("RibbonName") = "Empty"
  DoCmd.OpenForm "frmLogin", acNormal
  
  
End Function
