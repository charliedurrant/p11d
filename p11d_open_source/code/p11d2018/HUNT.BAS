Attribute VB_Name = "Hunt"
Option Explicit

Public Sub ShowHunt()
  On Error GoTo ShowHunt_ERR
  
  
  F_Hunt.Show vbModal
  
ShowHunt_END:
  'Set F_Hunt = Nothing
  Exit Sub
ShowHunt_ERR:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "ShowHunt"), "ShowHunt", "Error in ShowHunt")
  Resume ShowHunt_END
End Sub
