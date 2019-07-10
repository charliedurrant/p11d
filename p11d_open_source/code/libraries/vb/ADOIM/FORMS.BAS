Attribute VB_Name = "FormFunctions"
Option Explicit
Private mInUpdateType As Boolean

Public Sub FrameEnable(frm As Form, Fram As Frame, Enable As Boolean)
  Dim Ctrl As Control
  'On Error Resume Next
  For Each Ctrl In frm
    If (Not TypeOf Ctrl Is Data) And (Not TypeOf Ctrl Is Menu) Then
      If Ctrl.Container Is Fram Then
        Ctrl.Enabled = Enable
        'Call FrameEnable(Frm, Ctrl, Enable)
        If TypeOf Ctrl Is Frame Then Call FrameEnable(frm, Ctrl, Enable)
      End If
    End If
  Next Ctrl
End Sub

Public Function InUpdate() As Boolean
  InUpdate = mInUpdateType
End Function

Public Sub SetUpdateType(frmDest As Frm_Dest, ByVal UpdateType As IMPORT_UPDATE_TYPE, ByVal AllowChangeUpdateType As Boolean, Optional ByVal LockUpdateType As IMPORT_UPDATE_TYPE = [_NO_LOCK_UPDATES])
  Dim AllowAddNew As Boolean
    
  On Error Resume Next
  If mInUpdateType Then Exit Sub
  mInUpdateType = True
  AllowAddNew = Not ((UpdateType And UPDATEONLY_FLAG) = UPDATEONLY_FLAG)
  UpdateType = UpdateType And (Not UPDATEONLY_FLAG) ' Xor value
  If AllowAddNew Then
    frmDest.Chk_Add.Value = vbChecked
  Else
    frmDest.Chk_Add.Value = vbUnchecked
  End If
  If (LockUpdateType And UPDATEONLY_FLAG) <> UPDATEONLY_FLAG Then
    frmDest.Opt_Update(0).Enabled = False
  Else
    frmDest.Opt_Update(0).Enabled = AllowAddNew
  End If
  If UpdateType = NO_UPDATES Then
    frmDest.Opt_Update(0).Value = True
    frmDest.Chk_Add.Enabled = False
  End If
  If UpdateType = UPDATE_FIRST Then
    frmDest.Opt_Update(1).Value = True
  End If
  If UpdateType = UPDATE_ONLYFIRST Then
    frmDest.Opt_Update(2).Value = True
  End If
  If UpdateType = UPDATE_ALL Then
    frmDest.Opt_Update(3).Value = True
  End If
  
  Call FrameEnable(frmDest, frmDest.Fra_Update, AllowChangeUpdateType)
  frmDest.Fra_Update.Enabled = AllowChangeUpdateType
  If AllowChangeUpdateType Then
    frmDest.Chk_Add.Enabled = ((LockUpdateType And UPDATEONLY_FLAG) = UPDATEONLY_FLAG)
    frmDest.Opt_Update(0).Enabled = ((LockUpdateType And NO_UPDATES) = NO_UPDATES)
    frmDest.Opt_Update(1).Enabled = ((LockUpdateType And UPDATE_FIRST) = UPDATE_FIRST)
    frmDest.Opt_Update(2).Enabled = ((LockUpdateType And UPDATE_ONLYFIRST) = UPDATE_ONLYFIRST)
    frmDest.Opt_Update(3).Enabled = ((LockUpdateType And UPDATE_ALL) = UPDATE_ALL)
    If frmDest.Chk_Add.Value = vbUnchecked Then frmDest.Opt_Update(0).Enabled = False
    If frmDest.Opt_Update(0).Value Then frmDest.Chk_Add.Enabled = False
  End If
  mInUpdateType = False
End Sub

Public Function GetUpdateType(frmDest As Frm_Dest) As IMPORT_UPDATE_TYPE
  Dim AllowAddNew As Boolean
    
  If mInUpdateType Then Exit Function
  mInUpdateType = True
  If frmDest.Opt_Update(0).Value Then
    GetUpdateType = NO_UPDATES
  End If
  If frmDest.Opt_Update(1).Value Then
    GetUpdateType = UPDATE_FIRST
  End If
  If frmDest.Opt_Update(2).Value Then
    GetUpdateType = UPDATE_ONLYFIRST
  End If
  If frmDest.Opt_Update(3).Value Then
    GetUpdateType = UPDATE_ALL
  End If
  If (frmDest.Chk_Add.Value <> vbChecked) Then GetUpdateType = (GetUpdateType Or UPDATEONLY_FLAG)
  mInUpdateType = False
End Function


