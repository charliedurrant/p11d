Attribute VB_Name = "Display"
Option Explicit

Public Function ShowDebugPopupex() As Boolean
  On Error Resume Next
  Unload frmAbout
  Unload frmMSGOKCancel
  Unload frmErr
  Call frmDebugMenu.PopupMenu(frmDebugMenu.mnuDebug)
  ShowDebugPopupex = True
End Function

Public Sub ShowSystemInfo()
  Dim frmsys As Form
  Set frmsys = New frmEnvir
  If UpdateSys(frmsys) Then
    frmsys.Show vbModal
  End If
  Set frmsys = Nothing
End Sub

Public Sub ShowAppInfo()
  Dim li As ListItem
  Dim sItem As staticcls
  Dim sFile As String
  Dim sPath As String
  Dim sExe As String

  On Error GoTo ShowAppInfo_err
  frmAppInfo.lblSysInfo(0).Caption = mAppExeName & " (" & GetModuleName(ghInstance) & ")"
  frmAppInfo.lblSysInfo(1).Caption = mHomeDirectory
  frmAppInfo.lblSysInfo(2).Caption = mAppCmdParam
  frmAppInfo.lblSysInfo(3).Caption = mAppVersion
  frmAppInfo.lblSysInfo(4).Caption = mAppVersion
  sFile = GetModuleName(ghInstance, True)
  If FileExistsEx(sFile, False, False) Then
    frmAppInfo.lblSysInfo(5).Caption = Format$((FileLen(sFile)), "#,###,##0") & " Bytes "
    frmAppInfo.lblSysInfo(6).Caption = Format$(FileDateTime(sFile), "hh:mm:ss dd/mm/yyyy ")
  End If
  sExe = App.EXEName
  sPath = App.Path
  If right$(sPath, 1) = "\" Then
    sPath = left$(sPath, Len(sPath) - 1)
  End If
  frmAppInfo.lvStatics.ListItems.Clear
  For Each sItem In gstaticdata
    Set li = frmAppInfo.lvStatics.ListItems.Add(, , sItem.Name)
    li.SubItems(1) = sItem.Value
  Next
  Call AutoWidthListViewEx(frmAppInfo.lvStatics, True)
  frmAppInfo.Show vbModal
  
ShowAppInfo_end:
  Exit Sub
  
ShowAppInfo_err:
  Resume ShowAppInfo_end
End Sub

Public Sub CentreInFormEx(FormToCentre As Object, CentreOnForm As Object)
  Dim X As Single, Y As Single
  
  On Error Resume Next
  If Not CentreOnForm Is Nothing Then
    X = (CentreOnForm.Width - FormToCentre.Width) / 2
    Y = (CentreOnForm.Height - FormToCentre.Height) / 2
    If Not FormToCentre.MDIChild Then
      X = CentreOnForm.left + X
      Y = CentreOnForm.top + Y
    End If
  Else
    X = (VB.Screen.Width - FormToCentre.Width) / 2
    Y = (VB.Screen.Height - FormToCentre.Height) / 2
  End If
  If X < 0 Then X = 0
  If Y < 0 Then Y = 0
  FormToCentre.left = X
  FormToCentre.top = Y
End Sub

  'OutsideHours = do you need the password outside work hours
Public Function GetTCSPasswordEx(ByVal OutsideHours As Boolean, Contact As String, ByVal Reason As String) As Boolean
  Dim d0 As Date, nday As Integer, nTime As Double, AllowNoPassword As Boolean
  Static PrevPassword As Long
    
  On Error GoTo GetTCSPasswordEx_err
  d0 = Now
  AllowNoPassword = False
  If Not OutsideHours Then
    nday = Weekday(d0)
    AllowNoPassword = (nday = vbSunday) Or (nday = vbSaturday)
    nTime = DatePart("h", d0)
    nTime = nTime + (DatePart("n", d0) / 60)
    AllowNoPassword = AllowNoPassword Or ((nTime < 9) Or (nTime > 17.5))
    If AllowNoPassword Then PrevPassword = GetPassword_Daily(UNDATED)
  End If
  
  'Initialise the dialog box
  frmPassw.Caption = gPasswordTitle
  frmPassw.lblPrompt.Caption = gPasswordPrompt
  
  frmPassw.txtPassword = CStr(PrevPassword)
  frmPassw.txtPassword.SelStart = 0
  frmPassw.txtPassword.SelLength = Len(frmPassw.txtPassword.Text)
  frmPassw.lblInfoDate = "Please enter the password for " & Format$(d0, "dd mmmm yyyy") & "."
  frmPassw.lblContact = Contact
  If AllowNoPassword Then
    frmPassw.txtPassword.Enabled = False
    frmPassw.lblInfoDate = "Outside hours no password required"
  End If
 
  'Show password dialog box until correct password or
  'the dialog box is cancelled
  Do Until GetTCSPasswordEx
    frmPassw.Show vbModal
    If Not frmPassw.PasswordOk Then Exit Do
    
    'Check password is correct
    PrevPassword = -1
    If IsNumeric(frmPassw.txtPassword.Text) Then
      PrevPassword = CLng(frmPassw.txtPassword.Text)
    End If
    If PrevPassword = GetPassword_Daily(UNDATED) Then
      Reason = Trim$(Reason)
      If (Not m_INotifyTCSPassword Is Nothing) And (Len(Reason) > 0) Then
        Call m_INotifyTCSPassword.notify(-1, -1, Reason)
      End If
      GetTCSPasswordEx = True
    Else
      'If password is incorrect then show info box and then show
      'password dialog again
      Call ErrorMessageEx(ERR_INFO, Err, "GetTCSPassword", "Invalid Password", "The password you have entered is invalid", False)
      frmPassw.txtPassword.SelStart = 0
      frmPassw.txtPassword.SelLength = Len(frmPassw.txtPassword.Text)
    End If
  Loop
  
GetTCSPasswordEx_end:
  Unload frmPassw
  Set frmPassw = Nothing
  Exit Function
  
GetTCSPasswordEx_err:
  Resume GetTCSPasswordEx_end
End Function

Public Function DisplayMessageEx(InForm As Object, Message As String, Title As String, Optional ByVal OKText As String = "Ok", Optional ByVal CancelText As String = "Cancel") As Boolean
  On Error GoTo DisplayMessageEx_Err
  
  Load frmMSGOKCancel
  DisplayMessageEx = frmMSGOKCancel.displayMsg(InForm, Message, Title, OKText, CancelText)
  Unload frmMSGOKCancel
  Set frmMSGOKCancel = Nothing
  DoEvents
  
DisplayMessageEx_End:
  Exit Function

DisplayMessageEx_Err:
  DisplayMessageEx = False
  Resume DisplayMessageEx_End
End Function

Public Function isMDI_Minimized() As Boolean
  Dim i As Long, frm As Form
  
  On Error Resume Next
  isMDI_Minimized = False
  If vbg Is Nothing Then Exit Function
  For i = (vbg.Forms.Count - 1) To 0 Step -1
    Set frm = vbg.Forms(i)
    If TypeOf frm Is MDIForm Then
      isMDI_Minimized = (frm.WindowState = vbMinimized)
      Exit Function
    End If
 Next i
End Function
