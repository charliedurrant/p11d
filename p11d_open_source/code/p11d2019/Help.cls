VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ShownTopic As Boolean
Private m_HelpForms As Dictionary
Private m_DefaultLink As String
Private Sub Class_Initialize()
  Call PopulateHelpLinks
End Sub
Private Sub Class_Terminate()
  If Not m_HelpForms Is Nothing Then Set m_HelpForms = Nothing
  
  If m_ShownTopic Then
    ' close help file
    Call HtmlHelp(0, "", HH_CLOSE_ALL, 0)
    ' fix - give time for thread to return - otherwise causes GPF (ref: http://helpware.net/FAR/far_faq.htm#HH_CLOSE_ALL)
    Sleep (2000)
  End If
      
End Sub
Public Property Get HelpFile() As String
  HelpFile = FullPath(AppPathDos) & S_SYSTEMDIR_HELP & "P11D32.chm"
End Property
Public Property Get HelpForms() As Dictionary
  Set HelpForms = m_HelpForms
End Property
Public Sub ShiftControlF11()
  Dim sFrm As String, sCtrl As String
  
  On Error GoTo err_err
  
  Dim s As String
    
  Call ActiveObjName(sFrm, sCtrl)
    
  s = "Form:" & sFrm & vbCrLf
  
  If (Len(sCtrl) > 0) Then
    s = s & "Control:" & sCtrl & vbCrLf
  End If
  s = s & "Help link:" & ActiveHelpLink()
  Call MsgBox(s)
    
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "ShowHelp", "Show help", "Error displaying help.")
  Resume err_end
End Sub
Private Function ActiveHelpLink() As String
  Dim sActiveHelpLink As String
  Dim sFrm As String, sCtrl As String
  Dim cHF As HelpForm, cHC As HelpControl
  
  
  Call ActiveObjName(sFrm, sCtrl)
  ' check if active form's help exist
  If m_HelpForms.Exists(sFrm) Then
    ' get helpform
    Set cHF = m_HelpForms.Item(sFrm)
    ' check if active control help exist
    If cHF.Controls.Exists(sCtrl) Then
      ' get helpcontrol
      Set cHC = cHF.Controls.Item(sCtrl)
      ' check if helplink for control exist
      If Len(cHC.HelpLink) > 0 Then
        ' use control helplink
        sActiveHelpLink = Trim$(cHC.HelpLink)
      End If
    End If
    If Len(sActiveHelpLink) = 0 Then
      ' use form helplink as either control or its helplink did not exist
      sActiveHelpLink = Trim$(cHF.HelpLink)
    End If
    
    If Len(sActiveHelpLink) = 0 Then
      sActiveHelpLink = m_DefaultLink
    End If
  Else
    sActiveHelpLink = m_DefaultLink
  End If
  ActiveHelpLink = sActiveHelpLink
End Function
Public Sub ShowHelp(ByVal sDisplaySpecificHelpLink As String)
  On Error GoTo err_err
  
  Dim cHF As HelpForm
  Dim cHC As HelpControl
  
  Call xSet("ShowHelp")
    
  If Len(sDisplaySpecificHelpLink) = 0 Then
    'get active form&/control names
    sDisplaySpecificHelpLink = ActiveHelpLink()
  End If
  ' display the help via HelpLink control
  Call ShowHelpEx(sDisplaySpecificHelpLink)
  
err_end:
  Call xReturn("ShowHelp")
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "ShowHelp", "Show help", "Error displaying help.")
  Resume err_end
  Resume
End Sub

Public Sub ActiveObjName(ByRef sFrmName As String, ByRef sCtrlName As String)
  sFrmName = UCASE$(Screen.ActiveForm.Name)
  If Not Screen.ActiveControl Is Nothing Then
    sCtrlName = ControlName(Screen.ActiveControl)
  Else
    sCtrlName = ""
  End If
End Sub

Private Function ControlName(ByVal oControl As Control) As String
  Dim iIndex As Integer
  iIndex = -1
  ControlName = UCASE$(oControl.Name)
  On Error Resume Next
  If IsControlArray(oControl) Then
    iIndex = oControl.Index
  End If
  
  
  'iIndex = oControl.Index
  If iIndex >= 0 Then ControlName = ControlName & "(" & iIndex & ")"
End Function


Public Function IsControlArray(objCtrl As Object) As Boolean
    IsControlArray = Not objCtrl.Parent.Controls(objCtrl.Name) Is objCtrl
End Function

Public Sub PopulateHelpLinks()
  On Error GoTo PopulateHelpLinks_Err
  Dim s As String
  Dim db As Database
  Dim rsHelpLinks As Recordset
  Dim lLastFormPID As Long
  Dim cHF As HelpForm
  Dim cHC As HelpControl

  Call xSet("PopulateHelpLinks")

  Set db = p11d32.PDDB(p11d32.P11DWS)
  Set rsHelpLinks = db.OpenRecordset(sql.Queries(SELECT_HELPLINKS), dbOpenSnapshot, dbFailOnError)
  ' end sub if no records
  If (rsHelpLinks.EOF And rsHelpLinks.BOF) Then GoTo PopulateHelpLinks_End
  
  Set m_HelpForms = New Dictionary
  lLastFormPID = -1
  Do While Not rsHelpLinks.EOF
    If (lLastFormPID <> rsHelpLinks.Fields("form_pid")) Then
      ' create HelpForm and add to dictioanry
      Set cHF = New HelpForm
      cHF.HelpLink = rsHelpLinks.Fields("form_helplink")
      s = UCASE$(Trim$(rsHelpLinks.Fields("form_name")))
      Call m_HelpForms.Add(UCASE$(s), cHF)
      lLastFormPID = rsHelpLinks.Fields("form_pid")
    End If
    ' add form's control(s) if either helplink or tooltip for it exist
    If Len(rsHelpLinks.Fields("control_helplink")) > 0 Or Len(rsHelpLinks.Fields("tooltip")) > 0 Then
      Set cHC = New HelpControl
      cHC.HelpLink = "" & rsHelpLinks.Fields("control_helplink")
      cHC.Tooltip = "" & rsHelpLinks.Fields("tooltip")
      s = UCASE$(Trim$(rsHelpLinks.Fields("control_name")))
      Call cHF.Controls.Add(s, cHC)
    End If
    rsHelpLinks.MoveNext
  Loop

PopulateHelpLinks_End:
  Call xReturn("PopulateHelpLinks")
  Exit Sub
PopulateHelpLinks_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PopulateHelpLinks", "Populate HelpLinks", "Error setting up help links")
  Resume PopulateHelpLinks_End
End Sub

Private Sub AssignControlTooltip(ByVal oForm As Form, bCheckAssigned As Boolean)
  On Error GoTo err_err
  
  Dim ctrl As Control
  Dim sFormName As String
  Dim sCtrlName As String
  Dim cHF As HelpForm
  Dim cHC As HelpControl
  
  Call xSet("AssignControlTooltip")

  sFormName = UCASE$(oForm.Name)
  If Not m_HelpForms.Exists(sFormName) Then GoTo err_end
  Set cHF = m_HelpForms.Item(sFormName)
  If (bCheckAssigned And cHF.bIsAssigned) Then GoTo err_end

  For Each ctrl In oForm.Controls
    ' get ctrl name and index
    sCtrlName = ControlName(ctrl)
    ' chk if control exist in help control
    If cHF.Controls.Exists(sCtrlName) Then
      ' get help control
      Set cHC = cHF.Controls.Item(sCtrlName)
      ' chk if the tooltip exist for control
      If Len(cHC.Tooltip) > 0 Then
        ' apply toltip
        Call LetToolTipText(ctrl, cHC.Tooltip)
      End If
    End If
  Next ctrl
  cHF.bIsAssigned = True

err_end:
  Call xReturn("AssignControlTooltip")
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "AssignControlTooltip", "AssignControlTooltip", "Error assigning tooltip to form controls")
  Resume err_end
End Sub

Private Sub LetToolTipText(ByVal oControl As Control, ByVal sText As String)
  On Error Resume Next
  oControl.ToolTipText = sText
End Sub

Public Sub ShowForm(ByRef oForm As Form, Optional iFormModal As FormShowConstants)
  On Error GoTo err_err
  
  Call AssignControlTooltip(oForm, iFormModal = vbModeless)
  
      oForm.Show iFormModal
  
  
err_end:
  Call xReturn("ShowForm")
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "ShowForm", "ShowForm", "Error showing modal form")
  Resume err_end
End Sub
Public Sub ShowHelpEx(Optional ByVal sHelpPage As String)
  Dim sHelpFile As String
  
  
  If Len(HelpFile) = 0 Then Call Err.Raise(380, "ShowHelp", "Help file is not set.")
  
  sHelpPage = Trim$(sHelpPage)
  
  sHelpFile = HelpFile & ">main"
  
  Call HtmlHelpString(0, sHelpFile, HH_DISPLAY_TOPIC, sHelpPage)
  Call HtmlHelpString(0, sHelpFile, HH_SYNC, sHelpPage)
  
  m_ShownTopic = True
End Sub

Public Property Let DefaultLink(ByVal NewValue As String)
  m_DefaultLink = NewValue
End Property

