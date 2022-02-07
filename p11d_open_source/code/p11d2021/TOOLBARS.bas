Attribute VB_Name = "ToolBars"
Option Explicit

Public Enum ToolbarButtons
  TBR_OPEN_EMPLOYER = 1
  TBR_EDIT_EMPLOYER
  TBR_SEPERATOR1
  TBR_REFRESH_EMPLOYERS
  TBR_CONFIRM
  TBR_UNDO
  TBR_SEPERATOR2
  TBR_ADD_BENEFIT
  TBR_REMOVE_BENEFIT
  TBR_SEPERATOR3
  TBR_PRINT
  TBR_PREVIEW
  TBR_EMPLOYERSCREEN
  TBR_SEPERATOR4
  TBR_SHAREDVANS
  TBR_EMPLOYEESCREEN
End Enum

Public Enum MNU_FILE
  MNU_FILE_NEW = 0
  MNU_FILE_OPEN
  MNU_FILE_EDIT
  MNU_FILE_DELETE
  MNU_FILE_PRINT
  MNU_FILE_SEPERATOR1
  MNU_FILE_IMPORT
  MNU_FILE_ERROR_LOG
  MNU_FILE_ELECTRONIC_SUBMISSION
  MNU_FILE_REFRESHEMPLOYERS
  MNU_FILE_CHANGEDIRECTORY
  MNU_FILE_EMPLOYEE_LETTER
  MNU_FILE_EMPLOYER
  MNU_FILE_PASSWORD
  MNU_FILE_TOOLS
End Enum

Public Enum MNU_EMPLOYEE
  MNU_EMPLOYEE_CONFIRM = 0
  MNU_EMPLOYEE_UNDO
  MNU_EMPLOYEE_SEPERATOR1
  MNU_EMPLOYEE_DETAILS
  MNU_EMPLOYEE_ADD
  MNU_EMPLOYEE_DELETE
  MNU_EMPLOYEE_GOTO
End Enum

Public Enum BenList
  MNU_ADD = 0
  MNU_KILL
  MNU_COPY
  MNU_PASTE
  MNU_BENSEP1
End Enum

Public Enum BenList2
  MNU_A
  MNU_C
  MNU_D
  MNU_E
  MNU_F
  MNU_G
End Enum

Public Enum BenList3
  MNU_I
  MNU_J
  MNU_K
  MNU_L
  MNU_M
  MNU_N
End Enum

Public Enum BenListTB
  MNU_BLASSETST = 0
  MNU_BLPAYMENTS
  MNU_BLNOTIONAL
  MNU_BLVANS
  MNU_BLSERVICES
  MNU_BLASSETSP
  MNU_BLSHARES
  MNU_BLSUBS
  MNU_BLNURSERY
  MNU_BLTAXPAID
  MNU_BLOTHERITEMS
  MNU_BLTRAVEL
  MNU_BLENTS
  MNU_BLGEN
  MNU_BLCHAUFFEUR
  MNU_BLOTHERP
  MNU_BLTRANS
  MNU_BLRELOC
  MNU_BLEECAR
End Enum

Public Enum BenListN
  MNU_SUBS = 0
  MNU_NURS
  MNU_EDU
  MNU_TAXPAID
  MNU_SHARES
  MNU_OTHERS
End Enum

Public Enum BenListO
  MNU_TRAVEL = 0
  MNU_ENTS
  MNU_GENERAL
  MNU_HPHONE
  MNU_NQRELOC
  MNU_CHAUF
  MNU_O
End Enum

Public Enum DisplayType
  [_INVALID_DISPPLAY_TYPE] = 0
  D_EMPLOYER_ON = 1
  D_EMPLOYER_OFF
  D_BENEFIT
  D_EMPLOYEES
End Enum

Public Sub BenefitToolBar(i As Long, ObjIndex As Long)
  
  On Error GoTo benefittoolbar_Err
  Call xSet("benefittoolbar")
  
  
  If Not p11d32.CurrentEmployer.MoveMenuUpdateEmployee Then GoTo benefittoolbar_End
  
  Select Case i
    Case 1
      Call BenScreenSwitch(BC_ALL)
    Case 2
      Call BenScreenSwitch(BC_COMPANY_CARS_F)
    Case 3
      Call BenScreenSwitch(BC_EMPLOYEE_CAR_E)
    Case 4
      Call BenScreenSwitch(BC_PHONE_HOME_N)
    Case 5
      Call BenScreenSwitch(BC_PRIVATE_MEDICAL_I)
    Case 6
        Call BenScreenSwitch(BC_VOUCHERS_AND_CREDITCARDS_C)
    Case 7
      MDIMain.mnuHouseButton.Visible = True
      Call MDIMain.PopupMenu(MDIMain.mnuHouseButton)
      MDIMain.mnuHouseButton.Visible = False
    Case 8
      'AM
      Call BenScreenSwitch(BC_LOAN_OTHER_H)
'      Call MDIMain.PopupMenu(MDIMain.mnuLoans)
    Case 9
      MDIMain.mnuOtherButton.Visible = True
      Call MDIMain.PopupMenu(MDIMain.mnuOtherButton)
      MDIMain.mnuOtherButton.Visible = False
    Case Else
      ECASE "Unknown button"
  End Select
  
  
  
  
benefittoolbar_End:
  Call xReturn("benefittoolbar")
  Exit Sub

benefittoolbar_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToolBar", "Error in Benefit Toolbar", "Unable to complete action")
  Resume benefittoolbar_End
  Resume
End Sub

Public Function DisplayEx(ByVal dt As DisplayType) As Boolean
  Static DT_LAST As DisplayType
  Dim i As Long
  
  On Error GoTo DisplayEx_Err
  Call xSet("DisplayEx")
'MP RV notes - On MDIMain form there are
' 3 toolbars (tbrMain, tbrBenefits, tbrNavigate) handled in ToolBarDisplay proc
' 6 Menues - 3 (File, Employee, Help) handled in individual procs - so any ref to them was moved to individual proc
'            3 (Employer, View, Benefits) handled here in this proc

  If dt = DT_LAST Then GoTo DisplayEx_End
  'CAD trying to convert to something useable
  Call ToolBarDisplay(dt)
  Call DisplayFileMenu(dt)
  Call DisplayEmployeeMenu(dt)
  Call DisplayHelpMenu(dt)  'AM
  
  Select Case dt
    Case D_BENEFIT, D_EMPLOYEES
      MDIMain.mnuEmployer.Visible = True
'MP RV -set in DisplayHelpMenu, always is true      MDIMain.mnuHelp.Visible = True
      MDIMain.mnuBenefits.Visible = True
      MDIMain.mnuBenefitsToolsAbacusExport.Visible = p11d32.ReportPrint.AbacusUDM
      If dt = D_EMPLOYEES Then
        Call MDIMain.CutCopyPasteVisible(False)
        MDIMain.mnuEmployerTransferEmployees.Enabled = True
        MDIMain.mnuView.Visible = True
'MP RV - moved to ToolBarDisplay under Case D_Employees
'MP RV        MDIMain.cmdGoto.Visible = False
'MP RV        MDIMain.cmdDown.Visible = False
'MP RV        MDIMain.cmdUp.Visible = False
      Else
        MDIMain.mnuEmployerTransferEmployees.Enabled = False
'MP RV - moved to ToolBarDisplay under Case D_BENEFIT
'MP RV        MDIMain.cmdGoto.Visible = True
'MP RV        MDIMain.cmdDown.Visible = True
'MP RV        MDIMain.cmdUp.Visible = True
        MDIMain.mnuView.Visible = False
      End If
    Case D_EMPLOYER_ON, D_EMPLOYER_OFF
'MP RV      MDIMain.tbrBenefits.Visible = False
      MDIMain.mnuEmployer.Visible = False
'MP RV      MDIMain.mnuHelp.Visible = True
'MP RV      MDIMain.mnuView.Visible = False
'MP RV removed as was reset below      MDIMain.mnuBenefits.Visible = False
'MP RV      MDIMain.mnuHelp.Visible = True
'MP RV      MDIMain.mnuEmployee.Visible = False
      MDIMain.mnuView.Visible = False
      MDIMain.mnuBenefits.Visible = False
    Case Else
      ECASE ("Invalid display type")
  End Select

  DT_LAST = dt

DisplayEx_End:
  Call xReturn("DisplayEx")
  Exit Function

DisplayEx_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DisplayEx", "Display Ex", "Error setting menus and toolbars.")
  Resume DisplayEx_End
  Resume
End Function

Private Function ToolBarDisplay(ByVal dt As DisplayType) As Boolean
'MP RV changed scope to Private
  Dim i As Long
  Dim lStart As Long, lEnd As Long
  On Error GoTo ToolBarDisplay_Err
  Call xSet("ToolBarDisplay")

  'if display type added the update in UpdateDisplay
  With MDIMain
    Select Case dt
      Case D_EMPLOYER_ON
        .tbrBenefits.Visible = False
        .tbrNavigate.Visible = False
        
        .tbrMain.Buttons(TBR_OPEN_EMPLOYER).Visible = True
        .tbrMain.Buttons(TBR_OPEN_EMPLOYER).Enabled = True
        .tbrMain.Buttons(TBR_EDIT_EMPLOYER).Visible = True
        .tbrMain.Buttons(TBR_SEPERATOR1).Visible = True
        .tbrMain.Buttons(TBR_REFRESH_EMPLOYERS).Visible = True
        
        .tbrMain.Buttons(TBR_CONFIRM).Visible = False
        .tbrMain.Buttons(TBR_UNDO).Visible = False
        .tbrMain.Buttons(TBR_SEPERATOR2).Visible = True
        
        .tbrMain.Buttons(TBR_SEPERATOR3).Visible = True
        .tbrMain.Buttons(TBR_PRINT).Visible = False
        .tbrMain.Buttons(TBR_PREVIEW).Visible = False
        .tbrMain.Buttons(TBR_EMPLOYERSCREEN).Visible = False
        .tbrMain.Buttons(TBR_SEPERATOR4).Visible = False        'MP RV - 2 bars were showing
        .tbrMain.Buttons(TBR_SHAREDVANS).Visible = False
        .tbrMain.Buttons(TBR_EMPLOYEESCREEN).Visible = True
      Case D_EMPLOYER_OFF
        .tbrBenefits.Visible = False
        .tbrNavigate.Visible = False
        
        .tbrMain.Buttons(TBR_OPEN_EMPLOYER).Visible = True
        .tbrMain.Buttons(TBR_OPEN_EMPLOYER).Enabled = False
        .tbrMain.Buttons(TBR_EDIT_EMPLOYER).Visible = False
        .tbrMain.Buttons(TBR_SEPERATOR1).Visible = True
        .tbrMain.Buttons(TBR_REFRESH_EMPLOYERS).Visible = True
        
        .tbrMain.Buttons(TBR_CONFIRM).Visible = False
        .tbrMain.Buttons(TBR_UNDO).Visible = False
        .tbrMain.Buttons(TBR_SEPERATOR2).Visible = True
        
        .tbrMain.Buttons(TBR_SEPERATOR3).Visible = False
        .tbrMain.Buttons(TBR_PRINT).Visible = False
        .tbrMain.Buttons(TBR_PREVIEW).Visible = False
        .tbrMain.Buttons(TBR_EMPLOYERSCREEN).Visible = False
        .tbrMain.Buttons(TBR_SEPERATOR4).Visible = False
        .tbrMain.Buttons(TBR_SHAREDVANS).Visible = False
        .tbrMain.Buttons(TBR_EMPLOYEESCREEN).Visible = False
      Case D_EMPLOYEES
        .tbrBenefits.Visible = True
        .tbrNavigate.Visible = True
        .cmdGoto.Visible = False      'MP RV added here - removed from DisplayEx
        .cmdDown.Visible = False      'MP RV added here - removed from DisplayEx
        .cmdUp.Visible = False        'MP RV added here - removed from DisplayEx
        .chkMoveToNextEmployeeWithBenefit.Visible = False
        .tbrMain.Buttons(TBR_OPEN_EMPLOYER).Visible = False
        .tbrMain.Buttons(TBR_EDIT_EMPLOYER).Visible = False
        .tbrMain.Buttons(TBR_SEPERATOR1).Visible = False
        .tbrMain.Buttons(TBR_REFRESH_EMPLOYERS).Visible = False
        .tbrMain.Buttons(TBR_CONFIRM).Visible = True
        .tbrMain.Buttons(TBR_UNDO).Visible = True
        .tbrMain.Buttons(TBR_SEPERATOR2).Visible = True
        
        .tbrMain.Buttons(TBR_SEPERATOR3).Visible = True
        .tbrMain.Buttons(TBR_PRINT).Visible = True
        .tbrMain.Buttons(TBR_PREVIEW).Visible = True
        .tbrMain.Buttons(TBR_EMPLOYERSCREEN).Visible = True
        .tbrMain.Buttons(TBR_SEPERATOR4).Visible = True
        .tbrMain.Buttons(TBR_SHAREDVANS).Visible = True
        .tbrMain.Buttons(TBR_EMPLOYEESCREEN).Visible = False
      Case D_BENEFIT
        .tbrBenefits.Visible = True  'MP RV added here - removed from DisplayEx
        .tbrNavigate.Visible = True  'MP RV added
        .cmdGoto.Visible = True      'MP RV added here - removed from DisplayEx
        .cmdDown.Visible = True      'MP RV added here - removed from DisplayEx
        .cmdUp.Visible = True        'MP RV added here - removed from DisplayEx
        If CurrentForm Is F_AllBenefits Then
          .chkMoveToNextEmployeeWithBenefit.Visible = False
        Else
          .chkMoveToNextEmployeeWithBenefit.Visible = True
        End If
        .tbrMain.Buttons(TBR_EMPLOYEESCREEN).Visible = True
        
    End Select
  End With

ToolBarDisplay_End:
  Call xReturn("ToolBarDisplay")
  Exit Function

ToolBarDisplay_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ToolBarDisplay", "Tool Bar Display", "Error setting the toolbar.")
  Resume ToolBarDisplay_End
  Resume
End Function

Private Sub DisplayPassword()
  Dim ben As IBenefitClass
  
  On Error GoTo DisplayPassword_ERR
  
  If p11d32.CurrentEmployer Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "DisplayPassword", "The current employer is nothing.")
  
  Set ben = p11d32.CurrentEmployer
  MDIMain.mnuFileItems(MNU_FILE_PASSWORD).Visible = True
  MDIMain.mnuFilePasswordClear.Visible = Len(ben.value(employer_PassWord_db)) > 0
  
DisplayPassword_END:
  Exit Sub
DisplayPassword_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "DisplayPassWord", "Display Password", "Error displaying password menu.")
End Sub
Private Function DisplayFileMenu(ByVal dt As DisplayType) As Boolean
  
  On Error GoTo DisplayFileMenu_Err
  Call xSet("DisplayFileMenu")

  Call BringForwardMenu(dt)
  
  With MDIMain
    Select Case dt
      Case D_EMPLOYEES, D_BENEFIT
        .mnuFileItems(MNU_FILE_NEW).Visible = False
        .mnuFileItems(MNU_FILE_OPEN).Visible = False
        .mnuFileItems(MNU_FILE_EDIT).Visible = False
        .mnuFileItems(MNU_FILE_DELETE).Visible = False
        .mnuFileItems(MNU_FILE_PRINT).Visible = True
        .mnuFileItems(MNU_FILE_SEPERATOR1).Visible = True
        .mnuFileItems(MNU_FILE_IMPORT).Visible = False
        .mnuFileItems(MNU_FILE_ELECTRONIC_SUBMISSION).Visible = False
        .mnuFileItems(MNU_FILE_REFRESHEMPLOYERS).Visible = False
        .mnuFileItems(MNU_FILE_CHANGEDIRECTORY).Visible = False
        .mnuFileItems(MNU_FILE_EMPLOYER).Visible = True
        .mnuFileItems(MNU_FILE_TOOLS).Visible = False
        .mnuFileItems(MNU_FILE_ERROR_LOG).Visible = False
        Call DisplayPassword
      Case D_EMPLOYER_ON
        .mnuFileItems(MNU_FILE_NEW).Visible = True
        .mnuFileItems(MNU_FILE_OPEN).Visible = True
        .mnuFileItems(MNU_FILE_EDIT).Visible = True
        .mnuFileItems(MNU_FILE_DELETE).Visible = True
        .mnuFileItems(MNU_FILE_PRINT).Visible = False
        .mnuFileItems(MNU_FILE_SEPERATOR1).Visible = True
        .mnuFileItems(MNU_FILE_IMPORT).Visible = True
        .mnuFileItems(MNU_FILE_ELECTRONIC_SUBMISSION).Visible = True
        .mnuFileItems(MNU_FILE_ERROR_LOG).Visible = True
        .mnuFileItems(MNU_FILE_REFRESHEMPLOYERS).Visible = True
        .mnuFileItems(MNU_FILE_CHANGEDIRECTORY).Visible = True
        .mnuFileItems(MNU_FILE_EMPLOYER).Visible = False
        .mnuFileItems(MNU_FILE_PASSWORD).Visible = False
        .mnuFileItems(MNU_FILE_TOOLS).Visible = True
        .mnuFileToolsFindFiles.Visible = True
        .mnuFileToolsFindRepairCompactEmployer.Visible = True
      Case D_EMPLOYER_OFF
        .mnuFileItems(MNU_FILE_NEW).Visible = True
        .mnuFileItems(MNU_FILE_OPEN).Visible = False
        .mnuFileItems(MNU_FILE_EDIT).Visible = False
        .mnuFileItems(MNU_FILE_DELETE).Visible = False
        .mnuFileItems(MNU_FILE_PRINT).Visible = False
        .mnuFileItems(MNU_FILE_SEPERATOR1).Visible = True
        .mnuFileItems(MNU_FILE_IMPORT).Visible = False
        .mnuFileItems(MNU_FILE_ERROR_LOG).Visible = False
        .mnuFileItems(MNU_FILE_ELECTRONIC_SUBMISSION).Visible = True
        .mnuFileItems(MNU_FILE_REFRESHEMPLOYERS).Visible = True
        .mnuFileItems(MNU_FILE_CHANGEDIRECTORY).Visible = True
        .mnuFileItems(MNU_FILE_EMPLOYER).Visible = False
        .mnuFileItems(MNU_FILE_PASSWORD).Visible = False
        .mnuFileItems(MNU_FILE_TOOLS).Visible = True
        .mnuFileToolsFindFiles.Visible = True
        .mnuFileToolsFindRepairCompactEmployer.Visible = False
    End Select
  End With
  
DisplayFileMenu_End:
  Call xReturn("DisplayFileMenu")
  Exit Function

DisplayFileMenu_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DisplayFileMenu", "Display File Menu", "Error setting the file menu.")
  Resume DisplayFileMenu_End
  Resume
End Function

Private Sub BringForwardMenu(ByVal dt As DisplayType)
  Select Case dt
    Case D_EMPLOYER_ON, D_EMPLOYER_OFF
      MDIMain.mnuFileBringForward.Visible = True
      MDIMain.mnuFileBringForwardBreak.Visible = True
    Case Else
      MDIMain.mnuFileBringForward.Visible = False
      MDIMain.mnuFileBringForwardBreak.Visible = False
  End Select
End Sub

Private Function DisplayEmployeeMenu(ByVal dt As DisplayType) As Boolean

  On Error GoTo DisplayEmployeeMenu_Err
  Call xSet("DisplayEmployeeMenu")

  With MDIMain
    Select Case dt
      Case D_EMPLOYEES
        .mnuEmployee.Visible = True
        .mnuEmployeeItems(MNU_EMPLOYEE_CONFIRM).Visible = True
        .mnuEmployeeItems(MNU_EMPLOYEE_UNDO).Visible = True
        .mnuEmployeeItems(MNU_EMPLOYEE_SEPERATOR1).Visible = True
        .mnuEmployeeItems(MNU_EMPLOYEE_DETAILS).Visible = False
        .mnuEmployeeItems(MNU_EMPLOYEE_ADD).Visible = True
        .mnuEmployeeItems(MNU_EMPLOYEE_DELETE).Visible = True
        .mnuEmployeeItems(MNU_EMPLOYEE_GOTO).Visible = False
'MP RV removed as is handled in DisplayEx        .mnuBenefits.Visible = False
      Case D_BENEFIT
'MP RV removed as is handled in DisplayEx        .mnuBenefits.Visible = False
        .mnuEmployeeItems(MNU_EMPLOYEE_DETAILS).Visible = True
        .mnuEmployeeItems(MNU_EMPLOYEE_ADD).Visible = False
        .mnuEmployeeItems(MNU_EMPLOYEE_DELETE).Visible = False
        .mnuEmployeeItems(MNU_EMPLOYEE_GOTO).Visible = True
      Case D_EMPLOYER_ON
'MP RV removed as is handled in DisplayEx       .mnuBenefits.Visible = False
        .mnuEmployee.Visible = False
      Case D_EMPLOYER_OFF
'MP RV removed as is handled in DisplayEx       .mnuBenefits.Visible = False
        .mnuEmployee.Visible = False
    End Select
  End With
DisplayEmployeeMenu_End:
  Call xReturn("DisplayEmployeeMenu")
  Exit Function

DisplayEmployeeMenu_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DisplayEmployeeMenu", "Display Employee Menu", "Error setting the employee menu.")
  Resume DisplayEmployeeMenu_End
  Resume
End Function

Private Function DisplayHelpMenu(ByVal dt As DisplayType) As Boolean 'AM
'MP RV changed scope to Private
  On Error GoTo DisplayHelpMenu_Err
  Call xSet("DisplayHelpMenu")

  With MDIMain
    .mnuHelp.Visible = True
    .mnuHelpP11D.Visible = True
    .mnuSepX.Visible = True
    .mnuHelpAbout.Visible = True
    .mnuHelpFAQs.Visible = False
  End With
  
DisplayHelpMenu_End:
  Call xReturn("DisplayHelpMenu")
  Exit Function

DisplayHelpMenu_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DisplayHelpMenu", "Display Help Menu", "Error setting the help menu.")
  Resume DisplayHelpMenu_End
'MP RV why Resume below needed again, remove? If yes - remove elsewhere on this mod
  Resume
End Function
