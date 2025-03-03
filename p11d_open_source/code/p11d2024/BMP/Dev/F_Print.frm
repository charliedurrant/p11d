VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A7CE771F-05B2-43CF-9650-ED841A9049FA}#1.0#0"; "ATC3FolderBrowser.ocx"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_Print 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin atc3FolderBrowser.FolderBrowser fbUserReportsDirectory 
      Height          =   600
      Left            =   135
      TabIndex        =   33
      Top             =   5400
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1058
   End
   Begin VB.CommandButton cmdReportWizard 
      Caption         =   "&Wizard"
      Height          =   420
      Left            =   1380
      TabIndex        =   12
      ToolTipText     =   "Select a management/user report and shift click this button to load the report into the wizard"
      Top             =   6075
      Width           =   1050
   End
   Begin VB.CommandButton cmdPrintPreview 
      Caption         =   "&Preview"
      Default         =   -1  'True
      Height          =   420
      Index           =   1
      Left            =   3690
      TabIndex        =   1
      Top             =   6075
      Width           =   1050
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      Height          =   1155
      Left            =   5445
      TabIndex        =   11
      ToolTipText     =   $"F_Print.frx":0000
      Top             =   45
      Width           =   1965
      Begin VB.PictureBox pctFrame3 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   135
         ScaleHeight     =   825
         ScaleWidth      =   1635
         TabIndex        =   29
         Top             =   225
         Width           =   1635
         Begin VB.OptionButton optOrientation 
            Caption         =   "Portrait"
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   31
            Top             =   90
            Width           =   1185
         End
         Begin VB.OptionButton optOrientation 
            Caption         =   "Landscape"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   1230
         End
      End
   End
   Begin VB.Frame fraP46Options 
      Caption         =   "P46(Car) options"
      Height          =   2490
      Left            =   4005
      TabIndex        =   10
      Top             =   2655
      Width           =   3390
      Begin VB.PictureBox pctFrame1 
         BorderStyle     =   0  'None
         Height          =   2265
         Left            =   90
         ScaleHeight     =   2265
         ScaleWidth      =   3210
         TabIndex        =   19
         Top             =   180
         Width           =   3210
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 1"
            Height          =   330
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   0
            Width           =   3120
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 2"
            Height          =   330
            Index           =   1
            Left            =   90
            TabIndex        =   24
            Top             =   270
            Width           =   3120
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 3"
            Height          =   330
            Index           =   2
            Left            =   90
            TabIndex        =   23
            Top             =   540
            Width           =   3120
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 4"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   22
            Top             =   855
            Width           =   3120
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Range"
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   20
            Top             =   1260
            Width           =   2760
         End
         Begin atc2valtext.ValText txtP46Date 
            Height          =   285
            Index           =   0
            Left            =   1035
            TabIndex        =   21
            Top             =   1575
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "F_Print.frx":008C
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "05/04/1999"
            Minimum         =   "06/04/1998"
            AutoSelect      =   0
         End
         Begin atc2valtext.ValText txtP46Date 
            Height          =   285
            Index           =   1
            Left            =   1035
            TabIndex        =   26
            Top             =   1890
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "F_Print.frx":00A8
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "05/04/1999"
            Minimum         =   "06/04/1998"
            AutoSelect      =   0
         End
         Begin VB.Label lblDateFrom 
            Caption         =   "Date from"
            Height          =   195
            Left            =   135
            TabIndex        =   28
            Top             =   1620
            Width           =   690
         End
         Begin VB.Label lblDateTo 
            Caption         =   "Date to"
            Height          =   240
            Left            =   135
            TabIndex        =   27
            Top             =   1935
            Width           =   645
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   45
            X2              =   3015
            Y1              =   1170
            Y2              =   1170
         End
      End
   End
   Begin VB.Frame fraRange 
      Caption         =   "Print range"
      Height          =   1410
      Left            =   4005
      TabIndex        =   9
      Top             =   1215
      Width           =   3390
      Begin VB.PictureBox pctFrame 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   180
         ScaleHeight     =   1095
         ScaleWidth      =   3030
         TabIndex        =   15
         Top             =   225
         Width           =   3030
         Begin VB.OptionButton optSelection 
            Caption         =   "Current employee"
            Height          =   330
            Index           =   3
            Left            =   0
            TabIndex        =   32
            Top             =   810
            Width           =   2085
         End
         Begin VB.OptionButton optSelection 
            Caption         =   "Selected"
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   -45
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optSelection 
            Caption         =   "All"
            Height          =   330
            Index           =   2
            Left            =   0
            TabIndex        =   17
            Top             =   540
            Width           =   1275
         End
         Begin VB.OptionButton optSelection 
            Caption         =   "Inverse selected"
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   225
            Width           =   1590
         End
      End
   End
   Begin VB.Frame fmeDetails 
      Caption         =   "Reports"
      Height          =   3930
      Left            =   45
      TabIndex        =   8
      Top             =   1215
      Width           =   3885
      Begin MSComctlLib.TreeView tvwReports 
         Height          =   3660
         Left            =   90
         TabIndex        =   13
         Top             =   240
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   6456
         _Version        =   393217
         Indentation     =   0
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printer"
      Height          =   1140
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   5325
      Begin VB.Label lblWhere 
         Caption         =   "Where"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   810
         Width           =   4965
      End
      Begin VB.Label lblType 
         Caption         =   "Type"
         Height          =   240
         Left            =   225
         TabIndex        =   6
         Top             =   495
         Width           =   4965
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   225
         Width           =   4920
      End
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Op&tions"
      Height          =   420
      Left            =   2565
      TabIndex        =   0
      Top             =   6075
      Width           =   1050
   End
   Begin VB.CommandButton cmdPrintPreview 
      Caption         =   "&OK"
      Height          =   420
      Index           =   0
      Left            =   5175
      TabIndex        =   2
      Top             =   6075
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   6300
      TabIndex        =   3
      Top             =   6075
      Width           =   1050
   End
   Begin VB.Label lblUserReportsFolder 
      Caption         =   "User reports folder"
      Height          =   240
      Left            =   135
      TabIndex        =   14
      Top             =   5175
      Width           =   2400
   End
End
Attribute VB_Name = "F_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_NodeLastSelected As node
Private m_ReportsTreeView As TreeView
Public FromEmployeeLetter As Boolean

Private Sub cmdCancel_Click()
  Call UnLoadPrintForm
End Sub

Private Sub cmdOK_Click()
  Call DoPrint(PRINT_REPORT)
End Sub

Private Function ShowCheckDialog(Optional bCancel As Boolean) As Boolean
  
  On Error GoTo err_Err
  
  If p11d32.ReportPrint.CheckOptions < YES_ALWAYS Then Call F_PrintCheckOptions.Start
  
  bCancel = F_PrintCheckOptions.cancel
  
  Select Case p11d32.ReportPrint.CheckOptions
    Case NEVER, NO_THIS_TIME_ONLY
      ShowCheckDialog = False
    Case YES_ALWAYS, YES_THIS_TIME_ONLY
      ShowCheckDialog = True
    Case Else
      ECASE ("Unkown type")
  End Select
  
err_End:
  Exit Function
err_Err:
  Call ErrorMessage(Err.Number, Err, "ShowCheckDialog", "ShowCheckDialog", Err.Description)
  Resume err_End

End Function

Public Sub DoPrint(ByVal rTarget As REPORT_TARGET)
  Dim TempOL As ObjectList
  Dim rt As RPT_TYPE
  Dim sMessage As String
  Dim bCancel As Boolean, bRunChecks As Boolean
  On Error GoTo DoPrint_ERR
  
  Call xSet("DoPrint")
  'Confirm message for print
  If rTarget = PRINT_REPORT Then
    If p11d32.CurrentEmployer.GetEmployees(TempOL) Then
      Set p11d32.ReportPrint.SelectedEmployees = TempOL
      sMessage = "You are about to print the report:- " & vbCrLf & "'" & tvwReports.SelectedItem.Text & "'" _
        & vbCrLf & " for " & p11d32.ReportPrint.SelectedEmployees.Count & " employee(s)" & _
        vbCrLf & vbCrLf & "Do you wich to continue?"
      If p11d32.ReportPrint.DefaultReportIndex = RPT_EMPLOYEE_LETTER_EMAIL Then
        sMessage = sMessage & vbCrLf & vbCrLf & "This operation will also email the employees"
      End If
      Select Case MultiDialog("Warning", sMessage, "Continue", "Cancel")
        Case 2
          Exit Sub
      End Select
    End If
  End If
  
  Call SetCursor
  
  rt = p11d32.ReportPrint.ReportPrintPrapare(p11d32.ReportPrint.DefaultReportIndex)
  
  Select Case rt
    Case RPTT_STANDARD
      If p11d32.CurrentEmployer.GetEmployees(TempOL) Then
        Set p11d32.ReportPrint.SelectedEmployees = TempOL
          Call ClearAllCursors
          bRunChecks = ShowCheckDialog(bCancel)
          If bCancel Then GoTo DoPrint_END
          If bRunChecks Then
            Call F_CompanyCarCheckerWizard.CheckBeforePrint(p11d32.CurrentEmployer, TempOL)
            If Not F_CompanyCarCheckerWizard.ContinuePrint Then GoTo DoPrint_END
          End If
          Call SetCursor
          If p11d32.ReportPrint.DoStandardReport(p11d32.ReportPrint.DefaultReportIndex, rTarget) Then
            'cancel not pressed
            Call UnLoadPrintForm
          End If
      End If
    Case RPTT_MANAGEMENT
      Call UnLoadPrintForm
      'CADcheck
      If Not p11d32.ReportPrint.Destination = REPD_FILE_HTML Then 'EK Added to disable html reports
        Call p11d32.ReportPrint.DoWizardReport(p11d32.ReportPrint.DefaultReportIndex, rTarget)
      Else
        Call MsgBox("You are not able to print Management Reports to HTML. " & vbCrLf & "Please amend your destination selection on the print options form.", vbExclamation, "Functionality not available")
      End If
    Case RPTT_OTHER
      Call UnLoadPrintForm
      Call p11d32.ReportPrint.DoOtherReport(Nothing, p11d32.ReportPrint.DefaultReportIndex, rTarget)
    Case RPTT_MM
      Call UnLoadPrintForm
      Call p11d32.ReportPrint.DoOtherReport(Nothing, p11d32.ReportPrint.DefaultReportIndex, rTarget)
    Case RPTT_USER
      Call UnLoadPrintForm
      If Not p11d32.ReportPrint.Destination = REPD_FILE_HTML Then 'EK Added to disable html reports
        Call p11d32.ReportPrint.DoWizardReport(p11d32.ReportPrint.DefaultReportIndex, rTarget)
      Else
        Call MsgBox("You are not able to print Reports to HTML. " & vbCrLf & "Please amend your destination selection on the print options form.", vbExclamation, "Functionality not available")
      End If
    'FC - AE - 160801
    Case RPTT_ABACUSUDM
      Call UnLoadPrintForm
      Call p11d32.ReportPrint.DoWizardReport(p11d32.ReportPrint.DefaultReportIndex, rTarget)
  End Select
  
DoPrint_END:
  Call ClearAllCursors
  Call xReturn("DoPrint")
  Exit Sub
DoPrint_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "DoPrint", "Do Print", "Error determining how to print.")
  Resume DoPrint_END
  Resume
End Sub


Private Sub cmdOptions_Click()
'  F_PrintOptions.Show vbModal
  Call p11d32.Help.ShowForm(F_PrintOptions, vbModal)
End Sub

Public Sub cmdPrintPreview_Click(Index As Integer)

  Dim i As Integer
  
  If Not CheckP46Date Then Exit Sub
  
  Select Case Index
      Case 0
        Call DoPrint(PRINT_REPORT)
      Case 1
        Call DoPrint(PREPARE_REPORT)
  End Select

End Sub
Private Sub AddImageToNewNode(n As node, ByVal lReportIndex As Long)

  On Error GoTo AddImageToNewNode_ERR
  
  
AddImageToNewNode_END:
  Exit Sub
AddImageToNewNode_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddImageToNewNode", "Add Image To New Node", "Error adding image to new node in print tree.")
  Resume AddImageToNewNode_END
End Sub
Private Sub ReportsToTree()
  Dim i As Long, j As Long
  Dim n As node
  Dim sCurrentName As String
  Dim bFoundParent As Boolean
  Dim bSelected As Boolean
  
  On Error GoTo ReportsToTree_ERR
  
  Call xSet("ReportsToTree")
  
  'get the user reports
  'FC - AE - 160801
  ' Not All Reports
  If Not p11d32.ReportPrint.AllReports Then
    Call ReportToTree(tvwReports, RPT_P46CAR)
    Call ReportToTree(tvwReports, RPT_EMPLOYEE_LETTER)
    Call ReportToTree(tvwReports, RPT_EMPLOYEE_LETTER_EMAIL)
    Call ManagementReportsToTree(tvwReports)
    'Abacus Export
    If p11d32.ReportPrint.AbacusUDM Then
      For i = [RPT_ABACUSUDM] To [RPT_LAST_ABACUSUDM]
        Call ReportToTree(tvwReports, i)   'AM
      Next i
    End If
  ' All Reports
  Else
    For i = RPT_START To RPT_END
      If IsManagementReport(i) Then
        If Not Is83FileName(p11d32.ReportPrint.ManagementReportPathAndFile(i)) Then Call Err.Raise(ERR_FILE_INVALID, "ReportsToTree", "The management report file " & p11d32.ReportPrint.ManagementReportPathAndFile(i) & " is not 8.3 format.")
      End If
      'Abacus Export
      If Not p11d32.ReportPrint.AbacusUDM Then
        If Not i >= [RPT_ABACUSUDM] Or Not i <= [RPT_LAST_ABACUSUDM] Then
          Call ReportToTree(tvwReports, i)   'km
        End If
      Else
        Call ReportToTree(tvwReports, i) 'km
      End If
    Next i
  End If

  Call ReportsToTreeEndPrintDialog
  
ReportsToTree_END:
  Call xReturn("ReportsToTree")
  Exit Sub
ReportsToTree_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReportsToTree", "Reports To Tree", "Error placing the reports to the treeview.")
  Resume ReportsToTree_END
  Resume
End Sub

Public Function SettingsToScreen() As Boolean
  
  Dim i As Long
  
  On Error GoTo SettingsToScreen_Err
  Call xSet("SettingsToScreen")

  'default report set in InitPrintDialog
  Call ReportsImageListSet(tvwReports)
  
  Call ReportsToTree
  
  Call SetDefaultVTDate(txtP46Date(0))
  Call SetDefaultVTDate(txtP46Date(1))
  
  Call txtP46Date_LostFocus(-1)
  Call optQuarter_Click(-1)
  Call optSelection_Click(-1)
    
  fbUserReportsDirectory.Directory = p11d32.ReportPrint.ReportPathUser
  
SettingsToScreen_End:
  Call xReturn("SettingsToScreen")
  Exit Function

SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error placing the print options to the screen.")
  Resume SettingsToScreen_End
  Resume
End Function

Private Sub UnLoadPrintForm()
  If Not Me.FromEmployeeLetter Then
    Unload Me
    DoEvents
  End If
End Sub
Private Sub cmdReportWizard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim rt As RPT_TYPE
  Dim sPathAndFile As String
 On Error GoTo err_Err
  
  Call UnLoadPrintForm
  
  rt = p11d32.ReportPrint.ReportPrintPrapare(p11d32.ReportPrint.DefaultReportIndex)
  If (Shift And vbShiftMask) And ((rt = RPTT_MANAGEMENT) Or (rt = RPTT_USER)) Then
    'get the file name and pass as a template
    sPathAndFile = p11d32.ReportPrint.ManagementReportPathAndFile(p11d32.ReportPrint.DefaultReportIndex)
  Else
    sPathAndFile = p11d32.ReportPrint.TemplatePathAndFile
  End If
  Call p11d32.ReportPrint.StartReportWizard(sPathAndFile, REPORTW_LOAD_FILE)

err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MouseDown", "MouseDown", Err.Description)
  Resume err_End
End Sub
Private Sub ReportsToTreeEndPrintDialog()
  If ReportsToTreeEnd(tvwReports, m_NodeLastSelected) Then
    Call tvwReports_NodeClick(tvwReports.SelectedItem)
  End If
End Sub
  
Private Sub Form_Load()
  Call SettingsToScreen
End Sub
Private Sub optOrientation_Click(Index As Integer)
  Dim i As Long
  
  If tvwReports.SelectedItem.Tag = L_REPORT_USER_TAG Then Exit Sub
  
  Select Case Index
    Case 0
      p11d32.ReportPrint.Orientation(tvwReports.SelectedItem.Tag) = PORTRAIT
    Case 1
      p11d32.ReportPrint.Orientation(tvwReports.SelectedItem.Tag) = LANDSCAPE
  End Select
  
End Sub

Private Sub optQuarter_Click(Index As Integer)
  Dim i As Long, lCurrentQuarterIndex As Long
  Dim dQuarterEnd As Date, dQuarterStart As Date, dNow As Date
  Dim bSetValue As Boolean
  dNow = Now
  
  Select Case Index
    Case -1
      For i = 0 To 4
        
        If i < 4 Then
          
          Call p11d32.Rates.GetP46QuarterStartEnd(dQuarterStart, dQuarterEnd, i + 1)
          If DateInRange(dNow, dQuarterStart, dQuarterEnd) Then lCurrentQuarterIndex = i
          optQuarter(i).Caption = "Quarter " & CStr(i + 1) & " (" & DateValReadToScreen(dQuarterStart) & " - " & DateValReadToScreen(dQuarterEnd) & ")"
        End If
        If p11d32.ReportPrint.P46Range = i Then
          optQuarter(i) = True
          bSetValue = True
        End If
      Next
      If Not bSetValue Then
        optQuarter(lCurrentQuarterIndex) = True
        p11d32.ReportPrint.P46UserDateFrom = DateValReadToScreen(p11d32.ReportPrint.P46DateFrom)
        p11d32.ReportPrint.P46UserDateTo = DateValReadToScreen(p11d32.ReportPrint.P46DateTo)
      End If
    Case Is < 4
      Call p11d32.Rates.GetP46QuarterStartEnd(dQuarterStart, dQuarterEnd, Index + 1)
      p11d32.ReportPrint.P46DateFrom = dQuarterStart
      p11d32.ReportPrint.P46DateTo = dQuarterEnd
      txtP46Date(0).Enabled = False
      txtP46Date(1).Enabled = False
      txtP46Date(0).Validate = False
      txtP46Date(1).Validate = False
      p11d32.ReportPrint.P46Range = Index
    Case 4
      txtP46Date(0).Enabled = True
      txtP46Date(1).Enabled = True
      txtP46Date(0).Validate = True
      txtP46Date(1).Validate = True
      txtP46Date(0).AllowEmpty = False
      txtP46Date(1).AllowEmpty = False
      p11d32.ReportPrint.P46Range = Index
      Call SetRangeOnSelectRange
  End Select
  
End Sub
Private Sub SetRangeOnSelectRange()
  On Error GoTo err_Err
  
  Call txtP46Date_LostFocus(0)
  Call txtP46Date_LostFocus(1)

err_Err:
 Exit Sub
End Sub
Private Sub optSelection_Click(Index As Integer)
  Dim i As Long
  Dim es As EMPLOYEE_SELECTION
    
  If Index <= [_ES_LAST_ITEM] And Index >= [_ES_FIRST_ITEM] Then
    p11d32.ReportPrint.EmployeeSelection = Index
  ElseIf Index = -1 Then
    If (p11d32.ReportPrint.RemeberEmployeeSelection) Then
      es = p11d32.ReportPrint.EmployeeSelection
    Else
      es = ES_SELECTED
    End If
    For i = [_ES_FIRST_ITEM] To [_ES_LAST_ITEM]
      If i = es Then
        optSelection(i) = True
        Exit For
      End If
    Next
  Else
    Call ECASE("Invalid employee selection:" & Index)
  End If
  
End Sub

Public Function CheckP46Date() As Boolean
  Dim i As Long
  On Error GoTo CheckP46Date_Err
  Call xSet("CheckP46Date")

  CheckP46Date = True

  If fraP46Options.Enabled And p11d32.ReportPrint.P46Range = P46_USERRANGE Then
    'check the validity of the two txtP46DAte
    For i = 0 To 1
      If txtP46Date(i).FieldInvalid Then
        Call ErrorMessage(ERR_ERROR, Err, "CheckP46Date", "Check P46 Date", "The P46(Car) date range is invalid.")
        CheckP46Date = False
        txtP46Date(i).SetFocus
        Exit For
      End If
    Next
  End If


CheckP46Date_End:
  Call xReturn("CheckP46Date")
  Exit Function

CheckP46Date_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckP46Date", "Check P46 Date", "Error checking the user input P46 date range.")
  Resume CheckP46Date_End
End Function

Private Sub tvwReports_Collapse(ByVal node As MSComctlLib.node)
  node.Image = IMG_FOLDER_CLOSED
End Sub

Private Sub tvwReports_Expand(ByVal node As MSComctlLib.node)
  node.Image = IMG_FOLDER_OPEN
End Sub
Private Sub NodeClick(node As node)
  Dim bFrameOrientation As Boolean
  On Error GoTo NodeClick_ERR
  
  If node Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "NodeClick", "The node is nothing.")
  If StrComp(node.Text, p11d32.ReportPrint.Name(RPT_EMPLOYEE_LETTER_EMAIL)) = 0 Then
    cmdPrintPreview.Item(1).Enabled = False
  Else
    cmdPrintPreview.Item(1).Enabled = True
  End If
  
  If Len(node.Tag) > 0 Then
    'valid report not header
    Call ReportsSelectNodeImage(m_NodeLastSelected, node)
    Set tvwReports.SelectedItem = node
    
    p11d32.ReportPrint.DefaultReportIndex = node.Tag
    
    Select Case p11d32.ReportPrint.ReportType(node.Tag, True)
      Case RPTT_MANAGEMENT
        
      'FC - AE - 160801
      Case RPTT_ABACUSUDM
      Case RPTT_USER
        p11d32.ReportPrint.UserReportFileLessExtension = node.Text
      Case Else
        bFrameOrientation = True
        If p11d32.ReportPrint.Orientation(node.Tag) = LANDSCAPE Then
          optOrientation(1) = True
        Else
          optOrientation(0) = True
        End If
    End Select
      
    Call EnableFrame(Me, fraOrientation, bFrameOrientation)
            
    If node.Tag = P11D_REPORTS.RPT_P46CAR Then
      fraP46Options.Enabled = True
    Else
      fraP46Options.Enabled = False
    End If
  End If
  
  
NodeClick_END:
  Exit Sub
NodeClick_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NodeClick", "NodeClick", "Error clicking a node on the reports list.")
  Resume NodeClick_END
  Resume
End Sub
Private Sub tvwReports_NodeClick(ByVal node As MSComctlLib.node)
  Call NodeClick(node)
End Sub

Private Sub txtP46Date_LostFocus(Index As Integer)
  Select Case Index
    Case 0
      p11d32.ReportPrint.P46UserDateFrom = txtP46Date(0).Text
      p11d32.ReportPrint.P46DateFrom = TryConvertDateDMY(txtP46Date(Index).Text, UNDATED)
    Case 1
      p11d32.ReportPrint.P46UserDateTo = txtP46Date(1).Text
      p11d32.ReportPrint.P46DateTo = TryConvertDateDMY(txtP46Date(Index).Text, UNDATED)
    Case -1
      txtP46Date(0).Text = p11d32.ReportPrint.P46UserDateFrom
      txtP46Date(1).Text = p11d32.ReportPrint.P46UserDateTo
  End Select
End Sub
Public Property Get ReportsTreeView() As TreeView
  Set ReportsTreeView = m_ReportsTreeView
End Property
Public Property Let ReportsTreeView(ByVal tv As TreeView)
  m_ReportsTreeView = tv
End Property

Private Sub fbUserReportsDirectory_Ended()
  On Error GoTo err_Err:
  
  p11d32.ReportPrint.ReportPathUser = fbUserReportsDirectory.Directory
  Call ReportsToTreeEndPrintDialog
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UserReportsDirectory", "User Reports Directory", Err.Description)
  Resume err_End
End Sub
Private Sub fbUserReportsDirectory_Started()
  fbUserReportsDirectory.Directory = p11d32.ReportPrint.ReportPathUser
End Sub

