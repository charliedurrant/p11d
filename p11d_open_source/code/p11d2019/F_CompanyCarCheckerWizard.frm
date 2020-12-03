VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_DataCheckerWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data checker"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   -225
      TabIndex        =   0
      Top             =   6675
      Width           =   12135
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print results"
         Height          =   390
         Left            =   4950
         TabIndex        =   27
         Top             =   375
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnBack 
         Caption         =   "<  &Back"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         TabIndex        =   25
         Top             =   375
         Width           =   1230
      End
      Begin VB.CommandButton btnNext 
         Caption         =   "&Next >"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7875
         TabIndex        =   4
         Top             =   375
         Width           =   1230
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   9375
         TabIndex        =   3
         Top             =   375
         Width           =   1230
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   450
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin TabDlg.SSTab tabCheckWizard 
      Height          =   7350
      Left            =   0
      TabIndex        =   1
      Top             =   -375
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12965
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "F_CompanyCarCheckerWizard.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frame"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "F_CompanyCarCheckerWizard.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "tvwCheckResults"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "btnRefresh"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbOrderBy"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "dtgCheckWizard"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fraHeader"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "pctHeader"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "pctSpacer"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "pctInfo"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "iml"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "ctlOverlappingCars"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "F_CompanyCarCheckerWizard.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin P11D2019.CarCheckOverlap ctlOverlappingCars 
         Height          =   3915
         Left            =   3225
         TabIndex        =   28
         Top             =   2250
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   6906
      End
      Begin MSComctlLib.ImageList iml 
         Left            =   2100
         Top             =   3825
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   32
         ImageHeight     =   33
         MaskColor       =   16776960
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "F_CompanyCarCheckerWizard.frx":0054
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pctInfo 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   1350
         ScaleHeight     =   765
         ScaleWidth      =   9240
         TabIndex        =   23
         Top             =   6375
         Width           =   9240
         Begin P11D2019.TransparentPictureBox TransparentPictureBox1 
            Height          =   615
            Left            =   300
            TabIndex        =   26
            Top             =   0
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   1085
         End
         Begin VB.Label lblInfo 
            Caption         =   "lblInfo"
            Height          =   765
            Left            =   975
            TabIndex        =   24
            Top             =   0
            Width           =   8265
         End
      End
      Begin VB.PictureBox pctSpacer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   3225
         ScaleHeight     =   885
         ScaleWidth      =   7260
         TabIndex        =   20
         Top             =   1200
         Width           =   7290
         Begin VB.Label lblCheckType 
            BackColor       =   &H8000000E&
            Caption         =   "lblCheckType"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   150
            TabIndex        =   22
            Top             =   75
            Width           =   5655
         End
         Begin VB.Label lblEeName 
            BackColor       =   &H8000000E&
            Caption         =   "lblEeName"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   150
            TabIndex        =   21
            Top             =   450
            Width           =   5655
         End
      End
      Begin VB.PictureBox pctHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   10695
         TabIndex        =   16
         Top             =   390
         Width           =   10695
         Begin VB.Image imgMagnify 
            Height          =   735
            Left            =   9540
            Picture         =   "F_CompanyCarCheckerWizard.frx":08E8
            Top             =   15
            Width           =   870
         End
         Begin VB.Label lblResults 
            BackColor       =   &H80000009&
            Caption         =   "Anaylse results"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   405
            TabIndex        =   17
            Top             =   225
            Width           =   8055
         End
      End
      Begin VB.Frame fraHeader 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   945
         Width           =   11000
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   135
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   10695
      End
      Begin TrueDBGrid60.TDBGrid dtgCheckWizard 
         Bindings        =   "F_CompanyCarCheckerWizard.frx":10E2
         Height          =   3885
         Left            =   3225
         OleObjectBlob   =   "F_CompanyCarCheckerWizard.frx":10FF
         TabIndex        =   6
         Top             =   2250
         Visible         =   0   'False
         Width           =   7305
      End
      Begin VB.ComboBox cmbOrderBy 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton btnRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   6300
         Width           =   735
      End
      Begin MSComctlLib.TreeView tvwCheckResults 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   7858
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   0
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Frame frame 
         BackColor       =   &H80000009&
         Height          =   7695
         Left            =   -75000
         TabIndex        =   2
         Top             =   0
         Width           =   16215
         Begin VB.CheckBox chkAllChecks 
            BackColor       =   &H80000009&
            Caption         =   "Run all checks"
            Height          =   195
            Left            =   3600
            TabIndex        =   12
            Top             =   5100
            Width           =   7695
         End
         Begin MSComctlLib.ImageList imgCheckBox 
            Left            =   960
            Top             =   5280
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin MSComctlLib.TreeView tvwChecks 
            Height          =   3495
            Left            =   3600
            TabIndex        =   7
            Top             =   1560
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   6165
            _Version        =   393217
            HideSelection   =   0   'False
            LabelEdit       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Data dbCoCarChecker 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   480
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   6120
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Image Image1 
            Height          =   1020
            Left            =   9300
            Picture         =   "F_CompanyCarCheckerWizard.frx":3807
            Top             =   5250
            Width           =   990
         End
         Begin VB.Label lblDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "lblDescription"
            ForeColor       =   &H80000008&
            Height          =   990
            Left            =   3750
            TabIndex        =   15
            Top             =   5400
            Width           =   5520
         End
         Begin VB.Label labelIntroduction 
            BackColor       =   &H00FFFFFF&
            Caption         =   $"F_CompanyCarCheckerWizard.frx":48CF
            Height          =   1695
            Left            =   3600
            TabIndex        =   14
            Top             =   960
            Width           =   6855
         End
         Begin VB.Label lblCheckWizard 
            BackColor       =   &H80000009&
            Caption         =   "Welcome to the Data Checker wizard"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3600
            TabIndex        =   9
            Top             =   480
            Width           =   6495
         End
         Begin VB.Image imCog 
            Height          =   6405
            Left            =   0
            Picture         =   "F_CompanyCarCheckerWizard.frx":4979
            Top             =   360
            Width           =   3225
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Display by"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1230
         Width           =   2895
      End
   End
End
Attribute VB_Name = "F_DataCheckerWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_bInError
Private Const S_NODE_KEY As String = "CEHCKER_KEY_VALUE"
Private Const S_DATA_KEY As String = "DATA"
Private m_ey As Employer
Private m_CheckType As CHECKS
Private m_TempOL As ObjectList
Private m_NodesOL As ObjectList
Private m_bSelected As Boolean
Private m_bPrint As Boolean
Private m_bCheckDown As Boolean
Private m_tnNodeSelected As node
Private Sub btnBack_Click()
  tabCheckWizard.tab = 0
  btnBack.Enabled = False
  cmdPrint.Visible = False
  btnNext.Enabled = ChecksSelected(tvwChecks)
  btnNext.Caption = "&Next"
End Sub
Private Sub btnRefresh_Click()
  Call RefreshChecks
End Sub
Private Sub RefreshChecks()
  Dim tnNode As node
  Dim bFound As Boolean
  
  On Error GoTo err_err
  
'  Set m_tnNodeSelected = tvwCheckResults.SelectedItem
  
  LockWindowUpdate (tvwCheckResults.hwnd)
  Call tvwCheckResults.nodes.Clear
  Call CheckOverlapDirty
  Call DoChecks
  Call SelectFirstCheck
  
err_end:
  LockWindowUpdate (0)
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "btnRefresh_Click", "Failed to refresh", Err.Description)
  Resume err_end
  Resume
End Sub
Private Sub SelectFirstCheck()
  Dim tnNode As node
  Dim i As Integer
  Dim bFound As Boolean
  
  On Error GoTo err_err
  
  If Not m_bInError Then
    dtgCheckWizard.Visible = False
    ctlOverlappingCars.Visible = False
    Set tnNode = tvwCheckResults.nodes.Add(, , , "No errors found")
    lblEeName = ""
    lblCheckType = "No errors found"
    tnNode.Bold = True
    GoTo err_end
  End If
  bFound = False
  cmdPrint.Visible = False
  If Not m_tnNodeSelected Is Nothing Then
    For Each tnNode In tvwCheckResults.nodes
      If tnNode.Tag = m_tnNodeSelected.Tag Then
        bFound = True
        tnNode.Selected = True
        cmdPrint.Visible = True
        Call ChangeCheck(tnNode)
        Exit For
      End If
    Next
  Else
    cmdPrint.Visible = True
    
  End If
  
  If Not bFound Then
    tvwCheckResults.nodes(1).FirstSibling.child.Selected = True
    'Select the first check error
    Call ChangeCheck(tvwCheckResults.nodes(1).FirstSibling.child)
  End If
  
  If (Not tvwCheckResults.SelectedItem Is Nothing) Then
   If Not tvwCheckResults.SelectedItem.Parent Is Nothing Then
     Call tvwCheckResults.SelectedItem.Parent.EnsureVisible
   Else
     Call tvwCheckResults.SelectedItem.EnsureVisible
   End If
  End If
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "SelectFirstCheck", "SelectFirstCheck", Err.Description)
  Resume err_end
End Sub
Private Function DateWithinTaxYear(dtNew As String, Optional bJustDate As Boolean = False) As Boolean
  Dim dt As Variant
  Dim bInYear As Boolean
  
  bInYear = True
  dt = ScreenToDateVal(dtNew, STDV_UNDATED)
  bInYear = Not (dt = UNDATED Or dt = UNDATED)
  If (bInYear And (Not bJustDate)) Then
    bInYear = DateInRange(dt, p11d32.Rates.value(TaxYearStart), p11d32.Rates.value(TaxYearEnd))
  End If

  DateWithinTaxYear = bInYear
End Function


Private Function ValidateDataFields() As String
  Dim sMessage As String
  Dim i As Long
  
  On Error GoTo err_err
  
  sMessage = ""
  If dtgCheckWizard.DataChanged Then
    For i = 0 To dtgCheckWizard.Columns.Count - 1
      If dtgCheckWizard.Columns(i).Visible And Not dtgCheckWizard.Columns(i).Locked Then
        Select Case dtgCheckWizard.Columns(i).Caption
          
          Case p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableFrom_db), p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableTo_db)
            If Not DateWithinTaxYear(CStr(dtgCheckWizard.Columns(i).value)) Then
              sMessage = sMessage & dtgCheckWizard.Columns(i).Caption & _
                        "Should be between " & CStr(p11d32.Rates.value(TaxYearStart)) & _
                        " and " & CStr(p11d32.Rates.value(TaxYearEnd)) & vbCrLf
            End If
          Case p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, car_Registrationdate_db)
            If Not DateWithinTaxYear(CStr(dtgCheckWizard.Columns(i).value), True) Then
              sMessage = sMessage & dtgCheckWizard.Columns(i).Caption & " Invalid Date" & vbCrLf
            End If
        End Select
      End If
    Next i
  End If
  ValidateDataFields = sMessage
  
err_end:
  Exit Function
err_err:
  Call ErrorMessage(Err.Description, Err, "ValidateDataFields", "ValidateDataFields", Err.Description)
  Resume err_end
  Resume
End Function

Private Sub chkAllChecks_Click()
  Dim tnNode As node
  
  On Error GoTo err_err
  
  If m_bCheckDown Then Exit Sub
  
  For Each tnNode In tvwChecks.nodes
    tnNode.Checked = ChkBoxToBool(chkAllChecks)
  Next
  
  If ChkBoxToBool(chkAllChecks) Then
    p11d32.ReportPrint.ChecksSelected = -1
  Else
    p11d32.ReportPrint.ChecksSelected = 0
  End If
  
  btnNext.Enabled = ChecksSelected(tvwChecks)
   
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "chkAllChecks_Click", "chkAllChecks_Click", Err.Description)
  Resume err_end
End Sub

Private Sub cmbOrderBy_Click()
  
  On Error GoTo err_err
  
  Call ChangeNodeDisplay(cmbOrderBy.ListIndex)
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "cmbOrderBy_Change", "cmbOrderBy_Change", Err.Description)
  Resume err_end
End Sub

Private Sub cmdPrint_Click()
  Dim n As node
  Dim nChild As node
  Dim sKey As String, sData As String, sFieldCaption As String
  Dim sSQL As String
  Dim rsMaster As ADODB.Recordset, rs As ADODB.Recordset
  Dim ac As atc3GRID_ADO.AutoClass
  Dim iCheck As CHECKS, i As Long
  Dim rep As Reporter
  Dim f As ADODB.Field
  Dim sWhere As String
  Dim cn As ADODB.Connection
  Dim benEY As IBenefitClass
  
  On Error GoTo err_err
  
  Set rep = ReporterNew()
       
  Set benEY = m_ey
  If Not rep.InitReport("TITLE" & vbCrLf & vbCrLf, REPORT_TARGET.PREPARE_REPORT, LANDSCAPE, True) Then GoTo err_end
  Set cn = ADOConnect(ADOAccess4ConnectString(benEY.value(employer_FileName), p11d32.SystemMDWPath()), adUseClient)
  
  For Each n In tvwCheckResults.nodes
      
    If n.Children > 0 Then
      sKey = n.Tag
      iCheck = Replace(sKey, S_NODE_KEY, "")
      Set rsMaster = Nothing
      sWhere = ""
      For Each nChild In tvwCheckResults.nodes
        If nChild.Parent Is n Then
          sData = GetPropertyFromString(nChild.Tag, S_DATA_KEY)
          
          sSQL = GetNodeClickSQL(iCheck, sData)
          Set rs = cn.Execute(sSQL)
          
          If rsMaster Is Nothing Then
            Set rsMaster = New ADODB.Recordset
            For Each f In rs.Fields
              Call rsMaster.Fields.Append(f.Name, f.Type, f.DefinedSize, f.Attributes)
            Next
            Call rsMaster.Open
          End If
          
           Do While Not rs.EOF
             Call rsMaster.AddNew
             For Each f In rs.Fields
              rsMaster.Fields(f.Name).value = rs.Fields(f.Name).value
             Next
             Call rsMaster.Update
             Call rs.MoveNext
           Loop
        
        End If
      Next
      
       Set ac = New atc3GRID_ADO.AutoClass
       If Not ac.InitAutoData("ReportErrors", rsMaster) Then GoTo err_end
       
       For Each f In rsMaster.Fields
         '{CAPTION="Description"}
          sFieldCaption = ""
         Select Case UCASE$(f.Name)
           Case S_FIELD_PERSONEL_NUMBER
             sFieldCaption = p11d32.BenDataLinkUDMDisplayName(BC_EMPLOYEE, ee_PersonnelNumber_db)
             If iCheck <> CK_EC_NI Then
              Call ac.AddFieldFormat(f.Name, "{GROUP}")
             End If
           Case "DISPLAYNAME"
             Call ac.AddFieldFormat(f.Name, "{HIDE}")
           Case "AVAILTO"
              sFieldCaption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableTo_db)
           Case "AVAILFROM"
              sFieldCaption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableFrom_db)
           Case "REGREPLACED"
           
             Call ac.AddFieldFormat(f.Name, "{HIDE}")
           Case "REGDATE"
             sFieldCaption = "Date first registered"
           Case "REG"
            sFieldCaption = "Registration"
         End Select
           
         If Len(sFieldCaption) > 0 Then
          Call ac.AddFieldFormat(f.Name, "{CAPTION=""" & sFieldCaption & """}")
         End If
         
       Next
      
       ac.dateFormat = "DD/MM/YYYY"
       ac.ReportHeader = "{B+}" & CheckListCaption(iCheck, CMT_LIST_ITEM) & "{B-}" & vbCrLf
       
       
       Call ac.ShowReport(rep)
       Call rep.Out(vbCrLf & vbCrLf)
       
       
             
    End If
  Next
rep.EndReport
  Call rep.PreviewReport

  
  
err_end:
  If Not cn Is Nothing Then
    Call cn.Close
  End If
  
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "Print", "Print", "Failed to print report for data checker")
  Resume err_end
  Resume
End Sub

Private Sub dtgCheckWizard_Validate(Cancel As Boolean)
  Dim sMessage As String
  On Error GoTo err_err

  sMessage = ValidateDataFields
  If Len(sMessage) > 0 Then
    If MsgBox(sMessage, vbOKCancel, "Data Checker") = vbCancel Then
      dtgCheckWizard.ReBind
      Call FormatGrid
    End If
    Cancel = True
  Else
    Call dtgCheckWizard.Update
  End If
  
err_end:
  Exit Sub
err_err:
  Resume err_end
End Sub

Private Sub Form_Load()
  Dim ben As IBenefitClass
  On Error GoTo err_err
  lblInfo.Caption = ""
  lblResults.Caption = ""
  lblStatus.Caption = ""
  Me.Caption = S_DATA_CHECKER_WIZARD_NAME
  dbCoCarChecker.DatabaseName = m_ey.db.Name
  lblEeName.Caption = ""
  If Not m_ey Is Nothing Then
    Set ben = m_ey
   lblResults.Caption = "Analyse results for " & ben.value(ITEM_DESC)
  End If
    
  lblCheckType.Caption = "No check selected"
  Call FillOrderByComboBox
  tvwCheckResults.ImageList = MDIMain.imlTree
  Call imgCheckBox.ListImages.Add(1, , MDIMain.imlTree.ListImages(IMG_INFO).Picture)
  Call imgCheckBox.ListImages.Add(2, , MDIMain.imlTree.ListImages(IMG_UNSELECTED).Picture)
  Call imgCheckBox.ListImages.Add(3, , MDIMain.imlTree.ListImages(IMG_SELECTED).Picture)
  Me.tabCheckWizard.tab = 0
  btnNext.Caption = "&Next"
  btnBack.Enabled = False
  With dtgCheckWizard
    .Left = ctlOverlappingCars.Left
    .Top = ctlOverlappingCars.Top
    .width = ctlOverlappingCars.width
    .height = ctlOverlappingCars.height
  End With
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "Load", "Load", Err.Description)
  Resume err_end
End Sub

Public Function CheckBeforePrint(ey As Employer, TempOL As ObjectList, ByVal rt As REPORT_TARGET) As Boolean
  Dim i As CHECKS
  
  On Error GoTo err_err
  
  If ey Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "Start", "The employer is nothing.")
  Set m_ey = ey
  If TempOL Is Nothing Then
    Set m_TempOL = ey.employees
  Else
    Set m_TempOL = TempOL
  End If
  
  
  Call FillCheckTreeview(-1)
  
  m_bSelected = True
  Call CompanyCarCheckerFuncStart
  'Run employee checks
  Call EmployeeChecks(CK_ALL_CHECKS)
  'Run company car checks
  For i = [_CK_FIRST_ITEM] To [_CK_EC_START] - 1
    If i <> CK_CC_NOCARS Then 'ignore gaps between car useage
      RunCompanyCarChecker (i)
    End If
  Next i
  
  tabCheckWizard.tab = 1
  btnNext.Enabled = True
  If (rt = PREPARE_REPORT) Then
    btnNext.Caption = "&Preview"
  ElseIf rt = PRINT_REPORT Then
    btnNext.Caption = "&Print"
  Else
    btnNext.Caption = "&Continue"
  End If
  btnBack.Visible = False
  
  Call FormatTreeview(tvwCheckResults)
  Call ChangeNodeDisplay(p11d32.ReportPrint.CHECKORDERBY)
  If Not m_bInError Then
    CheckBeforePrint = True
    GoTo err_end
  End If
  Call SelectFirstCheck
  Call p11d32.Help.ShowForm(Me, vbModal)
  CheckBeforePrint = m_bPrint
err_end:
  Call CompanyCarCheckerFuncEnd
  Unload Me
  Exit Function
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error in Start of CompanyCarChecker.")
  Resume err_end
  Resume
End Function



Public Sub Start(ey As Employer)
  
  On Error GoTo Start_ERR
  
  If ey Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "Start", "The employer is nothing.")
  Set m_ey = ey
  Set m_TempOL = m_ey.employees
  
  Call CompanyCarCheckerFuncStart
  Call FillCheckTreeview(p11d32.ReportPrint.ChecksSelected)
  Call p11d32.Help.ShowForm(Me, vbModal)
  
Start_END:
  Call CompanyCarCheckerFuncEnd
  Exit Sub
Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error in Start of CompanyCarChecker.")
  Resume Start_END
End Sub
Private Sub FillCheckTreeview(iChecksSelected As Long)
' Fill treeview with available checks
  Dim tnRoot As node
  Dim tnChild As node
  Dim i As Integer
  Dim nFirst As node
  On Error GoTo FillCheckTreeview_ERR
  
  tvwChecks.Checkboxes = True
  Set tvwChecks.ImageList = MDIMain.imlTree
  
  'Fill with company car checks
  Set tnRoot = tvwChecks.nodes.Add(, , , S_COMPANY_CAR_CHECKS, IMG_FOLDER_CLOSED)
  tnRoot.ExpandedImage = IMG_FOLDER_OPEN
  For i = [_CK_FIRST_ITEM] To [_CK_EC_START] - 1
    Set tnChild = tvwChecks.nodes.Add(tnRoot, tvwChild, , CheckListCaption(i, CMT_LIST_ITEM), IMG_REPORT)
    tnChild.Tag = i
    If (nFirst Is Nothing) Then Set nFirst = tnChild
  Next i
  
  'Fill with employee checks
  Set tnRoot = tvwChecks.nodes.Add(, , , S_EMPLOYEE_CHECKS, IMG_FOLDER_CLOSED)
  tnRoot.ExpandedImage = IMG_FOLDER_OPEN
  For i = [_CK_EC_START] To [_CK_EC_END]
    Set tnChild = tvwChecks.nodes.Add(tnRoot, tvwChild, , CheckListCaption(i, CMT_LIST_ITEM), IMG_REPORT)
    tnChild.Tag = i
    If (nFirst Is Nothing) Then Set nFirst = tnChild
  Next i
  
  
  Call FormatTreeview(tvwChecks)
  chkAllChecks.value = BoolToChkBox((p11d32.ReportPrint.ChecksSelected = -1))
  
  Call SetIniSettingsInCheckTvw(iChecksSelected)
  btnNext.Enabled = ChecksSelected(tvwChecks)
  If (Not nFirst Is Nothing) Then
    nFirst.Selected = True
    Call tvwChecks_NodeClick(nFirst)
    lblDescription.Caption = CheckListCaption(nFirst.Tag, CMT_ALERT_MESSAGE_DESCRIPTION)
  Else
    lblDescription.Caption = ""
  End If


FillCheckTreeview_END:
  Exit Sub
FillCheckTreeview_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Fill Treeview Error", "Fill Checks Treeview Error", "Error on filling Treeview with availabe co car checks.")
  Resume FillCheckTreeview_END
  Resume
End Sub

Private Sub SetIniSettingsInCheckTvw(iChecksSelected As Long)
  Dim tnNode As node, tnChild As node
  
  On Error GoTo err_err

  For Each tnNode In tvwChecks.nodes
    If tnNode.Children = 0 Then
      If iChecksSelected = -1 Then
        tnNode.Checked = True
      ElseIf ((2 ^ tnNode.Tag) And iChecksSelected) = (2 ^ tnNode.Tag) Then
        tnNode.Checked = True
      End If
    End If
  Next
  
  For Each tnNode In tvwChecks.nodes
    If tnNode.Children > 0 Then
      tnNode.Checked = True
      For Each tnChild In tvwChecks.nodes
        If (tnChild.Children = 0) Then
          If tnChild.Parent = tnNode Then
            If tnChild.Checked = False Then
              tnNode.Checked = False
            End If
          End If
        End If
      Next
    End If
  Next
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "SetIniSettingsInCheckTvw", "SetIniSettingsInCheckTvw", Err.Description)
  Resume err_end
End Sub

Private Sub FillOrderByComboBox()
  Dim i As Long
  On Error GoTo err_err
  
  For i = [_ORDER_FIRST_ITEM] To [_ORDER_LAST_ITEM]
    Call cmbOrderBy.AddItem(GetOrderName(i), i)
  Next i
  
  cmbOrderBy.ListIndex = p11d32.ReportPrint.CHECKORDERBY
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Description, Err, "FillOrderByComboBox", "FillOrderByComboBox", Err.Description)
  Resume err_end
End Sub

Private Function GetOrderName(i As CHECKORDERBY) As String
  Dim sName As String
  On Error GoTo err_err
  
  Select Case i
      Case ORDER_PNUM
        sName = "Personnel Number"
      Case ORDER_SURNAME
        sName = "Surname"
      Case ORDER_FULLNAME
        sName = "Fullname"
      Case ORDER_NI
        sName = "NI Number"
    End Select
    GetOrderName = sName
err_err:
  Exit Function
err_end:
  Call ErrorMessage(Err.Number, Err, "GetOrderName", "GetOrderName", Err.Description)
  Resume err_err
End Function

Private Sub btnCancel_Click()
  Me.Hide
End Sub

Private Function ChecksSelected(tvw As TreeView) As Boolean
' Determine whether any checks have been selected
  Dim tnRoot As node
  Dim tnChild As node
  Dim i As Integer

  On Error GoTo ChecksSelected_ERR

  For Each tnRoot In tvw.nodes
    If tnRoot.Children > 0 Then
      Set tnChild = tnRoot.child
      For i = 1 To tnRoot.Children
        If tnChild.Checked = True Then
          ChecksSelected = True
          Exit For
        End If
        Set tnChild = tnChild.Next
      Next i
    End If
  Next

ChecksSelected_END:
Exit Function

ChecksSelected_ERR:
Call ErrorMessage(ERR_ERROR, Err, "CheckSelected", "Checks treeview", "Failed to find selected checks")
End Function


Private Sub btnNext_Click()
  On Error GoTo btnNext_ERR
      
  If btnNext.Caption = "&Next" Then
    If ChecksSelected(tvwChecks) = True Then
      tvwCheckResults.nodes.Clear
      Set m_tnNodeSelected = Nothing
      Call DoChecks
      tabCheckWizard.tab = 1
      btnBack.Enabled = True
      btnNext.Caption = "&Finish"
      btnCancel.Caption = "&Cancel"
      Call SelectFirstCheck
      
    Else
      Call MsgBox("Please select a check to run.", vbExclamation, "Company Car Checker")
    End If
  Else
    Call CheckOverlapDirty
    m_bPrint = True
    Unload Me
  End If
btnNext_END:
  Exit Sub
btnNext_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "button Next click", "button Next click", "Error on button next click.")
  Resume btnNext_END
  Resume
End Sub

Private Sub DoChecks()
  Dim i As Integer
  Dim j As Integer
  Dim lEEChecks As Long
  Dim tnRoot As node
  Dim tnChild As node
  
  On Error GoTo DoChecks_ERR
  
  Call SetCursor(vbHourglass)
  
  If m_NodesOL Is Nothing Then Set m_NodesOL = New ObjectList
  Call m_NodesOL.RemoveAll
  
  lblStatus.Visible = True
  lblStatus.Caption = "Starting checks"
  lblStatus.Refresh
  
  'Company car checks
  m_bInError = False
  
  For Each tnRoot In tvwChecks.nodes
    Set tnChild = tnRoot.child
    For i = 1 To tnRoot.Children
      If tnChild.Checked Then
        j = tnChild.Tag
        If j < [_CK_EC_START] And j >= [_CK_FIRST_ITEM] Then
          Call RunCompanyCarChecker(j)
        Else
          lEEChecks = lEEChecks + (2 ^ j)
        End If
      End If
      Set tnChild = tnChild.Next
    Next i
  Next
  
  If lEEChecks > 0 Then Call EmployeeChecks(lEEChecks)
  
  Call FormatTreeview(tvwCheckResults)
  
  Call ChangeNodeDisplay(p11d32.ReportPrint.CHECKORDERBY)
DoChecks_END:
  lblStatus.Visible = False
  Call ClearCursor
  Exit Sub
DoChecks_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Do Checks", "Do Checks", "Error in DoChecks Co Car Checker Wizard.")
  Resume DoChecks_END
  Resume
End Sub


Private Sub CompanyCarCheckerFuncEnd()
  On Error GoTo CompanyCarCheckerFuncEnd_END
  
  If Not m_ey.CurrentEmployee Is Nothing Then
    m_ey.CurrentEmployee.LoadBenefits (TBL_ALLBENEFITS)
    If IsBenefitForm(CurrentForm) And Not CurrentForm Is F_Employees Then
      Call BenScreenSwitchEnd(CurrentForm)
    End If
  End If
    
CompanyCarCheckerFuncEnd_END:
  Exit Sub
CompanyCarCheckerFuncEnd_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CompanyCarCheckerFuncEnd"), Err.Description)
End Sub

Private Sub CompanyCarCheckerFuncStart()
' Kill current benefits
  On Error GoTo CompanyCarCheckerFuncStart_ERR
  
  If Not m_ey.MoveMenuUpdateEmployee Then Call Err.Raise(ERR_MOVE_MENU_UPDATE_EMPLOYEE, "CompanyCarCheckerFuncStart", "Failed to update employee.")
  
  Call m_ey.KillEmployeesBenefits
  
  If m_ey.db Is Nothing Then Call Err.Raise(ERR_DB_IS_NOTHING, "IProgress_Progress", "Employer db is nothing when trying company car checker.")
  
CompanyCarCheckerFuncStart_END:
  Exit Sub
CompanyCarCheckerFuncStart_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CompanyCarCheckerFuncStart"), Err.Description)
  Resume CompanyCarCheckerFuncStart_END
End Sub
Private Function AddFolderNode(c As CHECKS) As node
  Dim n As node
  Dim sKey As String
  
  sKey = S_NODE_KEY & c
  
  
  Set n = InTreeViewByKey(tvwCheckResults, sKey)
  
  If (n Is Nothing) Then
    Set n = tvwCheckResults.nodes.Add(, , sKey, CheckListCaption(c, CMT_TREEVIEW_NODE_TITLE), IMG_FOLDER_CLOSED)
    n.Tag = sKey
    n.ExpandedImage = IMG_FOLDER_OPEN
    Set AddFolderNode = n
  Else
    Set AddFolderNode = n
  End If
End Function

Private Function InTreeViewByTag(tvw As TreeView, vTag As Variant) As node
  Dim n As node
  
  On Error Resume Next
  Set InTreeViewByTag = Nothing
  For Each n In tvw.nodes
    If InTreeViewByTagEx(n, vTag) Then
      Set InTreeViewByTag = n
      Exit Function
    End If
  Next
End Function
Private Function InTreeViewByKey(tvw As TreeView, ByVal vKey As Variant) As node
  On Error Resume Next
  Set InTreeViewByKey = tvw.nodes(vKey)
  
End Function

Private Function InTreeViewByTagEx(n As node, vTag As Variant) As node
  Dim nChild As node
  
  On Error Resume Next
  
  If (n.Tag = vTag) Then
    InTreeViewByTagEx = True
    Exit Function
  End If
  If Not n.child Is Nothing Then
    Set nChild = n.Next
    Do While Not n Is Nothing
      If nChild.Tag = vTag Then
        InTreeViewByTagEx = True
        Exit Function
      End If
      If InTreeViewByTagEx(nChild, vTag) Then
        InTreeViewByTagEx = True
        Exit Function
      End If
      Set nChild = nChild.Next
    
    Loop
  End If
  
  
End Function

Private Function GetFolderNode(c As CHECKS) As node
  Dim sKey As String
  sKey = S_NODE_KEY & c
  Set GetFolderNode = tvwCheckResults.nodes(sKey)
  
End Function
Private Function DoingCheck(c As CHECKS, iChecks As Long) As Boolean
  Dim i As Long
  i = 2 ^ c
  
  DoingCheck = (iChecks And i) = i
End Function
Private Sub EmployeeChecks(lChecks As Long)
  Dim n As node
  Dim i As Long
  Dim ee As Employee
  
  On Error GoTo EmployeeChecks_ERR
  
  lblStatus.Caption = "Calculating Employee checks..."
  lblStatus.Refresh
  For i = 1 To m_TempOL.Count
    Set ee = m_TempOL(i)
    Set n = Nothing
    If DoingCheck(CK_EC_NI, lChecks) Then
      Call CheckNINumbers(ee)
    End If
    
  Next i
 
EmployeeChecks_END:
  Exit Sub
EmployeeChecks_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "EmployeeChecks"), Err.Description)
  Resume EmployeeChecks_END
  Resume
End Sub
Private Function DataKey(ByVal sData As String) As String
  DataKey = SetPropertyFromString("", S_DATA_KEY, sData)
End Function
Private Function DataKeyEE(ByVal ee As Employee) As String
  DataKeyEE = SetPropertyFromString("", S_DATA_KEY, ee.PersonnelNumber)
End Function
Private Sub CheckNINumbers(ee As Employee)
  Dim child As node
  Dim tnCheckNI As node
  On Error GoTo CheckNINumbers_ERR
  
  If ee.NINumberValid Then GoTo CheckNINumbers_END
  m_bInError = True
  Set tnCheckNI = AddFolderNode(CK_EC_NI)
  Set child = tvwCheckResults.nodes.Add(tnCheckNI, tvwChild)
  child.Tag = DataKeyEE(ee)
  tnCheckNI.Sorted = True
  
CheckNINumbers_END:
  Exit Sub
CheckNINumbers_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CheckNINumbers"), Err.Description)
  Resume CheckNINumbers_END
End Sub

Private Function GetOrderDisplay(ben As IBenefitClass, iOrderBy As CHECKORDERBY) As String
  Dim sName As String
  
  On Error GoTo err_err
  
  Select Case iOrderBy
      Case ORDER_PNUM
        sName = ben.value(ee_PersonnelNumber_db)
      Case ORDER_SURNAME
        sName = ben.value(ee_Surname_db)
      Case ORDER_FULLNAME
        sName = ben.value(ee_FullName)
      Case ORDER_NI
        sName = ben.value(ee_NINumber_db)
    End Select
    If (Len(sName) = 0) Then
      sName = "'blank'"
    End If
    GetOrderDisplay = sName
err_err:
  Exit Function
err_end:
  Call ErrorMessage(Err.Number, Err, "GetOrderName", "GetOrderName", Err.Description)
  Resume err_err
End Function

Private Function GetOrderByName(i As CHECKORDERBY) As String
  Dim sName As String
  On Error GoTo err_err
  
  Select Case i
      Case ORDER_PNUM
        sName = S_FIELD_PERSONEL_NUMBER
      Case ORDER_SURNAME
        sName = "SURNAME"
      Case ORDER_FULLNAME
        sName = "FULLNAME"
      Case ORDER_NI
        sName = "NI"
    End Select
    GetOrderByName = sName
err_err:
  Exit Function
err_end:
  Call ErrorMessage(Err.Number, Err, "GetOrderName", "GetOrderName", Err.Description)
  Resume err_err
End Function

Private Sub ChangeNodeDisplay(cob As CHECKORDERBY)
  Dim tnNode As node
  Dim sDisplay As String, sData As String
  Dim ee As Employee
  
  On Error GoTo err_err
  
  Call CheckOverlapDirty
  
  p11d32.ReportPrint.CHECKORDERBY = cob
  
  For Each tnNode In tvwCheckResults.nodes
    If tnNode.Children = 0 Then
      If Len(tnNode.Tag) > 0 Then
        sData = GetPropertyFromString(tnNode.Tag, S_DATA_KEY)
        If Len(sData) > 0 Then
          Set ee = p11d32.CurrentEmployer.FindEmployee(sData)
          If (Not ee Is Nothing) Then
            tnNode.Text = GetOrderDisplay(ee, cob)
          Else
            tnNode.Text = sData
          End If
        Else
          Debug.Print 1
        End If
      Else
        Debug.Print 1
      End If
    End If
  Next
  For Each tnNode In tvwCheckResults.nodes
    If tnNode.Children > 0 Then tnNode.Sorted = True
  Next
  If Not m_tnNodeSelected Is Nothing Then
  
    lblEeName.Caption = m_tnNodeSelected.Text
  End If
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "ChangeNodeDisplay", "ChangeNodeDisplay", Err.Description)
  Resume err_end
  Resume
End Sub

'this is the recordset that fills the results tree
Private Function DataCheckerTreeViewRecordset(ByVal iCheck As CHECKS) As Recordset
  Dim rs As Recordset
  Dim sOrderBy As String
  
  If iCheck = CK_CARS_IN_USE_BY_MORE_THAN_ONE_EMPLOYEE Then
    sOrderBy = "DisplayName"
  Else
    sOrderBy = GetOrderByName(p11d32.ReportPrint.CHECKORDERBY)
  End If
  
  Set rs = m_ey.db.OpenRecordset(sql.Queries(GetCCSQL(iCheck), sOrderBy))
  Set DataCheckerTreeViewRecordset = rs
End Function
Private Sub RunCompanyCarChecker(ByVal iCheck As CHECKS)
  Dim root As node
  Dim child As node
  Dim rs As Recordset
  Dim i As Integer
  Dim sData As String, PNum As String
  Dim sDisplayName As String
  Dim sOrderBy As String, sKey As String
  Dim ee As Employee
  
  On Error GoTo err_err
  
  Set rs = DataCheckerTreeViewRecordset(iCheck)
  'No Records returned so leave function
  If Records(rs) = 0 Then GoTo err_end
  
  lblStatus.Caption = "Calculating " & CheckListCaption(iCheck, CMT_TREEVIEW_NODE_TITLE) & "..."
  lblStatus.Refresh
  
  
  If iCheck <> CK_CARS_IN_USE_BY_MORE_THAN_ONE_EMPLOYEE Then
    PNum = rs.Fields(S_FIELD_PERSONEL_NUMBER).value
  Else
    PNum = "ALL"
  End If
  
  Do While Not rs.EOF
    If (UserSelected(PNum)) Then
      
      Set root = AddFolderNode(iCheck)
      m_bInError = True
      sData = rs.Fields("Data")
      Set child = tvwCheckResults.nodes.Add(root, tvwChild)
      child.Tag = DataKey(sData)
      child.Text = sData
      i = i + 1
    End If
    rs.MoveNext
  Loop
  
  If (i > 0) Then
    root.Text = root.Text + " (" + CStr(i) + ")"
  End If
  tvwCheckResults.Sorted = True
err_end:
  Exit Sub
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CompanyCarCheckerRegDates end"), Err.Description)
  Resume err_end
  Resume
End Sub

Private Function UserSelected(sPNum As String) As Boolean
  Dim i As Integer
  Dim bFound As Boolean
  Dim ee As Employee
  
  On Error GoTo err_err
    
  If sPNum = "ALL" Then
    bFound = True
    GoTo err_end
  End If
    
  If m_bSelected Then
    For i = 1 To m_TempOL.Count
      Set ee = m_TempOL(i)
      If StrComp(sPNum, ee.PersonnelNumber, vbBinaryCompare) = 0 Then
        bFound = True
      End If
    Next
  Else
    bFound = True
  End If
  
err_end:
  UserSelected = bFound
  Exit Function
err_err:
  Call ErrorMessage(Err.Number, Err, "DeleteUnselected", "Could not delete selected records", Err.Description)
  Resume err_end
  Resume
End Function

Private Sub ShowOverlapData(sSQL As String, iCheckType As CHECKS)
  
  On Error GoTo ShowOverlapData_ERR
  
  m_CheckType = iCheckType

  ctlOverlappingCars.Visible = True
  dtgCheckWizard.Visible = False
  
  Call ctlOverlappingCars.DrawOverlaps(sSQL)

ShowOverlapData_END:
  Exit Sub
ShowOverlapData_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "Show Data in Grid"), Err.Description)
  Resume ShowOverlapData_END
  Resume
End Sub
Private Sub LockColumn(ByVal iColIndex As Long)
  dtgCheckWizard.Columns(iColIndex).Locked = True
  dtgCheckWizard.Columns(iColIndex).BackColor = RGB(180, 180, 180)
End Sub

Private Sub FormatGrid()
  
  On Error GoTo FormatGrid_ERR

  Select Case m_CheckType
    Case CK_CC_AVAILDATES
      dtgCheckWizard.Columns(0).Visible = False
      dtgCheckWizard.Columns(1).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, car_Registration_db)
      Call LockColumn(1)
      dtgCheckWizard.Columns(2).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableFrom_db)
      dtgCheckWizard.Columns(3).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableTo_db)
      
    Case CK_CC_REGDATES
      dtgCheckWizard.Columns(0).Visible = True
      Call LockColumn(0)
      dtgCheckWizard.Columns(0).Caption = p11d32.BenDataLinkUDMDisplayName(BC_EMPLOYEE, ee_PersonnelNumber_db)
      
      dtgCheckWizard.Columns(1).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, car_Registration_db)
      Call LockColumn(1)
      
      dtgCheckWizard.Columns(2).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableFrom_db)
       
      dtgCheckWizard.Columns(3).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableTo_db)
      dtgCheckWizard.Columns(4).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, car_Registrationdate_db)
      
    Case CK_CC_EE_AVAILDATES
      
      dtgCheckWizard.Columns(S_FIELD_CAR_REGISTRATION).Locked = True
      dtgCheckWizard.Columns(0).Visible = False
      dtgCheckWizard.Columns(1).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, car_Registration_db)
      Call LockColumn(1)
      dtgCheckWizard.Columns(2).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableFrom_db)
      dtgCheckWizard.Columns(3).Caption = p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableTo_db)
      'dtgCheckWizard.Columns(4).Caption = p11d32.BenDataLinkUDMDisplayName(BC_EMPLOYEE, ee_joined_db)
      
      Call LockColumn(4)
      'dtgCheckWizard.Columns(5).Caption = p11d32.BenDataLinkUDMDisplayName(BC_EMPLOYEE, ee_left_db)
      Call LockColumn(5)
      
    Case CK_EC_NI
      dtgCheckWizard.Columns(0).Visible = False
      dtgCheckWizard.Columns(1).Caption = p11d32.BenDataLinkUDMDisplayName(BC_EMPLOYEE, ee_NINumber_db)
    Case Else
      ECASE "Unknown checktype"
  End Select
  
  Call SetColumnWidths(dtgCheckWizard)
  
FormatGrid_END:
  Exit Sub
FormatGrid_ERR:
  Call ErrorMessage(Err.Number, Err, "FormatGrid", "FormatGrid", Err.Description)
  Resume FormatGrid_END
End Sub

Private Sub ShowDataInGrid(iCheckType As CHECKS, sSQL As String)
  Dim i As Integer
  Dim j As Integer
  On Error GoTo ShowDataInGrid_ERR
    
  dbCoCarChecker.RecordSource = sSQL
  dbCoCarChecker.Refresh
  dtgCheckWizard.Enabled = True
  
  m_CheckType = iCheckType
  
  Call FormatGrid
  
  dtgCheckWizard.Visible = True
  ctlOverlappingCars.Visible = False
  
ShowDataInGrid_END:
  Exit Sub
ShowDataInGrid_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "Show Data in Grid"), Err.Description)
  Resume ShowDataInGrid_END
  Resume
End Sub

Private Sub SetColumnWidths(dtgGrdid As TDBGrid)
  Dim i As Integer
  Dim j As Integer
  
  On Error GoTo err_err
  
  j = 0
  For i = 0 To dtgGrdid.Columns.Count - 1
    If dtgGrdid.Columns(i).Visible = True Then
      j = j + 1
    End If
  Next i
  
  For i = 0 To dtgGrdid.Columns.Count - 1
    dtgGrdid.Columns(i).width = (dtgGrdid.width - 500) / j
  Next i
  
err_end:
  Exit Sub
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "Setting col widths"), Err.Description)
  Resume err_end
End Sub

Private Sub SetCheckImageList(tvw As TreeView, iml As ImageList)
  Call TreeView_SetImageList(tvw.hwnd, iml.hImageList, TVSIL_STATE)
End Sub

Private Sub FormatTreeview(tvw As TreeView)
  Dim tnNode As node
  
  For Each tnNode In tvw.nodes
    If Len(GetPropertyFromString(tnNode.Tag, "DATA")) = 0 Then
      tnNode.Expanded = True
    End If
  Next
  
  Call SetCheckImageList(tvw, imgCheckBox)
  
On Error GoTo err_err
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "FormatTreeview", "Error formatting treeview" & tvw.Name, Err.Description)
  Resume err_end
End Sub


Private Sub btnSave_Click()
  
  On Error GoTo btnSave_ERR
  
  Call ctlOverlappingCars.SaveOverlappingCCC(m_ey.db)
  
  If p11d32.ReportPrint.ChecksAutoRefresh Then Call RefreshChecks
  
btnSave_END:
  Exit Sub
btnSave_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "node click error"), Err.Description)
  Resume btnSave_END
  Resume
End Sub

Private Sub dtgCheckWizard_BeforeUpdate(Cancel As Integer)

  On Error GoTo dtgCheckWizard_BeforeUpdate_ERR

  Select Case m_CheckType
    Case CK_CC_REGDATES, CK_CC_EE_AVAILDATES, CK_CC_AVAILDATES
      Call SaveCCCheckGrid
    Case CK_EC_NI
      Call SaveEeCheckGrid
  End Select
    
  Cancel = True
  If p11d32.ReportPrint.ChecksAutoRefresh Then Call RefreshChecks
dtgCheckWizard_BeforeUpdate_END:
  Exit Sub
dtgCheckWizard_BeforeUpdate_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "dtgCheckWizard_BeforeUpdate", "dtgCheckWizard_BeforeUpdate", Err.Description)
  Cancel = True
  Resume dtgCheckWizard_BeforeUpdate_END
End Sub


Private Sub SaveEeCheckGrid()
  Dim ben As IBenefitClass
  Dim ee As Employee
  Dim i As Integer
  Dim sNewPN As String
  Dim t As TableDef
    
  On Error GoTo SaveEeCheckGrid_ERR
  
  Set ben = m_ey.FindEmployee(CStr(dtgCheckWizard.Columns("P_Num").value))
  If Not ben Is Nothing Then
    Select Case m_CheckType
      Case CK_EC_NI
        ben.value(ee_NINumber_db) = CStr(dtgCheckWizard.Columns(p11d32.BenDataLinkUDMDisplayName(BC_EMPLOYEE, ee_NINumber_db)).value)
        ben.Dirty = True
        ben.writeDB
    End Select
    Call UpdateEeListview(ben, CStr(dtgCheckWizard.Columns("P_Num").value))
  End If
  
SaveEeCheckGrid_END:
  Exit Sub
SaveEeCheckGrid_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SaveEeCheckGrid", "Saving Employee info", Err.Description)
  Resume SaveEeCheckGrid_END
End Sub

Private Sub SaveCCCheckGrid()
  
  Dim ee As Employee
  Dim ben As IBenefitClass
  Dim i As Integer
  Dim j As Integer
  On Error GoTo SaveCCCheckGrid_ERR
  
  Set ee = m_ey.FindEmployee(CStr(dtgCheckWizard.Columns("P_Num").value))
  
  If Not ee Is Nothing Then
    'Load company car benefits
    ee.LoadBenefits (TBL_COMPANY_CARS)
    For j = 1 To ee.benefits.Count
      Set ben = ee.benefits(j)
      'get correct car benefit
      If ben.value(car_Registration_db) = dtgCheckWizard.Columns(p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, car_Registration_db)) Then
        Select Case m_CheckType
          Case CK_CC_AVAILDATES, CK_CC_EE_AVAILDATES
            ben.value(Car_AvailableFrom_db) = dtgCheckWizard.Columns(p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableFrom_db))
            ben.value(Car_AvailableTo_db) = dtgCheckWizard.Columns(p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableTo_db))
          Case CK_CC_REGDATES
            ben.value(Car_AvailableFrom_db) = dtgCheckWizard.Columns(p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableFrom_db))
            ben.value(Car_AvailableTo_db) = dtgCheckWizard.Columns(p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, Car_AvailableTo_db))
            ben.value(car_Registrationdate_db) = dtgCheckWizard.Columns(p11d32.BenDataLinkUDMDisplayName(BC_COMPANY_CARS_F, car_Registrationdate_db))
        End Select
        'write changes to db
        Call ben.writeDB
      End If
    Next j
  Else
    Call Err.Raise(Err.Number, ErrorSource(Err, "SaveOverlappingCCC"), "Could not find employee ")
  End If
  
SaveCCCheckGrid_END:
  Exit Sub
SaveCCCheckGrid_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SaveRegDatesCheckGrid", "Saving Reg Dates info", Err.Description)
  Resume SaveCCCheckGrid_END
  Resume
End Sub


Private Sub UpdateEeListview(ben As IBenefitClass, sPNum As String)
  
  Dim ibf As IBenefitForm2
  Dim i As Integer
  
  On Error GoTo UpdateEeListview_ERR
  
  Set ibf = F_Employees
  For i = 1 To ibf.lv.listitems.Count
    If ibf.lv.listitems(i).SubItems(1) = sPNum Then
      Call ibf.UpdateBenefitListViewItem(ibf.lv.listitems(i), ben)
    End If
  Next i
  
UpdateEeListview_END:
  Exit Sub
UpdateEeListview_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateEeListview", "UpdateEeListview", Err.Description)
  Resume UpdateEeListview_END
End Sub

Private Sub dtgCheckWizard_Error(ByVal DataError As Integer, Response As Integer)
'ignore this error - as provides a message box saying that the update has been cancelled
If DataError = 16389 Then Response = vbDataErrContinue
End Sub

Private Function GetNodeClickSQL(ByVal iCheck As CHECKS, sData As String) As String
  Dim QN As QUERY_NAMES
  On Error GoTo err_err
  
  Select Case iCheck
    Case CK_CARS_IN_USE_BY_MORE_THAN_ONE_EMPLOYEE
      QN = SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE
    Case CK_CC_OVERLAPS, CK_CC_NOCARS
      QN = SELECT_COMPANYCAR_CHECKER_CARS_OVERLAP
    Case CK_CC_SEQUENTIAL_NOT_MARKED_AS_REPLACED
      QN = SELECT_COMPANYCAR_CHECKER_CARS_OVERLAP
    Case CK_CC_AVAILDATES
      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_AVAILDATES
    Case CK_CC_REGDATES
      QN = SELECT_COMPANYCAR_CHECKER_REG_DATES_NODE_CLICK
    Case CK_CC_EE_AVAILDATES
      QN = SELECT_COMPANYCAR_CHECKER_CARS_EEE_AVAIL_DATES_TV
    Case CK_EC_NI
      QN = SELECT_EMPLOYEE_CHECK_NI_VALID_TV
    Case Else
      ECASE ("Unknown check value")
  End Select
    
  GetNodeClickSQL = sql.Queries(QN, sData)

err_end:
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "Could not find SQL statment to fetch data"), Err.Description)
  Resume err_end
End Function

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  
End Function

Private Sub CheckOverlapDirty()
  On Error GoTo err_err
  
  If ctlOverlappingCars.Visible And ctlOverlappingCars.Dirty Then
    Call ctlOverlappingCars.SaveOverlappingCCC(m_ey.db)
  End If
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "CheckOverlapDirty", "CheckOverlapDirty", Err.Description)
  Resume err_end
End Sub
Private Sub ChangeCheck(node As node)
  Dim sData As String
  Dim iCheckType As CHECKS
  Dim sKey As String
  Dim sSQL As String
  
  On Error GoTo ChangeCheck_ERR
  
  
  If tvwCheckResults.nodes.Count < 2 Then GoTo ChangeCheck_END
  Call CheckOverlapDirty
    
  If node.Children = 0 Then
    sKey = node.Parent.Tag
    iCheckType = Replace(sKey, S_NODE_KEY, "")
    sData = GetPropertyFromString(node.Tag, S_DATA_KEY)
    Me.ctlOverlappingCars.DragDropMode = (iCheckType = CK_CC_SEQUENTIAL_NOT_MARKED_AS_REPLACED)
        
    sSQL = GetNodeClickSQL(iCheckType, sData)
    
    Select Case iCheckType
      Case CK_CARS_IN_USE_BY_MORE_THAN_ONE_EMPLOYEE, CK_CC_NOCARS, CK_CC_OVERLAPS, CK_CC_SEQUENTIAL_NOT_MARKED_AS_REPLACED
        Call ShowOverlapData(sSQL, iCheckType)
      Case CK_CC_AVAILDATES, CK_CC_REGDATES, CK_CC_EE_AVAILDATES, CK_EC_NI
        Call ShowDataInGrid(iCheckType, sSQL)
      Case Else
        ECASE ("Unknown checktype")
    End Select
    
    If iCheckType = CK_CC_SEQUENTIAL_NOT_MARKED_AS_REPLACED Then
      pctInfo.Visible = True
      lblInfo.Caption = "Drag the end of a bar for a car to the start of another cars' bar to mark it as replaced. To delete a link press the - button to the right of the car, this clears the replacement flag. Only sequential cars can be linked. To change the dates for a car change the dates in the boxes."
    Else
      pctInfo.Visible = False
    End If
    
    
    lblCheckType.Caption = CheckListCaption(iCheckType, CMT_LIST_ITEM)
    lblEeName.Caption = node.Text
    Set m_tnNodeSelected = tvwCheckResults.SelectedItem
  Else
    If Not IsChildSelected(tvwCheckResults, node) Then
      'select the first child
      node.child.Selected = True
    End If
  End If
  
ChangeCheck_END:
  Exit Sub
ChangeCheck_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "node click error"), Err.Description)
  Resume ChangeCheck_END
  Resume
End Sub
Private Function IsChildSelected(tvw As TreeView, n As node)
  If Not tvw.SelectedItem Is Nothing Then
    If Not tvw.SelectedItem.Parent Is Nothing Then
      If tvw.SelectedItem.Parent Is n Then
        IsChildSelected = True
      End If
    End If
  End If
End Function

Private Sub tvwCheckResults_NodeClick(ByVal node As MSComctlLib.node)
  Call ChangeCheck(node)
End Sub

Private Sub NodesCheck(ByVal node As node)
  Dim tnRoot As node, tnChild As node
  Dim i As Integer
  Dim j As Integer
  Dim bSelect As Boolean
  
  On Error GoTo tvwChecks_NodeCheck_ERR
  m_bCheckDown = True
  If node.Children > 0 Then
    Set tnChild = node.child
    For i = 1 To node.Children
      tnChild.Checked = node.Checked
      Set tnChild = tnChild.Next
    Next i
  Else
    For Each tnRoot In tvwChecks.nodes
      If tnRoot.Children > 0 Then
        tnRoot.Checked = True
        For Each tnChild In tvwChecks.nodes
          If (tnChild.Children = 0) Then
            If tnChild.Parent = tnRoot Then
              If tnChild.Checked = False Then tnRoot.Checked = False
            End If
          End If
        Next
      End If
    Next
  End If
  chkAllChecks.value = BoolToChkBox(False)
  p11d32.ReportPrint.ChecksSelected = GetCheckedNodes
  btnNext.Enabled = ChecksSelected(tvwChecks)
  m_bCheckDown = False
tvwChecks_NodeCheck_END:
  Exit Sub
tvwChecks_NodeCheck_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "node check error"), Err.Description)
  Resume tvwChecks_NodeCheck_END
  Resume

End Sub


Private Sub tvwChecks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim n As node
  
  On Error GoTo err_err
  
  Set n = tvwChecks.HitTest(X, Y)
  If Not n Is Nothing Then
    If IsNumeric(n.Tag) Then
      lblDescription.Caption = CheckListCaption(n.Tag, CMT_ALERT_MESSAGE_DESCRIPTION)
    End If
    
  End If
err_end:
  Exit Sub
err_err:
  'no point on mouse move
  Resume err_end

  
End Sub

Private Sub tvwChecks_NodeCheck(ByVal node As MSComctlLib.node)
  Call NodesCheck(node)
End Sub

Private Function GetCheckedNodes() As Long
  Dim tnNode As node
  Dim i As Long
  Dim iChecks As Long
  On Error GoTo err_err
  
  For Each tnNode In tvwChecks.nodes
    If tnNode.Children = 0 And tnNode.Checked = True Then
      i = tnNode.Tag
      iChecks = iChecks + (2 ^ i)
    End If
  Next
  
  GetCheckedNodes = iChecks
  
err_end:
  Exit Function
err_err:
  Call ErrorMessage(Err.Number, Err, "GetCheckedNodes", "GetCheckedNodes", Err.Description)
  Resume err_end
End Function

Private Function GetCCSQL(iCheck As CHECKS, Optional sData As String) As QUERY_NAMES

  On Error GoTo err_err
  
  Select Case iCheck
    Case CK_CC_REGDATES
      GetCCSQL = SELECT_COMPANYCAR_CHECKER_CARS_REGDATES
    Case CK_CC_EE_AVAILDATES
      GetCCSQL = SELECT_COMPANYCAR_CHECKER_CARS_EEE_AVAIL_DATES
    Case CK_CC_NOCARS
      GetCCSQL = SELECT_COMPANYCAR_CHECKER_CARS_NOCARS
    Case CK_CC_AVAILDATES
      GetCCSQL = SELECT_COMPANYCAR_CHECKER_CARS_AVAILDATES_TV
    Case CK_CARS_IN_USE_BY_MORE_THAN_ONE_EMPLOYEE
      GetCCSQL = SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE_TV
    Case CK_CC_OVERLAPS
      GetCCSQL = SELECT_COMPANYCAR_CHECKER_CARS_OVERLAP_TV
    Case CK_CC_SEQUENTIAL_NOT_MARKED_AS_REPLACED
      GetCCSQL = SELECT_COMPANYCAR_CHECKER_SEQUENTIAL_NOT_MARKED_AS_REPLACED_TV
    Case Else
      Call Err.Raise(Err.Number, "Get CC SQL", "Unknown checktype")
  End Select
  
err_end:
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "GetCCSQL"), Err.Description)
  Resume err_end
End Function

Private Function GetGridSQL(iCheckType As CHECKS, sData As String) As String
  Dim iQuery As QUERY_NAMES
  
On Error GoTo err_err
  Select Case iCheckType
  
    Case CK_CC_AVAILDATES
      iQuery = SELECT_COMPANYCAR_CHECKER_CARS_LOG_AVAILDATES
    Case CK_CC_REGDATES
      iQuery = SELECT_COMPANYCAR_CHECKER_REG_DATES_NODE_CLICK
    Case CK_CC_EE_AVAILDATES
      iQuery = SELECT_COMPANYCAR_CHECKER_CARS_EEE_AVAIL_DATES_TV
    Case CK_EC_NI
      iQuery = SELECT_EMPLOYEE_CHECK_NI_VALID_TV
    Case Else
      Call Err.Raise(ERR_INVALID, "GetGridSQL", "Unknown checktype to be displayed in grid")
  End Select
  GetGridSQL = sql.Queries(iQuery, sData)
  
err_end:
  Exit Function
err_err:
  Call ErrorMessage(Err.Number, Err, "GetGridSQL", "Finding sql statement for grid", Err.Description)
  Resume err_end
End Function

Private Function CheckListCaption(ByVal CHK_TYPE As CHECKS, ByVal CMT As CHECK_MESSAGE_TYPE) As String
  Dim s As String
  On Error GoTo CheckListCaption_ERR
  
  Select Case CHK_TYPE
    Case CK_CARS_IN_USE_BY_MORE_THAN_ONE_EMPLOYEE
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "Cars in use by more than one employee"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Cars that are in use by two or more employees at the same time"
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Cars in use > 1 ee"
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "If a car with the same registration number is used by more than one employee at the same time this will be highlighted."
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
    
    Case CK_CC_OVERLAPS
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "Employees with two or more overlapping cars"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Employees who have two or more cars available to them on one day." & vbCrLf & vbCrLf & _
              "It will indicate P46(Car) flags that should be set whether the cars have registration dates after the date of first use."
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "If an employee has two cars available on the same day that overlap it is likely that the data is incorrect. It is common for fleet car data to have the wrong days recorded for cars. If a car ends on one day and another starts on the same day this will not be hilighted."
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Ee's with > 2 overlapping cars"
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
    Case CK_CC_SEQUENTIAL_NOT_MARKED_AS_REPLACED
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "Cars that are squential but not marked as replaced"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Employees who have two or more cars available to them on one day." & vbCrLf & vbCrLf & _
              "It will indicate P46(Car) flags that should be set whether the cars have registration dates after the date of first use."
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "If an employee has cars that have sequential dates and they are not marked as replacements for P46(Car) purposes they will be listed."
          
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Sequential cars not replaced"
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
    
    Case CK_CC_NOCARS
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "Employees with gaps between car usage"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Employees not having a car available for all days in the tax year."
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "If an employee has two cars where there is a gap between useage this is highlighted. Although it is not always the case that it is an error, the vast majority of cases that surface are caused by data entry issues."
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Ee's with gaps between car usage"
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
    
    Case CK_CC_AVAILDATES
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "Car's start date occurring after end date"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Any cars that have their 'Available to' date before their 'Available from' date"
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "If a car's start date occurs after the end date this is highlighted. It is not practical to prevent this when data is entered into the system hence this check is used."
          
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Cars start > end date"
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
      
    Case CK_CC_REGDATES
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "Registration dates of cars inconsistent"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Registration dates that are inconsistent with availability dates"
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "If a car's registration date is after the start date this is highlighted. This is a common error from data downloaded from fleet car systems."
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Car reg dates inconsistent"
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
    Case CK_CC_EE_AVAILDATES
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "Employee available dates inconsistent"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Check date leaving and joining of employee consistent with availability of car"
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "If a car's available dates are inconsistent with the employee's start or end dates this is highlighted."
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Ee avail dates for car"
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
    Case CK_EC_NI
      Select Case CMT
        Case CMT_LIST_ITEM
          s = "National Insurance numbers"
        Case CMT_ALERT_MESSAGE_CHANGE, CMT_ALERT_MESSAGE_CHECK
          s = "Check National Insurance numbers"
        Case CMT_ALERT_MESSAGE_DESCRIPTION
          s = "Checks for invalid NI numbers"
        Case CMT_TREEVIEW_NODE_TITLE
          s = "Invalid NI"
        Case Else
          ECASE ("Invalid Company car check message type = " & CMT)
      End Select
    Case Else
      Call ECASE("Invalid Company car check = " & CHK_TYPE)
  End Select
  
  
  Select Case CMT
    Case CMT_ALERT_MESSAGE_CHANGE
      s = S_CCCC_MESSAGE_PREFIX_CHANGE & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf & S_CCCC_MESSAGE_SUFFIX
    Case CMT_ALERT_MESSAGE_CHECK
      s = S_CCCC_MESSAGE_PREFIX_CHECK & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf & S_CCCC_MESSAGE_SUFFIX
  End Select

  CheckListCaption = s

CheckListCaption_END:
  Exit Function
CheckListCaption_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CheckListCaption", "CCC List Caption", "Error in CheckListCaption.")
  Resume CheckListCaption_END
End Function

Private Sub tvwChecks_NodeClick(ByVal node As MSComctlLib.node)
  If IsNumeric(node.Tag) Then
    node.Checked = Not node.Checked
    Call tvwChecks_NodeCheck(node)
  End If
End Sub
