VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A7CE771F-05B2-43CF-9650-ED841A9049FA}#1.0#0"; "ATC3FolderBrowser.ocx"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_PrintOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print options"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tab 
      Height          =   7110
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   12541
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "HMIT"
      TabPicture(0)   =   "F_PrintOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraReportTotalValue"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fmeHMITSectionChoice"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fmeHMITSections"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Employee Letter"
      TabPicture(1)   =   "F_PrintOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraEmployeeLetter"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraLetters"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "F_PrintOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDestination"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraOther"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraExportOptions"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "frmCheckOptions"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame frmCheckOptions 
         Caption         =   "Check Data"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   63
         Top             =   5640
         Width           =   5655
         Begin VB.CheckBox chkChecksAutoRefresh 
            Caption         =   "Auto refresh of checks"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   960
            Width           =   4575
         End
         Begin VB.CheckBox chkCheckDataBeforePrint 
            Caption         =   "Ask for checks before printing"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   2415
         End
         Begin VB.Frame frmCheckNeverAsk 
            Caption         =   "Frame1"
            Height          =   615
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   5295
            Begin VB.OptionButton optChkBeforePrint 
               Caption         =   "Never run checks"
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   67
               Top             =   240
               Width           =   1815
            End
            Begin VB.OptionButton optChkBeforePrint 
               Caption         =   "Always run checks"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   1815
            End
         End
      End
      Begin VB.Frame fraExportOptions 
         Caption         =   "Automated export / e-mail-attachment options"
         Height          =   1065
         Left            =   -74910
         TabIndex        =   31
         Top             =   2700
         Width           =   5700
         Begin VB.PictureBox pctFrame4 
            BorderStyle     =   0  'None
            Height          =   780
            Left            =   90
            ScaleHeight     =   780
            ScaleWidth      =   5505
            TabIndex        =   56
            Top             =   225
            Width           =   5505
            Begin VB.OptionButton optExportType 
               Caption         =   "HTML Internet explorer 4+"
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   59
               Top             =   0
               Width           =   2625
            End
            Begin VB.OptionButton optExportType 
               Caption         =   "HTML Netscape 4+"
               Height          =   375
               Index           =   1
               Left            =   0
               TabIndex        =   58
               Top             =   360
               Width           =   2625
            End
            Begin VB.OptionButton optExportType 
               Caption         =   "HTML Internet Explorer 5"
               Height          =   330
               Index           =   2
               Left            =   2790
               TabIndex        =   57
               Top             =   45
               Width           =   2085
            End
         End
      End
      Begin VB.Frame fraOther 
         Caption         =   "Other"
         Height          =   1830
         Left            =   -74910
         TabIndex        =   28
         Top             =   3780
         Width           =   5700
         Begin VB.CheckBox chkRememberEmployeeSelection 
            Caption         =   "Remember employee selection"
            Height          =   375
            Left            =   2520
            TabIndex        =   60
            Top             =   1035
            Width           =   2490
         End
         Begin VB.CheckBox chkHMITFieldTrim 
            Caption         =   "Trim fields on P11D"
            Height          =   375
            Left            =   2520
            TabIndex        =   37
            Top             =   630
            Width           =   2310
         End
         Begin VB.CheckBox chkDatesOnWorkingPaper 
            Caption         =   "Dates on working papers"
            Height          =   465
            Left            =   2520
            TabIndex        =   36
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkEmailTextAsHTML 
            Caption         =   "Email text as HTML"
            Height          =   240
            Left            =   135
            TabIndex        =   35
            Top             =   1485
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CheckBox chkSeparatePrintJobs 
            Caption         =   "Separate print jobs / duplex"
            Height          =   330
            Left            =   135
            TabIndex        =   34
            Top             =   1035
            Width           =   2895
         End
         Begin VB.CheckBox chkTimeStamp 
            Caption         =   "Time stamp"
            Height          =   240
            Left            =   135
            TabIndex        =   30
            Top             =   270
            Width           =   1950
         End
         Begin VB.CheckBox chkPrintedEmployees 
            Caption         =   "Employees printed report"
            Height          =   330
            Left            =   135
            TabIndex        =   29
            Top             =   630
            Width           =   2130
         End
      End
      Begin VB.Frame fraDestination 
         Caption         =   "Destination"
         Height          =   2250
         Left            =   -74910
         TabIndex        =   27
         Top             =   405
         Width           =   5700
         Begin VB.PictureBox pctFrame3 
            BorderStyle     =   0  'None
            Height          =   1995
            Left            =   90
            ScaleHeight     =   1995
            ScaleWidth      =   5550
            TabIndex        =   48
            Top             =   180
            Width           =   5550
            Begin atc3FolderBrowser.FolderBrowser fbExportDirectory 
               Height          =   555
               Left            =   45
               TabIndex        =   62
               Top             =   1440
               Width           =   5505
               _ExtentX        =   9710
               _ExtentY        =   979
            End
            Begin VB.OptionButton optDestination 
               Caption         =   "HTML file"
               Height          =   330
               Index           =   1
               Left            =   45
               TabIndex        =   53
               Top             =   285
               Width           =   1125
            End
            Begin VB.OptionButton optDestination 
               Caption         =   "Printer/Preview"
               Height          =   330
               Index           =   0
               Left            =   45
               TabIndex        =   52
               Top             =   0
               Width           =   1725
            End
            Begin VB.CheckBox chkAllowUserReportNameHTML 
               Caption         =   "Allow user export report name"
               Height          =   240
               Left            =   45
               TabIndex        =   51
               Top             =   990
               Width           =   2535
            End
            Begin VB.OptionButton optDestination 
               Caption         =   "PRN file"
               Height          =   330
               Index           =   2
               Left            =   45
               TabIndex        =   50
               Top             =   585
               Width           =   975
            End
            Begin atc2valtext.ValText txtPRNFileName 
               Height          =   315
               Left            =   2580
               TabIndex        =   49
               Top             =   570
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               BackColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   ""
               TypeOfData      =   3
               AllowEmpty      =   0   'False
               AutoSelect      =   0
            End
            Begin VB.Label lblExportTo 
               Caption         =   "Directory for export file:"
               Height          =   285
               Left            =   60
               TabIndex        =   55
               Top             =   1215
               Width           =   1665
            End
            Begin VB.Label Label1 
               Caption         =   "PRN file name"
               Height          =   375
               Left            =   1380
               TabIndex        =   54
               Top             =   630
               Width           =   1155
            End
         End
      End
      Begin VB.Frame fraLetters 
         Caption         =   "Employee letters"
         Height          =   4425
         Left            =   -74910
         TabIndex        =   25
         Top             =   360
         Width           =   5685
         Begin atc3FolderBrowser.FolderBrowser fbUserLetters 
            Height          =   555
            Left            =   90
            TabIndex        =   61
            Top             =   3420
            Width           =   5550
            _ExtentX        =   9790
            _ExtentY        =   979
         End
         Begin VB.CommandButton cmdEditEmployeeLetter 
            Caption         =   "Edit employee letter"
            Height          =   330
            Left            =   3825
            TabIndex        =   38
            Top             =   4005
            Width           =   1785
         End
         Begin MSComctlLib.TreeView tvwLetters 
            Height          =   3000
            Left            =   90
            TabIndex        =   26
            Top             =   225
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   5292
            _Version        =   393217
            Indentation     =   176
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin VB.Label lblUserLettersFolder 
            Caption         =   "User letters folder"
            Height          =   240
            Left            =   135
            TabIndex        =   39
            Top             =   3240
            Width           =   2085
         End
      End
      Begin VB.Frame fraEmployeeLetter 
         Caption         =   "Employee letter"
         Height          =   2160
         Left            =   -74910
         TabIndex        =   19
         Top             =   4815
         Width           =   5685
         Begin VB.TextBox txtEmailSubject 
            Height          =   405
            Left            =   135
            TabIndex        =   32
            Top             =   1380
            Width           =   5370
         End
         Begin VB.CommandButton cmdEmployeeLetterFont 
            Caption         =   "Font"
            Height          =   330
            Left            =   4680
            TabIndex        =   21
            Top             =   675
            Width           =   825
         End
         Begin VB.VScrollBar vsEmployeeLetterMargin 
            Height          =   330
            Left            =   5355
            Max             =   0
            Min             =   15
            TabIndex        =   20
            Top             =   255
            Value           =   15
            Width           =   150
         End
         Begin atc2valtext.ValText vtEmployeeLetterMargin 
            Height          =   330
            Left            =   4905
            TabIndex        =   22
            Top             =   255
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   582
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            Maximum         =   "15"
            Minimum         =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
            AutoSelect      =   0
         End
         Begin VB.Label lblEmailSubject 
            Caption         =   "Email subject"
            Height          =   240
            Left            =   180
            TabIndex        =   33
            Top             =   1155
            Width           =   1140
         End
         Begin VB.Label lblEmployeeLetterFont 
            Caption         =   "Font"
            Height          =   465
            Left            =   180
            TabIndex        =   24
            Top             =   675
            Width           =   1275
         End
         Begin VB.Label lblEmployeeLetterMargin 
            Caption         =   "Margin %"
            Height          =   345
            Left            =   180
            TabIndex        =   23
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Frame fmeHMITSections 
         Caption         =   "HMIT Sections"
         Height          =   5280
         Left            =   135
         TabIndex        =   4
         Top             =   1620
         Width           =   5595
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "A - Assets transferred"
            Height          =   330
            Index           =   1
            Left            =   135
            TabIndex        =   18
            Top             =   300
            Width           =   3525
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "B - Payments made on behalf of the employee"
            Height          =   330
            Index           =   2
            Left            =   135
            TabIndex        =   17
            Top             =   570
            Width           =   3795
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "C - Vouchers or credit cards"
            Height          =   330
            Index           =   3
            Left            =   135
            TabIndex        =   16
            Top             =   840
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "D - Living accommodation"
            Height          =   330
            Index           =   4
            Left            =   135
            TabIndex        =   15
            Top             =   1110
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "E - Mileage allowance"
            Height          =   330
            Index           =   5
            Left            =   135
            TabIndex        =   14
            Top             =   1380
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "F - Cars and fuel"
            Height          =   330
            Index           =   6
            Left            =   135
            TabIndex        =   13
            Top             =   1650
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "G - Vans"
            Height          =   330
            Index           =   7
            Left            =   135
            TabIndex        =   12
            Top             =   1920
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "H - Beneficial loans"
            Height          =   330
            Index           =   8
            Left            =   135
            TabIndex        =   11
            Top             =   2190
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "I - Private medical treatment or insurance"
            Height          =   330
            Index           =   9
            Left            =   135
            TabIndex        =   10
            Top             =   2460
            Width           =   3705
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "J - Qualifying relocation expenses"
            Height          =   330
            Index           =   10
            Left            =   135
            TabIndex        =   9
            Top             =   2730
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "K - Services supplied"
            Height          =   330
            Index           =   11
            Left            =   135
            TabIndex        =   8
            Top             =   3000
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "L - Assets placed at the employee's disposal"
            Height          =   330
            Index           =   12
            Left            =   135
            TabIndex        =   7
            Top             =   3270
            Width           =   3705
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "M - Other items"
            Height          =   330
            Index           =   13
            Left            =   135
            TabIndex        =   6
            Top             =   3540
            Width           =   3165
         End
         Begin VB.CheckBox chkHMITSections 
            Caption         =   "N - Expenses payments"
            Height          =   330
            Index           =   14
            Left            =   135
            TabIndex        =   5
            Top             =   3810
            Width           =   3165
         End
      End
      Begin VB.Frame fmeHMITSectionChoice 
         Caption         =   "HMIT Sections choice"
         Height          =   1140
         Left            =   135
         TabIndex        =   3
         Top             =   405
         Width           =   2175
         Begin VB.PictureBox pctFrame 
            BorderStyle     =   0  'None
            Height          =   825
            Left            =   135
            ScaleHeight     =   825
            ScaleWidth      =   1860
            TabIndex        =   40
            Top             =   225
            Width           =   1860
            Begin VB.OptionButton optHMITSectionChoice 
               Caption         =   "Selected"
               Height          =   285
               Index           =   2
               Left            =   0
               TabIndex        =   43
               Top             =   540
               Width           =   1365
            End
            Begin VB.OptionButton optHMITSectionChoice 
               Caption         =   "Relevant"
               Height          =   285
               Index           =   1
               Left            =   0
               TabIndex        =   42
               Top             =   270
               Width           =   1500
            End
            Begin VB.OptionButton optHMITSectionChoice 
               Caption         =   "All"
               Height          =   285
               Index           =   0
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   1365
            End
         End
      End
      Begin VB.Frame fraReportTotalValue 
         Caption         =   "Total value?"
         Height          =   1140
         Left            =   2385
         TabIndex        =   2
         Top             =   405
         Width           =   3345
         Begin VB.PictureBox pctFrame1 
            BorderStyle     =   0  'None
            Height          =   780
            Left            =   90
            ScaleHeight     =   780
            ScaleWidth      =   3120
            TabIndex        =   44
            Top             =   270
            Width           =   3120
            Begin VB.OptionButton optReportTotalValue 
               Caption         =   "Zero only"
               Height          =   240
               Index           =   2
               Left            =   90
               TabIndex        =   47
               Top             =   540
               Width           =   1500
            End
            Begin VB.OptionButton optReportTotalValue 
               Caption         =   "Non zero only"
               Height          =   240
               Index           =   1
               Left            =   90
               TabIndex        =   46
               Top             =   270
               Width           =   1545
            End
            Begin VB.OptionButton optReportTotalValue 
               Caption         =   "Non zero and zero"
               Height          =   240
               Index           =   0
               Left            =   90
               TabIndex        =   45
               Top             =   0
               Width           =   1635
            End
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   4815
      TabIndex        =   0
      Top             =   7290
      Width           =   1050
   End
End
Attribute VB_Name = "F_PrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_LastSelectedNode As Node
Private m_NodeCount As Long
Private Enum NODE_TYPES
  NODE_FOLDER = 0
  NODE_LETTER
End Enum
Private m_LastUserLetterCount As Long
Private Sub chkAllowUserReportNameHTML_Click()
  p11d32.ReportPrint.ExportAllowUserNameHTML = ChkBoxToBool(chkAllowUserReportNameHTML)
End Sub

Private Sub chkCheckDataBeforePrint_Click()
  Call EnableFrame(Me, frmCheckNeverAsk, Not ChkBoxToBool(chkCheckDataBeforePrint))
  If ChkBoxToBool(chkCheckDataBeforePrint) Then
    p11d32.ReportPrint.CheckOptions = YES_THIS_TIME_ONLY
  Else
    p11d32.ReportPrint.CheckOptions = YES_ALWAYS
  End If
  
  
End Sub

Private Sub chkChecksAutoRefresh_Click()
  p11d32.ReportPrint.ChecksAutoRefresh = ChkBoxToBool(chkChecksAutoRefresh)
End Sub

Private Sub chkDatesOnWorkingPaper_Click()
  p11d32.ReportPrint.DatesOnWorkingPaper = ChkBoxToBool(chkDatesOnWorkingPaper)
End Sub

Private Sub chkEmailTextAsHTML_Click()
  p11d32.ReportPrint.EmailTextAsHTML = ChkBoxToBool(chkEmailTextAsHTML)
End Sub

Private Sub chkHMITFieldTrim_Click()
  p11d32.ReportPrint.HMITFieldTrim = ChkBoxToBool(chkHMITFieldTrim)
End Sub

Private Sub chkPrintedEmployees_Click()
  p11d32.ReportPrint.PrintedEmployees = ChkBoxToBool(chkPrintedEmployees)
End Sub

Private Sub chkHMITSections_Click(Index As Integer)
  Dim i As Long
  
  On Error GoTo chkHMITSections_Click_ERR
  
  Call xSet("chkHMITSections_Click")
  
  Select Case Index
    Case -1
      'read them in
      For i = 1 To HMIT_SECTIONS.[_HMIT_COUNT] - 1
        If (p11d32.ReportPrint.HMITSections And (2 ^ i)) Or p11d32.ReportPrint.HMITSelectionChoice = HMIT_SC_ALL Then
          chkHMITSections(i) = vbChecked
        End If
      Next
    Case Else
      If chkHMITSections(Index) = vbChecked Then
        p11d32.ReportPrint.HMITSections = p11d32.ReportPrint.HMITSections Or 2 ^ Index
      Else
        p11d32.ReportPrint.HMITSections = p11d32.ReportPrint.HMITSections Xor 2 ^ Index
      End If
    End Select
    
chkHMITSections_Click_END:
  Call xReturn("chkHMITSections_Click")
  Exit Sub
chkHMITSections_Click_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "chkHMITSections_Click", "chk HMIT Sections Click", "Error clicking an HMIT sections checkbox index = " & Index & ".")
  Resume chkHMITSections_Click_END
  Resume
End Sub


Private Sub chkRememberEmployeeSelection_Click()
  p11d32.ReportPrint.RemeberEmployeeSelection = ChkBoxToBool(chkRememberEmployeeSelection)
End Sub

Private Sub chkSeparatePrintJobs_Click()
  p11d32.ReportPrint.SeparatePrintJobs = ChkBoxToBool(chkSeparatePrintJobs)
End Sub
Private Sub chkTimeStamp_Click()
  p11d32.ReportPrint.TimeStamp = chkTimeStamp.value
End Sub

'Private Sub chkUseNotesIfPossible_Click() 'RK Email 19/03/03
'  On Error GoTo chkUseNotesIfPossible_ERR
'
'  p11d32.ReportPrint.UseNotesForEmail = ChkBoxToBool(chkUseNotesIfPossible)
'
'  If p11d32.ReportPrint.UseNotesForEmail Then p11d32.ReportPrint.Mail.UseNotesAPI = True
'
'
'chkUseNotesIfPossible_END:
'  Exit Sub
'chkUseNotesIfPossible_ERR:
'  Call ErrorMessage(ERR_ERROR, Err, "UseNotesIfPossible", "UseNotesIfPossible", Err.Description)
'End Sub


Private Sub cmdEditEmployeeLetter_Click()
'  F_EmployeeLetter.Show vbModal
  Call p11d32.Help.ShowForm(F_EmployeeLetter, vbModal)
  ResetAllLetterNodes
End Sub

Private Sub cmdEmployeeLetterFont_Click()
  On Error GoTo cmdEmployeeLetterFont_ERR
  Call xSet("cmdEmployeeLetterFont")
    
  
  MDIMain.cdlg.Flags = MDIMain.cdlg.Flags Or cdlCFBoth Or cdlCFScalableOnly Or cdlCFWYSIWYG
  MDIMain.cdlg.CancelError = True
  MDIMain.cdlg.FontName = p11d32.ReportPrint.EmployeeLetterFontName
  MDIMain.cdlg.FontSize = p11d32.ReportPrint.EmployeeLetterFontSize
  Call MDIMain.cdlg.ShowFont
  If Not (cdlCFNoFaceSel And MDIMain.cdlg.Flags) Then
    p11d32.ReportPrint.EmployeeLetterFontName = MDIMain.cdlg.FontName
    p11d32.ReportPrint.EmployeeLetterFontSize = MDIMain.cdlg.FontSize
    Call EmployeeLetterFontToLabel
  End If
  
cmdEmployeeLetterFont_END:
  Call xReturn("cmdEmployeeLetterFont")
  Exit Sub
cmdEmployeeLetterFont_ERR:
  If Err.Number <> cdlCancel Then Call ErrorMessage(ERR_ERROR, Err, "cmdEmployeeLetterFont", "cmd Employee Letter Font", "Error changing the employee letter font.")
  Resume cmdEmployeeLetterFont_END
End Sub
Private Sub EmployeeLetterFontToLabel()
  On Error GoTo EmployeeLetterFontToLabel_ERR
   
  Call xSet("EmployeeLetterFontToLabel")
   
  lblEmployeeLetterFont.Font.Name = p11d32.ReportPrint.EmployeeLetterFontName
  lblEmployeeLetterFont.Font.SIZE = p11d32.ReportPrint.EmployeeLetterFontSize
  lblEmployeeLetterFont = p11d32.ReportPrint.EmployeeLetterFontName & "," & p11d32.ReportPrint.EmployeeLetterFontSize
  lblEmployeeLetterFont.ToolTipText = lblEmployeeLetterFont.Caption
  cmdEmployeeLetterFont.ToolTipText = lblEmployeeLetterFont.Caption
EmployeeLetterFontToLabel_END:
  Call xReturn("EmployeeLetterFontToLabel")
  Exit Sub
EmployeeLetterFontToLabel_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeLetterFontToLabel", "Employee Letter Font To Label", "Error setting the employee letter font to label.")
  Resume EmployeeLetterFontToLabel_END
  Resume
End Sub

Private Sub cmdHTMLOutputTo_Click()
  
End Sub




Private Sub cmdOK_Click()
  Dim Cancel As Boolean
  
  On Error GoTo cmdOK_ERR
  
  'If chkUseNotesIfPossible = vbChecked Then 'RK Email 19/03/03
  '  Cancel = ValidateFileFromTextBox(txtNotesIniFile, False, "Notes.ini")
  '  If Not Cancel Then
  '    p11d32.ReportPrint.NotesIniFile = txtNotesIniFile.Text
  '    Call p11d32.ReportPrint.UseNotes
  '    Unload Me
  '  End If
  'Else
    Unload Me
  'End If
  
cmdOK_END:
  Exit Sub
cmdOK_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "OK_Click", "OK_Click", Err.Description)
  Resume cmdOK_END
End Sub


Private Sub Form_Load()
  Set tvwLetters.ImageList = MDIMain.imlTree
  Set m_LastSelectedNode = Nothing
  
  Call SettingsToScreen
  
  Select Case F_Print.tvwReports.SelectedItem.Text  'RC - disable Print to HTML if P11D and working paper
    Case p11d32.ReportPrint.Name(RPT_HMIT)
      optDestination(1).Enabled = True
    Case p11d32.ReportPrint.Name(RPT_HMIT_PLUS_WORKING_PAPERS)
      optDestination(1).Enabled = True
    Case p11d32.ReportPrint.Name(RPT_HMIT_PLUS_WORKING_PAPERS)
      optDestination(1).Enabled = True
    Case p11d32.ReportPrint.Name(RPT_P46CAR)
      optDestination(1).Enabled = True
    Case Else
      If optDestination(1).value = True Then
        optDestination(0).value = True
      End If
      optDestination(1).Enabled = False
  End Select

End Sub

Public Function SettingsToScreen() As Boolean
  Dim i As Long
  Dim j As Long
  Dim iUserLetterCount As Long
  Dim sLetterFiles_System() As String
  Dim sLetterFiles_User() As String
  Dim sLetterFile_System As String
  Dim sLetterFile_User As String
  Dim sExt As String
  Dim bSelected As Boolean
  Dim n As Node
  
  On Error GoTo SettingsToScreen_Err
  Call xSet("SettingsToScreen")
  
  Me.tab.tab = 0
  
  tvwLetters.Nodes.Clear
  ' Add Top-level nodes for system and user letters
  Set n = tvwLetters.Nodes.Add(, , "SYS", "System letters")
  n.Tag = FIT_SYSTEM_DEFINED
  n.Image = IMG_FOLDER_CLOSED
  
  fbUserLetters.Directory = p11d32.ReportPrint.UserLettersPath
  'RK/EK Could split below into separate function
  
  j = p11d32.ReportPrint.GetLetterFiles(sLetterFiles_System)
  For i = 1 To j
    Call SplitPath(sLetterFiles_System(i), , sLetterFile_System, sExt)
    sLetterFile_System = sLetterFile_System & sExt
    Set n = tvwLetters.Nodes.Add("SYS", tvwChild, , sLetterFile_System)
    n.Tag = i + FIT_USER_DEFINED
  Next
  m_NodeCount = i
  
  Call UserLettersToScreen
'
'  iUserLetterCount = p11d32.GetUserLetterFiles(sLetterFiles_User)
'
'  If iUserLetterCount > 0 Then
'    For j = i + 1 To i + iUserLetterCount
'      Call SplitPath(sLetterFiles_User(j - i), , sLetterFile_User, sExt)
'      sLetterFile_User = sLetterFile_User & sExt
'      Set n = tvwLetters.Nodes.Add("USER", tvwChild, , sLetterFile_User)
'      n.Tag = \j + FIT_USER_DEFINED
'    Next
'  End If
'  m_NodeCount = i + iUserLetterCount

  Call DefaultLetterSelection

'  Call ResetAllLetterNodes
'
'  'Default if no Employee letter has been selected
'  If m_LastSelectedNode Is Nothing And tvwLetters.Nodes.Count > 0 Then
'    Set tvwLetters.SelectedItem = tvwLetters.Nodes(1)
'    Call SetAsSelectedNode(tvwLetters.Nodes(1))
'  End If
'
'  If p11d32.AppYear > 2000 Then
  chkDatesOnWorkingPaper.Visible = True
  
  Call EmployeeLetterFontToLabel
  'chkCheckDataBeforePrint = BoolToChkBox(p11d32.ReportPrint.RunChecks)
  chkCheckDataBeforePrint.value = (BoolToChkBox(Not (p11d32.ReportPrint.CheckOptions > NO_THIS_TIME_ONLY)))
  chkChecksAutoRefresh.value = BoolToChkBox(p11d32.ReportPrint.ChecksAutoRefresh)
  chkTimeStamp.value = BoolToChkBox(p11d32.ReportPrint.TimeStamp) 'JN code
  chkPrintedEmployees.value = BoolToChkBox(p11d32.ReportPrint.PrintedEmployees)
  optHMITSectionChoice_Click (-1)
  optReportTotalValue_Click (-1)
  chkHMITSections_Click (-1)
  vtEmployeeLetterMargin.Text = p11d32.ReportPrint.EmployeeLetterMargin
  vsEmployeeLetterMargin.value = p11d32.ReportPrint.EmployeeLetterMargin
  optDestination_Click (-1)
  optChkBeforePrint_Click (-1)
  optExportType_Click (-1)
  chkEmailTextAsHTML = BoolToChkBox(p11d32.ReportPrint.EmailTextAsHTML)
  chkAllowUserReportNameHTML.value = BoolToChkBox(p11d32.ReportPrint.ExportAllowUserNameHTML)
  txtEmailSubject = p11d32.ReportPrint.EmailSubject
  chkSeparatePrintJobs.value = BoolToChkBox(p11d32.ReportPrint.SeparatePrintJobs)
  'chkUseNotesIfPossible.value = BoolToChkBox(p11d32.ReportPrint.UseNotes(False)) 'RK Email 19/03/03
  chkDatesOnWorkingPaper.value = BoolToChkBox(p11d32.ReportPrint.DatesOnWorkingPaper)
  chkHMITFieldTrim.value = BoolToChkBox(p11d32.ReportPrint.HMITFieldTrim)
  
  
  
  chkRememberEmployeeSelection.value = BoolToChkBox(p11d32.ReportPrint.RemeberEmployeeSelection)
  
SettingsToScreen_End:
  Call xReturn("SettingsToScreen")
  Exit Function

SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error placing the print options to the screen.")
  Resume SettingsToScreen_End
  Resume
End Function
Private Sub UserLettersToScreen()
  Dim j As Long
  Dim iCount As Long
  Dim n As Node
  Dim sLetterFiles_User() As String
  Dim sLetterFile_User As String
  Dim sExt As String
  
On Error Resume Next
  tvwLetters.Nodes.Remove ("USER")
    
 On Error GoTo err_ERR
   
  m_NodeCount = m_NodeCount - m_LastUserLetterCount
   
   
  Set n = tvwLetters.Nodes.Add(, , "USER", "User letters")
  n.Tag = FIT_USER_DEFINED
  n.Image = IMG_FOLDER_CLOSED
  
  m_LastUserLetterCount = 0
  m_LastUserLetterCount = p11d32.ReportPrint.GetUserLetterFiles(sLetterFiles_User)
  If m_LastUserLetterCount > 0 Then
    For j = 1 To m_LastUserLetterCount
      Call SplitPath(sLetterFiles_User(j), , sLetterFile_User, sExt)
      sLetterFile_User = sLetterFile_User & sExt
      Set n = tvwLetters.Nodes.Add("USER", tvwChild, , sLetterFile_User)
      n.Tag = j + m_NodeCount + FIT_USER_DEFINED
    Next
  End If
  m_NodeCount = m_NodeCount + m_LastUserLetterCount
  
  
err_END:
  Exit Sub
err_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "UserLettersToScreen"), Err.Description)
End Sub
Private Sub DefaultLetterSelection()
  Call ResetAllLetterNodes
  
  'Default if no Employee letter has been selected
  If m_LastSelectedNode Is Nothing And tvwLetters.Nodes.Count > 0 Then
    Set tvwLetters.SelectedItem = tvwLetters.Nodes(1)
    Call SetAsSelectedNode(tvwLetters.Nodes(1))
  End If
  
End Sub

Private Sub optChkBeforePrint_Click(Index As Integer)
    Select Case Index
    Case -1
        Select Case p11d32.ReportPrint.CheckOptions
          Case YES_ALWAYS
            optChkBeforePrint(0) = True
          Case NEVER
            optChkBeforePrint(1) = True
          Case Else
            optChkBeforePrint(0) = False
            optChkBeforePrint(1) = False
        End Select
    Case Else
      Select Case Index
        Case 0
          p11d32.ReportPrint.CheckOptions = YES_ALWAYS
        Case 1
          p11d32.ReportPrint.CheckOptions = NEVER
      End Select
    End Select


End Sub

Private Sub optDestination_Click(Index As Integer)
  Dim i As Long
  Dim b As Boolean
  
  Select Case Index
    Case -1
      For i = REPD_PRINTER_FIRST_ITEM To REPD_PRINTER_LAST_ITEM
        If i = p11d32.ReportPrint.Destination Then
          optDestination(i) = True
          Exit For
        End If
      Next
      
      fbExportDirectory.Directory = p11d32.ReportPrint.ExportDirectory
      Me.txtPRNFileName.Text = p11d32.ReportPrint.PRNFileName & p11d32.ReportPrint.PRNFileExtension
    Case Else
      p11d32.ReportPrint.Destination = Index
      
      b = (Index = REPD_FILE_HTML) Or (Index = REPD_FILE_PRN)
      lblExportTo.Enabled = b
      
      fbExportDirectory.Enabled = b
      chkAllowUserReportNameHTML.Enabled = b
      
      b = (Index = REPD_FILE_PRN)
      txtPRNFileName.Enabled = b
  End Select
End Sub

Private Sub optExportType_Click(Index As Integer)
  
  Select Case Index
    Case -1
        Select Case p11d32.ReportPrint.ExportOption
          Case EXPORT_HTML_IE
            optExportType(0) = True
          Case EXPORT_HTML_NETSCAPE
            optExportType(1) = True
          Case EXPORT_HTML_INTEXP5  'km
            optExportType(2) = True
          Case Else
            Call ECASE("Invalid ExportOption, = " & p11d32.ReportPrint.ExportOption)
        End Select
    Case Else
      Select Case Index
        Case 0
          p11d32.ReportPrint.ExportOption = EXPORT_HTML_IE
        Case 1
          p11d32.ReportPrint.ExportOption = EXPORT_HTML_NETSCAPE
        Case 2  'km
          p11d32.ReportPrint.ExportOption = EXPORT_HTML_INTEXP5
      End Select
  End Select
End Sub

Private Sub optHMITSectionChoice_Click(Index As Integer)
  Dim i As Long
  
  Select Case Index
    Case -1
      For i = 0 To 2
        If i = p11d32.ReportPrint.HMITSelectionChoice Then
          optHMITSectionChoice(i) = True
          Exit For
        End If
      Next
    Case Else
      p11d32.ReportPrint.HMITSelectionChoice = Index
      If Index = HMIT_SC_SELECTED Then
        fmeHMITSections.Enabled = True
      Else
        fmeHMITSections.Enabled = False
      End If
  End Select
  
End Sub

Private Sub Option1_Click(Index As Integer)


End Sub

Private Sub optReportTotalValue_Click(Index As Integer)
  Dim i As Long
  
  Select Case Index
    Case -1
      For i = 0 To 2
        If i = p11d32.ReportPrint.P11DTotalValue Then
          optReportTotalValue(i) = True
          Exit For
        End If
      Next
    Case Else
      p11d32.ReportPrint.P11DTotalValue = Index
  End Select

End Sub

Private Sub tvwLetters_Collapse(ByVal Node As MSComctlLib.Node)
  If GetNodeType(Node) = NODE_FOLDER Then
    Node.Image = IMG_FOLDER_CLOSED
  End If
End Sub

Private Sub tvwLetters_NodeClick(ByVal Node As MSComctlLib.Node)
  Call SetAsSelectedNode(Node)
  Call ResetAllLetterNodes
  If GetNodeType(Node) = NODE_FOLDER Then
    Node.Expanded = Not (Node.Expanded)
    Node.Image = IIf(Node.Expanded, IMG_FOLDER_OPEN, IMG_FOLDER_CLOSED)
  End If
End Sub

Private Sub txtEmailSubject_Validate(Cancel As Boolean)
  p11d32.ReportPrint.EmailSubject = txtEmailSubject
End Sub
Private Sub txtPRNFileName_Validate(Cancel As Boolean)
  Dim sExt As String, sFileName As String, sPath As String
  Dim sFIle As String

  txtPRNFileName.Text = Trim$(txtPRNFileName.Text)
  sFileName = txtPRNFileName.Text
  If Len(sFileName) = 0 Then
    Cancel = True
    Exit Sub
  Else
    Call SplitPath(sFileName, sPath, sFIle, sExt)
    If Len(sPath) > 0 Then
      Cancel = True
      Exit Sub
    Else
      p11d32.ReportPrint.PRNFileName = sFIle
      p11d32.ReportPrint.PRNFileExtension = sExt
    End If
  End If
End Sub
Private Sub vsEmployeeLetterMargin_Change()
  vtEmployeeLetterMargin.Text = vsEmployeeLetterMargin.value
  p11d32.ReportPrint.EmployeeLetterMargin = vtEmployeeLetterMargin.Text
End Sub
Private Sub vtEmployeeLetterMargin_Validate(Cancel As Boolean)
  If vtEmployeeLetterMargin.FieldInvalid Then
    Cancel = True
  Else
    vsEmployeeLetterMargin.value = vtEmployeeLetterMargin.Text
    p11d32.ReportPrint.EmployeeLetterMargin = vtEmployeeLetterMargin.Text
  End If
End Sub

Private Function GetNodeType(MyNode As Node) As NODE_TYPES

  On Error GoTo GetNodeType_Err
  Call xSet("GetNodeType")
  If MyNode.Parent Is Nothing Then
    GetNodeType = NODE_FOLDER
  Else
    GetNodeType = NODE_LETTER
  End If
    
GetNodeType_End:
  Call xReturn("GetNodeType")
  Exit Function

GetNodeType_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetNodeType", "Error in GetNodeType", "Undefined error.")
  Resume GetNodeType_End
End Function

Private Function GetNodeFileType(MyNode As Node) As NODE_TYPES

  On Error GoTo GetNodeFileType_Err
  Call xSet("GetNodeFileType")
  If MyNode.Parent.Tag Then
    GetNodeFileType = FIT_USER_DEFINED
  Else
    GetNodeFileType = FIT_SYSTEM_DEFINED
  End If
    
GetNodeFileType_End:
  Call xReturn("GetNodeFileType")
  Exit Function

GetNodeFileType_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetNodeFileType", "Error in GetNodeFileType", "Undefined error.")
  Resume GetNodeFileType_End
End Function

Public Function AddNewLetterNode(ByVal sNewFileName As String, bSetAsSelected As Boolean) As Boolean
  'Called From F_Employeeletter.SaveAS
  Dim MyNode As Node
  On Error GoTo AddNewLetterNode_Err
  Call xSet("AddNewLetterNode")
  
  Call SplitPath(sNewFileName, , sNewFileName)
  sNewFileName = sNewFileName & S_EMPLOYEE_LETTER_FILE_EXTENSION
  'Add node if doesn't already exists
  Set MyNode = MatchLetterNode(sNewFileName)
  If MyNode Is Nothing Then
    m_NodeCount = m_NodeCount + 1
    Set MyNode = tvwLetters.Nodes.Add("USER", tvwChild, , sNewFileName)
    m_LastUserLetterCount = m_LastUserLetterCount + 1
    MyNode.Tag = m_NodeCount + FIT_USER_DEFINED
  End If
  If bSetAsSelected Then Call SetAsSelectedNode(MyNode)
  
  AddNewLetterNode = True
AddNewLetterNode_End:
  Call xReturn("AddNewLetterNode")
  Exit Function

AddNewLetterNode_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddNewLetterNode", "Error in AddNewLetterNode", "Undefined error.")
  AddNewLetterNode = False
  Resume AddNewLetterNode_End
  Resume
End Function
Public Function ResetAllLetterNodes() As Boolean
  'RK Reset treeview
  'MyNode.text used as filename may contain spaces
  Dim MyNode As Node
  Dim sMatchFile As String
  On Error GoTo ResetAllLetterNodes_Err
  Call xSet("ResetAllLetterNodes")
  sMatchFile = p11d32.ReportPrint.EmployeeLetterFile
  For Each MyNode In tvwLetters.Nodes
    If GetNodeType(MyNode) Then
      If StrComp(sMatchFile, MyNode.Text, vbTextCompare) = 0 Then
        Call SetAsSelectedNode(MyNode)
      Else
        MyNode.Selected = False
        MyNode.Image = IMG_LETTER_CLOSED
      End If
    End If
  Next MyNode
  ResetAllLetterNodes = True
ResetAllLetterNodes_End:
  Call xReturn("ResetAllLetterNodes")
  Exit Function

ResetAllLetterNodes_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ResetAllLetterNodes", "Error in ResetAllLetterNodes", "Undefined error.")
  ResetAllLetterNodes = False
  Resume ResetAllLetterNodes_End
End Function
Private Sub SetAsSelectedNode(ByVal Node As MSComctlLib.Node)
  Dim nt As NODE_TYPES
  nt = GetNodeType(Node)
  ' Modify letter nodes
  If nt = NODE_LETTER Then
    If GetNodeFileType(Node) = FIT_USER_DEFINED Then
      p11d32.ReportPrint.EmployeeLetterPath = p11d32.ReportPrint.UserLettersPath
    Else
      p11d32.ReportPrint.EmployeeLetterPath = FullPath(AppPath) & S_SYSTEMDIR_LETTERS
    End If
    p11d32.ReportPrint.EmployeeLetterFile = Node.Text
    Set m_LastSelectedNode = Node
    Set tvwLetters.SelectedItem = Node
    Node.Image = IMG_LETTER_OPEN
    Node.Selected = True
    Node.Parent.Expanded = True
    Node.Parent.Image = IMG_FOLDER_OPEN
  End If
End Sub

Public Function MatchLetterNode(sLetterFile As String) As Node
  Dim MyNode As Node
  Set MatchLetterNode = Nothing
  For Each MyNode In tvwLetters.Nodes
    If StrComp(sLetterFile, MyNode.Text, vbTextCompare) = 0 Then
        Set MatchLetterNode = MyNode
        Exit Function
    End If
  Next MyNode

MatchLetterNode_End:
  Call xReturn("MatchLetterNode")
  Exit Function

MatchLetterNode_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MatchLetterNode", "Error in MatchLetterNode", "Undefined error.")
  Set MatchLetterNode = Nothing
  Resume MatchLetterNode_End
End Function

Private Sub fbUserLetters_Ended()
  Dim s As String
    
    p11d32.ReportPrint.UserLettersPath = fbUserLetters.Directory
    Call UserLettersToScreen
    Call DefaultLetterSelection
End Sub

Private Sub fbUserLetters_Started()
  fbUserLetters.Directory = p11d32.ReportPrint.UserLettersPath
End Sub

Private Sub fbExportDirectory_Ended()
  p11d32.ReportPrint.ExportDirectory = fbExportDirectory.Directory
End Sub

Private Sub fbExportDirectory_Started()
  fbExportDirectory.Directory = p11d32.ReportPrint.ExportDirectory
End Sub


