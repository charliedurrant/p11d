VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "ATC2STAT.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form Frm_RepWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "abatec Report Wizard"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin atc2stat.TCSStatus SBar 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   59
      Top             =   7365
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Fra_Buttons 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   8115
      Begin VB.CommandButton Cmd_Export 
         Caption         =   "&Export"
         Height          =   360
         Left            =   2700
         TabIndex        =   12
         ToolTipText     =   "Export the report to a Word, Excel or plain text file"
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "&Print"
         Height          =   360
         Left            =   1920
         TabIndex        =   11
         ToolTipText     =   "Print out the report"
         Top             =   360
         Width           =   780
      End
      Begin VB.CheckBox chkPreviewLines 
         Caption         =   "All Lines"
         Height          =   270
         Left            =   90
         TabIndex        =   10
         Top             =   105
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.TextBox txtPreview 
         Alignment       =   1  'Right Justify
         Height          =   312
         Left            =   90
         TabIndex        =   8
         Top             =   375
         Width           =   675
      End
      Begin VB.CommandButton Cmd_LoadReport 
         Caption         =   "&Load"
         Height          =   360
         Left            =   4440
         TabIndex        =   6
         ToolTipText     =   "Load a previously save report specification from a file"
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Cmd_SaveReport 
         Caption         =   "&Save"
         Height          =   360
         Left            =   3660
         TabIndex        =   5
         ToolTipText     =   "Save the current report specification to a file"
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Cmd_Exit 
         Caption         =   "E&xit"
         Height          =   360
         Left            =   7080
         TabIndex        =   4
         ToolTipText     =   "Exit the Report Wizard"
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Cmd_Next 
         Caption         =   "&Next >"
         Height          =   360
         Left            =   6180
         TabIndex        =   2
         ToolTipText     =   "Move to the next screen"
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Cmd_Back 
         Caption         =   "< &Back"
         Height          =   360
         Left            =   5400
         TabIndex        =   3
         ToolTipText     =   "Move to the previous screen"
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Cmd_Preview 
         Caption         =   "Pre&view"
         Height          =   360
         Left            =   1140
         TabIndex        =   1
         ToolTipText     =   "Preview the report on screen"
         Top             =   360
         Width           =   780
      End
      Begin MSComCtl2.UpDown UpDownPreview 
         Height          =   315
         Left            =   795
         TabIndex        =   9
         Top             =   390
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   50
         BuddyControl    =   "txtPreview"
         BuddyDispid     =   196613
         OrigLeft        =   3000
         OrigTop         =   285
         OrigRight       =   3240
         OrigBottom      =   615
         Max             =   1000
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   7740
      Left            =   -10
      TabIndex        =   13
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   13653
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Selection"
      TabPicture(0)   =   "frm_rwiz.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_Fields"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Criteria/Formats"
      TabPicture(1)   =   "frm_rwiz.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Fra_Format"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Appearance"
      TabPicture(2)   =   "frm_rwiz.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Fra_Report"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "File Groups"
      TabPicture(3)   =   "frm_rwiz.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Fra_ARFileGroups"
      Tab(3).ControlCount=   1
      Begin VB.Frame Fra_ARFileGroups 
         BorderStyle     =   0  'None
         Caption         =   "Reporter (to be hidden)"
         Height          =   5592
         Left            =   -75000
         TabIndex        =   123
         Top             =   360
         Visible         =   0   'False
         Width           =   8160
         Begin VB.Frame fraFileGroups 
            Caption         =   "File groups"
            Height          =   2535
            Left            =   120
            TabIndex        =   124
            ToolTipText     =   "To edit a File Group select an item and then click it once"
            Top             =   480
            Width           =   7935
            Begin VB.OptionButton optFileGroupType 
               Caption         =   "Global"
               Height          =   255
               Index           =   2
               Left            =   6480
               TabIndex        =   145
               Top             =   2160
               Width           =   850
            End
            Begin VB.OptionButton optFileGroupType 
               Caption         =   "Local"
               Height          =   255
               Index           =   1
               Left            =   5520
               TabIndex        =   144
               Top             =   2160
               Width           =   850
            End
            Begin VB.OptionButton optFileGroupType 
               Caption         =   "System"
               Height          =   255
               Index           =   0
               Left            =   4560
               TabIndex        =   143
               Top             =   2160
               Width           =   850
            End
            Begin VB.TextBox txtFileGroup 
               Height          =   312
               Left            =   720
               TabIndex        =   129
               Top             =   2160
               Width           =   3615
            End
            Begin VB.CommandButton cmdRemoveFileGroup 
               Enabled         =   0   'False
               Height          =   300
               Left            =   7440
               Picture         =   "frm_rwiz.frx":0070
               Style           =   1  'Graphical
               TabIndex        =   128
               Top             =   600
               UseMaskColor    =   -1  'True
               Width           =   300
            End
            Begin VB.CommandButton cmdAddFileGroup 
               Height          =   300
               Left            =   7440
               Picture         =   "frm_rwiz.frx":03B2
               Style           =   1  'Graphical
               TabIndex        =   127
               Top             =   240
               UseMaskColor    =   -1  'True
               Width           =   300
            End
            Begin MSComctlLib.ListView lvwFileGroups 
               Height          =   1815
               Left            =   120
               TabIndex        =   125
               Top             =   240
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   3201
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Description"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label lblFileGroup 
               Caption         =   "Name:"
               Height          =   300
               Left            =   120
               TabIndex        =   130
               Top             =   2160
               Width           =   1215
            End
         End
         Begin TabDlg.SSTab ssTabReporterVersions 
            Height          =   3240
            Left            =   0
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   2880
            Width           =   8205
            _ExtentX        =   14473
            _ExtentY        =   5715
            _Version        =   393216
            TabOrientation  =   1
            Style           =   1
            TabHeight       =   520
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Standalone"
            TabPicture(0)   =   "frm_rwiz.frx":06F4
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fraFileGroupMember"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fraFileGroupMembers"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Administrator"
            TabPicture(1)   =   "frm_rwiz.frx":0710
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "Organiser"
            TabPicture(2)   =   "frm_rwiz.frx":072C
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            Begin VB.Frame fraFileGroupMembers 
               Caption         =   "File group members"
               Height          =   2475
               Left            =   120
               TabIndex        =   139
               Top             =   240
               Width           =   3255
               Begin VB.CommandButton cmdAddFileGroupMember 
                  Height          =   300
                  Left            =   2850
                  Picture         =   "frm_rwiz.frx":0748
                  Style           =   1  'Graphical
                  TabIndex        =   141
                  Top             =   240
                  UseMaskColor    =   -1  'True
                  Width           =   300
               End
               Begin VB.CommandButton cmdRemoveFileGroupMember 
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   2850
                  Picture         =   "frm_rwiz.frx":0A8A
                  Style           =   1  'Graphical
                  TabIndex        =   140
                  Top             =   600
                  UseMaskColor    =   -1  'True
                  Width           =   300
               End
               Begin MSComctlLib.ListView lvwFileGroupMembers 
                  Height          =   1995
                  Left            =   120
                  TabIndex        =   142
                  Top             =   240
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   3519
                  View            =   3
                  LabelEdit       =   1
                  Sorted          =   -1  'True
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  HideColumnHeaders=   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   1
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "Description"
                     Object.Width           =   2540
                  EndProperty
               End
            End
            Begin VB.Frame fraFileGroupMember 
               Height          =   2475
               Left            =   3480
               TabIndex        =   132
               Top             =   240
               Width           =   4580
               Begin VB.CheckBox chkDefinitionRecursive 
                  Caption         =   "Include sub-directories?"
                  Height          =   255
                  Left            =   2400
                  TabIndex        =   148
                  Top             =   180
                  Width           =   2000
               End
               Begin VB.Frame fraPackSelection 
                  BorderStyle     =   0  'None
                  Height          =   375
                  Left            =   120
                  TabIndex        =   135
                  Top             =   2040
                  Width           =   4335
                  Begin VB.CheckBox chkAllPacks 
                     Caption         =   "Select all"
                     Height          =   195
                     Left            =   3410
                     TabIndex        =   147
                     Top             =   150
                     Width           =   1515
                  End
                  Begin VB.OptionButton optPackArrange 
                     Caption         =   "Product"
                     Height          =   300
                     Index           =   0
                     Left            =   1800
                     TabIndex        =   137
                     Top             =   120
                     Width           =   855
                  End
                  Begin VB.OptionButton optPackArrange 
                     Caption         =   "FY"
                     Height          =   300
                     Index           =   1
                     Left            =   2760
                     TabIndex        =   136
                     Top             =   120
                     Value           =   -1  'True
                     Width           =   615
                  End
                  Begin VB.Label lblPackArrange 
                     Caption         =   "View pack selection by:"
                     Height          =   255
                     Left            =   0
                     TabIndex        =   146
                     Top             =   120
                     Width           =   1815
                  End
               End
               Begin VB.OptionButton optDefinition 
                  Caption         =   "Directory"
                  Height          =   375
                  Index           =   0
                  Left            =   1080
                  TabIndex        =   134
                  Top             =   120
                  Width           =   1095
               End
               Begin VB.OptionButton optDefinition 
                  Caption         =   "File"
                  Height          =   375
                  Index           =   1
                  Left            =   200
                  TabIndex        =   133
                  Top             =   120
                  Width           =   855
               End
               Begin MSComctlLib.TreeView tvwPacks 
                  Height          =   1245
                  Left            =   120
                  TabIndex        =   138
                  Top             =   840
                  Width           =   4305
                  _ExtentX        =   7594
                  _ExtentY        =   2196
                  _Version        =   393217
                  LabelEdit       =   1
                  Style           =   7
                  ImageList       =   "ImL"
                  Appearance      =   1
               End
               Begin VB.Label lblPlaceHolder 
                  Caption         =   "FolderBrowser place holder"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   151
                  Top             =   480
                  Width           =   4245
               End
            End
         End
         Begin VB.Label lblFileGroups 
            Caption         =   "Use this screen to specify the file groups to be reported on."
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   240
            Width           =   7455
         End
      End
      Begin VB.Frame Fra_Report 
         BorderStyle     =   0  'None
         Height          =   5592
         Left            =   -74940
         TabIndex        =   81
         Top             =   300
         Width           =   7995
         Begin VB.Frame Fra_Orient 
            Caption         =   "Page Orientation"
            Height          =   972
            Left            =   120
            TabIndex        =   119
            Top             =   480
            Width           =   1512
            Begin VB.OptionButton Opt_Orient 
               Caption         =   "Portrait"
               Height          =   192
               Index           =   0
               Left            =   180
               TabIndex        =   121
               Top             =   360
               Width           =   1092
            End
            Begin VB.OptionButton Opt_Orient 
               Caption         =   "Landscape"
               Height          =   192
               Index           =   1
               Left            =   180
               TabIndex        =   120
               Top             =   660
               Width           =   1092
            End
         End
         Begin VB.Frame Fra_HF 
            Caption         =   "Headers and Footers"
            Height          =   2415
            Left            =   120
            TabIndex        =   102
            Top             =   1620
            Width           =   7815
            Begin VB.CheckBox Chk_PageHead 
               Caption         =   "Page Header"
               Height          =   315
               Left            =   180
               TabIndex        =   114
               Top             =   1380
               Width           =   1395
            End
            Begin VB.CheckBox Chk_PageFoot 
               Caption         =   "Page Footer"
               Height          =   375
               Left            =   180
               TabIndex        =   113
               Top             =   1920
               Width           =   1395
            End
            Begin VB.CheckBox Chk_RepHead 
               Caption         =   "Report Header"
               Height          =   435
               Left            =   180
               TabIndex        =   112
               Top             =   780
               Width           =   1395
            End
            Begin VB.TextBox Txt_PageFoot 
               Height          =   312
               Index           =   2
               Left            =   5820
               TabIndex        =   111
               Top             =   1920
               Width           =   1692
            End
            Begin VB.TextBox Txt_PageFoot 
               Height          =   312
               Index           =   1
               Left            =   3660
               TabIndex        =   110
               Top             =   1920
               Width           =   1692
            End
            Begin VB.TextBox Txt_PageHead 
               Height          =   312
               Index           =   2
               Left            =   5820
               TabIndex        =   109
               Top             =   1380
               Width           =   1692
            End
            Begin VB.TextBox Txt_PageHead 
               Height          =   312
               Index           =   1
               Left            =   3660
               TabIndex        =   108
               Top             =   1380
               Width           =   1692
            End
            Begin VB.TextBox Txt_RepHead 
               Height          =   312
               Index           =   2
               Left            =   5820
               TabIndex        =   107
               Top             =   840
               Width           =   1692
            End
            Begin VB.TextBox Txt_RepHead 
               Height          =   312
               Index           =   1
               Left            =   3660
               TabIndex        =   106
               Top             =   840
               Width           =   1692
            End
            Begin VB.TextBox Txt_PageFoot 
               Height          =   312
               Index           =   0
               Left            =   1620
               TabIndex        =   105
               Top             =   1920
               Width           =   1692
            End
            Begin VB.TextBox Txt_PageHead 
               Height          =   312
               Index           =   0
               Left            =   1620
               TabIndex        =   104
               Top             =   1380
               Width           =   1692
            End
            Begin VB.TextBox Txt_RepHead 
               Height          =   312
               Index           =   0
               Left            =   1620
               TabIndex        =   103
               Top             =   840
               Width           =   1692
            End
            Begin VB.Label Lbl_Right 
               Caption         =   "Right"
               Height          =   255
               Left            =   6420
               TabIndex        =   118
               Top             =   600
               Width           =   915
            End
            Begin VB.Label Lbl_Centre 
               Caption         =   "Centre"
               Height          =   315
               Left            =   4260
               TabIndex        =   117
               Top             =   600
               Width           =   795
            End
            Begin VB.Label Lbl_Left 
               Caption         =   "Left"
               Height          =   255
               Left            =   2280
               TabIndex        =   116
               Top             =   600
               Width           =   795
            End
            Begin VB.Label Lbl_HFInst 
               Caption         =   "Right-click twice in a text box to change the font for that text, or to add a control code."
               Height          =   255
               Left            =   180
               TabIndex        =   115
               Top             =   300
               Width           =   6975
            End
         End
         Begin VB.Frame Fra_Widthing 
            Caption         =   "Auto Widthing"
            Height          =   975
            Left            =   1860
            TabIndex        =   94
            Top             =   480
            Width           =   6075
            Begin VB.CheckBox Chk_InclColHeaders 
               Caption         =   "Include headings in auto-widthing"
               Height          =   252
               Left            =   420
               TabIndex        =   100
               Top             =   660
               Value           =   1  'Checked
               Width           =   2895
            End
            Begin VB.TextBox Txt_PrevLines 
               Height          =   312
               Left            =   2220
               TabIndex        =   99
               Top             =   285
               Width           =   690
            End
            Begin VB.CheckBox Chk_AutoWidth 
               Caption         =   "Number of lines to base auto-widthing on:"
               Height          =   375
               Left            =   180
               TabIndex        =   98
               Top             =   240
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.CheckBox Chk_FitToPage 
               Caption         =   "Fit fields to page"
               Height          =   252
               Left            =   3660
               TabIndex        =   97
               Top             =   300
               Width           =   1755
            End
            Begin VB.CheckBox Chk_TrimHeadings 
               Caption         =   "Trim headings"
               Height          =   252
               Left            =   3660
               TabIndex        =   96
               Top             =   600
               Width           =   1392
            End
            Begin MSComCtl2.UpDown UpDownPL 
               Height          =   312
               Left            =   2671
               TabIndex        =   95
               Top             =   285
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "Chk_AutoWidth"
               BuddyDispid     =   196656
               OrigLeft        =   3000
               OrigTop         =   285
               OrigRight       =   3240
               OrigBottom      =   615
               Max             =   1000
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Lbl_AutoWidth 
               Caption         =   "Number of lines to base auto-widthing on:"
               Height          =   435
               Left            =   420
               TabIndex        =   101
               Top             =   240
               Width           =   1875
            End
         End
         Begin VB.Frame Fra_Misc 
            Caption         =   "Miscellaneous Options"
            Height          =   1410
            Left            =   120
            TabIndex        =   82
            Top             =   4125
            Width           =   7815
            Begin VB.ComboBox Cbo_GHDelimiter 
               Height          =   315
               Left            =   6720
               TabIndex        =   91
               Top             =   840
               Width           =   915
            End
            Begin VB.ComboBox Cbo_GHSeparator 
               Height          =   315
               Left            =   6720
               TabIndex        =   90
               Top             =   300
               Width           =   915
            End
            Begin VB.CheckBox Chk_GroupHeaders 
               Caption         =   "Hide group header type"
               Height          =   252
               Left            =   120
               TabIndex        =   89
               Top             =   780
               Width           =   2055
            End
            Begin VB.CheckBox Chk_DisplayRecordCount 
               Caption         =   "Display record count"
               Height          =   252
               Left            =   120
               TabIndex        =   88
               Top             =   540
               Width           =   1995
            End
            Begin VB.CheckBox Chk_RHonAllPages 
               Caption         =   "Display report header on every page"
               Height          =   252
               Left            =   120
               TabIndex        =   87
               Top             =   300
               Width           =   3015
            End
            Begin VB.CheckBox Chk_CollapseAll 
               Caption         =   "Summary report"
               Height          =   252
               Left            =   120
               TabIndex        =   86
               Top             =   1020
               Width           =   1515
            End
            Begin VB.CheckBox Chk_AlignHeadings 
               Caption         =   "Align field headers as per data"
               Height          =   375
               Left            =   3600
               TabIndex        =   85
               Top             =   180
               Width           =   1695
            End
            Begin VB.CommandButton Cmd_DataFontSet 
               Caption         =   "Global Field Font"
               Height          =   360
               Left            =   3600
               TabIndex        =   84
               Top             =   900
               Width           =   1572
            End
            Begin VB.CheckBox Chk_IncBlankLines 
               Caption         =   "Include blank lines"
               Height          =   252
               Left            =   3600
               TabIndex        =   83
               Top             =   600
               Width           =   1632
            End
            Begin VB.Label Lbl_GHDelimiter 
               Caption         =   "Group header delimiter"
               Height          =   375
               Left            =   5640
               TabIndex        =   93
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label Lbl_GHSep 
               Caption         =   "Group header separator"
               Height          =   375
               Left            =   5640
               TabIndex        =   92
               Top             =   300
               Width           =   1095
            End
         End
         Begin VB.Label Lbl_ReportInst 
            Caption         =   "Use this screen to specify the report-wide design of your report."
            Height          =   252
            Left            =   240
            TabIndex        =   122
            Top             =   120
            Width           =   4512
         End
      End
      Begin VB.Frame Fra_Format 
         BorderStyle     =   0  'None
         Height          =   5472
         Left            =   120
         TabIndex        =   23
         Top             =   420
         Width           =   7935
         Begin VB.CheckBox Chk_Display 
            Caption         =   "Display field"
            Height          =   315
            Left            =   360
            TabIndex        =   152
            Top             =   2575
            Width           =   1215
         End
         Begin VB.Frame Fra_FldCriteria 
            Enabled         =   0   'False
            Height          =   2412
            Left            =   240
            TabIndex        =   24
            Top             =   3000
            Visible         =   0   'False
            Width           =   7092
            Begin VB.CommandButton Cmd_FldCritDel 
               Caption         =   "Delete criterion"
               Height          =   372
               Left            =   240
               TabIndex        =   29
               Top             =   1440
               Width           =   1212
            End
            Begin VB.CommandButton Cmd_FldCritAdd 
               Caption         =   "Add Criterion"
               Height          =   372
               Left            =   240
               TabIndex        =   28
               Top             =   1920
               Width           =   1212
            End
            Begin VB.CheckBox Chk_FldCritCase 
               Caption         =   "Selection is case sensitive"
               Height          =   372
               Left            =   5520
               TabIndex        =   27
               Top             =   1920
               Width           =   1392
            End
            Begin VB.ComboBox Cbo_FldCritOp 
               Height          =   315
               Left            =   1860
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   1020
               Width           =   3492
            End
            Begin VB.CheckBox Chk_AllCriteriaRequired 
               Caption         =   $"frm_rwiz.frx":0DCC
               Height          =   615
               Left            =   1860
               TabIndex        =   25
               Top             =   1680
               Width           =   3375
            End
            Begin atc2valtext.ValText Txt_FldCritValue 
               Height          =   312
               Left            =   5520
               TabIndex        =   149
               Top             =   1020
               Width           =   1392
               _ExtentX        =   2461
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "frm_rwiz.frx":0E58
               Text            =   ""
               TXTAlign        =   2
               AutoSelect      =   0
            End
            Begin atc2valtext.ValText Txt_FldCritValue2 
               Height          =   312
               Left            =   5520
               TabIndex        =   150
               Top             =   1500
               Width           =   1392
               _ExtentX        =   2461
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MouseIcon       =   "frm_rwiz.frx":0E74
               Text            =   ""
               TXTAlign        =   2
               AutoSelect      =   0
            End
            Begin VB.Label Lbl_FldCrit 
               Caption         =   "Select only values in this field which are"
               Height          =   372
               Left            =   300
               TabIndex        =   33
               Top             =   900
               Width           =   1452
            End
            Begin VB.Label Lbl_FldCritAnd 
               Alignment       =   1  'Right Justify
               Caption         =   "and"
               Height          =   252
               Left            =   5940
               TabIndex        =   32
               Top             =   1320
               Width           =   432
            End
            Begin VB.Label Lbl_CriteriaInst 
               BackStyle       =   0  'Transparent
               Caption         =   "Click on a cell in one of the criteria rows in the grid to add / alter / delete a criterion."
               Height          =   312
               Index           =   0
               Left            =   180
               TabIndex        =   31
               Top             =   180
               Width           =   6612
            End
            Begin VB.Label Lbl_CriteriaInst 
               BackStyle       =   0  'Transparent
               Caption         =   $"frm_rwiz.frx":0E90
               Height          =   492
               Index           =   1
               Left            =   180
               TabIndex        =   30
               Top             =   360
               Width           =   6792
            End
         End
         Begin VB.Frame Fra_Group 
            Caption         =   "Group / Sorting Options"
            Height          =   2292
            Left            =   5625
            TabIndex        =   35
            Top             =   2580
            Width           =   1815
            Begin VB.CheckBox Chk_Group 
               Caption         =   "Group by this field"
               Height          =   315
               Left            =   120
               TabIndex        =   40
               Top             =   840
               Width           =   1560
            End
            Begin VB.CheckBox Chk_GroupOpt 
               Caption         =   "Show in header"
               Height          =   252
               Index           =   0
               Left            =   90
               TabIndex        =   39
               Top             =   1440
               Width           =   1620
            End
            Begin VB.CheckBox Chk_GroupOpt 
               Caption         =   "Page Break after"
               Height          =   252
               Index           =   1
               Left            =   90
               TabIndex        =   38
               Top             =   1725
               Width           =   1512
            End
            Begin VB.ComboBox Cbo_Sorting 
               Height          =   315
               ItemData        =   "frm_rwiz.frx":0F20
               Left            =   105
               List            =   "frm_rwiz.frx":0F22
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   495
               Width           =   1455
            End
            Begin VB.CheckBox Chk_GroupOpt 
               Caption         =   "Sub total"
               Height          =   252
               Index           =   2
               Left            =   90
               TabIndex        =   36
               Top             =   2010
               Width           =   1512
            End
            Begin VB.Label Lbl_Group 
               Caption         =   "For each group:"
               Height          =   255
               Left            =   105
               TabIndex        =   42
               Top             =   1215
               Width           =   1575
            End
            Begin VB.Label Lbl_Sorting 
               Caption         =   "Sorting"
               Height          =   225
               Left            =   120
               TabIndex        =   41
               Top             =   255
               Width           =   615
            End
         End
         Begin MSFlexGridLib.MSFlexGrid FlG_Fields 
            DragIcon        =   "frm_rwiz.frx":0F24
            Height          =   2055
            Left            =   180
            TabIndex        =   34
            Top             =   420
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   3625
            _Version        =   393216
            Rows            =   6
            Cols            =   1
            FocusRect       =   0
            SelectionMode   =   2
            AllowUserResizing=   1
         End
         Begin VB.HScrollBar HSc_FGLeft 
            Height          =   375
            Left            =   60
            TabIndex        =   78
            Top             =   420
            Width           =   375
         End
         Begin VB.HScrollBar HSc_FGRight 
            Height          =   375
            Left            =   7380
            TabIndex        =   77
            Top             =   420
            Width           =   375
         End
         Begin VB.Frame Fra_FmtDDetail 
            BorderStyle     =   0  'None
            Height          =   2952
            Left            =   -15
            TabIndex        =   43
            Top             =   2565
            Width           =   7332
            Begin VB.Frame Fra_FmtDataDisp 
               Caption         =   "Data Display"
               Height          =   672
               Left            =   120
               TabIndex        =   68
               Top             =   1140
               Width           =   5415
               Begin VB.TextBox Txt_Prefix 
                  Height          =   264
                  Left            =   600
                  TabIndex        =   72
                  Top             =   240
                  Width           =   732
               End
               Begin VB.TextBox Txt_Suffix 
                  Height          =   264
                  Left            =   1860
                  TabIndex        =   71
                  Top             =   240
                  Width           =   732
               End
               Begin VB.CommandButton Cmd_FontData 
                  Caption         =   "Font"
                  Height          =   360
                  Left            =   4320
                  TabIndex        =   70
                  Top             =   240
                  Width           =   960
               End
               Begin VB.ComboBox Cbo_Alignment 
                  Height          =   315
                  Left            =   3120
                  Style           =   2  'Dropdown List
                  TabIndex        =   69
                  Top             =   240
                  Width           =   1035
               End
               Begin VB.Label Lbl_Prefix 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Prefix"
                  Height          =   252
                  Left            =   60
                  TabIndex        =   75
                  Top             =   300
                  Width           =   492
               End
               Begin VB.Label Lbl_Suffix 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Suffix"
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   74
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Lbl_Alignment 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Align"
                  Height          =   255
                  Left            =   2340
                  TabIndex        =   73
                  Top             =   300
                  Width           =   735
               End
            End
            Begin VB.Frame Fra_FmtData 
               Caption         =   "Data Format"
               Height          =   1032
               Left            =   135
               TabIndex        =   51
               Top             =   1860
               Width           =   5412
               Begin VB.ComboBox Cbo_Format 
                  Height          =   315
                  Left            =   2940
                  TabIndex        =   58
                  Text            =   "Combo1"
                  Top             =   255
                  Width           =   2340
               End
               Begin VB.CheckBox Chk_Trunc 
                  Caption         =   "Truncate to"
                  Height          =   192
                  Left            =   2610
                  TabIndex        =   57
                  Top             =   315
                  Visible         =   0   'False
                  Width           =   1275
               End
               Begin VB.TextBox Txt_FW 
                  Height          =   288
                  Left            =   1020
                  TabIndex        =   56
                  Top             =   660
                  Visible         =   0   'False
                  Width           =   732
               End
               Begin VB.ComboBox Cbo_BooleanFalse 
                  Height          =   315
                  Left            =   4140
                  TabIndex        =   55
                  Text            =   "Combo1"
                  Top             =   660
                  Width           =   1152
               End
               Begin VB.ComboBox Cbo_BooleanTrue 
                  Height          =   315
                  Left            =   4140
                  TabIndex        =   54
                  Text            =   "Combo1"
                  Top             =   240
                  Width           =   1152
               End
               Begin VB.TextBox Txt_Trunc 
                  Height          =   288
                  Left            =   3900
                  TabIndex        =   53
                  Top             =   255
                  Visible         =   0   'False
                  Width           =   585
               End
               Begin VB.ComboBox Cbo_DataType 
                  Height          =   315
                  Left            =   540
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Label Lbl_Trunc 
                  Caption         =   "characters"
                  Height          =   252
                  Left            =   4560
                  TabIndex        =   67
                  Top             =   300
                  Width           =   792
               End
               Begin VB.Label Lbl_BooleanFalse 
                  Alignment       =   1  'Right Justify
                  Caption         =   "False text"
                  Height          =   255
                  Left            =   2500
                  TabIndex        =   66
                  Top             =   660
                  Width           =   1400
               End
               Begin VB.Label Lbl_DataType 
                  Caption         =   "Type"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   65
                  Top             =   300
                  Width           =   375
               End
               Begin VB.Label Lbl_BooleanTrue 
                  Alignment       =   1  'Right Justify
                  Caption         =   "True text"
                  Height          =   255
                  Left            =   2500
                  TabIndex        =   64
                  Top             =   240
                  Width           =   1400
               End
               Begin VB.Label Lbl_DataType1 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   252
                  Left            =   540
                  TabIndex        =   63
                  Top             =   300
                  Width           =   1332
               End
               Begin VB.Label Lbl_FW 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Fixed Width Characters"
                  Height          =   372
                  Left            =   60
                  TabIndex        =   62
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   912
               End
               Begin VB.Label lbl_Format 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   2940
                  TabIndex        =   61
                  ToolTipText     =   "Shows the effect of the formatting on the sample data item"
                  Top             =   615
                  Width           =   2340
               End
               Begin VB.Label Lbl_Fmt 
                  Caption         =   "Format"
                  Height          =   255
                  Left            =   2055
                  TabIndex        =   60
                  Top             =   285
                  Width           =   810
               End
               Begin VB.Label Lbl_FmtSample 
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   7
                  ToolTipText     =   "Sample data item"
                  Top             =   600
                  Width           =   810
               End
            End
            Begin VB.ComboBox Cbo_Sum 
               Height          =   315
               ItemData        =   "frm_rwiz.frx":106E
               Left            =   5700
               List            =   "frm_rwiz.frx":1070
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   2565
               Width           =   1455
            End
            Begin VB.Frame Fra_Header 
               Caption         =   "Header Display"
               Height          =   675
               Left            =   120
               TabIndex        =   46
               Top             =   420
               Width           =   5430
               Begin VB.ComboBox Cbo_FieldName 
                  Height          =   315
                  Left            =   600
                  TabIndex        =   48
                  Top             =   240
                  Width           =   3315
               End
               Begin VB.CommandButton Cmd_FontHead 
                  Caption         =   "Font"
                  Height          =   360
                  Left            =   4320
                  TabIndex        =   47
                  Top             =   240
                  Width           =   960
               End
               Begin VB.Label Lbl_FieldName 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Text"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   49
                  Top             =   300
                  Width           =   435
               End
            End
            Begin VB.CheckBox Chk_Wrap 
               Caption         =   "Wrap field text"
               Height          =   315
               Left            =   1800
               TabIndex        =   45
               Top             =   0
               Width           =   1395
            End
            Begin VB.CheckBox Chk_NoSquash 
               Caption         =   "Do not squash field text"
               Height          =   315
               Left            =   3480
               TabIndex        =   44
               Top             =   0
               Width           =   2355
            End
            Begin MSComDlg.CommonDialog CD_Fonts 
               Left            =   180
               Top             =   2340
               _ExtentX        =   688
               _ExtentY        =   688
               _Version        =   393216
            End
            Begin VB.Label lbl_Sum 
               Caption         =   "Sum"
               Height          =   210
               Left            =   5700
               TabIndex        =   76
               Top             =   2310
               Width           =   1215
            End
         End
         Begin VB.Label Lbl_CriteriaLbl 
            Caption         =   "Field details:"
            Height          =   252
            Left            =   60
            TabIndex        =   80
            Top             =   180
            Visible         =   0   'False
            Width           =   1032
         End
         Begin VB.Label Lbl_FormatInst 
            Caption         =   $"frm_rwiz.frx":1072
            Height          =   495
            Left            =   60
            TabIndex        =   79
            Top             =   0
            Visible         =   0   'False
            Width           =   7275
         End
      End
      Begin VB.Frame Fra_Fields 
         BorderStyle     =   0  'None
         Height          =   5472
         Left            =   -74865
         TabIndex        =   14
         Top             =   420
         Width           =   7920
         Begin VB.ListBox lstFields 
            Height          =   3870
            IntegralHeight  =   0   'False
            Left            =   4980
            TabIndex        =   16
            Top             =   1080
            Width           =   2835
         End
         Begin VB.CommandButton Cmd_Clear 
            Caption         =   "&Clear All"
            Height          =   375
            Left            =   7065
            TabIndex        =   15
            ToolTipText     =   "Remove all selected fields and reinitialise report defaults"
            Top             =   5040
            Width           =   780
         End
         Begin MSComctlLib.ImageList ImL 
            Left            =   7200
            Top             =   240
            _ExtentX        =   794
            _ExtentY        =   794
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   18
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":1149
                  Key             =   "EMPTY_BOX"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":149D
                  Key             =   "CHECK_BOX_BLACK"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":17F1
                  Key             =   "EMPTY"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":1B45
                  Key             =   "CHECK_BLACK"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":1E99
                  Key             =   "CROSS_BOX_BLACK"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":21ED
                  Key             =   "CROSS_BLACK"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":2541
                  Key             =   "CHECK_BOX_BLUE"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":2895
                  Key             =   "CROSS_BOX_RED"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":2BE9
                  Key             =   "CHECK_BLUE"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":2F3D
                  Key             =   "CROSS_RED"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":3291
                  Key             =   "CHECK_BOX_RED"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":35E5
                  Key             =   "CROSS_BOX_BLUE"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":3939
                  Key             =   "CHECK_RED"
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":3C8D
                  Key             =   "CROSS_BLUE"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":3FE1
                  Key             =   "FOLDER_CLOSED"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":40F3
                  Key             =   "FOLDER_OPEN"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":4205
                  Key             =   "DISABLED_NODE"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frm_rwiz.frx":4557
                  Key             =   "FOLDER_DISABLED"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TrV_Fields 
            Height          =   3870
            Left            =   60
            TabIndex        =   17
            Top             =   1080
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   6826
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   617
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImL"
            Appearance      =   1
         End
         Begin VB.Label Lbl_FieldsInst 
            Caption         =   "Click in the left-hand pane on each field you want to add to or remove from your report."
            Height          =   252
            Index           =   0
            Left            =   180
            TabIndex        =   22
            Top             =   0
            Width           =   6732
         End
         Begin VB.Label Lbl_FieldsInst 
            Caption         =   $"frm_rwiz.frx":4669
            Height          =   495
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   240
            Width           =   5835
         End
         Begin VB.Label Lbl_FieldTree 
            Caption         =   "Fields available:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label Lbl_FieldsSelected 
            Caption         =   "Fields selected:"
            Height          =   255
            Left            =   5040
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lbl_SpecFilename 
            Height          =   465
            Left            =   60
            TabIndex        =   18
            Top             =   4995
            Width           =   6900
         End
      End
   End
   Begin VB.Menu Mnu_RepWiz 
      Caption         =   "RepWizMenu"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Font 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu_ChangeFont 
         Caption         =   "Change Font"
      End
      Begin VB.Menu Mnu_InsLineBreak 
         Caption         =   "Insert Line Break"
      End
      Begin VB.Menu Mnu_InsPageNumber 
         Caption         =   "Insert Page Number"
      End
      Begin VB.Menu Mnu_InsDate 
         Caption         =   "Insert Date"
      End
      Begin VB.Menu Mnu_InsTime 
         Caption         =   "Insert Time"
      End
      Begin VB.Menu Mnu_InsUser 
         Caption         =   "Insert User"
      End
      Begin VB.Menu Mnu_InsAppName 
         Caption         =   "Insert Application Name"
      End
      Begin VB.Menu Mnu_InsAppVer 
         Caption         =   "Insert Application Version"
      End
      Begin VB.Menu mnu_InsSpecFilename 
         Caption         =   "Insert Report Filename"
      End
      Begin VB.Menu Mnu_UserDefined 
         Caption         =   "(User Defined)"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Frm_RepWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IReportForm
Implements IFolderBrowserControlEvents

#If AbacusReporter Then
  Private m_fb As FolderBrowserControl
#End If


Private mReportWiz As ReportWizard
Private mReportDataSets As DataSetCollection
Private mReportDetails As ReportDetails
Private mReportFields As Collection
Private mReportParser As Parser
Private mCurrentField As ReportField

Private HFTextBoxWithFocus As TextBox

Private hwndTV As Long

Private DragFromCol As Long
Private DragToCol As Long

Public InFillFormat As Boolean
Private InFillRepFormat As Boolean

Private Enum FLEXGRID_ROWS
  ROW_NAME = 1
  ROW_DATASET = 2
  ROW_DATATYPE = 3
  ROW_GROUPBY = 4
  ROW_HIDE = 5
  ROW_WIDTH_TYPE = 6
  ROW_KEYSTRING = 7
  ROW_FIRST_CRITERIA = 8
End Enum

Private mMaxCriteria As Long
Private InNodeClick As Boolean

Public CancelReport As Boolean
Private mCancelCursorState As Long

Public ActiveListItem_FileGroup As ListItem
Public ActiveListItem_FileGroupMember As ListItem

Public Function BeginWiz() As Long
  hwndTV = TrV_Fields.hwnd
  BeginWiz = hwndTV
  If mReportWiz.AR Then 'set this to 1 fram past where you want to be
    mReportWiz.CurrentFrame = "Fra_Fields"
  Else
    mReportWiz.CurrentFrame = "Fra_Format"
  End If
  Set mReportParser = mReportWiz.ReportParser
  Call PrepareControls
  Call FillReportDetails
   
  
  Call Cmd_Back_Click
End Function

Private Sub Cbo_BooleanFalse_LostFocus()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_BooleanTrue_LostFocus()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_Format_Change()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_Format_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_Format_LostFocus()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_GHDelimiter_Change()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Cbo_GHSeparator_Change()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Cbo_Sorting_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Chk_AlignHeadings_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_CollapseAll_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_DisplayRecordCount_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_FitToPage_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_GroupHeaders_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_IncBlankLines_Click()
  Static Inchange As Boolean
  
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_InclColHeaders_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_NoSquash_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Chk_PageFoot_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_PageHead_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_RepHead_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Cbo_Sum_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Chk_RHonAllPages_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_Wrap_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub chkPreviewLines_Click()
  Me.txtPreview.Enabled = (Me.chkPreviewLines.Value = vbUnchecked)
  Me.UpDownPreview.Enabled = (Me.chkPreviewLines.Value = vbUnchecked)
End Sub

Private Sub Cmd_Clear_Click()
  If DisplayMessage(Me, "Are you sure you want to clear all the selected fields?", "Clear Selected Fields", "Yes", "No") Then
    Call ClearAllFields
  End If
End Sub

Public Sub ClearAllFields()
  Dim rFld As ReportField
  TrV_Fields.Nodes.Clear
  lstFields.Clear
  Call FillTreeView(TrV_Fields, mReportDataSets)
  For Each rFld In mReportFields
    rFld.Selected = False
  Next rFld
  Call ClearCollection(mReportFields)
  Cmd_Preview.Enabled = False
  Cmd_Print.Enabled = False
  Cmd_Export.Enabled = False
  Cmd_SaveReport.Enabled = False
  Call SetButtons
End Sub

Private Sub Cmd_DataFontSet_Click()
  Call SetFont(mReportDetails.DataFont, "Set Global Field Font")
End Sub

Private Sub Cmd_DataFontSet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Cmd_DataFontSet.ToolTipText = mReportDetails.DataFont.FontDescription
End Sub

Private Sub Cmd_Exit_Click()
  'RK QUERY: Prompt Save of ARConfigFile based on RepWiz.ARConfigDirty here?
  #If AbacusReporter Then
    If mReportWiz.FileGroupContainer.Dirty Then
      Call mReportWiz.FileGroupContainer.Save(Me, mReportDetails)
    End If
  #End If
  Call mReportWiz.ExitWizard
End Sub

Private Sub Cmd_Export_Click()
  Dim LineCount As Long
  On Error GoTo Err_Err
  
  LineCount = -1
  If Me.chkPreviewLines.Value = vbUnchecked Then LineCount = CLng(Me.txtPreview.Text)
  Call mReportWiz.PrepareReport(LineCount, REPORTW_PREPARE_EXPORT)
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "Cmd_Export_Click", "Error in exporting", Err.Description)
    Resume Err_End
    Resume
End Sub

Private Sub Cmd_FontData_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Cmd_FontData.ToolTipText = mCurrentField.DataFont.FontDescriptionRestricted(mReportDetails.DataFont.Name, mReportDetails.DataFont.Size)
End Sub

Private Sub Cmd_FontHead_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Cmd_FontHead.ToolTipText = mCurrentField.DataFont.FontDescriptionRestricted(mReportDetails.DataFont.Name, mReportDetails.DataFont.Size)
End Sub

Private Sub Cmd_Print_Click()
  Dim LineCount As Long
  On Error GoTo Err_Err
  LineCount = -1
  If Me.chkPreviewLines.Value = vbUnchecked Then LineCount = CLng(Me.txtPreview.Text)
  Call mReportWiz.PrepareReport(LineCount, REPORTW_PRINT)
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "Cmd_Print_Click", "Error in printing", Err.Description)
    Resume Err_End
    Resume
End Sub

Private Sub Cmd_SaveReport_Click()
  On Error GoTo Err_Err
  Cmd_SaveReport.Enabled = False
  #If AbacusReporter Then
    If CheckFileGroupSettings(mReportWiz, Me, mReportDetails) Then
      Call mReportWiz.SaveReport 'RK 10/02/05
      Cmd_SaveReport.Enabled = True
    End If
  #Else
    Call mReportWiz.SaveReport
    Cmd_SaveReport.Enabled = True
  #End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "Cmd_SaveReport_Click", "Error saving report", Err.Description)
    Resume Err_End
    Resume
End Sub

Private Sub HideTabs()
  'RK Hide tabs - used for wizard screen manipulation
  'Me.Height = 7335
  'Me.Width = 8250
  Me.ssTab.top = Me.ssTab.top - (Me.ssTab.TabHeight + 100)
  Me.Height = Me.Height - (Me.ssTab.TabHeight + 100)
  'Me.ssTabReporterVersions.Height = Me.ssTabReporterVersions.Height + Fra_Buttons.Height
End Sub

Private Sub Form_Load()
 
 #If AbacusReporter Then
  Set m_fb = New FolderBrowserControl
  Call m_fb.Setup(Me, lblPlaceHolder, "fb")
  m_fb.Visible = True
 #End If

End Sub

Private Property Get IReportForm_FormType() As REPORTW_GOTOFORM
  IReportForm_FormType = TCSREPWIZ
End Property

Private Property Set IReportForm_ReportDataSets(RHS As DataSetCollection)
  Set mReportDataSets = RHS
End Property

Private Property Set IReportForm_ReportDetails(RHS As ReportDetails)
  Set mReportDetails = RHS
End Property

Private Property Set IReportForm_ReportFields(RHS As Collection)
  Set mReportFields = RHS
End Property

Private Property Set IReportForm_ReportWizard(RHS As ReportWizard)
  Set mReportWiz = RHS
End Property

Public Property Get FieldSelected() As ReportField
  Set FieldSelected = mCurrentField
End Property

Public Property Set FieldSelected(ByVal NewValue As ReportField)
  Set mCurrentField = NewValue
End Property

Private Sub Cbo_Alignment_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_DataType_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Call FillFormatDefaults(mCurrentField.DataType)
    Inchange = False
  End If
End Sub

Private Sub Cbo_FieldName_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_FieldName_LostFocus()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cbo_FldCritOp_Click()
  Dim cOp As CRITERION_COMPARISONS
  Dim sText As String
  
  cOp = Cbo_FldCritOp.ListIndex + 1
  If (cOp = CRITERION_COMPARISON_BW) Or (cOp = CRITERION_COMPARISON_BWEQ) Then
    Lbl_FldCritAnd.Visible = True
    Txt_FldCritValue2.Visible = True
  Else
    Lbl_FldCritAnd.Visible = False
    Txt_FldCritValue2.Visible = False
  End If
  If Txt_FldCritValue.Enabled And Fra_FldCriteria.Enabled = True Then
    Call Txt_FldCritValue.SetFocus
  End If
End Sub

Private Sub Chk_AutoWidth_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_Display_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Chk_FldCritCase_Click()
  Txt_FldCritValue.SetFocus
End Sub

Private Sub Chk_Group_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Chk_GroupOpt_Click(Index As Integer)
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Chk_TrimHeadings_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Chk_Trunc_Click()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Cmd_Back_Click()
  On Error GoTo Err_Err
  Select Case mReportWiz.CurrentFrame
    Case "Fra_Format"
      If Fra_FldCriteria.Visible Or Not Fra_FldCriteria.Enabled Then
        Call ChangeScreen("Fra_Fields")
      Else
        Call ChangeScreen("Fra_Criteria")
      End If
    Case "Fra_Report"
      Call ChangeScreen("Fra_Format")
    #If AbacusReporter Then
      Case "Fra_Fields"
        Call ChangeScreen("Fra_ARFileGroups")
    #End If
  End Select
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "Cmd_Back_Click", "Error moving back", Err.Description)
    Resume Err_End
    Resume
End Sub


Private Sub Cmd_FldCritAdd_Click()
  Dim Crit As Criterion, i As Long, Index As Long
  
  On Error GoTo CriteriaAdd_Err
  Set Crit = New Criterion
  Index = FlG_Fields.row - ROW_FIRST_CRITERIA + 1
  Crit.Comparison = Cbo_FldCritOp.ListIndex + 1
  If Crit.Comparison = CRITERION_COMPARISON_UNDATED Or Crit.Comparison = CRITERION_COMPARISON_DATED Then
    Crit.Value = ""
  Else
    'RK QUERY: This converts incorrectly for dates - does the value have to be type with a ValText box
    If mCurrentField.DataType = TYPE_DATE Then
      Crit.Value = TryConvertDateDMY(Txt_FldCritValue.Text)
    Else
      Crit.Value = GetTypedValue(Txt_FldCritValue.Text, mCurrentField.DataType)
    End If
  End If
  If Crit.Comparison = CRITERION_COMPARISON_BW Or Crit.Comparison = CRITERION_COMPARISON_BWEQ Then
    If mCurrentField.DataType = TYPE_DATE Then
      Crit.Value2 = TryConvertDateDMY(Txt_FldCritValue2.Text)
    Else
      Crit.Value2 = GetTypedValue(Txt_FldCritValue2.Text, mCurrentField.DataType)
    End If
  End If
  If Index <= mCurrentField.Criteria.Count Then Call mCurrentField.Criteria.Remove(Index)
  Call mCurrentField.Criteria.AddIndex(Crit, Index)
  Call DisplayCriteria
  Txt_FldCritValue.SetFocus
  
CriteriaAdd_End:
  Exit Sub
  
CriteriaAdd_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddCriteria", "Unable to add criterion", "Error adding criterion.")
  Resume CriteriaAdd_End
  Resume
End Sub


Private Sub Cmd_FldCritDel_Click()
  Dim Index As Long
  
  FlG_Fields.TextMatrix(FlG_Fields.row, FlG_Fields.col) = ""
  Index = FlG_Fields.row - ROW_FIRST_CRITERIA + 1
  Call mCurrentField.Criteria.Remove(Index)
  Call DisplayCriteria
  Txt_FldCritValue.SetFocus
End Sub


Private Sub Cmd_FontData_Click()
  Dim Apply As Boolean
  Dim Fld As ReportField
  
  'cad this is wrong???
  'the Apply thing is wrong
  'need common function with that of below
  
  
  If SetFont(mCurrentField.DataFont, "Set Field Data Font", True, Apply) Then
    If Not Apply Then Apply = DisplayMessage(Me, "Do you want to apply this font setting to all the other field data ?", "Data Font", "Yes", "No")
    If Apply Then
      For Each Fld In mReportFields
        If Not Fld Is mCurrentField Then
          Fld.DataFont.Bold = mCurrentField.DataFont.Bold
          Fld.DataFont.ForeColor = mCurrentField.DataFont.ForeColor
          Fld.DataFont.Italic = mCurrentField.DataFont.Italic
          Fld.DataFont.Name = mCurrentField.DataFont.Name
          Fld.DataFont.Size = mCurrentField.DataFont.Size
          Fld.DataFont.Strikethru = mCurrentField.DataFont.Strikethru
          Fld.DataFont.Underline = mCurrentField.DataFont.Underline
        End If
      Next Fld
    End If
  End If
End Sub


Private Sub Cmd_FontHead_Click()
  Dim Apply As Boolean
  Dim Fld As ReportField
  If SetFont(mCurrentField.HeadingFont, "Set Field Header Font", True, Apply) Then
    If Not Apply Then
      If MsgBox("Do you want to apply this font setting to all the other field headings ?", vbYesNo, "Heading Font") = vbYes Then Apply = True
    End If
    If Apply Then
      For Each Fld In mReportFields
        If Not Fld Is mCurrentField Then
          Fld.HeadingFont.Bold = mCurrentField.HeadingFont.Bold
          Fld.HeadingFont.ForeColor = mCurrentField.HeadingFont.ForeColor
          Fld.HeadingFont.Italic = mCurrentField.HeadingFont.Italic
          Fld.HeadingFont.Name = mCurrentField.HeadingFont.Name
          Fld.HeadingFont.Size = mCurrentField.HeadingFont.Size
          Fld.HeadingFont.Strikethru = mCurrentField.HeadingFont.Strikethru
          Fld.HeadingFont.Underline = mCurrentField.HeadingFont.Underline
        End If
      Next Fld
    End If
  End If
End Sub

Private Sub Cmd_LoadReport_Click()
  Dim FileName As String
  On Error GoTo Err_Err
  Cmd_LoadReport.Enabled = False
  FileName = FileOpenDlg("Open Report", "Report Wizard Files (*.rep)|*.rep|All Files (*.*)|*.*", mReportWiz.ReportFilesPath)
  If Len(FileName) > 0 Then
    mReportWiz.ReportFileName = FileName
    If LoadReportDetails(RepParser, mReportWiz, Me, mReportDataSets, mReportDetails, mReportFields, mReportWiz.ReportFileName) Then
      mReportWiz.Status = "Specification successfully loaded"
    End If
  End If
  Cmd_LoadReport.Enabled = True
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "Cmd_LoadReport_Click", "Error Loading Report", Err.Description)
    Resume Err_End
    Resume
End Sub

Private Sub Cmd_Next_Click()
  On Error GoTo Err_Err
  
  Select Case mReportWiz.CurrentFrame
    Case "Fra_Fields"
      Call ChangeScreen("Fra_Criteria")
    Case "Fra_Format"
      If Fra_FldCriteria.Visible Then
        Call ChangeScreen("Fra_Format")
      Else
        Call ChangeScreen("Fra_Report")
      End If
    #If AbacusReporter Then
    Case "Fra_ARFileGroups"
        If CheckFileGroupSettings(mReportWiz, Me, mReportDetails) Then
          Call ChangeScreen("Fra_Fields")
        End If
    #End If
  End Select
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "Cmd_Next_Click", "Error moving next", Err.Description)
    Resume Err_End
    Resume
End Sub

Public Sub Cmd_Preview_Click()
  Dim LineCount As Long
  On Error GoTo Err_Err
  
  LineCount = -1
  If Me.chkPreviewLines.Value = vbUnchecked Then LineCount = CLng(Me.txtPreview.Text)
  Call mReportWiz.PrepareReport(LineCount, REPORTW_PREVIEW)

Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "Cmd_Preview_Click", "Error in previewing report", Err.Description)
    Resume Err_End
    Resume
End Sub

Private Sub FlG_Fields_DragDrop(Source As Control, x As Single, y As Single)
  Dim rFld As ReportField, i As Long, nod As node
   
  DragToCol = FlG_Fields.MouseCol
  If Source Is FlG_Fields Then
    If DragToCol < FlG_Fields.FixedCols Then DragToCol = FlG_Fields.FixedCols
    If DragFromCol = DragToCol Then Exit Sub
    FlG_Fields.ColPosition(DragFromCol) = DragToCol
    FlG_Fields.col = DragToCol
    
    If DragToCol < DragFromCol Then
      For i = DragToCol To (DragFromCol - 1)
        Set rFld = mReportFields.Item(i)
        rFld.Order = rFld.Order + 1
      Next i
    Else
      For i = (DragFromCol + 1) To DragToCol
        Set rFld = mReportFields.Item(i)
        rFld.Order = rFld.Order - 1
      Next i
    End If
    Set rFld = mReportFields.Item(DragFromCol)
    rFld.Order = DragToCol
    Call UpdateFieldOrders(mReportFields)  ' order has been preserved
    Call RefreshList
    Call FlG_Fields_SelChange
  End If
End Sub

Private Sub FlG_Fields_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not Fra_FldCriteria.Visible Then
    If FlG_Fields.MouseCol >= FlG_Fields.FixedCols And FlG_Fields.MouseRow < FlG_Fields.FixedRows Then
      DragFromCol = FlG_Fields.MouseCol
      Call FlG_Fields_SelChange
      FlG_Fields.Drag vbBeginDrag
    End If
  End If
End Sub

Private Sub FlG_Fields_SelChange()
  Static Inchange As Boolean
  Dim Crit As Criterion, i As Long, Index As Long
  Dim NewSelection As Boolean
  
  On Error Resume Next
  If Inchange Then Exit Sub
  Inchange = True
  If Fra_FldCriteria.Visible Then
    NewSelection = False
    FlG_Fields.Redraw = False
    If FlG_Fields.ColSel <> FlG_Fields.col Then
      FlG_Fields.col = FlG_Fields.col
      NewSelection = True
    End If
    If FlG_Fields.row < ROW_FIRST_CRITERIA Then
      FlG_Fields.row = ROW_FIRST_CRITERIA
    End If
    FlG_Fields.Redraw = True
    If NewSelection Or True Then
      Call FillFormatDetails(SCREEN_CRITERIA)
      Call FillCritOpCbo(mCurrentField.DataType)
      'RK TODO: 11/03/05 Attempted to use ValTxt box replaced text box.
      Txt_FldCritValue.TypeOfData = CValDataTypes(mCurrentField.DataType)
      Txt_FldCritValue.Text = ""
      Txt_FldCritValue2.TypeOfData = CValDataTypes(mCurrentField.DataType)
      Txt_FldCritValue2.Text = ""
      Chk_FldCritCase.Value = vbUnchecked
      Cmd_FldCritDel.Enabled = False
      Cmd_FldCritAdd.Caption = "Add criterion"
      Index = FlG_Fields.row - ROW_FIRST_CRITERIA + 1
      If Index <= mCurrentField.Criteria.Count Then
        Set Crit = mCurrentField.Criteria.Item(Index)
        If Not Crit Is Nothing Then
          Cbo_FldCritOp.Text = Cbo_FldCritOp.List(Crit.Comparison - 1)
          If mCurrentField.DataType = TYPE_DATE Then
            Txt_FldCritValue.Text = DateValReadToScreen(Crit.Value)
            Txt_FldCritValue2.Text = DateValReadToScreen(Crit.Value2)
          Else
            Txt_FldCritValue.Text = Crit.Value
            Txt_FldCritValue2.Text = Crit.Value2
          End If
          Chk_FldCritCase.Value = vbUnchecked
          Cmd_FldCritDel.Enabled = True
          Cmd_FldCritAdd.Caption = "Alter criterion"
        End If
      End If
    End If
  Else
    FlG_Fields.Redraw = False
    FlG_Fields.row = FlG_Fields.FixedRows
    FlG_Fields.col = FlG_Fields.col
    FlG_Fields.RowSel = FlG_Fields.Rows - 1
    FlG_Fields.Redraw = True
    Call ReadFormatDetails
    
    Cbo_FieldName.Clear
    If Len(mCurrentField.FieldName) Then
      Cbo_FieldName.AddItem mCurrentField.FieldName
    End If
    Cbo_FieldName.AddItem mCurrentField.Description  'cad p11d
    
    Call FillFormatDetails(SCREEN_FORMATS)
    Call FillFormatDefaults(mCurrentField.DataType)
  End If
  Inchange = False
End Sub
'

Private Sub lstFields_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim nod As node
  
  If lstFields.ListIndex >= 0 Then
    Set nod = TrV_Fields.Nodes(lstFields.ItemData(lstFields.ListIndex))
    If Button = vbRightButton Then
      nod.Selected = True
      nod.EnsureVisible
    End If
    lstFields.ToolTipText = "Data Set: " & nod.Parent.Text
  Else
    lstFields.ToolTipText = ""
  End If
End Sub
#If AbacusReporter Then
  
  Private Sub chkAllPacks_Validate(Cancel As Boolean)
    If Not IsFileGroupMemberEdit(Me) Then
      Cancel = SaveFileGroupMember(mReportWiz, Me, ActiveListItem_FileGroupMember)
    End If
  End Sub

  Public Property Get fb() As FolderBrowserControl
    Set fb = m_fb
  End Property
  
  Private Sub lvwFileGroupMembers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Check Member definition changes are valid
    If SaveFileGroupMember(mReportWiz, Me, Me.ActiveListItem_FileGroupMember) Then
      Call LoadFileGroupMember(mReportWiz, Me, Item)
    End If
  End Sub
  
  Private Sub lvwFileGroupMembers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim li As ListItem
    Set li = lvwFileGroupMembers.HitTest(x, y)
    If Not li Is Nothing Then li.ToolTipText = li.Text
  
  End Sub
 
  Private Sub lvwFileGroups_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo Err_Err
    
      Call LoadlvwFileGroupsItem(Item)
      Call SetButtons
    
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "lvwFileGroups_ItemCheck", "Error editing File Group", Err.Description)
    Item.Checked = Not Item.Checked
    Resume Err_End
    Resume
  End Sub
  
  Private Sub LoadlvwFileGroupsItem(ByVal Item As MSComctlLib.ListItem)
    Dim MyFileGroup As ARFileGroup
    Call ConvertListItem_FileGroup(mReportWiz, lvwFileGroups.SelectedItem.Text, MyFileGroup)
    Call mReportWiz.FileGroupContainer.MakeDirty(Me, True)
    
    'Validate and load
    If ValidateFileGroup(mReportWiz, Me) Then
      If ValidateFileGroupMember(mReportWiz, Me) Then
        Call LoadFileGroup(mReportWiz, Me, Item)
      End If
    End If

  End Sub
  
  Private Sub lvwFileGroups_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LoadlvwFileGroupsItem(Item)
  End Sub

  Private Sub lvwFileGroups_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim li As ListItem
    Set li = lvwFileGroups.HitTest(x, y)
    If Not li Is Nothing Then li.ToolTipText = li.Text
    
  End Sub
  
  Private Sub optDefinition_LostFocus(Index As Integer)
    'Only validate when focus is set outside of OptionButton/FolderBrowsers
    'Validation is carried here as the Option Button is really part of the fb UserControl
    If Not IsFileGroupMemberEdit(Me) Then
      Select Case Index
        Case DEFINITION_DIRECTORY
          Call SaveFileGroupMember(mReportWiz, Me, ActiveListItem_FileGroupMember)
        Case DEFINITION_FILE
          Call SaveFileGroupMember(mReportWiz, Me, ActiveListItem_FileGroupMember)
        Case Else
          Call Err.Raise(ERR_NOT_SUPPORTED, , "DefinitionType: " & Index & " is not supported.")
      End Select
    End If
  End Sub
  Private Sub optFileGrouptype_Click(Index As Integer)
    Dim MyFileGroup As ARFileGroup
    On Error GoTo Err_Err
    
    'If Me.ActiveControl Is optFileGroupType Then 'RK QUERY: Why does this not work
    If StrComp(Me.ActiveControl.Name, "optFileGroupType", vbTextCompare) = 0 Then
      'Apply change to FileGroup Member
      Call mReportWiz.FileGroupContainer.MakeDirty(Me, True)
      Call ConvertListItem_FileGroup(mReportWiz, Me.ActiveListItem_FileGroup.Text, MyFileGroup)
      MyFileGroup.GroupType = Index
      Call SetButtons
    End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "optFileGrouptype_Click", "Error editing File Group.", Err.Description)
    Resume Err_End
    Resume
  End Sub
  
  Private Sub optPackArrange_Click(Index As Integer)
    Call LoadAvailablePacks(mReportWiz, Me)
    If fb.Style = FolderBrowser Then
      Call ApplyPackSelection(mReportWiz, Me)
    Else 'If fb.Style = FileBrowser Then
      Call DisablePackSelection(Me)
    End If
    g_AbacusReporter.ArrangePacksSelection = Index
  End Sub

  Private Sub txtFileGroup_Validate(Cancel As Boolean)
    On Error GoTo Err_Err
    'Update FileGroup object name directly
    Dim MyFileGroup As ARFileGroup
    Dim li As ListItem
    Dim sOldString As String
  
  '  If Not IsFileGroupEdit(Me) Then
    If lvwFileGroups.ListItems.Count > 0 Then
      If ValidateFileGroup(mReportWiz, Me) Then
        sOldString = lvwFileGroups.SelectedItem.Text
        Call ConvertListItem_FileGroup(mReportWiz, sOldString, MyFileGroup)
        MyFileGroup.Name = txtFileGroup.Text
        lvwFileGroups.SelectedItem.Text = txtFileGroup.Text
        Call mReportWiz.FileGroupContainer.MakeDirty(Me, True)
        
        'Update tag (parentage) for FileGroupMembers list items
        For Each li In lvwFileGroupMembers.ListItems
          If StrComp(li.Tag, sOldString, vbTextCompare) = 0 Then
            li.Tag = txtFileGroup.Text
          End If
        Next li
      Else
        Cancel = True
      End If
    End If
   ' End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "txtFileGroup_Validate", "Error editing File Group", Err.Description)
    Resume Err_End
    Resume
  End Sub
  
  
  Private Sub chkAllPacks_Click()
    On Error GoTo Err_Err
    If Me.ActiveControl Is chkAllPacks Then
      If chkAllPacks.Value = vbChecked Then
        Call SetAllPackSelection(Me, SELECTED_NODE, False)
      Else
        Call SetAllPackSelection(Me, NOTSELECTED_NODE, False)
      End If
    End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(ERR_ERROR, Err, "chkAllPacks_Click", "Error updating pack selection", Err.Description)
    GoTo Err_End
    Resume
  End Sub
  
  Private Sub chkDefinitionRecursive_Click()
    Dim MyFileGroupMember As ARFileGroupMember
    On Error GoTo Err_Err
    If Me.ActiveControl Is chkDefinitionRecursive Then
    'If StrComp(Me.ActiveControl.Name, "chkDefinitionRecursive", vbTextCompare) = 0 Then
      Call ConvertListItem_FileGroupMember(Me, mReportWiz, Me.lvwFileGroupMembers.SelectedItem, Nothing, MyFileGroupMember)
      MyFileGroupMember.Recursive = chkDefinitionRecursive.Value
      Call mReportWiz.FileGroupContainer.MakeDirty(Me, True)
    End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(ERR_ERROR, Err, "chkDefinitionRecursive_Click", "Error updating File Group member", Err.Description)
    GoTo Err_End
    Resume
  End Sub
    
  Private Sub cmdAddFileGroup_Click()
    On Error GoTo Err_Err
    If lvwFileGroups.ListItems.Count > 0 Then
      'Validate existing data
      If ValidateFileGroupMember(mReportWiz, Me) Then
        If ValidateFileGroup(mReportWiz, Me) Then
          Call AddFileGroup(mReportWiz, Me)
          Call AddFileGroupMember(mReportWiz, Me)
        End If
      End If
    Else
      Call AddFileGroup(mReportWiz, Me)
      Call EnableFileGroupMembers(Me, True)
      Call AddFileGroupMember(mReportWiz, Me)
    End If
    Call SetButtons
    
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "cmdAddFileGroup_Click", "Error adding File Group", Err.Description)
    Resume Err_End
    Resume
  End Sub

  Private Sub cmdAddFileGroupMember_Click()
    On Error GoTo Err_Err
    If ValidateFileGroupMember(mReportWiz, Me) Then
      Call AddFileGroupMember(mReportWiz, Me)
    End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "cmdAddFileGroupMember_Click", "Error adding File Group Member", Err.Description)
    Resume Err_End
    Resume
  End Sub
  
  Private Sub cmdAddFileGroupMember_LostFocus()
    On Error GoTo Err_Err
    Dim i As Long
    'Force validation of new items
    If Not IsFileGroupMemberEdit(Me) Then
      Call ValidateFileGroupMember(mReportWiz, Me)
    End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "cmdAddFileGroupMember_LostFocus", "Error adding File Group Member", Err.Description)
    Resume Err_End
    Resume
  End Sub
  
  Private Sub cmdRemoveFileGroup_Click()
    Call RemoveFileGroup(mReportWiz, Me)
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "cmdRemoveFileGroup_Click", "Error removing File Group", Err.Description)
    Resume Err_End
    Resume
  End Sub
  
  Private Sub cmdRemoveFileGroupMember_Click()
    On Error GoTo Err_Err
    Call RemoveFileGroupMember(mReportWiz, Me)
  '  If lvwFileGroupMembers.ListItems.Count = 0 Then fb.Enabled = False
    
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "cmdRemoveFileGroupMember_Click", "Error removing File Group Member", Err.Description)
    Resume Err_End
    Resume
  End Sub
  
  Private Sub optDefinition_Click(Index As Integer)
    'If Option item has changed then reset fb object and pack selection
    Dim MyFileGroup As ARFileGroup
    Dim MyFileGroupMember As ARFileGroupMember
    
    If Not Me.ActiveControl Is Nothing Then
      If StrComp(Me.ActiveControl.Name, "optDefinition", vbTextCompare) = 0 Then
        Call ConvertListItem_FileGroupMember(Me, mReportWiz, Me.lvwFileGroupMembers.SelectedItem, MyFileGroup, MyFileGroupMember)
        'Save existing pack selection
        If fb.Style = FolderBrowser Then
          Call SaveFileGroupMember(mReportWiz, Me, Me.ActiveListItem_FileGroupMember)
        End If
        
        'Toggle definition
        fb.Style = Index
        If MyFileGroupMember.DefinitionType <> Index Then
          fb.Directory = "" 'RK Set to blank string as DirectoryBrowser converts to Fullpath
        Else
          fb.Directory = MyFileGroupMember.Definition
        End If
        
        'Apply pack selection
        If fb.Style = FileBrowser Then
          Call DisablePackSelection(Me)
        Else
          Call ApplyPackSelection(mReportWiz, Me)
        End If
      End If
    End If
  End Sub
  Private Sub tvwPacks_LostFocus()
    If Not IsFileGroupMemberEdit(Me) Then
      Call SaveFileGroupMember(mReportWiz, Me, ActiveListItem_FileGroupMember)
    End If
  End Sub
  
  Private Sub tvwPacks_NodeClick(ByVal node As MSComctlLib.node)
    Dim bNotAllPacks As Boolean
    On Error GoTo Err_Err
    
    'Toggle image
    If node.Image = NOTSELECTED_NODE Then
      node.Image = SELECTED_NODE
    ElseIf node.Image = SELECTED_NODE Then
      node.Image = NOTSELECTED_NODE
    End If
    
    'Set All checkbox if appropriate
    If node.Image = SELECTED_NODE Then
      chkAllPacks.Value = CChecked(CheckAllPacksSelected(Me))
    ElseIf node.Image = NOTSELECTED_NODE Then
      chkAllPacks.Value = vbUnchecked
    End If
Err_End:
    Exit Sub
Err_Err:
    Call ErrorMessage(Err.Number, Err, "tvwPacks_NodeClick", "Error editing pack selection.", Err.Description)
    Resume Err_End
    Resume
  End Sub
  
  Private Sub tvwPacks_Collapse(ByVal node As MSComctlLib.node)
    node.Image = FOLDER_CLOSED
  End Sub
  
  Private Sub tvwPacks_Expand(ByVal node As MSComctlLib.node)
    node.Image = FOLDER_OPEN
  End Sub

#End If

Private Sub Mnu_ChangeFont_Click()
  Select Case HFTextBoxWithFocus
    Case Txt_RepHead(0)
      Call SetFont(mReportDetails.RepHeaderFontL, "Set Report Header (Left) Font")
    Case Txt_RepHead(1)
      Call SetFont(mReportDetails.RepHeaderFontC, "Set Report Header (Centre) Font")
    Case Txt_RepHead(2)
      Call SetFont(mReportDetails.RepHeaderFontR, "Set Report Header (Right) Font")
    Case Txt_PageHead(0)
      Call SetFont(mReportDetails.PageHeaderFontL, "Set Page Header (Left) Font")
    Case Txt_PageHead(1)
      Call SetFont(mReportDetails.PageHeaderFontC, "Set Page Header (Centre) Font")
    Case Txt_PageHead(2)
      Call SetFont(mReportDetails.PageHeaderFontR, "Set Page Header (Right) Font")
    Case Txt_PageFoot(0)
      Call SetFont(mReportDetails.PageFooterFontL, "Set Page Footer (Left) Font")
    Case Txt_PageFoot(1)
      Call SetFont(mReportDetails.PageFooterFontC, "Set Page Footer (Centre) Font")
    Case Txt_PageFoot(2)
      Call SetFont(mReportDetails.PageFooterFontR, "Set Page Footer (Right) Font")
    Case Else
  End Select
End Sub

Private Sub Mnu_InsAppName_Click()
  HFTextBoxWithFocus.SelText = "{APPLICATION}"
End Sub

Private Sub Mnu_InsAppVer_Click()
  HFTextBoxWithFocus.SelText = "{VERSION}"
End Sub

Private Sub Mnu_InsDate_Click()
  HFTextBoxWithFocus.SelText = "{DATE}"
End Sub

Private Sub Mnu_InsLineBreak_Click()
  HFTextBoxWithFocus.SelText = "~"
End Sub

Private Sub Mnu_InsPageNumber_Click()
  HFTextBoxWithFocus.SelText = "{PAGE}"
End Sub

Private Sub mnu_InsSpecFilename_Click()
  HFTextBoxWithFocus.SelText = "{SPECFILENAME}"
End Sub

Private Sub Mnu_InsTime_Click()
  HFTextBoxWithFocus.SelText = "{TIME}"
End Sub

Private Sub Mnu_InsUser_Click()
  HFTextBoxWithFocus.SelText = "{USER}"
End Sub

Private Sub Mnu_UserDefined_Click(Index As Integer)
  Dim i As Long, s As String
  
  s = Mnu_UserDefined(Index).Caption
  For i = 1 To UBound(mReportWiz.CtrlCodesDesc)
    If s = mReportWiz.CtrlCodesDesc(i) Then
      HFTextBoxWithFocus.SelText = mReportWiz.CtrlCodes(i)
      Exit For
    End If
  Next i
End Sub

Private Sub Opt_Orient_Click(Index As Integer)
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Friend Sub SetButtons()
  'RK Added in override for AbacusReporter implementations
  If mReportWiz.AR And ssTab.Tab = GetScreenTab(SCREEN_FILE_GROUPS) Then
    #If AbacusReporter Then
      Cmd_Next.Enabled = CheckOneFileGroupSelected(Me)
      Cmd_Clear.Enabled = (mReportFields.Count > 0)
      Cmd_Preview.Enabled = (mReportFields.Count > 0) And CheckOneFileGroupSelected(Me)
      Cmd_Print.Enabled = (mReportFields.Count > 0) And CheckOneFileGroupSelected(Me)
      Cmd_Export.Enabled = (mReportFields.Count > 0) And CheckOneFileGroupSelected(Me)
      Cmd_SaveReport.Enabled = mReportWiz.FileGroupContainer.Dirty
    #End If
  Else
    Cmd_Next.Enabled = (mReportFields.Count > 0)
    Cmd_Clear.Enabled = (mReportFields.Count > 0)
    Cmd_Preview.Enabled = (mReportFields.Count > 0)
    Cmd_Print.Enabled = (mReportFields.Count > 0)
    Cmd_Export.Enabled = (mReportFields.Count > 0)
    Cmd_SaveReport.Enabled = (mReportFields.Count > 0)
  End If
End Sub

Public Sub FillFieldFlG()
  Dim Fld As ReportField, i As Long
  
  With FlG_Fields
    .Redraw = False
    If .Cols - 1 <> mReportFields.Count Then
      .Cols = mReportFields.Count + 1
    End If
    For Each Fld In mReportFields
      i = Fld.Order
      .ColAlignment(i) = 0
      .TextMatrix(ROW_NAME, i) = Fld.Description 'cad p11d
      .TextMatrix(ROW_DATASET, i) = Fld.DataSet.Name
      .TextMatrix(ROW_DATATYPE, i) = DataTypeName(Fld.DataType)
      .TextMatrix(ROW_GROUPBY, i) = YesNo(Fld.Group)
      .TextMatrix(ROW_HIDE, i) = YesNo(Not Fld.Hide)
      If mReportDetails.AutoWidth Then
        .TextMatrix(ROW_WIDTH_TYPE, i) = "Auto"
      Else
        .TextMatrix(ROW_WIDTH_TYPE, i) = CStr(Fld.Width) & "%"
      End If
      If (Fld.Group And Fld.GroupHeader) Or Fld.Hide Then
        .TextMatrix(ROW_WIDTH_TYPE, i) = "N/A"
      End If
      .TextMatrix(ROW_KEYSTRING, i) = Fld.KeyString
    Next Fld
    .Redraw = True
    .Refresh
  End With

End Sub

Public Function FillFormatDetails(ByVal CurrentScreen As REPORT_SCREEN) As Boolean
  Dim rFld As ReportField, NewSelection As Boolean
  Dim i As Long, j As Long
  Dim s As String
  
  On Error Resume Next
  
  mReportWiz.Status = ""
    
  Call FillFieldFlG
  InFillFormat = True
  
  Set rFld = GetSelectedField()
  If Not mCurrentField Is rFld Then NewSelection = True
  Set mCurrentField = rFld
  
  If CurrentScreen = SCREEN_CRITERIA Then
    If NewSelection Then
      If mCurrentField.DataType = TYPE_STR Then
        Chk_FldCritCase.Visible = True
      Else
        Chk_FldCritCase.Visible = False
      End If
    End If
  ElseIf CurrentScreen = SCREEN_FORMATS Then
    If mCurrentField.Group Then
      Chk_Group.Value = vbChecked
      Call FrameEnable(Me, Fra_Group, True)
    Else
      Chk_Group.Value = vbUnchecked
      Call FrameEnable(Me, Fra_Group, False)
    End If
    
    If Not mCurrentField.Hide Then
      Chk_Display.Value = vbChecked
      Fra_FmtDDetail.Enabled = True
      Call FrameEnable(Me, Fra_FmtDDetail, True)
      Fra_Group.Enabled = True
      Chk_Group.Enabled = True
      Lbl_Sorting.Enabled = True
      Cbo_Sorting.Enabled = True
    Else
      Chk_Display.Value = vbUnchecked
      Fra_FmtDDetail.Enabled = False
      Call FrameEnable(Me, Fra_FmtDDetail, False)
      Fra_Group.Enabled = False
      Call FrameEnable(Me, Fra_Group, False)
    End If
         
    If mCurrentField.Wrap Then
      Chk_Wrap.Value = vbChecked
    Else
      Chk_Wrap.Value = vbUnchecked
    End If
         
    If mCurrentField.NoSquash Then
      Chk_NoSquash.Value = vbChecked
    Else
      Chk_NoSquash.Value = vbUnchecked
    End If
  
    If mCurrentField.GroupHeader Then
      Chk_GroupOpt(0).Value = vbChecked
    Else
      Chk_GroupOpt(0).Value = vbUnchecked
    End If
        
    If mCurrentField.GroupPageBreak Then
      Chk_GroupOpt(1).Value = vbChecked
    Else
      Chk_GroupOpt(1).Value = vbUnchecked
    End If
      
    If mCurrentField.GroupTotal Then
      Chk_GroupOpt(2).Value = vbChecked
    Else
      Chk_GroupOpt(2).Value = vbUnchecked
    End If
        
    Cbo_DataType.Text = DataTypeName(mCurrentField.DataType)
    Lbl_DataType1.Caption = DataTypeName(mCurrentField.DataType)
    Cbo_Sorting.Text = SortingName(mCurrentField.Sort)
    Cbo_Sum.Text = SumName(mCurrentField.SumType)
    Cbo_Alignment.Text = AlignmentName(mCurrentField.Alignment)
    Cbo_BooleanTrue.Text = mCurrentField.BooleanTrue
    Cbo_BooleanFalse.Text = mCurrentField.BooleanFalse
  
'    Cbo_FieldName.Clear
'    Cbo_FieldName.AddItem mCurrentField.Name
'    Cbo_FieldName.AddItem ""
    s = mCurrentField.FieldName
    If Len(s) = 0 Then
      s = mCurrentField.Description
    End If
    Cbo_FieldName.Text = s
  
    Cbo_Sum.Enabled = False
    lbl_Sum.Enabled = False
    If mCurrentField.DataType <> TYPE_STR Then
      Chk_Trunc.Visible = False
      Txt_Trunc.Visible = False
      Lbl_Trunc.Visible = False
    End If
    If mCurrentField.DataType <> TYPE_BOOL Then
      Lbl_BooleanTrue.Visible = False
      Cbo_BooleanTrue.Visible = False
      Lbl_BooleanFalse.Visible = False
      Cbo_BooleanFalse.Visible = False
    End If
    Select Case mCurrentField.DataType
      Case TYPE_STR
        Cbo_Format.Visible = False
        lbl_Format.Visible = False
        Lbl_Fmt.Visible = False
        Lbl_FmtSample.Visible = False
        If mCurrentField.TextWidth > 0 Then
          Chk_Trunc.Value = vbChecked
          Txt_Trunc.Text = CStr(mCurrentField.TextWidth)
          Txt_Trunc.Enabled = True
          Lbl_Trunc.Enabled = True
        Else
          Chk_Trunc.Value = vbUnchecked
          Txt_Trunc.Text = ""
          Txt_Trunc.Enabled = False
          Lbl_Trunc.Enabled = False
        End If
        Chk_Trunc.Visible = True
        Txt_Trunc.Visible = True
        Lbl_Trunc.Visible = True
      Case TYPE_LONG, TYPE_DOUBLE, TYPE_DATE
        Cbo_Sum.Enabled = (mCurrentField.DataType <> TYPE_DATE)
        lbl_Sum.Enabled = (mCurrentField.DataType <> TYPE_DATE)
        Cbo_Format.Visible = True
        lbl_Format.Visible = True
        Lbl_Fmt.Visible = True
        Lbl_FmtSample.Visible = True
      Case TYPE_BOOL
        Cbo_Format.Visible = False
        lbl_Format.Visible = False
        Lbl_Fmt.Visible = False
        Lbl_FmtSample.Visible = False
        Cbo_BooleanTrue.Text = mCurrentField.BooleanTrue
        Lbl_BooleanTrue.Visible = True
        Cbo_BooleanTrue.Visible = True
        Cbo_BooleanFalse.Text = mCurrentField.BooleanFalse
        Lbl_BooleanFalse.Visible = True
        Cbo_BooleanFalse.Visible = True
    End Select
    Txt_Suffix.Text = mCurrentField.Suffix
    Txt_Prefix.Text = mCurrentField.Prefix
    
    Cbo_Format.Text = mCurrentField.Format
    Lbl_FmtSample.Caption = Format$(lbl_Format.Tag)
    lbl_Format.Caption = Format$(lbl_Format.Tag, mCurrentField.Format)
    
    'Txt_Width.Text = CStr(CDbl(Val(mCurrentField.Width)))
    'If mReportDetails.AutoWidth Or mCurrentField.Hide Then
      'Lbl_Width.Visible = False
      'Txt_Width.Visible = False
    'Else
      'Lbl_Width.Visible = True
      'Txt_Width.Visible = True
    'End If
    
    Txt_FW.Text = CStr(CDbl(Val(mCurrentField.FixedWidth)))
  End If
  FillFormatDetails = NewSelection
  InFillFormat = False
End Function

Public Sub ReadFormatDetails()
  Dim i As Long, j As Long
  Dim s As String
  
  If mCurrentField Is Nothing Then
    Call ECASE("ReadFormatDetails - no current field")
    Set mCurrentField = GetSelectedField
  End If
  
  If Chk_Display = vbChecked Then
    If mCurrentField.Hide Then
      mCurrentField.Hide = False
      'Call mReportWiz.SetEqualFieldWidths
      'Txt_Width.Text = CStr(CDbl(Val(mCurrentField.Width)))
    End If
  Else
    If Not mCurrentField.Hide Then
      mCurrentField.Hide = True
      'Call mReportWiz.SetEqualFieldWidths
      'Txt_Width.Text = CStr(CDbl(Val(mCurrentField.Width)))
    End If
  End If
    
  mCurrentField.Wrap = (Chk_Wrap = vbChecked)
  mCurrentField.NoSquash = (Chk_NoSquash = vbChecked)
  
  mCurrentField.Group = (Chk_Group = vbChecked)
  mCurrentField.GroupHeader = mCurrentField.Group And (Chk_GroupOpt(0) = vbChecked)
  mCurrentField.GroupPageBreak = mCurrentField.Group And (Chk_GroupOpt(1) = vbChecked)
  mCurrentField.GroupTotal = mCurrentField.Group And (Chk_GroupOpt(2) = vbChecked)

  mCurrentField.DataType = Cbo_DataType.ItemData(Cbo_DataType.ListIndex)
  mCurrentField.Sort = Cbo_Sorting.ItemData(Cbo_Sorting.ListIndex)
  mCurrentField.Alignment = Cbo_Alignment.ItemData(Cbo_Alignment.ListIndex)
  mCurrentField.SumType = Cbo_Sum.ItemData(Cbo_Sum.ListIndex)
  
  mCurrentField.BooleanTrue = Cbo_BooleanTrue.Text
  mCurrentField.BooleanFalse = Cbo_BooleanFalse.Text

  If Cbo_Format.Text = "(none)" Then
    mCurrentField.Format = ""
  Else
    mCurrentField.Format = Cbo_Format.Text
  End If
  mCurrentField.Suffix = Txt_Suffix.Text
  mCurrentField.Prefix = Txt_Prefix.Text

  mCurrentField.FieldName = Trim$(Cbo_FieldName.Text)

  If Chk_Trunc.Value = vbChecked Then
    mCurrentField.TextWidth = isLongEx(Txt_Trunc.Text, DEFAULT_TRUNC_WIDTH)
    If mCurrentField.TextWidth <= 0 Then mCurrentField.TextWidth = DEFAULT_TRUNC_WIDTH
  Else
    mCurrentField.TextWidth = -1
  End If
  mCurrentField.FixedWidth = CLng(Val(Txt_FW.Text))
  If mCurrentField.FixedWidth <= 0 Then
    mCurrentField.FixedWidth = 1
  End If
End Sub

Public Function GetSelectedField() As ReportField
  Dim rFld As ReportField, fCol As Long
  
  On Error GoTo GetSelectedField_Err
  fCol = FlG_Fields.col
  For Each rFld In mReportFields
    If rFld.Order = fCol Then
      Set GetSelectedField = rFld
      Exit For
    End If
  Next rFld
  If GetSelectedField Is Nothing Then Call ECASE("GetSelectedField: Selected field could not be found in field list")
  
GetSelectedField_End:
  Exit Function
  
GetSelectedField_Err:
  Resume GetSelectedField_End
End Function

Public Sub PrepareControls()
  Dim i As Long, j As Long
  Dim FlGWidth As Long
  Dim RSel As Long, RCount As Long
  Dim DTypeName As String
  
  FlG_Fields.AllowUserResizing = flexResizeBoth
  FlG_Fields.WordWrap = True
  
  FlG_Fields.Redraw = False
  FlG_Fields.Rows = ROW_FIRST_CRITERIA + mMaxCriteria
  FlGWidth = FlG_Fields.Width
  FlG_Fields.TextMatrix(0, 0) = ""
  FlG_Fields.TextMatrix(ROW_NAME, 0) = "FIELD"
  FlG_Fields.TextMatrix(ROW_DATASET, 0) = "DATA SET"
  FlG_Fields.TextMatrix(ROW_DATATYPE, 0) = "DATA TYPE"
  FlG_Fields.TextMatrix(ROW_GROUPBY, 0) = "GROUPING"
  FlG_Fields.TextMatrix(ROW_HIDE, 0) = "DISPLAY"
  FlG_Fields.TextMatrix(ROW_WIDTH_TYPE, 0) = "WIDTH"
  FlG_Fields.TextMatrix(ROW_KEYSTRING, 0) = "(KEY)"
  FlG_Fields.ColWidth(0) = FlGWidth * 0.15
  For i = 4 To FlG_Fields.Rows - 1
    FlG_Fields.RowHeight(i) = 0
  Next i
  FlG_Fields.Redraw = True
  FlG_Fields.Refresh

  If mReportWiz.AllowDataTypeChange Then
    Lbl_DataType1.Visible = False
    Cbo_DataType.Visible = True
  Else
    Cbo_DataType.Visible = False
    Lbl_DataType1.Visible = True
  End If
  Cbo_DataType.Clear
  For i = TYPE_STR To TYPE_BOOL
    Cbo_DataType.AddItem DataTypeName(i), i - 1
    Cbo_DataType.ItemData(Cbo_DataType.NewIndex) = i
  Next i
  Lbl_Sorting.Visible = True
  Cbo_Sorting.Visible = True
  Cbo_Sorting.Clear
  For i = SORT_NONE To SORT_DESCENDING
    Cbo_Sorting.AddItem SortingName(i)
    Cbo_Sorting.ItemData(Cbo_Sorting.NewIndex) = i
  Next i
  Cbo_Alignment.Clear
  For i = ALIGN_LEFT To ALIGN_CENTER
    Cbo_Alignment.AddItem AlignmentName(i)
    Cbo_Alignment.ItemData(Cbo_Alignment.NewIndex) = i
  Next i
  Cbo_Sum.Clear
  For i = TYPE_NOSUM To TYPE_MEAN
    Cbo_Sum.AddItem SumName(i)
    Cbo_Sum.ItemData(Cbo_Sum.NewIndex) = i
  Next i
    
  Cbo_BooleanTrue.Clear
  Cbo_BooleanTrue.AddItem "True"
  Cbo_BooleanTrue.AddItem "Yes"
  Cbo_BooleanTrue.AddItem "False"
  Cbo_BooleanTrue.AddItem "No"
  Cbo_BooleanFalse.Clear
  Cbo_BooleanFalse.AddItem "False"
  Cbo_BooleanFalse.AddItem "No"
  Cbo_BooleanFalse.AddItem "True"
  Cbo_BooleanFalse.AddItem "Yes"
  
  Cbo_GHDelimiter.Clear
  Cbo_GHDelimiter.AddItem ";  "
  Cbo_GHDelimiter.AddItem ",  "
  Cbo_GHDelimiter.AddItem "   "
  
  Cbo_GHSeparator.Clear
  Cbo_GHSeparator.AddItem " = "
  Cbo_GHSeparator.AddItem ": "
  Cbo_GHSeparator.AddItem ""
  
  If Not IsEmpty(mReportWiz.CtrlCodes) Then
    Mnu_UserDefined(0).Caption = mReportWiz.CtrlCodesDesc(1)
    Mnu_UserDefined(0).Visible = True
    For i = 2 To UBound(mReportWiz.CtrlCodes)
      Load Mnu_UserDefined(i - 1)
      Mnu_UserDefined(i - 1).Visible = True
      Mnu_UserDefined(i - 1).Caption = mReportWiz.CtrlCodesDesc(i)
    Next i
  End If
  
  txtPreview = Trim$(CStr(DEFAULT_PREVIEW_LINES))
  
  'RK 26/10/04
  Call HideTabs
  #If AbacusReporter Then
    If mReportWiz.AR Then
      fb.Style = FileBrowser
      Call fb.AddFileExtension("Abacus files (*.abc)", "*.abc", True, True)
    End If
  #End If
End Sub

Public Function ChangeScreen(ToScreen As String, Optional FromScreen As String)
  Dim i As Long, j As Long
  'LockWindowUpdate Me.hWnd
  mReportWiz.Status = ""
  Select Case ToScreen
    #If AbacusReporter Then
      Case "Fra_ARFileGroups"
        'RK 25/10/04
        'Fra_Fields.ZOrder 0
        Call EnableFileGroupMembers(Me, False)
        ssTab.Tab = GetScreenTab(SCREEN_FILE_GROUPS)
        Me.Caption = mReportWiz.Title & "File Selection"
        Cmd_Back.Enabled = False
        Fra_ARFileGroups.Visible = True
        fraFileGroups.Visible = True
        ssTabReporterVersions.Visible = True
        ssTabReporterVersions.ForeColor = Me.BackColor  '&H8000000F
        If mReportWiz.FileGroupContainer.Loaded Then
          Call LoadAvailablePacks(mReportWiz, Me)
          Call LoadAvailableFileGroups(mReportWiz, Me, mReportDetails)
          Call ApplyFileGroupSelection(mReportWiz, Me, mReportDetails)
          Call SetButtons
        End If
    #End If
    Case "Fra_Fields"
      'RK 25/10/04
      'Fra_Fields.ZOrder 0
      ssTab.Tab = GetScreenTab(SCREEN_SELECTION)
      Me.Caption = mReportWiz.Title & "Field Selection"
      Cmd_Back.Enabled = mReportWiz.AR
      Cmd_LoadReport.Enabled = True
      If mReportFields.Count = 0 Then
        Cmd_Preview.Enabled = False
        Cmd_Print.Enabled = False
        Cmd_Export.Enabled = False
        Cmd_SaveReport.Enabled = False
      Else
        Cmd_Preview.Enabled = True
        Cmd_Print.Enabled = True
        Cmd_Export.Enabled = True
        Cmd_SaveReport.Enabled = True
      End If
      Call SetButtons
    Case "Fra_Format"
      FlG_Fields.Redraw = False
      'RK 25/10/04
      'Fra_Format.ZOrder 0
      ssTab.Tab = GetScreenTab(SCREEN_FORMATS)
      Fra_FldCriteria.Visible = False
      Fra_FmtDDetail.Visible = True
      Chk_Display.Visible = True
      Fra_Group.Visible = True
      Me.Caption = mReportWiz.Title & "Apply Field Formats"
      Lbl_CriteriaLbl.Visible = False
      Lbl_FormatInst.Visible = True
      Cmd_Back.Enabled = True
      Cmd_LoadReport.Enabled = False
      Cmd_Next.Enabled = True
      For i = ROW_DATATYPE To ROW_WIDTH_TYPE
        FlG_Fields.RowHeight(i) = FlG_Fields.RowHeight(1)
      Next i
      For i = 0 To mMaxCriteria + 1
        FlG_Fields.RowHeight(ROW_FIRST_CRITERIA + i - 1) = 0
      Next i
      For i = ROW_GROUPBY To ROW_WIDTH_TYPE
        FlG_Fields.RowHeight(i) = FlG_Fields.RowHeight(1)
      Next i
      FlG_Fields.FixedRows = 1
      FlG_Fields.AllowBigSelection = True
      FlG_Fields.SelectionMode = flexSelectionByColumn
      Call FillFormatDetails(SCREEN_FORMATS)
      FlG_Fields.Redraw = True
      FlG_Fields.Refresh
      Call FlG_Fields_SelChange
    Case "Fra_Criteria"
      Fra_FldCriteria.Enabled = True
      FlG_Fields.Redraw = True
      ToScreen = "Fra_Format"
      'RK 25/10/04
      'Fra_Format.ZOrder 0
      ssTab.Tab = GetScreenTab(SCREEN_CRITERIA)
      Fra_FmtDDetail.Visible = False
      Chk_Display.Visible = False
      Fra_Group.Visible = False
      Fra_FldCriteria.Visible = True
      Me.Caption = mReportWiz.Title & "Apply Selection Criteria"
      Lbl_FormatInst.Visible = False
      Lbl_CriteriaLbl.Visible = True
      Cmd_Back.Enabled = True
      Cmd_LoadReport.Enabled = False
      Cmd_Next.Enabled = True
      FlG_Fields.FixedRows = 4
      FlG_Fields.AllowBigSelection = False
      FlG_Fields.SelectionMode = flexSelectionFree
      Call FillFormatDetails(SCREEN_CRITERIA)
      Call DisplayCriteria
      For i = ROW_GROUPBY To (ROW_FIRST_CRITERIA - 1)
        FlG_Fields.RowHeight(i) = 0
      Next i
      For i = 0 To mMaxCriteria
        FlG_Fields.RowHeight(ROW_FIRST_CRITERIA + i) = FlG_Fields.RowHeight(1)
      Next i
      FlG_Fields.Redraw = True
      FlG_Fields.Refresh
      Call FlG_Fields_SelChange
    Case "Fra_Report"
      'RK 25/10/04
      'Fra_Report.ZOrder 0
      ssTab.Tab = GetScreenTab(SCREEN_REPORT)
      Me.Caption = mReportWiz.Title & "Design Report Appearance"
      Cmd_Back.Enabled = True
      Cmd_LoadReport.Enabled = False
      Cmd_Next.Enabled = False
  End Select
  mReportWiz.CurrentFrame = ToScreen
  'LockWindowUpdate 0
End Function

Public Sub ActionCurrentScreen(ByVal InPrint As Boolean)
  Fra_Fields.Enabled = Not InPrint
  Fra_FldCriteria.Enabled = Not InPrint
  Fra_Format.Enabled = Not InPrint
  Fra_Report.Enabled = Not InPrint
End Sub

Public Function SetFont(FontObj As FontDetails, DialogBoxName As String, Optional PromptIfFontOrSizeChanged As Boolean = False, Optional WasPrompted As Boolean) As Boolean
  Dim s As String
  
  With CD_Fonts
    .FontBold = FontObj.Bold
    .FontItalic = FontObj.Italic
    .FontName = FontObj.Name
    .FontSize = FontObj.Size
    .FontStrikethru = FontObj.Strikethru
    .FontUnderline = FontObj.Underline
    .Color = FontObj.ForeColor
    .Flags = cdlCFEffects Or cdlCFBoth 'Or cdlCFTTOnly 'Or cdlCFANSIOnly
    .CancelError = True
    .DialogTitle = DialogBoxName
    On Error Resume Next
    .ShowFont
    If Err.Number = cdlCancel Then Exit Function
    If PromptIfFontOrSizeChanged Then
      If (.FontName <> FontObj.Name) Or (.FontSize <> FontObj.Size) Then
        WasPrompted = True
        's = "You have changed the font or font size for this data column from:" & vbCrLf & _
            mReportDetails.DataFont.Name & ", Size " & Trim$(CStr(mReportDetails.DataFont.Size)) & _
            "  to  " & _
            .FontName & ", Size " & Trim$(CStr(.FontSize)) & vbCrLf
        s = "You have changed the font or font size for this data column." & vbCrLf
        s = s & vbCrLf
        s = s & "To keep the change, the font / font size of all the other data columns will be set to the same value." '& vbCrLf & vbCrLf & _
        '     "Press OK to make font / font size the same for all columns, or Cancel to cancel the changes."
        'cad
        If DisplayMessage(Me, s, "Font Name or Size Changed", "Ok", "Cancel") Then
          mReportDetails.DataFont.Name = .FontName
          mReportDetails.DataFont.Size = .FontSize
        Else
          Exit Function
        End If
        .FontName = FontObj.Name
        .FontSize = FontObj.Size
      End If
    End If
    FontObj.Bold = .FontBold
    FontObj.Italic = .FontItalic
    FontObj.Name = .FontName
    FontObj.Size = .FontSize
    FontObj.Strikethru = .FontStrikethru
    FontObj.Underline = .FontUnderline
    FontObj.ForeColor = .Color
  End With
  SetFont = True
End Function



Public Sub TrV_Fields_NodeClick(ByVal node As MSComctlLib.node)
  Dim SelAction As SELECTITEM_ACTION
  Dim dSet As ReportDataSet
  Dim IsItemSelected As Boolean
        
  mReportWiz.Status = ""
  IsItemSelected = (node.Image = SELECTED_NODE Or node.Image = SELECTED_PARENT)
  If node.Children = 0 And Not node.Bold Then
    SelAction = mReportWiz.SelectItem(node.Key, IsItemSelected)
    If SelAction = SELECTITEM_ERROR Then
      Beep
      mReportWiz.Status = "You cannot add the selected field to your report"
      Exit Sub
    End If
      
    If IsItemSelected Then
      node.Image = NOTSELECTED_NODE
    Else
      node.Image = SELECTED_NODE
    End If
    Call UpdateFieldOrders(mReportFields)
    Call RefreshList
    Call SetButtons
  End If
End Sub

Private Sub Txt_FldCritValue_Change()
  Me.Cmd_FldCritAdd.Enabled = True 'Len(Trim$(Me.Txt_FldCritValue.Text)) > 0
End Sub

Private Sub Txt_FW_LostFocus()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub TB_SetFocus(txt As TextBox)
  On Error Resume Next
  Debug.Print "TextBox: " & txt.Name & "(" & txt.Index & ")"
  Set HFTextBoxWithFocus = txt
  Call HFTextBoxWithFocus.SetFocus
End Sub

Private Sub Txt_PageFoot_LostFocus(Index As Integer)
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Txt_PageHead_LostFocus(Index As Integer)
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Txt_PrevLines_Change()
  txtPreview.Text = Txt_PrevLines.Text
End Sub

Private Sub Txt_RepHead_LostFocus(Index As Integer)
  Static Inchange As Boolean
  On Error Resume Next
  
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub Txt_RepHead_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim txt As TextBox
  
  Set txt = Txt_RepHead(Index)
  Call TB_SetFocus(txt)
  If Button = vbRightButton Then
    Call LockWindowUpdate(txt.hwnd)
    Frm_RepWiz.SetFocus
    txt.Enabled = False
    Call ShowPopUpHFMenu
    txt.Enabled = True
    Call LockWindowUpdate(0)
  End If
End Sub

Private Sub Txt_PageHead_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim txt As TextBox
  
  Set txt = Me.Txt_PageHead(Index)
  Call TB_SetFocus(txt)
  If Button = vbRightButton Then
    Call LockWindowUpdate(txt.hwnd)
    Frm_RepWiz.SetFocus
    txt.Enabled = False
    Call ShowPopUpHFMenu
    txt.Enabled = True
    Call LockWindowUpdate(0)
  End If
End Sub

Private Sub Txt_PageFoot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim txt As TextBox
  
  Set txt = Txt_PageFoot(Index)
  Call TB_SetFocus(txt)
  If Button = vbRightButton Then
    Call LockWindowUpdate(txt.hwnd)
    Frm_RepWiz.SetFocus
    txt.Enabled = False
    Call ShowPopUpHFMenu
    txt.Enabled = True
    Call LockWindowUpdate(0)
  End If
End Sub

Private Sub Txt_Prefix_Change()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Txt_PrevLines_LostFocus()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillRepFormat Then
    Inchange = True
    Call ReadReportDetails
    Call FillReportDetails
    Inchange = False
  End If
End Sub

Private Sub ShowPopUpHFMenu()
  Dim s As String
  
  On Error GoTo ShowPopUpHFMenu_Err
  Select Case HFTextBoxWithFocus
    Case Txt_RepHead(0)
      s = mReportDetails.RepHeaderFontL.FontDescription
    Case Txt_RepHead(1)
      s = mReportDetails.RepHeaderFontC.FontDescription
    Case Txt_RepHead(2)
      s = mReportDetails.RepHeaderFontR.FontDescription
    Case Txt_PageHead(0)
      s = mReportDetails.PageHeaderFontL.FontDescription
    Case Txt_PageHead(1)
      s = mReportDetails.PageHeaderFontC.FontDescription
    Case Txt_PageHead(2)
      s = mReportDetails.PageHeaderFontR.FontDescription
    Case Txt_PageFoot(0)
      s = mReportDetails.PageFooterFontL.FontDescription
    Case Txt_PageFoot(1)
      s = mReportDetails.PageFooterFontC.FontDescription
    Case Txt_PageFoot(2)
      s = mReportDetails.PageFooterFontR.FontDescription
    Case Else
  End Select
  'Mnu_Font.Caption = "Font: " & s
  Mnu_Font.Visible = False
  Mnu_ChangeFont.Caption = "Change font from: " & s
  Call PopupMenu(Mnu_RepWiz, vbPopupMenuLeftAlign, HFTextBoxWithFocus.left + HFTextBoxWithFocus.Width + 150)
ShowPopUpHFMenu_End:
  Exit Sub
  
ShowPopUpHFMenu_Err:
  Resume ShowPopUpHFMenu_End
End Sub

Private Sub Txt_Suffix_Change()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Private Sub Txt_Trunc_Change()
  Static Inchange As Boolean
  On Error Resume Next
  If Not Inchange And Not InFillFormat Then
    Inchange = True
    Call ReadFormatDetails
    Call FillFormatDetails(SCREEN_FORMATS)
    Inchange = False
  End If
End Sub

Public Sub ReadReportDetails()
  
  With mReportDetails
         
    .LeftJoinLeaves = (Chk_IncBlankLines.Value = vbChecked)
    .IncludeHeaders = (Chk_InclColHeaders.Value = vbChecked)
    .SummaryReport = (Chk_CollapseAll.Value = vbChecked)
    .AlignHeaders = (Chk_AlignHeadings.Value = vbChecked)
    .FitToPage = (Chk_FitToPage.Value = vbChecked)
    .ReportHeaderOnAllPages = (Chk_RHonAllPages.Value = vbChecked)
    .PrintRecordCount = (Chk_DisplayRecordCount.Value = vbChecked)
    .HideGroupHeaderTypes = (Chk_GroupHeaders.Value = vbChecked)
    .GroupHeaderDelimiter = Cbo_GHDelimiter.Text
    .GroupHeaderSeparator = Cbo_GHSeparator.Text
  
    .TrimHeadings = (Chk_TrimHeadings.Value = vbChecked)
  
    .AutoWidth = (Chk_AutoWidth.Value = vbChecked)
    
    .PreviewLines = Max(1, CLng(Val(Txt_PrevLines.Text)))
    
    .IncludeRepHeader = (Chk_RepHead.Value = vbChecked)
    .IncludePageHeader = (Chk_PageHead.Value = vbChecked)
    .IncludePageFooter = (Chk_PageFoot.Value = vbChecked)
    
    If Opt_Orient(0).Value = True Then
      .Orientation = PORTRAIT
    Else
      .Orientation = LANDSCAPE
    End If

    .RepHeaderL = Txt_RepHead(0).Text
    .RepHeaderC = Txt_RepHead(1).Text
    .RepHeaderR = Txt_RepHead(2).Text
    .PageHeaderL = Txt_PageHead(0).Text
    .PageHeaderC = Txt_PageHead(1).Text
    .PageHeaderR = Txt_PageHead(2).Text
    .PageFooterL = Txt_PageFoot(0).Text
    .PageFooterC = Txt_PageFoot(1).Text
    .PageFooterR = Txt_PageFoot(2).Text

  End With
End Sub

Public Sub FillReportDetails()
  Dim i As Long

  InFillRepFormat = True

  With mReportDetails
         
    Chk_IncBlankLines.Value = -.LeftJoinLeaves
    Chk_InclColHeaders.Value = -.IncludeHeaders
    Chk_CollapseAll.Value = -.SummaryReport
    Chk_AlignHeadings.Value = -.AlignHeaders
    Chk_FitToPage.Value = -.FitToPage
    Chk_RHonAllPages.Value = -.ReportHeaderOnAllPages
    Chk_DisplayRecordCount.Value = -.PrintRecordCount
    Chk_GroupHeaders.Value = -.HideGroupHeaderTypes
    Cbo_GHDelimiter.Text = .GroupHeaderDelimiter
    Cbo_GHSeparator.Text = .GroupHeaderSeparator
    
    Chk_TrimHeadings.Value = -.TrimHeadings
    
    Chk_AutoWidth.Value = -.AutoWidth
    Txt_PrevLines.Enabled = .AutoWidth
    
    Txt_PrevLines = CStr(.PreviewLines)
     
    If .Orientation = PORTRAIT Then
      Opt_Orient(0).Value = True
    Else
      Opt_Orient(1).Value = True
    End If
        
    Chk_RepHead.Value = -.IncludeRepHeader
    Chk_PageHead.Value = -.IncludePageHeader
    Chk_PageFoot.Value = -.IncludePageFooter
    For i = 0 To 2
      Txt_RepHead(i).Enabled = .IncludeRepHeader
      Txt_PageHead(i).Enabled = .IncludePageHeader
      Txt_PageFoot(i).Enabled = .IncludePageFooter
    Next i
    
    Txt_RepHead(0).Text = .RepHeaderL
    Txt_RepHead(1).Text = .RepHeaderC
    Txt_RepHead(2).Text = .RepHeaderR
    Txt_PageHead(0).Text = .PageHeaderL
    Txt_PageHead(1).Text = .PageHeaderC
    Txt_PageHead(2).Text = .PageHeaderR
    Txt_PageFoot(0).Text = .PageFooterL
    Txt_PageFoot(1).Text = .PageFooterC
    Txt_PageFoot(2).Text = .PageFooterR
    
  End With
  
  InFillRepFormat = False

End Sub

Public Sub FillCritOpCbo(ByVal DataType As DATABASE_FIELD_TYPES)
  Cbo_FldCritOp.Clear
  Select Case DataType
    Case TYPE_STR
      Cbo_FldCritOp.AddItem "the same as" '"="
      Cbo_FldCritOp.AddItem "different to" '"<>"
      Cbo_FldCritOp.AddItem "alphabetically before" '"<"
      Cbo_FldCritOp.AddItem "alphabetically before or the same as" '"<="
      Cbo_FldCritOp.AddItem "alphabetically after" '">"
      Cbo_FldCritOp.AddItem "alphabetically after or the same as" '">="
      Cbo_FldCritOp.AddItem "alphabetically between" '"> & <"
      Cbo_FldCritOp.AddItem "alphabetically between or the same as" '">= & <="
      Cbo_FldCritOp.AddItem "like" '"like"
    Case TYPE_LONG
      Cbo_FldCritOp.AddItem "equal to" '"="
      Cbo_FldCritOp.AddItem "not equal to" '"<>"
      Cbo_FldCritOp.AddItem "less than" '">"
      Cbo_FldCritOp.AddItem "less than or equal to" '">="
      Cbo_FldCritOp.AddItem "greater than" '"<"
      Cbo_FldCritOp.AddItem "greater than or equal to" '"<="
      Cbo_FldCritOp.AddItem "between" '"> & <"
      Cbo_FldCritOp.AddItem "between or equal to" '">= & <="
    Case TYPE_DOUBLE
      Cbo_FldCritOp.AddItem "equal to" '"="
      Cbo_FldCritOp.AddItem "not equal to" '"<>"
      Cbo_FldCritOp.AddItem "less than" '">"
      Cbo_FldCritOp.AddItem "less than or equal to" '">="
      Cbo_FldCritOp.AddItem "greater than" '"<"
      Cbo_FldCritOp.AddItem "greater than or equal to" '"<="
      Cbo_FldCritOp.AddItem "between" '"> & <"
      Cbo_FldCritOp.AddItem "between or equal to" '">= & <="
    Case TYPE_DATE
      Cbo_FldCritOp.AddItem "the same as" '"="
      Cbo_FldCritOp.AddItem "different to" '"<>"
      Cbo_FldCritOp.AddItem "before" '">"
      Cbo_FldCritOp.AddItem "before or the same as" '">="
      Cbo_FldCritOp.AddItem "after" '"<"
      Cbo_FldCritOp.AddItem "after or the same as" '"<="
      Cbo_FldCritOp.AddItem "between" '"> & <"
      Cbo_FldCritOp.AddItem "between or the same as" '">= & <="
      Cbo_FldCritOp.AddItem "like"
      Cbo_FldCritOp.AddItem "undated"
      Cbo_FldCritOp.AddItem "dated"
    Case TYPE_BOOL
      Cbo_FldCritOp.AddItem "equal to" '"="
      Cbo_FldCritOp.AddItem "not equal to" '"<>"
    Case Else
  End Select
  Cbo_FldCritOp.Text = Cbo_FldCritOp.List(0)
End Sub

Public Sub FillFormatDefaults(ByVal DataType As DATABASE_FIELD_TYPES)
  With Cbo_Format
  .Clear
  .AddItem "(none)"
  Select Case DataType
    Case TYPE_LONG
      .AddItem "#"
      .AddItem "#0"
      .AddItem "#,###;(#,###)"
      .AddItem "#,##0;(#,##0)"
      lbl_Format.Tag = "-9836"
    Case TYPE_DOUBLE
      .AddItem "#.##"
      .AddItem "#.00"
      .AddItem "#,###.##;(#,###.##)"
      .AddItem "#,##0.00;(#,##0.00)"
      .AddItem "#,##0;(#,##0)"
      lbl_Format.Tag = "-9836.003"
    Case TYPE_DATE
      .AddItem "DD/MM/YY"
      .AddItem "DD/MM/YYYY"
      .AddItem "DD MMM YYYY"
      .AddItem "DDDD DD MMMM YYYY"
      lbl_Format.Tag = "8 May 1990"
    Case Else
  End Select
  .Text = ""
  End With
  Call FillFormatDetails(SCREEN_FORMATS)
End Sub

Public Sub DisplayCriteria()
  Dim rFld As ReportField, Crit As Criterion
  Dim i As Long, j As Long, Index As Long
    
  On Error GoTo DisplayCriteria_Err
  mMaxCriteria = CompactCriteria(mReportFields)
  FlG_Fields.Rows = ROW_FIRST_CRITERIA + mMaxCriteria + 1
  For i = ROW_FIRST_CRITERIA To (ROW_FIRST_CRITERIA + mMaxCriteria)
    FlG_Fields.TextMatrix(i, 0) = "CRITERIA #" & Trim$(CStr(i - ROW_FIRST_CRITERIA + 1))
    Index = i - ROW_FIRST_CRITERIA + 1
    For j = 1 To mReportFields.Count
      Set rFld = mReportFields.Item(j)
      FlG_Fields.TextMatrix(i, rFld.Order) = ""
      If Index <= rFld.Criteria.Count Then
        Set Crit = rFld.Criteria.Item(Index)
        If Not Crit Is Nothing Then
          FlG_Fields.TextMatrix(i, rFld.Order) = Crit.AsString
        End If
      End If
    Next j
  Next i
  Call FlG_Fields_SelChange
DisplayCriteria_End:
  Exit Sub
  
DisplayCriteria_Err:
  Resume DisplayCriteria_End
End Sub


Private Sub RefreshList()
  Dim nod As node, rFld As ReportField
  
  lstFields.Clear
  For Each rFld In mReportFields
    Set nod = TrV_Fields.Nodes(rFld.KeyString)
    Call lstFields.AddItem(rFld.Description)  'cad p11d
    lstFields.ItemData(lstFields.NewIndex) = nod.Index
  Next rFld
End Sub


Private Sub txtPreview_Change()
  Dim s As String
  
  s = Trim$(CStr(Max(0, isLongEx(txtPreview.Text, DEFAULT_PREVIEW_LINES))))
  If txtPreview.Text <> s Then txtPreview.Text = s
  Txt_PrevLines.Text = s
End Sub

Private Sub HSc_FGLeft_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  
  If Me.FlG_Fields.LeftCol >= 2 Then
    Me.FlG_Fields.LeftCol = Me.FlG_Fields.LeftCol - 1
  End If
End Sub

Private Sub HSc_FGRight_DragOver(Source As Control, x As Single, y As Single, State As Integer)

  If Me.FlG_Fields.LeftCol <= Me.FlG_Fields.Cols Then
    Me.FlG_Fields.LeftCol = Me.FlG_Fields.LeftCol + 1
  End If
End Sub

'
'Private Sub Cbo_Include1_Click(Index As Integer)
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillRepFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Cbo_Include2_Click(Index As Integer)
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillRepFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Cbo_Include3_Click(Index As Integer)
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillRepFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Chk_Include_Click(Index As Integer)
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillRepFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Chk_PageNos_Click()
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Opt_Incl1_Click(Index As Integer)
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillRepFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Opt_Incl2_Click(Index As Integer)
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillRepFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Opt_Incl3_Click(Index As Integer)
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillRepFormat Then
'    Inchange = True
'    Call ReadReportDetails
'    Call FillReportDetails
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Text1_Change()
'
'End Sub
'
'Private Sub Txt_BooleanFalse_LostFocus()
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillFormat Then
'    Inchange = True
'    Call ReadFormatDetails
'    Call FillFormatDetails(SCREEN_FORMATS)
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Txt_BooleanTrue_LostFocus()
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillFormat Then
'    Inchange = True
'    Call ReadFormatDetails
'    Call FillFormatDetails(SCREEN_FORMATS)
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Txt_DP_LostFocus()
'  Static Inchange As Boolean
'  On Error Resume Next
'  If Not Inchange And Not InFillFormat Then
'    Inchange = True
'    Call ReadFormatDetails
'    Call FillFormatDetails(SCREEN_FORMATS)
'    Inchange = False
'  End If
'End Sub
'
'Private Sub Txt_Width_LostFocus()
'  Dim Width As Double, Fld As ReportField, WidthReduce As Double
'  Dim TotalWidth As Double, WidthToRight As Double, WidthToLeft As Double
'
'  'apf Width = ConvertStrToDataType(Txt_Width.Text, FIELD_DATATYPE_DOUBLE)
'  Call FieldWidthLimits(mReportFields, FieldSelected.Order, TotalWidth, WidthToRight)
'  WidthToLeft = TotalWidth - WidthToRight - mCurrentField.Width
'  If Width < 0 Then Width = 0
'  If WidthToLeft + Width > 100 Then
'    Width = 100 - WidthToLeft
'  End If
'  If TotalWidth - mCurrentField.Width + Width > 100 Then
'    WidthReduce = (100 - WidthToLeft - Width) / WidthToRight
'    For Each Fld In mReportFields
'      If Fld.Order > mCurrentField.Order Then
'        Fld.Width = PercentDP(Val(Fld.Width * WidthReduce), 1)
'      End If
'    Next Fld
'  End If
'  mCurrentField.Width = PercentDP(Val(Width), 1)
'  Call FillFormatDetails(SCREEN_FORMATS)
'End Sub
'
Private Sub IFolderBrowserControlEvents_Ended(ByVal id As Variant)
  On Error GoTo Err_Err
  
  #If AbacusReporter Then
    'Apply Default pack selection of all to new/unspecified folders
    If fb.Style = FolderBrowser Then
      If Not CheckOnePackSelected(Me) Then
        Call SetAllPackSelection(Me, SELECTED_NODE, True)
        chkAllPacks.Value = vbChecked
      End If
    End If
    
    Call SaveFileGroupMember(mReportWiz, Me, Me.ActiveListItem_FileGroupMember)
  #End If
Err_End:
  Exit Sub
Err_Err:
  Call ErrorMessage(Err.Number, Err, "fb_Ended", "Error with FileBrowser", Err.Description)
  Resume Err_End
  Resume
End Sub

Private Sub IFolderBrowserControlEvents_Started(ByVal id As Variant)
  On Error GoTo Err_Err
  #If AbacusReporter Then
    Call DisablePackSelection(Me)
  #End If
Err_End:
  Exit Sub
Err_Err:
  Call ErrorMessage(Err.Number, Err, "fb_Ended", "Error with FileBrowser", Err.Description)
  Resume Err_End
  Resume
End Sub

Private Sub IFolderBrowserControlEvents_Validate(ByVal id As Variant, Cancel As Boolean)

End Sub

