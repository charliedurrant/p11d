VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "atc2stat.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   ClientHeight    =   6960
   ClientLeft      =   1350
   ClientTop       =   1125
   ClientWidth     =   8880
   Icon            =   "Mdimain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerVersionCheck 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1200
      Top             =   4800
   End
   Begin MSComctlLib.ImageList imlListViewBenefits 
      Left            =   2400
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":030A
            Key             =   "Employers"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":065C
            Key             =   "Employees"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":09AE
            Key             =   "Other"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":0F00
            Key             =   "Vouchers"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":1452
            Key             =   "Accommodation"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":17A4
            Key             =   "EmployeeCar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":1CF6
            Key             =   "CompanyCar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":2248
            Key             =   "Loan"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":259A
            Key             =   "Medical"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":28EC
            Key             =   "Relocation"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":2C3E
            Key             =   "Phone"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":3190
            Key             =   "SharedVan"
         EndProperty
      EndProperty
   End
   Begin atc2stat.TCSStatus sts 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   6630
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   582
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
   Begin MSComctlLib.Toolbar tbrBenefits 
      Align           =   4  'Align Right
      Height          =   5640
      Left            =   8310
      TabIndex        =   1
      Top             =   990
      Visible         =   0   'False
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   9948
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "imgBenefits"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Object.ToolTipText     =   "All benefits"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Company car benefits"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employee owned car benefits"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Telephone benefits"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Medical benefits"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Credit card benefits"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Accommodation / relocation benefits"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Loan benefits"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Other benefits"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Employer"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit Employer Details"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh Employers"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Confirm changes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Throw away changes"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Print reports"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Preview report"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "View employer screen"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Display employer van pool"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employees"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrNavigate 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      Begin VB.CheckBox chkMoveToNextEmployeeWithBenefit 
         Caption         =   "Move to next with benefit"
         Height          =   240
         Left            =   8010
         TabIndex        =   12
         Top             =   90
         Width           =   3000
      End
      Begin VB.TextBox txtEmployeeOfTotal 
         Height          =   285
         Left            =   4095
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   75
         Width           =   1635
      End
      Begin VB.CommandButton cmdUp 
         Height          =   285
         Left            =   7200
         MaskColor       =   &H8000000F&
         Picture         =   "Mdimain.frx":34E2
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   735
      End
      Begin VB.CommandButton cmdGoto 
         Caption         =   "Goto"
         Height          =   285
         Left            =   5715
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   75
         Width           =   735
      End
      Begin VB.CommandButton cmdDown 
         Height          =   285
         Left            =   6465
         MaskColor       =   &H8000000F&
         Picture         =   "Mdimain.frx":3A10
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Width           =   735
      End
      Begin VB.TextBox txtReference 
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   1305
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   2805
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList imgBenefits 
      Left            =   4680
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":3F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":4B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":53E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":5C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":6486
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":70D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":792A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":857C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":91CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgToolbar 
      Left            =   960
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":9A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":9B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":9C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":9D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":9EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":A412
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":A6A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":A936
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":AAA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":AC0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":AD74
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":AEDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":B048
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":B1B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":B31C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTree 
      Left            =   2760
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":B47A
            Key             =   "PRINT_CROSS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":B7CC
            Key             =   "PRINT_TICK"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":BB1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":BC30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":BD42
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":BE54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":BF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":C078
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":C18A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":C29C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdimain.frx":C3AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   180
      Top             =   3015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Create Employer"
         Index           =   0
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Open Employer"
         Index           =   1
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Edit Employer Details"
         Index           =   2
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Delete Employer"
         Index           =   3
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Print Report"
         Index           =   4
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Import"
         Index           =   6
         Begin VB.Menu mnuFileImportBegin 
            Caption         =   "&Begin Import"
         End
         Begin VB.Menu mnuFileImportTracking 
            Caption         =   "&Tracking"
         End
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "Error &Logs"
         Index           =   7
         Begin VB.Menu mnuFileErrorLogsImport 
            Caption         =   "&Import"
         End
         Begin VB.Menu mnuFileErrorLogsMagneticMedia 
            Caption         =   "&Magnetic media"
         End
         Begin VB.Menu mnuFileErrorLogsPayeOnline 
            Caption         =   "&PAYE Online"
         End
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "Electronic &Submission"
         Index           =   8
         Begin VB.Menu mnuMagneticMedia 
            Caption         =   "&Magnetic media"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPAYEOnline 
            Caption         =   "&PAYE Online"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuIntranet 
            Caption         =   "&Intranet"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Refresh Employers"
         Index           =   9
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "Change Directory"
         Index           =   10
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "Employee &Letter"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Employer List"
         Index           =   12
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "Pass&word"
         Index           =   13
         Begin VB.Menu mnuFilePasswordSet 
            Caption         =   "&Set"
         End
         Begin VB.Menu mnuFilePasswordClear 
            Caption         =   "&Clear"
         End
      End
      Begin VB.Menu mnuFileItems 
         Caption         =   "&Tools"
         Index           =   14
         Begin VB.Menu mnuFileToolsFindRepairCompactEmployer 
            Caption         =   "&Repair/compact employer"
         End
         Begin VB.Menu mnuFileToolsFindFiles 
            Caption         =   "&Find Files"
         End
         Begin VB.Menu mnuFileToolsZipEmployer 
            Caption         =   "&Zip employer file"
         End
      End
      Begin VB.Menu mnuFileResetSettings 
         Caption         =   "&Reset Settings"
      End
      Begin VB.Menu mnuFileBringForwardBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBringForward 
         Caption         =   "&Bring forward"
      End
      Begin VB.Menu mnuExitBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEmployer 
      Caption         =   "E&mployer"
      Begin VB.Menu mnuEmployerCDB 
         Caption         =   "Company defined &benefits"
      End
      Begin VB.Menu mnuEmployerCDC 
         Caption         =   "&Company defined categories"
      End
      Begin VB.Menu mnuEmployerMileageSchemes 
         Caption         =   "Com&pany mileage schemes"
      End
      Begin VB.Menu mnuEmployerSharedVans 
         Caption         =   "&Shared vans"
      End
      Begin VB.Menu mnuSeper0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployerTransferEmployees 
         Caption         =   "&Transfer employees"
      End
   End
   Begin VB.Menu mnuEmployee 
      Caption         =   "&Employee"
      Begin VB.Menu mnuEmployeeItems 
         Caption         =   "&Confirm changes to employee"
         Index           =   0
      End
      Begin VB.Menu mnuEmployeeItems 
         Caption         =   "&Undo changes to employee"
         Index           =   1
      End
      Begin VB.Menu mnuEmployeeItems 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEmployeeItems 
         Caption         =   "&Employee Details Page"
         Index           =   3
      End
      Begin VB.Menu mnuEmployeeItems 
         Caption         =   "&Add Employee"
         Index           =   4
      End
      Begin VB.Menu mnuEmployeeItems 
         Caption         =   "&Delete Employee"
         Index           =   5
      End
      Begin VB.Menu mnuEmployeeItems 
         Caption         =   "&Goto"
         Index           =   6
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmployeeNameOrder 
         Caption         =   "&Name order"
      End
      Begin VB.Menu mnuEmployeeSortEmployeeReferenceAsNumber 
         Caption         =   "&Sort employee reference as number"
      End
      Begin VB.Menu mnuEmployeeValidateOnscreenNINumber 
         Caption         =   "&Validate onscreen NI number"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSortByName 
         Caption         =   "Sort by &Name"
      End
      Begin VB.Menu mnuViewSortByPersonnelNumber 
         Caption         =   "Sort by &Personal reference"
      End
      Begin VB.Menu mnuViewSortByNINumber 
         Caption         =   "Sort by N&I number"
      End
      Begin VB.Menu mnuViewSortByStatus 
         Caption         =   "Sort by &Status"
      End
      Begin VB.Menu mnuViewGroup 
         Caption         =   "Sort by &Group"
         Begin VB.Menu mnuViewGroupSortByGroup1 
            Caption         =   "&Group 1"
         End
         Begin VB.Menu mnuViewGroupSortByGroup2 
            Caption         =   "&Group 2"
         End
         Begin VB.Menu mnuViewGroupSortByGroup3 
            Caption         =   "&Group 3"
         End
      End
      Begin VB.Menu mnuselect 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelect 
         Caption         =   "&Select"
         Begin VB.Menu mnuViewSelectAll 
            Caption         =   "&Select all"
         End
         Begin VB.Menu mnuViewSelectUnselectAll 
            Caption         =   "&Unselect all"
         End
         Begin VB.Menu mnuViewSelectReverse 
            Caption         =   "&Reverse selection"
         End
         Begin VB.Menu mnuViewSelectSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewSelectGroup1 
            Caption         =   "Group &1"
         End
         Begin VB.Menu mnuViewSelectGroup2 
            Caption         =   "Group &2"
         End
         Begin VB.Menu mnuViewSelectGroup3 
            Caption         =   "Group &3"
         End
         Begin VB.Menu mnuViewSelectSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewSelectHasEmail 
            Caption         =   "&Has email address"
         End
         Begin VB.Menu mnuViewSelectBlankEmail 
            Caption         =   "&Blank email address"
         End
         Begin VB.Menu mnuViewSelectSep5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewSelectLeftEmployees 
            Caption         =   "&Left employees"
         End
         Begin VB.Menu mnuViewSelectCurrentEmployees 
            Caption         =   "&Current employees"
         End
         Begin VB.Menu mnuViewSelectSep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewSelectEmployeeAlphabetically 
            Caption         =   "&Select &alphabetically by surname"
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&A - C"
               Index           =   1
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&D - F"
               Index           =   2
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&G - I"
               Index           =   3
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&J - L"
               Index           =   4
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&M - O"
               Index           =   5
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&P - R"
               Index           =   6
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&S - U"
               Index           =   7
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&V - X"
               Index           =   8
            End
            Begin VB.Menu mnuViewSelectEmployeeAlphabeticallyLetter 
               Caption         =   "&Y - Z + other"
               Index           =   9
            End
         End
         Begin VB.Menu mnuViewSelectSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewSelectEmployeeByReport 
            Caption         =   "&Select by &report"
         End
      End
   End
   Begin VB.Menu mnuBenefits 
      Caption         =   "&Benefits"
      Begin VB.Menu mnuBenfitsAdd 
         Caption         =   "&Add - Ctrl+Insert"
      End
      Begin VB.Menu mnuBenfitsDelete 
         Caption         =   "&Delete - Ctrl+Delete"
      End
      Begin VB.Menu mnuBenefitsCopy 
         Caption         =   "&Copy - Ctrl+Shift+C"
      End
      Begin VB.Menu mnuBenfitsPaste 
         Caption         =   "&Paste - Ctrl+Shift+V"
      End
      Begin VB.Menu mnuBenefitsCut 
         Caption         =   "C&ut - Ctrl+Shift+X"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBenefitsTools 
         Caption         =   "&Tools"
         Begin VB.Menu mnuBenefitsToolsDataChecker 
            Caption         =   "&Data checker"
         End
         Begin VB.Menu mnuBenefitsToolsFindCompanyCar 
            Caption         =   "&Find Company Car"
         End
         Begin VB.Menu mnuBenefitsToolsAbacusExport 
            Caption         =   "&Abacus Export"
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAllBens 
         Caption         =   "All Benefit&s"
      End
      Begin VB.Menu mnuAllBenSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssetsTransferred 
         Caption         =   "&A Assets transferred"
      End
      Begin VB.Menu mnuPayments 
         Caption         =   "&B Payments made on behalf of the employee"
         Begin VB.Menu mnuPaymentOnBehalf 
            Caption         =   "&Payments on behalf of the employee"
         End
         Begin VB.Menu mnuTaxOnNotionalPayments 
            Caption         =   "&Tax on notional payents"
         End
      End
      Begin VB.Menu mnuVouchersAndCredit 
         Caption         =   "&C Vouchers and credit cards"
      End
      Begin VB.Menu mnuAccomodation 
         Caption         =   "&D Living Accommodation"
      End
      Begin VB.Menu mnuEmployeeCars 
         Caption         =   "&E Mileage allowance"
      End
      Begin VB.Menu mnuCompanyCars 
         Caption         =   "&F Cars and car fuel"
      End
      Begin VB.Menu mnuVans 
         Caption         =   "&G Vans"
      End
      Begin VB.Menu mnuLoans 
         Caption         =   "&H Beneficial loans"
      End
      Begin VB.Menu mnuMedical 
         Caption         =   "&I Private medical treatment or insurance"
      End
      Begin VB.Menu mnuRelocation 
         Caption         =   "&J Qualifying relocation expenses"
      End
      Begin VB.Menu mnuServicesProvided 
         Caption         =   "&K Services provided"
      End
      Begin VB.Menu mnuAssetsAtDisposal 
         Caption         =   "&L Assets placed at employee's disposal"
      End
      Begin VB.Menu mnuOther 
         Caption         =   "&M Other items"
         Begin VB.Menu mnuSubscriptions 
            Caption         =   "Subscriptions and professional fees"
         End
         Begin VB.Menu mnuNursery 
            Caption         =   "Nursery places"
         End
         Begin VB.Menu mnuTaxPaidNotDeducted 
            Caption         =   "Tax paid but not deducted "
         End
      End
      Begin VB.Menu mnuExpenses 
         Caption         =   "&N Expenses payments"
         Begin VB.Menu mnuTravelAndSubsistence 
            Caption         =   "Travelling and subsistence"
         End
         Begin VB.Menu mnuEntertainment 
            Caption         =   "Entertainment"
         End
         Begin VB.Menu mnuGeneralExpensesBusinessTravel 
            Caption         =   "General expenses"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHomePhones 
            Caption         =   "&Home telephone"
         End
         Begin VB.Menu mnuNonQualifyingRelocation 
            Caption         =   "Non-qualifying relocation expenses"
         End
         Begin VB.Menu mnuChauffeur 
            Caption         =   "Chauffeur Expenses"
         End
         Begin VB.Menu mnuOOther 
            Caption         =   "Other items"
         End
      End
      Begin VB.Menu mnuCreditButton 
         Caption         =   "CreditButton"
         Visible         =   0   'False
         Begin VB.Menu mnuCreditButtonCreditCards 
            Caption         =   "CreditCards"
         End
         Begin VB.Menu mnuCreditButtonSubscriptions 
            Caption         =   "Subscriptions"
         End
      End
      Begin VB.Menu mnuHouseButton 
         Caption         =   "HouseButton"
         Visible         =   0   'False
         Begin VB.Menu mnuHouseButtonAccomodation 
            Caption         =   "Accomodation"
         End
         Begin VB.Menu mnuHouseButtonQualRelocation 
            Caption         =   "QualRelocation"
         End
         Begin VB.Menu mnuHouseButtonNonQualRelocation 
            Caption         =   "NonQualRelocation"
         End
      End
      Begin VB.Menu mnuOtherButton 
         Caption         =   "OtherButton"
         Visible         =   0   'False
         Begin VB.Menu mnuOtherButtonAssetsTransferred 
            Caption         =   "AssetsTransferred"
         End
         Begin VB.Menu mnuOtherButtonPayB 
            Caption         =   "PaymentsOfBehalf"
            Begin VB.Menu mnuOtherButtonPayBPayB 
               Caption         =   "PayemtsOnBehalf"
            End
            Begin VB.Menu mnuOtherButtonPayBTax 
               Caption         =   "TaxOnNotiional"
            End
         End
         Begin VB.Menu mnuOtherButtonVans 
            Caption         =   "Vans"
         End
         Begin VB.Menu mnuOtherButtonServices 
            Caption         =   "ServicesSupplied"
         End
         Begin VB.Menu mnuOtherButtonAssetsAtDisposal 
            Caption         =   "AssetsAtDisposal"
         End
         Begin VB.Menu mnuOtherButtonOther 
            Caption         =   "Other"
            Begin VB.Menu mnuOtherButtonOtherSubscriptions 
               Caption         =   "Subscriptions"
            End
            Begin VB.Menu mnuOtherButtonOtherNursery 
               Caption         =   "Nursery"
            End
            Begin VB.Menu mnuOtherButtonOtherIncomeTax 
               Caption         =   "IncomeTax"
            End
         End
         Begin VB.Menu mnuOtherButtonPExpenses 
            Caption         =   "OExpenses"
            Begin VB.Menu mnuOtherButtonOTravel 
               Caption         =   "Travel"
            End
            Begin VB.Menu mnuOtherButtonOEntertainment 
               Caption         =   "Entertainment"
            End
            Begin VB.Menu mnuOtherButtonOGeneral 
               Caption         =   "General"
            End
            Begin VB.Menu mnuOtherButtonOPhoneHome 
               Caption         =   "Phone Home"
            End
            Begin VB.Menu mnuOtherButtonONonQual 
               Caption         =   "NonQual"
            End
            Begin VB.Menu mnuOtherButtonOChauffeur 
               Caption         =   "Chauffeur"
            End
            Begin VB.Menu mnuOtherButtonOOther 
               Caption         =   "Other"
            End
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpP11D 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpFAQs 
         Caption         =   "&KnowledgeBase"
      End
      Begin VB.Menu mnuSepX 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F11}
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents clsEvent As TCSEventClass
Attribute clsEvent.VB_VarHelpID = -1
'AUTOINSERT_BEGIN


Private Sub chkMoveToNextEmployeeWithBenefit_Click()
    p11d32.MoveToNextEmployeeWithBenefit = ChkBoxToBool(chkMoveToNextEmployeeWithBenefit)
End Sub

Private Sub clsEvent_DebugMenuItem(Name As String, Index As Long, Parent As TCSMenuItems)
  Dim ibf As IBenefitForm2
  
On Error GoTo clsEvent_DebugMenuItem_err
  Select Case Parent
    Case MNU_DATABASE
      Call OnlyFromForm(F_Employers)
      Select Case UCASE$(Name)
        Case UCASE$(S_MNU_F12_DEBUG_SQL)
          Call p11d32.DebugSQL
        Case UCASE$(S_MNU_F12_UPDATE_FIX)
          Call p11d32.UpdateFixLevel
        Case UCASE$(S_MNU_F12_SPLIT_NAMES)
          Call p11d32.SplitNamesRun
        Case UCASE$(S_MNU_F12_DELETE_ALL_CDBS)
          Call p11d32.DeleteAllCDBS
        Case "REPAIR_COMPACT"
          Call p11d32.RepairAndCompactGeneral
  
'MP RV TTP#320
'        Case UCASE$(S_MNU_F12_SET_ACTUAL_MILES)
'          Call p11d32.SetActualMiles
      End Select
    Case MNU_APPLICATION
      Select Case UCASE$(Name)
        Case UCASE$(S_MNU_F12_MAGNETIC_MEDIA_ERROR_LOGGING)
          p11d32.MagneticMedia.ErrorLogging = Not p11d32.MagneticMedia.ErrorLogging
        Case UCASE$(S_MNU_F12_MAGNETIC_MEDIA_USER_DATA_SIZE)
          F_Input.ValText.TypeOfData = VT_LONG
          F_Input.ValText.Minimum = p11d32.MagneticMedia.MinDiskFreeSpace + 1
          If F_Input.Start("Magnetic media", "Enter the min (" & p11d32.MagneticMedia.MinDiskFreeSpace & ") file size in bytes.", F_Input.ValText.Minimum) Then
            p11d32.MagneticMedia.UserDataSize = F_Input.ValText.Text
          End If
        Case UCASE$(S_MNU_F12_UPDATE_LIST_ITEM)
          If IsBenefitForm(CurrentForm) Then
            Set ibf = CurrentForm
            Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ibf.benefit)
          End If
        Case UCASE$(S_MNU_F12_SHOWEMPLOYERS_FIX_LEVEL)
          Call FixLevelShowFunction(LEF_LISTVIEW_COLUMN, Not p11d32.FixLevelsShow)
          If F_Employers Is CurrentForm Then
            Call FixLevelShowFunction(LEF_SIZE_COLUMNS, p11d32.FixLevelsShow)
            Call BenScreenSwitchEnd(F_Employers)
          End If
          
        'Case UCase$(S_MNU_F12_ENABLE_EMAIL_REPORTS)
        '  P11d32.ReportPrint.EnableEmailReports = Not P11d32.ReportPrint.EnableEmailReports
        Case UCASE$(S_MNU_F12_EMAIL_DEBUG)
          Call p11d32.ReportPrint.EmailSettingsShow
          
        Case UCASE$(S_MNU_F12_DISPLAY_INVALIDFIELDS)
          p11d32.DisplayInvalidFields = Not p11d32.DisplayInvalidFields
        Case UCASE$(S_MNU_F12_KILL_BENEFITS)
          p11d32.KillBenefits = Not p11d32.KillBenefits
        Case UCASE$(S_MNU_F12_ENTER_SERIAL_NUMBER)
          Call EnterSerialNumber
        Case UCASE$(S_MNU_F12_PAYE_ONLINE_SHOW_EXTRA_SUBMISSION_PROPERTIES_MENU)
          p11d32.PAYEonline.ExtraSubmissionPropertiesMenu = Not p11d32.PAYEonline.ExtraSubmissionPropertiesMenu
        Case UCASE$(S_MNU_F12_VIEW_PROCEED_BUTTON_IF_ERRORS)
          p11d32.PAYEonline.ViewProceedButtonIfErrors = Not p11d32.PAYEonline.ViewProceedButtonIfErrors
        Case UCASE$(S_MNU_F12_PAYE_REFERENCE_ANY_FORMAT)
          p11d32.PayeReferenceAnyFormat = Not p11d32.PayeReferenceAnyFormat
        Case UCASE$(S_MNU_F12_DATA_TYPE_LIST_VIEW_SORTING)
          p11d32.DataTypeListViewSorting = Not p11d32.DataTypeListViewSorting
        Case UCASE$(S_MNU_F12_REREAD_CONTEXT_SENSITIVE_HELP_LINKS)
          p11d32.Help.PopulateHelpLinks
        Case UCASE$(S_MNU_F12_F12_SETTINGS_FORM)
          Call F_f12_Settings.Show(1)
        Case UCASE$(S_MNU_F12_SORT_OTHER_ALPHABETICALLY)
          p11d32.ReportPrint.SortOtherTypeBenefitsAlphabetically = Not p11d32.ReportPrint.SortOtherTypeBenefitsAlphabetically
        Case UCASE$(S_MNU_F12_QA_MANAGEMENT_REPORTS)
          Call QAManagementReports
        Case UCASE$(S_MNU_F12_VERSION_CHECK)
          p11d32.VersionCheckEnabled = Not p11d32.VersionCheckEnabled
        End Select
    Case MNU_BREAK
      Call ExitApp(True)
    Case Else
      Call ECASE("Unhandled F12 menu")
  End Select
clsEvent_DebugMenuItem_end:
  Call DisplayDebugVars
  Exit Sub
clsEvent_DebugMenuItem_err:
  Call ErrorMessage(ERR_ERROR, Err, "DebugMenuItem", "ERR_DEBUGMENU", "Error processing the debug menu event " & Name & ".")
  Resume clsEvent_DebugMenuItem_end
  Resume
End Sub

Private Sub HelpLink_AppKeyDownF1()
'  Call DisplayHelp
  Call p11d32.Help.ShowHelp("")
End Sub

Private Sub HelpLink_AppKeyDownShiftCtrlF11()
  Dim sFrm As String, sCtrl As String
  Dim s As String
    
'MP CodeRv - added condition to chk lic type
  If p11d32.LicenceType = LT_DEMO Or p11d32.LicenceType = LT_UNLICENSED Then
    Call p11d32.Help.ShiftControlF11
    
  End If
End Sub
Public Sub CutCopyPasteVisible(bVisible As Boolean)
  mnuBenefitsCopy.Visible = bVisible
  mnuBenefitsCut.Visible = bVisible
  mnuBenfitsPaste.Visible = bVisible
End Sub

Private Sub MDIForm_Load()
  Set clsEvent = gEvents
  'User stuff here
  Call sts.AddPanel(20, "", Down3D, S_P1, Me.imlTree.ListImages(IMG_SELECTED_STATUS).Picture)
  Call sts.AddPanel(50, "", Down3D, S_P2, Me.imlTree.ListImages(IMG_INFO).Picture)
  
  'set up benefit mnu captions
  Call BenefitMnuCaptionShortCut(mnuAssetsTransferred, BC_ASSETSTRANSFERRED_A)
  Call BenefitMnuCaptionShortCut(mnuPayments, BC_PAYMENTS_ON_BEFALF_B)
  Call BenefitMnuCaptionShortCut(mnuPaymentOnBehalf, BC_PAYMENTS_ON_BEFALF_B, True)
  Call BenefitMnuCaptionShortCut(mnuTaxOnNotionalPayments, BC_TAX_NOTIONAL_PAYMENTS_B, True)
  Call BenefitMnuCaptionShortCut(mnuVouchersAndCredit, BC_VOUCHERS_AND_CREDITCARDS_C)
  Call BenefitMnuCaptionShortCut(mnuAccomodation, BC_LIVING_ACCOMMODATION_D)
  Call BenefitMnuCaptionShortCut(mnuEmployeeCars, BC_EMPLOYEE_CAR_E)
  Call BenefitMnuCaptionShortCut(mnuCompanyCars, BC_COMPANY_CARS_F)
  Call BenefitMnuCaptionShortCut(mnuVans, BC_nonSHAREDVAN_G)
  
  'AM
  Call BenefitMnuCaptionShortCut(mnuLoans, BC_LOAN_OTHER_H)
  
  Call BenefitMnuCaptionShortCut(mnuHomePhones, BC_PHONE_HOME_N)
  Call BenefitMnuCaptionShortCut(mnuMedical, BC_PRIVATE_MEDICAL_I)
  Call BenefitMnuCaptionShortCut(mnuRelocation, BC_QUALIFYING_RELOCATION_J)
  Call BenefitMnuCaptionShortCut(mnuServicesProvided, BC_SERVICES_PROVIDED_K)
  Call BenefitMnuCaptionShortCut(mnuAssetsAtDisposal, BC_ASSETSATDISPOSAL_L)
  
  mnuOther.Caption = "&" & p11d32.Rates.BenClassTo(BC_INCOME_TAX_PAID_NOT_DEDUCTED_M, BCT_HMIT_SECTION_STRING) & " - Other"
  Call BenefitMnuCaptionShortCut(mnuSubscriptions, BC_CLASS_1A_M, True)
  Call BenefitMnuCaptionShortCut(mnuNursery, BC_NON_CLASS_1A_M, True)
  Call BenefitMnuCaptionShortCut(mnuTaxPaidNotDeducted, BC_INCOME_TAX_PAID_NOT_DEDUCTED_M, True)
  mnuExpenses.Caption = "&" & p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_HMIT_SECTION_STRING) & S_O_EXPENSES_CAPTION
  Call BenefitMnuCaptionShortCut(mnuEntertainment, BC_ENTERTAINMENT_N, True)
  'Call BenefitMnuCaptionShortCut(mnuGeneralExpensesBusinessTravel, BC_GENERAL_EXPENSES_BUSINESS_N, True)
  Call BenefitMnuCaptionShortCut(mnuHomePhones, BC_PHONE_HOME_N, True)
  Call BenefitMnuCaptionShortCut(mnuNonQualifyingRelocation, BC_NON_QUALIFYING_RELOCATION_N, True)
  Call BenefitMnuCaptionShortCut(mnuTravelAndSubsistence, BC_TRAVEL_AND_SUBSISTENCE_N, True)
  Call BenefitMnuCaptionShortCut(mnuChauffeur, BC_CHAUFFEUR_OTHERO_N, True)
  
  Call BenefitMnuCaptionShortCut(mnuOOther, BC_OOTHER_N, True)
  
  Call BenefitMnuCaptionShortCut(mnuCreditButtonCreditCards, BC_VOUCHERS_AND_CREDITCARDS_C)
  mnuCreditButtonSubscriptions.Enabled = False  'km
  mnuCreditButtonSubscriptions.Visible = False  'km
  Call BenefitMnuCaptionShortCut(mnuHouseButtonAccomodation, BC_LIVING_ACCOMMODATION_D)
  Call BenefitMnuCaptionShortCut(mnuHouseButtonNonQualRelocation, BC_NON_QUALIFYING_RELOCATION_N)
  Call BenefitMnuCaptionShortCut(mnuHouseButtonQualRelocation, BC_QUALIFYING_RELOCATION_J)
  
  Call BenefitMnuCaptionShortCut(mnuOtherButtonAssetsTransferred, BC_ASSETSTRANSFERRED_A)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonPayB, BC_PAYMENTS_ON_BEFALF_B)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonPayBPayB, BC_PAYMENTS_ON_BEFALF_B, True)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonPayBTax, BC_TAX_NOTIONAL_PAYMENTS_B, True)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonVans, BC_NONSHAREDVANS_G)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonServices, BC_SERVICES_PROVIDED_K)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonAssetsAtDisposal, BC_ASSETSATDISPOSAL_L)
  'Call BenefitMnuCaptionShortCut(mnuOtherButtonShares, BC_SHARES_M)
  
  mnuOtherButtonOther.Caption = mnuOther.Caption
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOtherSubscriptions, BC_CLASS_1A_M, True) 'sub is now c1a
  
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOtherNursery, BC_NON_CLASS_1A_M, True)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOtherIncomeTax, BC_INCOME_TAX_PAID_NOT_DEDUCTED_M, True)
  mnuOtherButtonPExpenses.Caption = mnuExpenses.Caption
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOTravel, BC_TRAVEL_AND_SUBSISTENCE_N, True)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOEntertainment, BC_ENTERTAINMENT_N, True)
  'Call BenefitMnuCaptionShortCut(mnuOtherButtonOGeneral, BC_GENERAL_EXPENSES_BUSINESS_N, True)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonONonQual, BC_NON_QUALIFYING_RELOCATION_N, True)
  
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOPhoneHome, BC_PHONE_HOME_N, True)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOChauffeur, BC_CHAUFFEUR_OTHERO_N, True)
  Call BenefitMnuCaptionShortCut(mnuOtherButtonOOther, BC_OOTHER_N, True)
  
  'Enable intranet option if licensed
  If p11d32.LicenceType = LT_INTRANET Or p11d32.LicenceType = LT_DEMO Then
    mnuIntranet.Enabled = True
  End If
  mnuPAYEOnline.Enabled = True  ' MPS - now always enabled
  'Disable Magnetic media error logs if short version
  If p11d32.LicenceType = LT_SHORT Then
    mnuFileErrorLogsMagneticMedia.Enabled = False
  End If
  chkMoveToNextEmployeeWithBenefit.value = BoolToChkBox(p11d32.MoveToNextEmployeeWithBenefit)
  ' setup app help hook
  mnuEmployeeSortEmployeeReferenceAsNumber.Checked = p11d32.SortEmployeeReferenceAsNumber
  mnuEmployeeValidateOnscreenNINumber.Checked = p11d32.ValidateNINumberOnEmployeeScreen
  
  
  Call p11d32.VersionCheck
End Sub
Private Sub BenefitMnuCaptionShortCut(mnu As Menu, bc As BEN_CLASS, Optional bSubItem As Boolean = False)
  mnu.Caption = "&" & p11d32.Rates.BenefitMenuCaption(bc, bSubItem)
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim doExitApp As Boolean
  
  gbAllowAppExit = True
  If (UnloadMode = vbAppTaskManager) Or (UnloadMode = vbAppWindows) Then
    gbForceExit = True
  End If
  doExitApp = gbForceExit
  If Not doExitApp Then
    If Me.WindowState = vbMinimized Then Me.WindowState = vbMaximized
    doExitApp = DisplayMessage(MDIMain, "Are you sure you want to exit " & AppName & "?", AppName, "Yes", "No")
  End If
  If doExitApp Then
    doExitApp = UserAppShutDown
    If doExitApp Then Call ExitApp
  End If
  If Not doExitApp Then
    Cancel = True
    gbAllowAppExit = False
  End If
End Sub
'AUTOINSERT_END

Private Sub cmdDown_Click()
  Call MoveEmployee(True)
End Sub

Private Sub cmdGoto_Click()
  Call GotoScreen
End Sub

Private Sub cmdUp_Click()
  Call MoveEmployee(False)
End Sub

Private Sub mnuaccommodation_Click()
  Call BenScreenSwitch(BC_LIVING_ACCOMMODATION_D)
End Sub

Private Sub mnuAccomodation_Click()
  Call BenScreenSwitch(BC_LIVING_ACCOMMODATION_D)
End Sub

Private Sub mnuAllBens_Click()
  Call BenScreenSwitch(BC_ALL)
End Sub

Private Sub mnuAssets_Click(Index As Integer)
  
End Sub

Private Sub mnuAssetsAtDisposal_Click()
  Call BenScreenSwitch(BC_ASSETSATDISPOSAL_L)
End Sub
Private Sub mnuAssetsTransferred_Click()
  Call BenScreenSwitch(BC_ASSETSTRANSFERRED_A)
End Sub

Private Sub mnuBenefitsToolsAbacusExport_Click()
  Call p11d32.ReportPrint.DoExportAbacusReportWizard
End Sub

Private Sub mnuBenefitsToolsCompanyCarChecker_Click()

End Sub

Private Sub mnuBenfitsDelete_Click()
  Call ToolBarButton(TBR_REMOVE_BENEFIT, BenefitFormSelectedIndex)
End Sub

Private Sub mnuBenfitsPaste_Click()
  Call LVKeyDown(vbKeyV, vbShiftMask Or vbCtrlMask)
End Sub
Private Sub mnuBenefitsCopy_Click()
  Call LVKeyDown(vbKeyC, vbShiftMask Or vbCtrlMask, CurrentForm)
End Sub
Private Sub mnuBenefitsCut_Click()
  Call LVKeyDown(vbKeyX, vbShiftMask Or vbCtrlMask, CurrentForm)
End Sub
Private Sub mnuBenefitsToolsDataChecker_Click()
    
  If Not p11d32.CurrentEmployer Is Nothing Then
    Call p11d32.CurrentEmployer.DataChecker
  End If
End Sub
Private Sub mnuBenefitsToolsFindCompanyCar_Click()
  Call p11d32.CurrentEmployer.FindCompanyCar
End Sub

Private Sub mnuBenfitsAdd_Click()
  Call ToolBarButton(TBR_ADD_BENEFIT, GetEmployeeIndexFromSelectedEmployee)
End Sub

Private Sub mnuBringForward_Click()

End Sub

Private Sub mnuChauffeur_Click()
  Call BenScreenSwitch(BC_CHAUFFEUR_OTHERO_N)
End Sub

Private Sub mnuCompanyCars_Click()
  Call BenScreenSwitch(BC_COMPANY_CARS_F)
End Sub
Private Sub mnuCreditButton_Click()
'  If p11d32.AppYear > 2000 Then
    Call BenScreenSwitch(BC_VOUCHERS_AND_CREDITCARDS_C)
'  End If
End Sub
Private Sub mnuCreditButtonCreditCards_Click()
  Call BenScreenSwitch(BC_VOUCHERS_AND_CREDITCARDS_C)
End Sub

Private Sub mnuCreditButtonSubscriptions_Click()
  Call BenScreenSwitch(BC_CLASS_1A_M)
End Sub


Private Sub mnuEmployeeCars_Click()
  Call BenScreenSwitch(BC_EMPLOYEE_CAR_E)
End Sub

Private Sub mnuEmployeeItems_Click(Index As Integer)
  Select Case Index
    Case MNU_EMPLOYEE_CONFIRM
      Call ToolBarButton(TBR_CONFIRM, GetEmployeeIndexFromSelectedEmployee)
    Case MNU_EMPLOYEE_UNDO
      Call ToolBarButton(TBR_UNDO, GetEmployeeIndexFromSelectedEmployee)
    Case MNU_EMPLOYEE_DETAILS
      Call ToolBarButton(TBR_EMPLOYEESCREEN, GetEmployeeIndexFromSelectedEmployee)
    Case MNU_EMPLOYEE_ADD
      Call ToolBarButton(TBR_ADD_BENEFIT, GetEmployeeIndexFromSelectedEmployee)
    Case MNU_EMPLOYEE_DELETE
      Call ToolBarButton(TBR_REMOVE_BENEFIT, GetEmployeeIndexFromSelectedEmployee)
    Case MNU_EMPLOYEE_GOTO
      Call GotoScreen
    Case Else
      ECASE "Unknown menu item selected"
  End Select
End Sub

Private Sub mnuEmployeeNameOrder_Click()
  Call SetNameOrder
End Sub

Private Sub mnuEmployeeSortEmployeeReferenceAsNumber_Click()
  p11d32.SortEmployeeReferenceAsNumber = Not p11d32.SortEmployeeReferenceAsNumber
  mnuEmployeeSortEmployeeReferenceAsNumber.Checked = p11d32.SortEmployeeReferenceAsNumber
  Dim ibf As IBenefitForm2
  Set ibf = F_Employees
  If ibf.lv.SortKey = LV_EE_PERSONNEL_NUMBER Then
    Call SetEmployeesSortOrder(LV_EE_PERSONNEL_NUMBER)
  End If
End Sub

Private Sub mnuEmployeeValidateOnscreenNINumber_Click()
   p11d32.ValidateNINumberOnEmployeeScreen = Not p11d32.ValidateNINumberOnEmployeeScreen
   mnuEmployeeValidateOnscreenNINumber.Checked = p11d32.ValidateNINumberOnEmployeeScreen
   Call F_Employees.TB_Data(L_NI_NUMBER_TEXT_BOX_INDEX).lValidate
End Sub

Private Sub mnuEmployerCDB_Click()
  p11d32.CurrentEmployer.EditCompanyDefinedBenefits
End Sub

Private Sub mnuEmployerCDC_Click()
  p11d32.CurrentEmployer.EditCDC
End Sub

Private Sub mnuEmployerItems_Click(Index As Integer)
  
End Sub

Private Sub mnuEmployerMileageSchemes_Click()
  p11d32.CurrentEmployer.EditFPCS
End Sub

Private Sub mnuEmployerSharedVans_Click()
  p11d32.CurrentEmployer.EditSharedVans
End Sub

Private Sub mnuEmployerTransferEmployees_Click()
  If Not p11d32.CurrentEmployer Is Nothing Then Call p11d32.CurrentEmployer.TransferEmployees
End Sub

Private Sub mnuEntertainment_Click()
  Call BenScreenSwitch(BC_ENTERTAINMENT_N)
End Sub

Private Sub mnuExit_Click()
  Call ExitApp(True)
End Sub

Private Sub mnuFileBringForward_Click()
  Call p11d32.BringForward.Initialise
End Sub

Private Sub mnuFileErrorLogsImport_Click()
  Call p11d32.ViewErrors(VET_IMPORTING)
End Sub

Private Sub mnuFileErrorLogsMagneticMedia_Click()
  Call p11d32.ViewErrors(VET_MAGNETIC_MEDIA)
End Sub

Private Sub mnuFileErrorLogsPAYEOnline_Click()
  Call p11d32.ViewErrors(VET_PAYEONLINE_VALIDATION)
End Sub


Private Sub mnuFileImportBegin_Click()
  On Error GoTo err_err
  
  Call p11d32.Importing.InitImport
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "FileImportBegin", "File Import Begin", Err.Description)
  Resume err_end
End Sub

Private Sub mnuFileImportTracking_Click()
  On Error GoTo err_err:
  Call p11d32.Help.ShowForm(F_ImportTracking, vbModal)
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "FileImportTracking", "File Import Tracking ", "Failed to setp import tracking")
  Resume err_end
End Sub
Private Function SelectedIndex() As Long
  Dim i As Long
  
  i = -1

  If Not Me.ActiveForm Is Nothing Then
    If Not Me.ActiveForm.LB.SelectedItem Is Nothing Then
      If Me.ActiveForm.LB.SelectedItem.Selected Then
        i = CLng(Me.ActiveForm.LB.SelectedItem.Tag)
      End If
    End If
  End If

  SelectedIndex = i
  
End Function
Private Sub mnuFileItems_Click(Index As Integer)
  Dim i As Long ', cDir As String
  Dim s As String
  
  i = SelectedIndex()
  
  Select Case Index
    Case MNU_FILE_OPEN
      Call ToolBarButton(TBR_OPEN_EMPLOYER, i)
    Case MNU_FILE_EDIT
      Call ToolBarButton(TBR_EDIT_EMPLOYER, i)
    Case MNU_FILE_NEW
      Call ToolBarButton(TBR_ADD_BENEFIT, i)
    Case MNU_FILE_DELETE
      Call ToolBarButton(TBR_REMOVE_BENEFIT, i)
    Case MNU_FILE_EMPLOYER
      Call ToolBarButton(TBR_EMPLOYERSCREEN, i)
    Case MNU_FILE_REFRESHEMPLOYERS
      Call ToolBarButton(TBR_REFRESH_EMPLOYERS, i)
    Case MNU_FILE_ERROR_LOG
    Case MNU_FILE_ELECTRONIC_SUBMISSION
    Case MNU_FILE_IMPORT
    'MP DB ToDo - this never get called as MNU_FILE_IMPORT=7 (ref to Menu Editor) - not clickable
    '           - chk resulting impact as p11d32.Importing.InitImport is called from just here
      
    Case MNU_FILE_PRINT
      Call p11d32.ReportPrint.InitPrintDialog
    Case MNU_FILE_CHANGEDIRECTORY
      s = p11d32.WorkingDirectory
      Call p11d32.CreateAndSetWorkingDirectory(MDIMain, p11d32.WorkingDirectory, True)
      If (StrComp(s, p11d32.WorkingDirectory, vbTextCompare) <> 0) Then
        Call ToolBarButton(TBR_REFRESH_EMPLOYERS, i)
      End If
    'Case MNU_FILE_EMPLOYEE_LETTER RK Removed 28/02/03
      'F_EmployeeLetter.Show vbModal
    Case MNU_FILE_PASSWORD, MNU_FILE_TOOLS
      'nothing as have sub menus
    Case Else
      Call ECASE("Invalid file item menu.")
  End Select
End Sub

Private Sub mnuFilePasswordClear_Click()
  mnuFilePasswordClear.Visible = Not PassWordWrite(p11d32.CurrentEmployer, "")
End Sub

Private Sub mnuFilePasswordSet_Click()
  If Not mnuFilePasswordClear.Visible Then
    mnuFilePasswordClear.Visible = F_PassWord.PassWord(p11d32.CurrentEmployer, PWM_SET)
  Else
    Call F_PassWord.PassWord(p11d32.CurrentEmployer, PWM_SET)
  End If
  Set F_PassWord = Nothing
End Sub


Private Sub mnuFileToolsFindFiles_Click()
  Call InitFindFiles
End Sub

Private Sub mnuFileToolsFindRepairCompactEmployer_Click()
  Call RepairCompactEmployer
End Sub

Private Sub mnuFileToolsImportTracking_Click()
End Sub

Private Sub mnuFileResetSettings_Click()
  On Error GoTo err_err
  
  Call FlushIniBuffer(p11d32.IniPathAndFile)
  Call xKill(p11d32.IniPathAndFile)
  Call p11d32.IniSettings(Ini_read)
  
err_err:
  Exit Sub
err_end:
  Call ErrorMessage(ERR_ERROR, Err, "ResetSettings", "ResetSettings", Err.Description)
  Resume err_end
  
End Sub

Private Sub mnuFileToolsZipEmployer_Click()
  Dim iEmployerIndex As Long
  Dim ben As IBenefitClass
  Dim sFileName As String, s As String
  On Error GoTo err_err

  iEmployerIndex = SelectedIndex()

  If iEmployerIndex = -1 Then GoTo err_end
    
  Set ben = p11d32.Employers(iEmployerIndex)
  
  sFileName = ben.value(employer_FileName)
  Call SplitPath(sFileName, , sFileName)
  s = FileSaveAsDlg("Zip employer", "Zip files (*.zip)|*.zip", CurDir$, sFileName & ".zip")
  If Len(s) = 0 Then GoTo err_end
  If FileExists(s) Then Call xKill(s)
  Call p11d32.Importing.Zip(ben, s)

err_err:
  Exit Sub
err_end:
  Call ErrorMessage(ERR_ERROR, Err, "Zip Employer", "Zip Employer", Err.Description)
  Resume err_end
End Sub

Private Sub mnuGeneralExpensesBusinessTravel_Click()
  Call BenScreenSwitch(BC_GENERAL_EXPENSES_BUSINESS_N)
End Sub
Private Sub mnuGroupItems_Click(Index As Integer)
  Select Case Index
    Case 0
      F_Employees.LB.SortKey = 4
    Case 1
      F_Employees.LB.SortKey = 5
    Case 2
      F_Employees.LB.SortKey = 6
  End Select
End Sub
Private Sub mnuHelpAbout_Click()
  Call DisplayDebugVars
  Call AppAbout
End Sub

Private Function CreateIE() As SHDOCVW.InternetExplorer
  On Error GoTo CreateIE_ERR
  
  Set CreateIE = New SHDOCVW.InternetExplorer
  
  Exit Function
CreateIE_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CreateIE"), "This functioniality is only available with Microsoft Internet Explorer 4 or above. The FAQs are available at " & p11d32.Rates.value(WebSite))
End Function

Private Sub mnuHelpFAQs_Click()
  Dim myIE As SHDOCVW.InternetExplorer
  
  On Error GoTo mnuHelpFAQs_Click_Err
    
  Set myIE = CreateIE
  Call myIE.Navigate(p11d32.Rates.value(WebSite))
  myIE.Visible = True
  
mnuHelpFAQs_Click_End:
  Exit Sub
mnuHelpFAQs_Click_Err:
  Call ErrorMessage(ERR_ERROR, Err, "mnuHelpFAQs_Click", "Error opening Internet Explorer", Err.Description)
  Resume mnuHelpFAQs_Click_End
  Resume
End Sub

Private Sub mnuHelpP11D_Click()
'  Call DisplayHelp(True)
  Call p11d32.Help.ShowHelp(S_DEFAULT_HELP_PAGE)
End Sub

Private Sub mnuHomePhones_Click()
  Call BenScreenSwitch(BC_PHONE_HOME_N)
End Sub
Private Sub mnuPassword_Click()

End Sub

Private Sub mnuHouseButtonAccomodation_Click()
   Call BenScreenSwitch(BC_LIVING_ACCOMMODATION_D)
End Sub


Private Sub mnuHouseButtonNonQualRelocation_Click()
  Call BenScreenSwitch(BC_NON_QUALIFYING_RELOCATION_N)
End Sub

Private Sub mnuHouseButtonQualRelocation_Click()
  Call BenScreenSwitch(BC_QUALIFYING_RELOCATION_J)
End Sub

Private Sub mnuIntranet_Click()
  p11d32.Intranet.Start
End Sub

'AM To be removed
'Private Sub mnuLoanHome_Click()
''  If p11d32.AppYear <= 2001 Then          'km - home loans no longer exist after v2001
''   Call BenScreenSwitch(BC_LOAN_HOME_H)
''  Else
'    Call BenScreenSwitch(BC_LOAN_OTHER_H) 'km - direct to benefit loans screen
''  End If
'End Sub
'AM To be removed
'Private Sub mnuLoanOther_Click()
'  Call BenScreenSwitch(BC_LOAN_OTHER_H)
'End Sub

Private Sub mnuLoans_Click()
  Call BenScreenSwitch(BC_LOAN_OTHER_H)
End Sub

Private Sub mnuMagneticMedia_Click()
  Call MsgBox("HMRC are no longer allowing software developers to test magnetic media." & vbCrLf & vbCrLf & "We have withdrawn support for Magnetic Media and have not updated the software to support the new" & vbCrLf & "van fuel fields for 2008. Please register for online filing as an alternative.")
  p11d32.MagneticMedia.Start
End Sub

Private Sub mnuMedical_Click()
  Call BenScreenSwitch(BC_PRIVATE_MEDICAL_I)
End Sub

Private Sub mnuNonQualifyingRelocation_Click()
  Call BenScreenSwitch(BC_NON_QUALIFYING_RELOCATION_N)
End Sub

Private Sub mnuNursery_Click()
  Call BenScreenSwitch(BC_NON_CLASS_1A_M)
End Sub

Private Sub mnuOtherButtonAssetsAtDisposal_Click()
  Call BenScreenSwitch(BC_ASSETSATDISPOSAL_L)
End Sub

Private Sub mnuOtherButtonAssetsTransferred_Click()
  Call BenScreenSwitch(BC_ASSETSTRANSFERRED_A)
End Sub

Private Sub mnuOtherButtonOChauffeur_Click()
  Call BenScreenSwitch(BC_CHAUFFEUR_OTHERO_N)
End Sub


Private Sub mnuOtherButtonOtherIncomeTax_Click()
  Call BenScreenSwitch(BC_INCOME_TAX_PAID_NOT_DEDUCTED_M)
End Sub

Private Sub mnuOtherButtonOtherNursery_Click()
  Call BenScreenSwitch(BC_NON_CLASS_1A_M)
End Sub

Private Sub mnuOtherButtonOtherSubscriptions_Click()
  Call BenScreenSwitch(BC_CLASS_1A_M)
End Sub

Private Sub mnuOtherButtonPayBPayB_Click()
  Call BenScreenSwitch(BC_PAYMENTS_ON_BEFALF_B)
End Sub

Private Sub mnuOtherButtonPayBTax_Click()
  Call BenScreenSwitch(BC_TAX_NOTIONAL_PAYMENTS_B)
End Sub


Private Sub mnuOtherButtonOEntertainment_Click()
  Call BenScreenSwitch(BC_ENTERTAINMENT_N)
End Sub

Private Sub mnuOtherButtonOGeneral_Click()
  Call BenScreenSwitch(BC_GENERAL_EXPENSES_BUSINESS_N)
End Sub

Private Sub mnuOtherButtonOOther_Click()
  Call BenScreenSwitch(BC_OOTHER_N)
End Sub

Private Sub mnuOtherButtonOPhoneHome_Click()
  Call BenScreenSwitch(BC_PHONE_HOME_N)
End Sub

Private Sub mnuOtherButtonONonQual_Click()
  Call BenScreenSwitch(BC_NON_QUALIFYING_RELOCATION_N)
End Sub


Private Sub mnuOtherButtonOTravel_Click()
  Call BenScreenSwitch(BC_TRAVEL_AND_SUBSISTENCE_N)
End Sub

Private Sub mnuOtherButtonServices_Click()
  Call BenScreenSwitch(BC_SERVICES_PROVIDED_K)
End Sub

Private Sub mnuOtherButtonShares_Click()
  'Call BenScreenSwitch(BC_SHARES_M)
End Sub

Private Sub mnuOtherButtonVans_Click()
  Call BenScreenSwitch(BC_NONSHAREDVANS_G)
End Sub

Private Sub mnuPAYEOnline_Click()
    p11d32.PAYEonline.Start
End Sub

Private Sub mnuPaymentOnBehalf_Click()
  Call BenScreenSwitch(BC_PAYMENTS_ON_BEFALF_B)
End Sub

Private Sub mnuPaymentsTax_Click()
  Call BenScreenSwitch(BC_TAX_NOTIONAL_PAYMENTS_B)
End Sub

Private Sub mnuPhoneHome_Click()
  Call BenScreenSwitch(BC_PHONE_HOME_N)
End Sub

Private Sub mnuPhones_Click()
  Call BenScreenSwitch(BC_PHONE_HOME_N)
End Sub

Private Sub mnuOOther_Click()
  Call BenScreenSwitch(BC_OOTHER_N)
End Sub

Private Sub mnuRelocation_Click()
  Call BenScreenSwitch(BC_QUALIFYING_RELOCATION_J)
End Sub

Private Sub mnuServicesProvided_Click()
  Call BenScreenSwitch(BC_SERVICES_PROVIDED_K)
End Sub

Private Sub mnuShares_Click()
  'Call BenScreenSwitch(BC_SHARES_M)
End Sub

Private Sub mnuSubscriptions_Click()
  Call BenScreenSwitch(BC_CLASS_1A_M)
End Sub

Private Sub mnuTaxOnNotionalPayments_Click()
  Call BenScreenSwitch(BC_TAX_NOTIONAL_PAYMENTS_B)
End Sub

Private Sub mnuTaxPaidNotDeducted_Click()
  Call BenScreenSwitch(BC_INCOME_TAX_PAID_NOT_DEDUCTED_M)
End Sub

Private Sub mnuTravelAndsubsistence_Click()
  Call BenScreenSwitch(BC_TRAVEL_AND_SUBSISTENCE_N)
End Sub

Private Sub mnuTravelAndSubsistance_Click()
  Call BenScreenSwitch(BC_TRAVEL_AND_SUBSISTENCE_N)
End Sub

Private Sub mnuVans_Click()
  Call BenScreenSwitch(BC_NONSHAREDVANS_G)
End Sub

Private Sub mnuViewItems_Click(Index As Integer)
  Select Case Index
    Case 0
      F_Employees.LB.SortKey = 0
    Case 1
      F_Employees.LB.SortKey = 1
    Case 2
      F_Employees.LB.SortKey = 2
    Case 3
      F_Employees.LB.SortKey = 3
  End Select
End Sub


Private Sub mnuViewGroupSortByGroup1_Click()
  Call SetEmployeesSortOrder(LV_EE_GROUP1)
End Sub

Private Sub mnuViewGroupSortByGroup2_Click()
  Call SetEmployeesSortOrder(LV_EE_GROUP2)
End Sub

Private Sub mnuViewGroupSortByGroup3_Click()
  Call SetEmployeesSortOrder(LV_EE_GROUP3)
End Sub


Private Sub mnuViewSelectAll_Click()
  Call SelectItems(F_Employees.LB, SELECT_ALL)
End Sub

Private Sub mnuViewSelectBlankEmail_Click()
  Call SelectItems(F_Employees.LB, SELECT_NO_EMAIL)
End Sub

Private Sub mnuViewSelectCurrentEmployees_Click()
  Call SelectItems(F_Employees.LB, SELECT_CURRENT_EMPLOYED)
End Sub

Private Sub mnuViewSelectEmployeeAlphabetically_Click()
  On Error GoTo err_err
  
  
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "mnuViewSelectEmployeeAlphabetically_Click", "Select Employees by Report", Err.Description)
  Resume err_end
  
End Sub

Private Sub mnuViewSelectEmployeeAlphabeticallyLetter_Click(Index As Integer)
  Call SelectItems(F_Employees.LB, Index)
End Sub

Private Sub mnuViewSelectEmployeeByReport_Click()
  On Error GoTo err_err
  
  Call p11d32.Help.ShowForm(F_SelectEmployeesByReport, vbModal)
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "mnuViewSelectEmployeeByReport_Click", "Select Employees by Report", Err.Description)
  Resume err_end
End Sub

Private Sub mnuViewSelectGroup1_Click()
  Call F_Employees.cmdSelectByGroup_Click(0)
End Sub

Private Sub mnuViewSelectGroup2_Click()
  Call F_Employees.cmdSelectByGroup_Click(1)
End Sub

Private Sub mnuViewSelectGroup3_Click()
  Call F_Employees.cmdSelectByGroup_Click(2)
End Sub

Private Sub mnuViewSelectHasEmail_Click()
  Call SelectItems(F_Employees.LB, SELECT_EMAIL)
End Sub

Private Sub mnuViewSelectLeftEmployees_Click()
  Call SelectItems(F_Employees.LB, SELECT_LEFT)
End Sub

Private Sub mnuViewSelectReverse_Click()
  Call SelectItems(F_Employees.LB, SELECT_REVERSE)
End Sub


Private Sub mnuViewSelectUnselectAll_Click()
  Call SelectItems(F_Employees.LB, SELECT_NONE)
End Sub

Private Sub mnuViewSortByName_Click()
  Call SetEmployeesSortOrder(LV_EE_NAME)
End Sub
Private Sub SetEmployeesSortOrder(ByVal LV_EE As LV_EE_ITEMS)
  Dim ibf As IBenefitForm2
  Dim lSortOrder As Long
  
  Set ibf = F_Employees
  
  lSortOrder = F_Employees.SortOrder
  Call SetSortOrder(ibf.lv, ibf.lv.ColumnHeaders(LV_EE + 1), lSortOrder)
  F_Employees.SortOrder = lSortOrder
  
End Sub

Private Sub mnuViewSortByNINumber_Click()
  Call SetEmployeesSortOrder(LV_EE_NI_NUMBER)
End Sub

Private Sub mnuViewSortByPersonnelNumber_Click()
  Call SetEmployeesSortOrder(LV_EE_PERSONNEL_NUMBER)
End Sub

Private Sub mnuViewSortByStatus_Click()
  Call SetEmployeesSortOrder(LV_EE_STATUS)
End Sub

Private Sub mnuVouchersAndCredit_Click()
  Call BenScreenSwitch(BC_VOUCHERS_AND_CREDITCARDS_C)
End Sub

Private Sub tbrBenefits_ButtonClick(ByVal Button As MSComctlLib.Button)
  Static bInFunc As Boolean
  
  If bInFunc Then Exit Sub
  bInFunc = True
  Call BenefitToolBar(Button.Index, GetEmployeeIndexFromSelectedEmployee)
  bInFunc = False
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  
  
  Call ToolBarButton(Button.Index, BenefitFormSelectedIndex())
End Sub

Public Sub SetConfirmUndo()
  On Error GoTo SetConfirmUndo_Err
  Call xSet("SetConfirmUndo")
  
    Me.tbrMain.Buttons(TBR_CONFIRM).Enabled = True
    Me.tbrMain.Buttons(TBR_UNDO).Enabled = True
    Me.mnuEmployeeItems(MNU_EMPLOYEE_CONFIRM).Enabled = True
    Me.mnuEmployeeItems(MNU_EMPLOYEE_UNDO).Enabled = True

SetConfirmUndo_End:
  Call xReturn("SetConfirmUndo")
  Exit Sub
SetConfirmUndo_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetConfirmUndo", "Set Confirm Undo", "Error setting the confirm/undo tool bar buttons to ON.")
  Resume SetConfirmUndo_End
End Sub

Public Sub SetUndo()
  On Error GoTo SetUndo_Err
  Call xSet("SetUndo")
    
  Me.tbrMain.Buttons(8).Enabled = True
  Me.mnuEmployeeItems(MNU_EMPLOYEE_UNDO).Enabled = True
  
SetUndo_End:
  Call xReturn("SetUndo")
  Exit Sub
SetUndo_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetUndo", "Set Undo", "Error setting the undo toolbar button to ON.")
  Resume SetUndo_End
End Sub

Public Sub ClearConfirmUndo()
  On Error GoTo ClearConfirmUndo_Err
  Call xSet("ClearConfirmUndo")
  
  tbrMain.Buttons(TBR_CONFIRM).Enabled = False
  tbrMain.Buttons(TBR_UNDO).Enabled = False
  mnuEmployeeItems(MNU_EMPLOYEE_CONFIRM).Enabled = False
  mnuEmployeeItems(MNU_EMPLOYEE_UNDO).Enabled = False
  
ClearConfirmUndo_End:
  Call xReturn("ClearConfirmUndo")
  Exit Sub
ClearConfirmUndo_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearConfirmUndo", "Clear Confirm Undo", "Error setting the confirm/undo toolbar buttons to OFF.")
  Resume ClearConfirmUndo_End
End Sub

Public Sub NavigateBarUpdate(emp As IBenefitClass)
  Dim ibf As IBenefitForm2
  Dim lEmployeeListIndex As Long, lEmployeeListIndexCount As Long
  
  On Error GoTo NavigateBarUpdate_Err
  Call xSet("NavigateBarUpdate")
  
  If emp Is Nothing Then
    txtName = ""
    txtReference = ""
    txtEmployeeOfTotal = ""
  Else
    txtName = emp.value(ee_FullName)
    txtReference = emp.value(ee_PersonnelNumber_db)
    Set ibf = F_Employees
    If Not ibf.lv.SelectedItem Is Nothing Then
      lEmployeeListIndexCount = ibf.lv.listitems.Count
      lEmployeeListIndex = ibf.lv.SelectedItem.Index
      txtEmployeeOfTotal.Text = lEmployeeListIndex & " of " & lEmployeeListIndexCount
      If lEmployeeListIndexCount > 1 Then
        If lEmployeeListIndex = 1 Then
          cmdDown.Enabled = False
          If cmdUp.Enabled = False Then cmdUp.Enabled = True 'saves redraw
        ElseIf lEmployeeListIndex = lEmployeeListIndexCount Then
          cmdUp.Enabled = False
          If cmdDown.Enabled = False Then cmdDown.Enabled = True
        ElseIf lEmployeeListIndexCount > 0 Then
          If cmdDown.Enabled = False Then cmdDown.Enabled = True
          If cmdUp.Enabled = False Then cmdUp.Enabled = True
        End If
      Else
        cmdDown.Enabled = False
        cmdUp.Enabled = False
      End If
    Else
      cmdDown.Enabled = False
      cmdUp.Enabled = False
    End If
  End If
  
NavigateBarUpdate_End:
  Call xReturn("NavigateBarUpdate")
  Exit Sub
  
NavigateBarUpdate_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NavigateBarUpdate", "Navigate Bar Update", "Error updating the navigate bar.")
  Resume NavigateBarUpdate_End
End Sub

Public Sub ClearDelete()
  tbrMain.Buttons(TBR_REMOVE_BENEFIT).Enabled = False
  mnuBenfitsDelete.Enabled = False
  
End Sub
Public Sub SetDelete()
  tbrMain.Buttons(TBR_REMOVE_BENEFIT).Enabled = True
  mnuBenfitsDelete.Enabled = True
End Sub

Public Sub ClearAdd()
  tbrMain.Buttons(TBR_ADD_BENEFIT).Enabled = False
  mnuBenfitsAdd.Enabled = False
End Sub

Public Sub SetAdd()
  tbrMain.Buttons(TBR_ADD_BENEFIT).Enabled = True
  mnuBenfitsAdd.Enabled = True
End Sub

Private Sub DisplayDebugVars()
  Dim s As String
  
  s = "Application settings:" & vbCrLf
  s = s & "  MagneticMedia Error Logging = " & p11d32.MagneticMedia.ErrorLogging & vbCrLf
  s = s & "  MagneticMedia Disk Size = " & Min(p11d32.MagneticMedia.UserDataSize, p11d32.MagneticMedia.MinDiskFreeSpace) & vbCrLf
  s = s & "  Show Employer fix level = " & p11d32.FixLevelsShow & vbCrLf
  s = s & "  Display Invalid Fields = " & p11d32.DisplayInvalidFields & vbCrLf
  s = s & "  Kill Benefits on printing = " & p11d32.KillBenefits & vbCrLf
  s = s & "  " & S_MNU_F12_PAYE_ONLINE_SHOW_EXTRA_SUBMISSION_PROPERTIES_MENU & " = " & p11d32.PAYEonline.ExtraSubmissionPropertiesMenu & vbCrLf
  
  Call SetHelpAboutText(s)
End Sub

Private Sub TimerVersionCheck_Timer()
  On Error GoTo err_err
  
  If p11d32.VersionCheckFinished Then
    TimerVersionCheck.Enabled = False
     If ((p11d32.VersionCheckResult <> "1") And (p11d32.VersionCheckResult <> S_VERSION_CHECK_UNKNOWN)) Then
        Call MsgBox("A new version (" & p11d32.VersionCheckResult & ") of the software is available for download at: " & vbCrLf & vbCrLf & S_URL_DOWNLOADS & vbCrLf & vbCrLf & "Please contact our support line on " & S_TELEPHONE & " if you have any questions.", vbInformation)
      End If
  End If
  
err_end:
  Exit Sub
err_err:
  Resume err_end
End Sub
