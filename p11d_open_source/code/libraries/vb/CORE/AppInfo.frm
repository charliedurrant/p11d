VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application Information"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvStatics 
      CausesValidation=   0   'False
      Height          =   2745
      Left            =   90
      TabIndex        =   15
      Top             =   3255
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4842
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Application Information"
      Height          =   2745
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.Label lblSysInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   1560
         TabIndex        =   17
         Top             =   2280
         Width           =   3330
      End
      Begin VB.Label lblCaptions 
         Caption         =   "File Version"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1643
         Width           =   1365
      End
      Begin VB.Label lblSysInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   1560
         TabIndex        =   13
         Top             =   1960
         Width           =   3330
      End
      Begin VB.Label lblCaptions 
         Caption         =   "Command line"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1009
         Width           =   1365
      End
      Begin VB.Label lblSysInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   1560
         TabIndex        =   11
         Top             =   1643
         Width           =   3330
      End
      Begin VB.Label lblCaptions 
         Caption         =   "Size"
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   1960
         Width           =   1365
      End
      Begin VB.Label lblSysInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   1560
         TabIndex        =   9
         Top             =   1326
         Width           =   3330
      End
      Begin VB.Label lblCaptions 
         Caption         =   "Time and Date"
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label lblCaptions 
         Caption         =   "Name"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   375
         Width           =   1365
      End
      Begin VB.Label lblCaptions 
         Caption         =   "Location"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   692
         Width           =   1365
      End
      Begin VB.Label lblSysInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   375
         Width           =   3330
      End
      Begin VB.Label lblSysInfo 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   692
         Width           =   3330
      End
      Begin VB.Label lblCaptions 
         Caption         =   "Program Version"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1326
         Width           =   1365
      End
      Begin VB.Label lblSysInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Top             =   1009
         Width           =   3330
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   6030
      Width           =   1260
   End
   Begin VB.Label lbl1 
      Caption         =   "Application Globals"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2925
      Width           =   3255
   End
End
Attribute VB_Name = "frmAppInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call Me.lvStatics.Refresh
End Sub

