VERSION 5.00
Begin VB.Form frmEnvir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviroment"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLocalInfo 
      Caption         =   "Locale Information"
      Height          =   1905
      Left            =   75
      TabIndex        =   24
      Top             =   4185
      Width           =   5415
      Begin VB.Label lblLocalUser 
         Height          =   1545
         Left            =   2970
         TabIndex        =   26
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label lblLocaleSys 
         Height          =   1545
         Left            =   90
         TabIndex        =   25
         Top             =   270
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4230
      TabIndex        =   8
      Top             =   6120
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Caption         =   "Memory Information"
      Height          =   1515
      Left            =   75
      TabIndex        =   5
      Top             =   1185
      Width           =   5415
      Begin VB.Label lblSysInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   8
         Left            =   3855
         TabIndex        =   19
         Top             =   1125
         Width           =   1410
      End
      Begin VB.Label lblInformation 
         Caption         =   "Overall memory usage"
         Height          =   270
         Index           =   10
         Left            =   150
         TabIndex        =   18
         Top             =   1125
         Width           =   2520
      End
      Begin VB.Label lblSysInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3855
         TabIndex        =   13
         Top             =   750
         Width           =   1410
      End
      Begin VB.Label lblSysInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3855
         TabIndex        =   12
         Top             =   375
         Width           =   1410
      End
      Begin VB.Label lblInformation 
         Caption         =   "Free physical memory available"
         Height          =   270
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   750
         Width           =   2520
      End
      Begin VB.Label lblInformation 
         Caption         =   "Total  physical memory available"
         Height          =   270
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   375
         Width           =   2520
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Information"
      Height          =   1065
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   5415
      Begin VB.Label lblSysInfo 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   675
         Width           =   3105
      End
      Begin VB.Label lblSysInfo 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2160
         TabIndex        =   10
         Top             =   300
         Width           =   3105
      End
      Begin VB.Label lblInformation 
         Caption         =   "Windows version"
         Height          =   270
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   675
         Width           =   1965
      End
      Begin VB.Label lblInformation 
         Caption         =   "Processor type"
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   1935
      End
   End
   Begin VB.Frame fmeSystem 
      Caption         =   "Disk Space Information"
      Height          =   1395
      Left            =   75
      TabIndex        =   0
      Top             =   2745
      Width           =   5415
      Begin VB.Label lblStatic 
         Caption         =   "Disk Free"
         Height          =   240
         Index           =   3
         Left            =   3105
         TabIndex        =   23
         Top             =   975
         Width           =   720
      End
      Begin VB.Label lblStatic 
         Caption         =   "Disk Free"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   22
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblStatic 
         Caption         =   "Disk Size"
         Height          =   225
         Index           =   1
         Left            =   3105
         TabIndex        =   21
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblStatic 
         Caption         =   "Disk Size"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lblSysInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   3840
         TabIndex        =   17
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label lblSysInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   1305
         TabIndex        =   16
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label lblSysInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3840
         TabIndex        =   15
         Top             =   615
         Width           =   1410
      End
      Begin VB.Label lblSysInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1305
         TabIndex        =   14
         Top             =   630
         Width           =   1410
      End
      Begin VB.Label lblInformation 
         Caption         =   "Current drive"
         Height          =   270
         Index           =   6
         Left            =   3105
         TabIndex        =   2
         Top             =   270
         Width           =   2025
      End
      Begin VB.Label lblInformation 
         Caption         =   "Application drive"
         Height          =   270
         Index           =   5
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmEnvir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub lblLocalInfoUser_Click()

End Sub
