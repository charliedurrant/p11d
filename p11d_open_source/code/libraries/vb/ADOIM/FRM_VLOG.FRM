VERSION 5.00
Begin VB.Form Frm_VLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Log File"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TCSIMP.FWCtrl FW_ViewSpec 
      Height          =   1812
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   7092
      _ExtentX        =   12515
      _ExtentY        =   3201
   End
   Begin TCSIMP.FWCtrl FW_ViewLog 
      Height          =   1812
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   7092
      _ExtentX        =   12515
      _ExtentY        =   3201
   End
   Begin VB.Frame Fra_View 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   7332
      Begin VB.CommandButton Cmd_ClearFile 
         Caption         =   "Clear File"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_ToggleEdit 
         Caption         =   "Toggle Edit"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Lbl_Edit 
         Alignment       =   1  'Right Justify
         Height          =   372
         Left            =   5880
         TabIndex        =   6
         Top             =   0
         Width           =   1452
      End
   End
   Begin VB.CommandButton Cmd_ViewSpec 
      Caption         =   "View Spec File"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_ViewLog 
      Caption         =   "View Log File"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Finish 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "Frm_VLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard
Implements IImportForm

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_VLOG
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cmd_ClearFile_Click()
  Dim FreeFil As Long

  FreeFil = FreeFile
  Open AppPath & "\IMPWIZ.LOG" For Output As #FreeFil
  Print #FreeFil, "Import Wizard Log File"
  Print #FreeFil, Time$; " on "; Date$
  Close #FreeFil
  
  m_ImpWiz.ReCalc_VLog

End Sub

Private Sub Cmd_Finish_Click()
  Call SwitchForm(Me, TCSIMP_DEST)
End Sub

Private Sub Cmd_ToggleEdit_Click()

'RTF_File.SaveFile AppPath & "\IMPWIZ.LOG", rtfText
If Lbl_Edit = "Edit ON" Then
  Lbl_Edit = "Edit OFF"
  Cmd_ClearFile.Enabled = False
Else
  Lbl_Edit = "Edit ON"
  Cmd_ClearFile.Enabled = True
End If

End Sub

Private Sub Cmd_ViewLog_Click()
  Call m_ImpWiz.ReCalc_VLog
End Sub

Private Sub Cmd_ViewSpec_Click()
  Call m_ImpWiz.ReCalc_VLog
End Sub
