VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm_Dest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Destination Recordset"
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
   Begin VB.Frame Fra_Update 
      Caption         =   "Type of import"
      Height          =   1572
      Left            =   4980
      TabIndex        =   9
      Top             =   60
      Width           =   2292
      Begin VB.OptionButton Opt_Update 
         Caption         =   "Update all matches"
         Height          =   192
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   2112
      End
      Begin VB.OptionButton Opt_Update 
         Caption         =   "Update only first match"
         Height          =   192
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   2112
      End
      Begin VB.OptionButton Opt_Update 
         Caption         =   "Update first match"
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2112
      End
      Begin VB.OptionButton Opt_Update 
         Caption         =   "No updating"
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2112
      End
      Begin VB.CheckBox Chk_Add 
         Caption         =   "Add records"
         Height          =   192
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   2052
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlG_Dest 
      Bindings        =   "Frm_dest.frx":0000
      Height          =   3255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5741
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin VB.ComboBox Cbo_Table 
      Height          =   288
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Cmd_Next 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Back 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Lbl_TableInst2 
      Height          =   372
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   4092
   End
   Begin VB.Label Lbl_NumRecs 
      Height          =   432
      Left            =   300
      TabIndex        =   6
      Top             =   4980
      Width           =   4992
   End
   Begin VB.Label Lbl_TabContents 
      Caption         =   "Contents of selected recordset:"
      Height          =   252
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   2592
   End
   Begin VB.Label Lbl_TableInst 
      Caption         =   "Your data will be imported into the recordset shown below."
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   4152
   End
End
Attribute VB_Name = "Frm_Dest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard
Implements IImportForm

Private Sub Chk_Add_Click()
  Dim tmp As IMPORT_UPDATE_TYPE
    
  On Error Resume Next
  If InUpdate Then Exit Sub
  tmp = GetUpdateType(Me)
  Call SetUpdateType(Me, tmp, m_ImpWiz.AllowChangeUpdateType, m_ImpWiz.CurrentDest.LockUpdate)
  m_ImpWiz.ImpParent.UpdateType = tmp
  m_ImpWiz.CurrentDest.UpdateType = m_ImpWiz.ImpParent.UpdateType
End Sub

Private Sub Cmd_SeeContents_Click()
  m_ImpWiz.DestPreviewRecLimit = 1000000000
  Set m_ImpWiz.CurrentDest = Nothing
  If Not m_ImpWiz.ReCalc_Dest Then Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub Command1_Click()
  MsgBox CStr(m_ImpWiz.ImpParent.UpdateType)
End Sub

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_DEST
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cbo_Table_Click()
  Static inTableClick As Boolean
  
  On Error Resume Next
  If inTableClick Then Exit Sub
  inTableClick = True
  If Not m_ImpWiz.CurrentDest Is Nothing Then
    If StrComp(Cbo_Table.Text, m_ImpWiz.CurrentDest.DisplayName, vbTextCompare) <> 0 Then
      If Not m_ImpWiz.ReCalc_Dest Then Call SwitchForm(Me, TCSIMP_CANCEL)
    End If
  End If
  inTableClick = False
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub Cmd_Next_Click()
  Call m_ImpWiz.ReCalc_Src(Nothing)
  Call SwitchForm(Me, TCSIMP_SOURCE)
End Sub

Private Sub Opt_Update_Click(Index As Integer)
  Dim tmp As IMPORT_UPDATE_TYPE
  
  On Error Resume Next
  If InUpdate Then Exit Sub
  tmp = GetUpdateType(Me)
  Call SetUpdateType(Me, tmp, m_ImpWiz.AllowChangeUpdateType, m_ImpWiz.CurrentDest.LockUpdate)
  m_ImpWiz.ImpParent.UpdateType = tmp
  m_ImpWiz.CurrentDest.UpdateType = m_ImpWiz.ImpParent.UpdateType
End Sub
