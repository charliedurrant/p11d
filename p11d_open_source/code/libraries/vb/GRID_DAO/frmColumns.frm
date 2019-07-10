VERSION 5.00
Begin VB.Form frmColumns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select columns?"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   2610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1575
      TabIndex        =   2
      Top             =   2925
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   525
      TabIndex        =   1
      Top             =   2925
      Width           =   960
   End
   Begin VB.ListBox lst 
      Height          =   2790
      Left            =   45
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   90
      Width           =   2490
   End
End
Attribute VB_Name = "frmColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_bCancelPressed As Boolean
Private m_RetVal As Long
Private m_NumberOfColumnsRequired As Long
Private m_ac As AutoClass

Private Sub cmdCancel_Click()
  m_RetVal = -1
  Me.Hide
End Sub

Private Sub OKClick()
  Dim aCol As AutoCol
  Dim i As Long
  
  On Error GoTo OKClick_ERR
  If (m_NumberOfColumnsRequired = -1) And (lst.SelCount = 0) Then
    m_RetVal = -1
    Me.Hide
  ElseIf (m_NumberOfColumnsRequired = -1) Or (m_NumberOfColumnsRequired = lst.SelCount) Then
    For i = 0 To lst.ListCount - 1
      If lst.Selected(i) Then
        If GetAColByGridIndexEx(lst.ItemData(i), aCol, m_ac) Then aCol.ClipboardColumn = True
      End If
    Next
    m_RetVal = lst.SelCount
    Me.Hide
  Else
    Call Err.Raise(ERR_NUMBER_OF_COLUMNS, "OKClick", "You must select " & m_NumberOfColumnsRequired & "!")
  End If
    
OKClick_END:
  Exit Sub
  
OKClick_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "OKClick", "OK Click", "Error in OKClick")
  Resume OKClick_END
End Sub

Private Sub cmdOK_Click()
  Call OKClick
End Sub

Public Function GetColumns(ByVal ac As AutoClass, ByVal NumberOfColumnsRequired) As Long
  Dim i As Long
  Dim grid As TDBGrid
  Dim Col As Column
  Dim aCol As AutoCol
  
  On Error GoTo GetColumns_ERR
  Set grid = ac.grid.TDBGrid
  For Each Col In grid.Columns
    If Col.Visible Then
      Call GetAColByGridIndexEx(grid.Col, aCol, ac)
      If Not aCol.NoCopy Then
        lst.AddItem Col.Caption
        lst.ItemData(lst.ListCount - 1) = Col.ColIndex
      End If
    End If
  Next
  Set m_ac = ac
  m_NumberOfColumnsRequired = NumberOfColumnsRequired
  If NumberOfColumnsRequired < 0 Then
    Me.Caption = "Select columns ..."
  Else
    Me.Caption = "Select " & NumberOfColumnsRequired & " columns ..."
  End If
  Me.cmdOK.Enabled = False
  Me.Show 1
    
GetColumns_END:
  Set m_ac = Nothing
  m_bCancelPressed = False
  lst.Clear
  GetColumns = m_RetVal
  Exit Function
  
GetColumns_ERR:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "GetColumns"), "Get Columns", "Error in GetColumns")
  Resume GetColumns_END
End Function

Private Sub lst_Click()
  Me.cmdOK.Enabled = (m_NumberOfColumnsRequired = -1) Or (m_NumberOfColumnsRequired = lst.SelCount)
End Sub
