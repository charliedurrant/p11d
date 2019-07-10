VERSION 5.00
Object = "*\AUBTG.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Display errors"
      Height          =   240
      Left            =   1890
      TabIndex        =   2
      Top             =   2430
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   180
      TabIndex        =   1
      Top             =   2250
      Width           =   915
   End
   Begin UBTG.TCSUBGRD TCSUBTDBG1 
      Height          =   1905
      Left            =   225
      TabIndex        =   0
      Top             =   90
      Width           =   3885
      _ExtentX        =   5980
      _ExtentY        =   2646
      DisplayErrors   =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OL  As ObjectList


Private Sub Check1_Click()
  Form1.TCSUBTDBG1.DisplayErrors = Check1.Value
End Sub

Private Sub Form_Load()
  Dim cl As c, i As Long
  Dim c As TrueDBGrid50.Column
  
  Set OL = New ObjectList
  For i = 1 To 10
    Set cl = New c
    cl.Name = "Fred" & i
    cl.Number = i
    Call OL.Add(cl)
  Next
  
  Set c = TCSUBTDBG1.Grid.Columns.Add(0)
  c.Caption = "NAME"
  c.Visible = True
  Set c = TCSUBTDBG1.Grid.Columns.Add(1)
  c.Caption = "number"
  c.Visible = True
  
  TCSUBTDBG1.ObjectList = OL
  TCSUBTDBG1.Grid.ReBind
End Sub


Private Sub TCSUBTDBG1_AddData(ByVal RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As TCSCOREDLL.ObjectList, NewObjectListIndex As Long)
  Dim cl As New c
  
  cl.Name = RowBuf.Value(RowBufRowIndex, 0)
  cl.Number = RowBuf.Value(RowBufRowIndex, 1)
  NewObjectListIndex = ObjectList.Add(cl)
End Sub

Private Sub TCSUBTDBG1_ReadData(RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As TCSCOREDLL.ObjectList, ByVal ObjectListIndex As Long)
  Dim i As Long
  Dim cl As c
  
  Set cl = ObjectList(ObjectListIndex)
  RowBuf.Value(RowBufRowIndex, 0) = cl.Name
  RowBuf.Value(RowBufRowIndex, 1) = cl.Number
End Sub

Private Sub TCSUBTDBG1_Validate(FirstColIndexInError As Variant, ErrorMessage As String, ByVal RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long)
  If Not IsNull(RowBuf.Value(RowBufRowIndex, 1)) Then 'ha is changed
    If Not IsNumeric(RowBuf.Value(RowBufRowIndex, 1)) Then
      FirstColIndexInError = 1
      ErrorMessage = "Error in number"
    End If
  End If
End Sub


Private Sub TCSUBTDBG1_WriteData(ByVal RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As TCSCOREDLL.ObjectList, ObjectListIndex As Long, ByVal NewRow As Boolean)
Dim cl As c
  
  If NewRow Then
    Set cl = New c
    ObjectListIndex = ObjectList.Add(cl)
  Else
    Set cl = ObjectList(ObjectListIndex)
  End If
  
  If Not IsNull(RowBuf.Value(RowBufRowIndex, 0)) Then cl.Name = RowBuf.Value(RowBufRowIndex, 0)
  If Not IsNull(RowBuf.Value(RowBufRowIndex, 1)) Then cl.Number = RowBuf.Value(RowBufRowIndex, 1)
End Sub
