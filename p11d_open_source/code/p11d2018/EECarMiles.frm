VERSION 5.00
Begin VB.Form F_EECarMiles 
   Caption         =   "Employee Owned Car Mileage"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton B_Ok 
      Caption         =   "&OK"
      Height          =   405
      Left            =   4680
      TabIndex        =   0
      Top             =   3840
      Width           =   1755
   End
   Begin VB.Label lblTotalMiles 
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "F_EECarMiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public benefit As IBenefitClass
Private Sub B_OK_Click()
  Me.benefit.Dirty = True
  If benefit.Dirty Then
   Call MDIMain.SetConfirmUndo
  End If
  F_BenEECar.TB_Data(9).Text = SumMiles
  LoadingMiles = False
  Me.Hide
End Sub

'Private Sub _FieldInvalid(Valid As Boolean, Message As String)
'  Call MDIMain.sts.SetStatus(0, Message)
'End Sub

'Private Sub efgMiles_UserChangedCell(Row As Long, Column As Long)
'
'  If Column = 0 Then _
'    Me.lblTotalMiles.Caption = SumMiles
'  Me.benefit.dirty = true

'End Sub


Private Function SumMiles() As String
  Dim lTotal As Long, I As Long, s As String
  On Error GoTo SumMiles_Err
  Call xSet("SumMiles")
  
  For I = 1 To EEMilesGrid.NoRows
    s = EEMilesGrid.TextMatrix(3, I)
    If Len(s) > 0 Then
      lTotal = lTotal + CLng(s)
    End If
  Next I
  SumMiles = CStr(lTotal)
  
SumMiles_End:
  Call xReturn("SumMiles")
  Exit Function
SumMiles_Err:
  SumMiles = S_ERROR
  Resume SumMiles_End
End Function



Public Function SortGrid() As Boolean
  Dim I As Long, j As Long
  Dim sNext As String
  Dim sThis As String
  Dim q(5) As String
  Dim X As Long
  Dim ThisDate As Date, NextDate As Date
  On Error GoTo SortGrid_Err
  Call xSet("SortGrid")
    X = EEMilesGrid.NoRows
    
    For I = 1 To EEMilesGrid.NoRows - 2
      For j = 1 To EEMilesGrid.NoRows - 2
      sThis = EEMilesGrid.TextMatrix(j, 0)
      sNext = EEMilesGrid.TextMatrix(j + 1, 0)
      If IsDate(sThis) And IsDate(sNext) Then
        ThisDate = CDate(sThis)
        NextDate = CDate(sNext)
        If NextDate < ThisDate Then
          'swap the elements
          q(0) = EEMilesGrid.TextMatrix(j + 1, 0)
          q(1) = EEMilesGrid.TextMatrix(j + 1, 1)
          q(2) = EEMilesGrid.TextMatrix(j + 1, 2)
          q(3) = EEMilesGrid.TextMatrix(j + 1, 3)
          
          'EEMilesGrid.TextMatrix(j + 1, 0) = EEMilesGrid.TextMatrix(j, 0)
          'EEMilesGrid.TextMatrix(j + 1, 1) = EEMilesGrid.TextMatrix(j, 1)
          'EEMilesGrid.TextMatrix(j + 1, 2) = EEMilesGrid.TextMatrix(j, 2)
          'EEMilesGrid.TextMatrix(j + 1, 3) = EEMilesGrid.TextMatrix(j, 3)
          
          'EEMilesGrid.TextMatrix(j, 0) = q(0)
          'EEMilesGrid.TextMatrix(j, 1) = q(1)
          'EEMilesGrid.TextMatrix(j, 2) = q(2)
          'EEMilesGrid.TextMatrix(j, 3) = q(3)
          If X = j Then
            X = X + 1
          ElseIf X = (j + 1) Then
            X = X - 1
          End If
        End If
      End If
      Next j
    Next I
  'EEMilesGrid.Row = x

SortGrid_End:
  Call xReturn("SortGrid")
  Exit Function

SortGrid_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SortGrid", "ERR_UNDEFINED", "Undefined error.")
  Resume SortGrid_End
End Function

Private Sub Form_Load()

End Sub
