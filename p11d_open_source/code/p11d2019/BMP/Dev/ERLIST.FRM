VERSION 2.00
Begin Form F_ErList 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Er List"
   ClientHeight    =   3615
   ClientLeft      =   1920
   ClientTop       =   4305
   ClientWidth     =   6465
   Height          =   4020
   KeyPreview      =   -1  'True
   Left            =   1860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   3615
   ScaleWidth      =   6465
   Top             =   3960
   Width           =   6585
   Begin CommandButton B_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3180
      Width           =   1245
   End
   Begin CommandButton B_Ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3870
      TabIndex        =   1
      Top             =   3180
      Width           =   1245
   End
   Begin Frame Panel3D1 
      AutoSize        =   3  'AutoSize Child To Panel
      BevelOuter      =   1  'Inset
      Height          =   2535
      Left            =   60
      TabIndex        =   3
      Top             =   570
      Width           =   6375
      Begin ListBox LB 
         Prop110         =   ERLIST.FRX:0000
         AllowAdditions  =   0   'False
         AllowDeletions  =   0   'False
         AllowUpdates    =   0   'False
         BackColor       =   &H00FFFFFF&
         Caption         =   "ListBox1"
         CaptionType     =   3  'No Caption
         ColHeadingAlignment=   2  'Center
         Cols            =   2
         DividerType     =   0  'None
         FixedCols       =   2
         ForeColor       =   &H00800000&
         HdgForeColor    =   &H00800000&
         Height          =   2505
         Left            =   15
         MousePointer    =   1  'Arrow
         RowHeight       =   150
         RowLabels       =   0   'False
         ScrollBars      =   0  'None
         SelectionTypeCol=   0  'None
         SelectionTypeRow=   1  'Single
         TabIndex        =   0
         Top             =   15
         VariableColWidth=   0   'False
         Width           =   6345
      End
   End
   Begin Label L_ErList 
      BackStyle       =   0  'Transparent
      Caption         =   "Select an employer"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6345
   End
End
Option Explicit

Sub B_Cancel_Click ()

  ER_SELECTED = "[NULL]"
  Unload Me
  
End Sub

Sub B_OK_Click ()


  If LB.Row = -1 Then
    ErrorMessage "Employer list",Err,  "Either select an employer from the list, or press cancel to stop the selection", 0, "", ERR_INFO
  Else
    ER_SELECTED = LB.ColText(0)
    Unload Me
  End If
  
End Sub

Sub Form_KeyDown (keycode As Integer, Shift As Integer)
processkeys keycode, Shift
Select Case keycode
  Case KEY_ESCAPE:
    Call B_Cancel_Click
End Select
End Sub

Sub Form_Load ()
  Call centreform(Me)
  LB.Rows = N_ER

End Sub

Sub LB_DblClick ()

  ER_SELECTED = LB.ColText(0)
  Unload Me

End Sub

Sub LB_KeyDown (keycode As Integer, Shift As Integer)
  Select Case keycode
  Case 13: SendKeys ("{TAB}")
  Case Else
  End Select
End Sub

Sub LB_RowLoaded (BookMark As String, RowNum As Long)
  'If UCase$(ER_DB) <> UCase$(Employers(RowNum + 1).FName & ".MDB") Then
    LB.ColText(0) = Employers(RowNum + 1).FName
    LB.ColText(1) = Employers(RowNum + 1).NAme
  'Else
  '  LB.Rows = LB.Rows - 1
  'End If
End Sub

