VERSION 5.00
Object = "{E297AE83-F913-4A8C-873C-EDEAC00CB9AC}#2.1#0"; "atc3ubgrd.ocx"
Begin VB.UserControl P11DbAdjustmentEditor 
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   ScaleHeight     =   3315
   ScaleWidth      =   5205
   Begin atc3ubgrd.UBGRD grid 
      Height          =   1230
      Left            =   495
      TabIndex        =   0
      Top             =   345
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   2170
   End
End
Attribute VB_Name = "P11DbAdjustmentEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Property Set Adjustments(ByVal value As ObjectList)
  Set grid.ObjectList = value
End Property
Public Property Get Adjustments() As ObjectList
  Set Adjustments = grid.ObjectList
End Property
Private Sub grid_ReadData(RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ATC2CORE.ObjectList, ByVal ObjectListIndex As Long)
  Dim p11dbAdjust As P11DbAdjustment
  Dim i As Long
  
  Set p11dbAdjust = ObjectList(ObjectListIndex)
  
  For i = 0 To (RowBuf.ColumnCount - 1)
    Select Case i
      Case 0
        RowBuf.value(RowBufRowIndex, i) = p11dbAdjust.caption
      Case 1
        RowBuf.value(RowBufRowIndex, i) = p11dbAdjust.value
      Case Else
        ECASE ("Invalid column ubgrd read data.")
    End Select
  Next
End Sub
Private Sub grid_DeleteData(ObjectList As ATC2CORE.ObjectList, ObjectListIndex As Long)
  Call ObjectList.Remove(ObjectListIndex)
End Sub

Private Sub grid_ValidateTCS(FirstColIndexInError As Long, ValidateMessage As String, ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ByVal ObjectListIndex As Long)
  Dim l As Long
  Dim caption As Variant
  Dim ol As ObjectList
  Dim i As Long
  Dim adjustment As P11DbAdjustment
  
  On Error GoTo err_err
  
  Call xSet("P11dbAdjustmentsValidate")
  
  With RowBuf
    For l = 0 To RowBuf.ColumnCount - 1
      Select Case l
        Case 0
          'description
          caption = RowBuf.value(RowBufRowIndex, l)
          If (Len(caption) = 0) Then
            FirstColIndexInError = l
            ValidateMessage = "Value must not be blank"
            GoTo err_end
          End If
          
          If GrisIsTooLong(ValidateMessage, RowBuf, RowBufRowIndex, l, 100) Then
            FirstColIndexInError = l
            GoTo err_end
          End If
          
          Set ol = grid.ObjectList
          For i = 1 To ol.count
            Set adjustment = ol(i)
            If ObjectListIndex <> i Then
              If (adjustment.caption = caption) Then
                FirstColIndexInError = l
                ValidateMessage = "Description must be unique"
                GoTo err_end
              End If
            End If
          Next
        Case 1
          'no of miles
          If GridIsNotNumericOrLong(ValidateMessage, RowBuf.value(RowBufRowIndex, l), ObjectListIndex, False) Then
            FirstColIndexInError = l
            GoTo err_end
          End If
      End Select
    Next
  End With

err_end:
  Call xReturn("P11dbAdjustmentsValidate")
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "P11dbAdjustmentsValidate", "P11db Adjustments Validate", "Error validating a P11Db adjustment.")
  Resume err_end
End Sub
Private Sub grid_WriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ATC2CORE.ObjectList, ObjectListIndex As Long)
  Dim p11dbAdjust As P11DbAdjustment
  
  If ObjectListIndex = -1 Then
    Set p11dbAdjust = New P11DbAdjustment
    ObjectListIndex = ObjectList.Add(p11dbAdjust)
  Else
    Set p11dbAdjust = ObjectList(ObjectListIndex)
  End If
       
  With p11dbAdjust
    If Not IsNull(RowBuf.value(RowBufRowIndex, 0)) Then .caption = RowBuf.value(RowBufRowIndex, 0)
    If Not IsNull(RowBuf.value(RowBufRowIndex, 1)) Then .value = CLng(RowBuf.value(RowBufRowIndex, 1))
  End With
End Sub
Private Sub UserControl_Initialize()
  Call AddUBGRDStandardColumn(grid.grid, 0, 4300, "Description", "")
  Call AddUBGRDStandardColumn(grid.grid, 1, 1000, "Value", "")
End Sub

Private Sub UserControl_Resize()
  grid.Top = 0
  grid.Left = 0
  grid.width = UserControl.width
  grid.height = UserControl.height
End Sub
