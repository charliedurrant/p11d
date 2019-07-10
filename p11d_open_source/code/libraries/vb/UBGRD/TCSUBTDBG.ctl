VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.UserControl UBGRD 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "TCSUBTDBG.ctx":0000
   Begin TrueDBGrid50.TDBGrid Grid_i 
      Height          =   2850
      Left            =   180
      OleObjectBlob   =   "TCSUBTDBG.ctx":00FA
      TabIndex        =   0
      Top             =   225
      Width           =   4335
   End
End
Attribute VB_Name = "UBGRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Grid As TrueDBGrid50.TDBGrid

Public Event ReadData(RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ByVal ObjectListIndex As Long)
Public Event WriteData(ByVal RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long, ByVal NewRow As Boolean)
Public Event Validate(FirstColIndexInError, ErrorMessage As String, ByVal RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long)

Private m_OL As ObjectList
Private m_bAllowReadData As Boolean
Private m_bDisplayErrors As Boolean
Private m_bHideGridSystemErrors As Boolean
Private Const S_DISPLAYERRORS As String = "DisplayErrors"
Private Const B_DISPLAYERRORS_DEF As Boolean = True

Private Sub UserControl_Initialize()
  Set Grid = Grid_i
End Sub

Private Sub UserControl_InitProperties()
  m_bDisplayErrors = B_DISPLAYERRORS_DEF
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    m_bDisplayErrors = .ReadProperty(S_DISPLAYERRORS, B_DISPLAYERRORS_DEF)
  End With
End Sub

Private Sub UserControl_Resize()
  With Grid_i
    .Left = 0
    .Top = 0
    .Width = UserControl.Width
    .Height = UserControl.Height
  End With
End Sub

Public Property Get DisplayErrors() As Boolean
  DisplayErrors = m_bDisplayErrors
End Property

Public Property Let DisplayErrors(NewValue As Boolean)
  m_bDisplayErrors = NewValue
  Call PropertyChanged(S_DISPLAYERRORS)
End Property

Public Property Let ObjectList(NewValue As ObjectList)
  Set m_OL = NewValue
  Grid_i.ReBind
End Property

Public Property Get ObjectList() As ObjectList
  Set ObjectList = m_OL
End Property

Private Sub grid_i_Error(ByVal DataError As Integer, Response As Integer)
  Response = Not m_bHideGridSystemErrors
  m_bHideGridSystemErrors = Not m_bHideGridSystemErrors
End Sub

Private Sub grid_i_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal offset As Long, ApproximatePosition As Long)
  Dim ObjectListIndex As Long, RowBufRowIndex As Long
  Dim i As Long
  
  '//first set if there are any valid objects in the list
  If m_OL Is Nothing Then Exit Sub
  For i = 1 To m_OL.Count
    If Not m_OL(i) Is Nothing Then Exit For
  Next
  
  If i > m_OL.Count Then Exit Sub
  
  'this gets the fist object list index
  ObjectListIndex = GetObjectListIndex(m_OL, RowBuf, StartLocation, offset)
  If ObjectListIndex < 0 Then Exit Sub
  
  RowBufRowIndex = 0
  
  Do While (ObjectListIndex > 0) And (RowBufRowIndex < RowBuf.RowCount)
    ' Fill Row Buffer
    RowBuf.Bookmark(RowBufRowIndex) = CStr(ObjectListIndex)
    If RowBuf.ColumnCount > 0 Then RaiseEvent ReadData(RowBuf, RowBufRowIndex, m_OL, ObjectListIndex)
    RowBufRowIndex = RowBufRowIndex + 1
    If RowBufRowIndex = 1 Then ApproximatePosition = ObjectListIndex
    ObjectListIndex = GetObjectListIndex(m_OL, RowBuf, ObjectListIndex, 1)
  Loop
  
  RowBuf.RowCount = RowBufRowIndex
  
End Sub

Private Sub grid_i_UnboundAddData(ByVal RowBuf As TrueDBGrid50.RowBuffer, NewRowBookmark As Variant)
  Dim NewObjectListIndex As Long
  Dim ErrorMessage As String
  Dim FirstColIndexInError As Long
  
  FirstColIndexInError = -1
  RaiseEvent Validate(FirstColIndexInError, ErrorMessage, RowBuf, 0)
  If FirstColIndexInError = -1 Then
    RaiseEvent WriteData(RowBuf, 0, m_OL, NewObjectListIndex, True)
    If NewObjectListIndex >= 1 And NewObjectListIndex <= m_OL.Count Then
      NewRowBookmark = CStr(NewObjectListIndex)
    Else
      Call Err.Raise(ERR_UNBOUNDADD, "UnboundAddData", "Error setting the new rows bookmark to the object list index of value " & NewObjectListIndex & ".")
    End If
  Else
    m_bHideGridSystemErrors = True
    If m_bDisplayErrors Then
      'Albert ZZZZZ
      'Call tcscoredll.ErrorMessage(ERR_INFO, Err, "UnboundAddData", "Error in column " & FirstColIndexInError & ". ", ErrorMessage)
    End If
    
  End If
End Sub

Private Sub grid_i_UnboundWriteData(ByVal RowBuf As TrueDBGrid50.RowBuffer, WriteLocation As Variant)
  Dim FirstColIndexInError As Long
  Dim ObjectListIndex As Long
  Dim ErrorMessage As String
  
  FirstColIndexInError = -1
  If IsNumeric(WriteLocation) Then
    ObjectListIndex = CLng(WriteLocation)
    RaiseEvent Validate(FirstColIndexInError, ErrorMessage, RowBuf, 0)
    If FirstColIndexInError = -1 Then
      RaiseEvent WriteData(RowBuf, 0, m_OL, ObjectListIndex, False)
    Else
      
      If m_bDisplayErrors Then
        'Albert ZZZZZ
        'Call tcscoredll.ErrorMessage(ERR_ERROR, Err, "UnboundAddData", "Error in column " & FirstColIndexInError & ". ", ErrorMessage)
      End If
    End If
  Else
    Call Err.Raise(ERR_UNBOUNDWRITE, "UnboundWriteData", "The object list index is no numeric.")
  End If
End Sub

Private Function GetObjectListIndex(oList As ObjectList, RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal offset As Long) As Long
  Dim FirstObject As Long
  On Error GoTo GetObjectListIndex_Err
    
  Call xSet("GetObjectListIndex")
  GetObjectListIndex = -1
  
  If IsNull(StartLocation) Then
    If offset > 0 Then
      FirstObject = GetObjectListOffset(oList, 0, 1)
      offset = offset - 1
      If FirstObject < 0 Then RowBuf.RowCount = 0
    ElseIf offset < 0 Then
      FirstObject = GetObjectListOffset(oList, oList.Count + 1, -1)
    Else
      GetObjectListIndex = -1
      Exit Function
    End If
    If FirstObject < 0 Then Exit Function
  Else
    FirstObject = StartLocation
  End If
  GetObjectListIndex = GetObjectListOffset(oList, FirstObject, offset)
  
GetObjectListIndex_End:
  Call xReturn("GetObjectListIndex")
  Exit Function

GetObjectListIndex_Err:
  GetObjectListIndex = -1
  Resume GetObjectListIndex_End
End Function

Private Function GetObjectListOffset(oList As ObjectList, ByVal Start As Long, ByVal offset As Long) As Long
  Dim i As Long, xOffset As Long
    
  GetObjectListOffset = Start
  If offset > 0 Then
    xOffset = 0: GetObjectListOffset = -1
    For i = Start + 1 To oList.Count
      If Not oList(i) Is Nothing Then
        xOffset = xOffset + 1
        If offset = xOffset Then
          GetObjectListOffset = i
          Exit For
        End If
      End If
    Next i
  ElseIf offset < 0 Then
    offset = offset * -1
    xOffset = 0: GetObjectListOffset = -1
    For i = (Start - 1) To 1 Step -1
      If Not oList(i) Is Nothing Then
        xOffset = xOffset + 1
        If offset = xOffset Then
          GetObjectListOffset = i
          Exit For
        End If
      End If
    Next i
  End If
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty(S_DISPLAYERRORS, m_bDisplayErrors, B_DISPLAYERRORS_DEF)
  End With
End Sub
