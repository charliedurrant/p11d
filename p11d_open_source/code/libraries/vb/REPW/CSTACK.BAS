Attribute VB_Name = "CoordStack"
Option Explicit
Public Enum PUSH_TYPE
  PUSH_BOTH = 1
  PUSH_X
  PUSH_Y
End Enum

Private m_STACKTOP As Long

Private m_STACKMAX As Long
Private Const INCREMENT As Long = 20
Private m_CoordStack() As Coordinate

Public Sub SetStackX(ByVal xCoord As Single)
  Dim i As Long
  For i = 1 To m_STACKTOP
    m_CoordStack(m_STACKTOP).x = xCoord
  Next i
End Sub

Public Sub SetStackY(ByVal yCoord As Single)
  Dim i As Long
  For i = 1 To m_STACKTOP
    m_CoordStack(m_STACKTOP).y = yCoord
  Next i
End Sub

Public Sub PushCoord(Optional ByVal pType As PUSH_TYPE = PUSH_BOTH)
  Call xSet("PushCoord")
  
  m_STACKTOP = m_STACKTOP + 1
  If m_STACKTOP > m_STACKMAX Then
    m_STACKMAX = m_STACKMAX + INCREMENT
    ReDim Preserve m_CoordStack(1 To m_STACKMAX) As Coordinate
  End If
  If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
    m_CoordStack(m_STACKTOP).x = Printer.CurrentX
    m_CoordStack(m_STACKTOP).y = Printer.CurrentY
  ElseIf IsExportHTML Then
    m_CoordStack(m_STACKTOP).x = ReportControl.HTML.CurrentX
    m_CoordStack(m_STACKTOP).y = ReportControl.HTML.CurrentY
  Else
    m_CoordStack(m_STACKTOP).x = ReportControl.Preview.CurrentX
    m_CoordStack(m_STACKTOP).y = ReportControl.Preview.CurrentY
  End If
  If pType = PUSH_X Then
    m_CoordStack(m_STACKTOP).y = -1
  End If
  If pType = PUSH_Y Then
    m_CoordStack(m_STACKTOP).x = -1
  End If
  m_CoordStack(m_STACKTOP).FirstX = ReportControl.FirstX
  Call xReturn("PushCoord")
End Sub

Public Sub PopCoord()
  Call xSet("PopCoord")
  If m_STACKTOP > 0 Then
    If (ReportControl.rTarget = RPT_PRINTER) Or (ReportControl.rTarget = RPT_PREVIEW_PRINT) Then
      If m_CoordStack(m_STACKTOP).y <> -1 Then Printer.CurrentY = m_CoordStack(m_STACKTOP).y
      If m_CoordStack(m_STACKTOP).x <> -1 Then Printer.CurrentX = m_CoordStack(m_STACKTOP).x
    ElseIf IsExportHTML Then
      If m_CoordStack(m_STACKTOP).x <> -1 Then ReportControl.HTML.CurrentX = m_CoordStack(m_STACKTOP).x
      If m_CoordStack(m_STACKTOP).y <> -1 Then ReportControl.HTML.CurrentY = m_CoordStack(m_STACKTOP).y
    Else
      If m_CoordStack(m_STACKTOP).y <> -1 Then ReportControl.Preview.CurrentY = m_CoordStack(m_STACKTOP).y
      If m_CoordStack(m_STACKTOP).x <> -1 Then ReportControl.Preview.CurrentX = m_CoordStack(m_STACKTOP).x
    End If
    ReportControl.FirstX = m_CoordStack(m_STACKTOP).FirstX
    m_STACKTOP = m_STACKTOP - 1
  Else
    Call Err.Raise(ERR_POPCOORDSTACK, "PopCoord", "Pop without Push!")
  End If
  Call xReturn("PopCoord")
End Sub

Public Function StackTopX() As Single
  Call xSet("StackTopX")
  If m_STACKTOP > 0 Then
    StackTopX = m_CoordStack(m_STACKTOP).x
  Else
    Call Err.Raise(ERR_POPCOORDSTACK, "StackTopX", "No Top X - Stack empty!")
  End If
  Call xReturn("StackTopX")
End Function

Public Function StackTopY() As Single
  Call xSet("StackTopY")
  If m_STACKTOP > 0 Then
    StackTopY = m_CoordStack(m_STACKTOP).y
  Else
    Call Err.Raise(ERR_POPCOORDSTACK, "StackTopY", "No Top Y - Stack empty!")
  End If
  Call xReturn("StackTopY")
End Function

