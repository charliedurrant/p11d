VERSION 5.00
Begin VB.UserControl Monitor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2805
   ScaleWidth      =   4320
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   1935
      Top             =   360
   End
End
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Event GetValue(Value As Single, ToolTip As String)
Private Const L_GRID_SQUARE_WIDTH_PIXELS As Long = 12
Private Type Size
  cx As Long
  cy As Long
End Type
Private Type GridPoint
  Value As Single
  ToolTip As String
  Marked As Boolean
End Type
Private m_MemoryUseage() As GridPoint 'percentages of free memory
Private m_CurrentMemoryUseageFreeSlot As Long
Private m_xGridOffset As Long
Private m_szUserControlPixels As Size
Private Const L_STEP As Long = 2
Private m_MemoryUseageMaxSlots As Long
Private m_sngMaxValue As Single
Private m_Caption As String
Private m_XLast As Single
Private m_BeepIfGetHigherThan As Single
Public Property Let Max(ByVal NewValue As Single)
  m_sngMaxValue = NewValue
End Property
Public Property Get Max() As Single
  Max = m_sngMaxValue
End Property
Public Property Get Interval() As Long
  Interval = tmr.Interval
End Property
Public Property Let Interval(ByVal NewValue As Long)
  tmr.Interval = NewValue
End Property
Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Let Caption(ByVal NewValue As String)
  m_Caption = NewValue
End Property
Private Sub Sound(ByVal sngValue)
  'each percentage point = 100Hz
  Dim sng As Single
  Dim iHZ As Long
  
  If m_BeepIfGetHigherThan <> -1 Then
    If m_BeepIfGetHigherThan < sngValue Then
      sng = sngValue - m_BeepIfGetHigherThan
      sng = Round((sng / m_BeepIfGetHigherThan), 2) * 100
      iHZ = 1000 + (sng * 300!)
      Call Beep(iHZ, 100)
    End If
  End If
  
  
End Sub
Private Sub UpdateMemoryUseageArray()
  Dim i As Long
  Dim sngPhysicalTotal As Single, sngPhysicalAvailable As Single
  Dim sngPercentage As Single
  Dim sngValue As Single
  Dim sToolTip As String
  RaiseEvent GetValue(sngValue, sToolTip)
  
  Call Sound(sngValue)
  sngPercentage = sngValue / m_sngMaxValue
  UserControl.CurrentX = 0
  UserControl.CurrentY = 0
  
  UserControl.Print m_Caption
  
  If m_CurrentMemoryUseageFreeSlot > m_MemoryUseageMaxSlots Then
    'shift the elements down
    For i = 1 To m_MemoryUseageMaxSlots
      If i < m_MemoryUseageMaxSlots Then
        m_MemoryUseage(i) = m_MemoryUseage(i + 1)
      End If
    Next
    m_MemoryUseage(m_MemoryUseageMaxSlots).Value = sngPercentage
    m_MemoryUseage(m_MemoryUseageMaxSlots).ToolTip = sToolTip
  Else
    m_MemoryUseage(m_CurrentMemoryUseageFreeSlot).Value = sngPercentage
    m_MemoryUseage(m_CurrentMemoryUseageFreeSlot).ToolTip = sToolTip
    m_CurrentMemoryUseageFreeSlot = m_CurrentMemoryUseageFreeSlot + 1
  End If
End Sub

Private Sub tmr_Timer()
  Call Draw
End Sub

Private Sub Draw()
  If Not UserControl.Enabled Then Exit Sub
  'update the
  
  UserControl.BackColor = vbBlack
  Call UserControl.Cls
  
  If m_sngMaxValue = 0 Then Exit Sub

  Call DrawGrid
  UserControl.ForeColor = vbGreen
  Call UpdateMemoryUseageArray
  
  
  Call DrawGraphLines
End Sub
Private Sub DrawGrid()
  Dim iCurrent As Long
  Dim i As Long
  
  UserControl.ForeColor = RGB(0, 128, 64)
  'draw the vertical lines
  iCurrent = m_xGridOffset
  Do While iCurrent < m_szUserControlPixels.cx
    i = iCurrent * Screen.TwipsPerPixelX
    UserControl.Line (i, 0)-(i, UserControl.Height)
    iCurrent = iCurrent + L_GRID_SQUARE_WIDTH_PIXELS
  Loop
  If (m_CurrentMemoryUseageFreeSlot > m_MemoryUseageMaxSlots) Then
    m_xGridOffset = m_xGridOffset - 1
    If m_xGridOffset < ((-1 * L_GRID_SQUARE_WIDTH_PIXELS) + 1) Then
      m_xGridOffset = 0
    End If
  End If
  'draw the horizontal lines
  iCurrent = 0
  Do While iCurrent < m_szUserControlPixels.cy
    i = iCurrent * Screen.TwipsPerPixelY
    UserControl.Line (0, i)-(UserControl.Width, i)
    iCurrent = iCurrent + L_GRID_SQUARE_WIDTH_PIXELS
  Loop
End Sub
Private Sub DrawGraphLines()
  Dim y1 As Long, y2 As Long
  Dim x1 As Long, x2 As Long
  Dim i As Long
  Dim iTwips As Long
  
  Const L_BOX_MARK_WIDTH As Long = 5
  UserControl.ForeColor = vbGreen
  For i = 2 To m_CurrentMemoryUseageFreeSlot - 1
    y1 = UserControl.Height - (UserControl.Height * m_MemoryUseage(i - 1).Value)
    y2 = UserControl.Height - (UserControl.Height * m_MemoryUseage(i).Value)
    x1 = ((i - 2) * L_STEP) * Screen.TwipsPerPixelX
    x2 = ((i - 1) * L_STEP) * Screen.TwipsPerPixelX
    
    If m_MemoryUseage(i).Marked Then
      iTwips = 1 * Screen.TwipsPerPixelY
      UserControl.Line (x1, y1 - iTwips)-(x2, y2 + iTwips), vbRed, BF
    Else
      UserControl.Line (x1, y1)-(x2, y2)
    End If
  Next
End Sub

Private Sub UserControl_Initialize()
  m_CurrentMemoryUseageFreeSlot = 1
  m_BeepIfGetHigherThan = -1
  m_XLast = -1!
  Call DrawGrid
End Sub
Public Property Get BeepIfGetHigherThan() As Single
  BeepIfGetHigherThan = m_BeepIfGetHigherThan
End Property
Public Property Let BeepIfGetHigherThan(ByVal NewValue As Single)
  m_BeepIfGetHigherThan = NewValue
End Property
Private Sub ToolTipSet(ByVal X As Single)
  Dim iX As Long
  Dim sToolTip As String
  Dim i As Long, x1 As Long, x2 As Long
  Dim sCurentToolTip As String
  Dim bInFunc As Boolean
  
  If bInFunc Then Exit Sub
  
  If X = m_XLast Then Exit Sub
  
  bInFunc = True
  m_XLast = X
  iX = X / CSng(Screen.TwipsPerPixelX)
  'scale the mouse x position
  sCurentToolTip = UserControl.Extender.ToolTipText
  For i = 2 To m_CurrentMemoryUseageFreeSlot - 1
    x1 = ((i - 2) * L_STEP)
    x2 = ((i - 1) * L_STEP)
    If (iX >= x1 And iX <= x2) Then
       sToolTip = m_MemoryUseage(i).ToolTip
       m_MemoryUseage(i).Marked = True
    Else
      m_MemoryUseage(i).Marked = False
    End If
  Next
  If (Len(sToolTip) > 0) Then
    If (StrComp(sCurentToolTip, sToolTip) <> 0) Then
      UserControl.Extender.ToolTipText = sToolTip
    End If
  Else
    If Len(sCurentToolTip) > 0 Then
      UserControl.Extender.ToolTipText = sToolTip
    End If
  End If
  bInFunc = False
  
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call ToolTipSet(X)
End Sub
Private Sub UserControl_Resize()
  Dim i As Long
  
  m_szUserControlPixels.cx = UserControl.Width / Screen.TwipsPerPixelX
  m_szUserControlPixels.cy = UserControl.Height / Screen.TwipsPerPixelY
  
  'resize the control to an even number of pixels as we need this for the step
  If m_szUserControlPixels.cx Mod L_STEP <> 0 Then
    m_szUserControlPixels.cx = m_szUserControlPixels.cx + 1
    UserControl.Width = m_szUserControlPixels.cx * Screen.TwipsPerPixelX
    Exit Sub
  End If
  m_MemoryUseageMaxSlots = m_szUserControlPixels.cx / 2
  ReDim Preserve m_MemoryUseage(1 To m_MemoryUseageMaxSlots)
  If m_CurrentMemoryUseageFreeSlot > m_MemoryUseageMaxSlots Then
    m_CurrentMemoryUseageFreeSlot = m_MemoryUseageMaxSlots + 1
  End If
End Sub

