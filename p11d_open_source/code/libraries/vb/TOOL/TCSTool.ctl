VERSION 5.00
Begin VB.UserControl Tool 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TCSTool.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   480
   ToolboxBitmap   =   "TCSTool.ctx":0312
End
Attribute VB_Name = "Tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Enumerations
' ------------

Public Enum TCSTOOLERRORS
  ERR_CREATETOOLTIPWINDOW
  ERR_ADDTOOL
  ERR_REMOVETOOL
End Enum

Public Enum ttDelayTimeConstants
  ttDefault = TTDT_AUTOMATIC '= 0
  ttInitial = TTDT_INITIAL   '= 3 Defaults to 500ms
  ttShow = TTDT_AUTOPOP      '= 2 Defaults to 5000ms
  ttReshow = TTDT_RESHOW     '= 1 Defaults to 100ms
  ttMask = 3
End Enum

Public Enum ttStyle
  Normal = 0
  Balloon
End Enum

' Private member variables
' ------------------------

Private m_hwndTV As Long
Private m_hwndTT As Long
Private m_lngMaxTip As Long
Private m_Style As ttStyle
Private m_bkColor As OLE_COLOR
Private m_txtColor As OLE_COLOR
Private m_Left As Long
Private m_Right As Long
Private m_Top As Long
Private m_Bottom As Long
Private m_Width As Long


' Public property get and lets
' ----------------------------
Public Property Get hWndTT() As Long
Attribute hWndTT.VB_Description = "Read only value of tooltip window handle"
  hWndTT = m_hwndTT
End Property

Public Property Get hWndTreeView() As Long
Attribute hWndTreeView.VB_Description = "Treeview window handle - need to assign before creating window"
  hWndTreeView = m_hwndTV
End Property

Public Property Let hWndTreeView(ByVal NewValue As Long)
  m_hwndTV = NewValue
End Property

Public Property Get TooltipStyle() As ttStyle
Attribute TooltipStyle.VB_Description = "Sets tooltip window style (balloon style requires comctl32.dll version 5.80 or higher"
  TooltipStyle = m_Style
End Property

Public Property Let TooltipStyle(ByVal NewValue As ttStyle)
  m_Style = NewValue
  Call PropertyChanged("TooltipStyle")
End Property

Public Property Get bkColor() As OLE_COLOR
Attribute bkColor.VB_Description = "Sets the background colour of tooltip window (set as OLE_COLOR)"
  If (m_hwndTT = 0) Then Exit Property
  If m_bkColor = 0 Then
    bkColor = SendMessage(m_hwndTT, TTM_GETTIPBKCOLOR, 0, 0)
  Else
    bkColor = m_bkColor
  End If
End Property

Public Property Let bkColor(ByVal NewValue As OLE_COLOR)
  If (m_hwndTT = 0) Then Exit Property
  Call SendMessage(m_hwndTT, TTM_SETTIPBKCOLOR, NewValue, 0)
  m_bkColor = NewValue
  Call PropertyChanged("bkColor")
End Property

Public Property Get DelayTime(ByVal dwType As ttDelayTimeConstants) As Long
Attribute DelayTime.VB_Description = "Sets the initial display, duration and reshow times for tooltip window in milliseconds"
  If (m_hwndTT = 0) Then Exit Property
  DelayTime = SendMessage(m_hwndTT, TTM_GETDELAYTIME, (dwType And ttMask), 0&)
End Property

Public Property Let DelayTime(ByVal dwType As ttDelayTimeConstants, ByVal NewValue As Long)
  If (m_hwndTT = 0) Then Exit Property
  Call SendMessage(m_hwndTT, TTM_SETDELAYTIME, (dwType And ttMask), NewValue)  ' no rtn val
End Property

Public Property Get txtColor() As OLE_COLOR
Attribute txtColor.VB_Description = "Sets the text colour for the tooltip window (as OLE_COLOR)"
  If (m_hwndTT = 0) Then Exit Property
  If m_txtColor = 0 Then
    txtColor = SendMessage(m_hwndTT, TTM_GETTIPTEXTCOLOR, 0, 0)
  Else
    txtColor = m_txtColor
  End If
End Property

Public Property Let txtColor(ByVal NewValue As OLE_COLOR)
  If (m_hwndTT = 0) Then Exit Property
  Call SendMessage(m_hwndTT, TTM_SETTIPTEXTCOLOR, NewValue, 0)
  m_txtColor = NewValue
  Call PropertyChanged("txtColor")
End Property

Public Property Get MarginLeft() As Long
Attribute MarginLeft.VB_Description = "Sets the left margin of the tooltip window (in twips)"
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelX
  If (m_hwndTT = 0) Then Exit Property
  If m_Left = 0 Then
    Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
    MarginLeft = (r.Left) * i
  Else
    MarginLeft = m_Left
  End If
End Property

Public Property Let MarginLeft(ByVal NewValue As Long)
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelX
  If (m_hwndTT = 0) Then Exit Property
  Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
  r.Left = NewValue / i
  Call SendMessage(m_hwndTT, TTM_SETMARGIN, 0, r)
  m_Left = NewValue / i
  Call PropertyChanged("MarginLeft")
  
End Property

Public Property Get MarginRight() As Long
Attribute MarginRight.VB_Description = "Sets the right margin of the tooltip window (in twips)"
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelX
  If (m_hwndTT = 0) Then Exit Property
  If m_Right = 0 Then
    Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
    MarginRight = (r.Right) * i
  Else
    MarginRight = m_Right
  End If
End Property

Public Property Let MarginRight(ByVal NewValue As Long)
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelX
  If (m_hwndTT = 0) Then Exit Property
  Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
  r.Right = NewValue / i
  Call SendMessage(m_hwndTT, TTM_SETMARGIN, 0, r)
  m_Right = NewValue / i
  Call PropertyChanged("MarginRight")

End Property

Public Property Get MarginTop() As Long
Attribute MarginTop.VB_Description = "Sets the top margin of the tooltip window (in twips)"
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelY
  If (m_hwndTT = 0) Then Exit Property
  If m_Top = 0 Then
    Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
    MarginTop = (r.Top) * i
  Else
    MarginTop = m_Top
  End If
End Property

Public Property Let MarginTop(ByVal NewValue As Long)
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelY
  If (m_hwndTT = 0) Then Exit Property
  Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
  r.Top = NewValue / i
  Call SendMessage(m_hwndTT, TTM_SETMARGIN, 0, r)
  m_Top = NewValue / i
  Call PropertyChanged("MarginTop")
  
End Property

Public Property Get MarginBottom() As Long
Attribute MarginBottom.VB_Description = "Sets the bottom margin of the tooltip window (in twips)"
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelY
  If (m_hwndTT = 0) Then Exit Property
  If m_Bottom = 0 Then
    Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
    MarginBottom = (r.Bottom) * i
  Else
    MarginBottom = m_Bottom
  End If
End Property

Public Property Let MarginBottom(ByVal NewValue As Long)
  Dim r As RECT
  Dim i As Long
  
  i = Screen.TwipsPerPixelY
  If (m_hwndTT = 0) Then Exit Property
  Call SendMessage(m_hwndTT, TTM_GETMARGIN, 0, r)
  r.Bottom = NewValue / i
  Call SendMessage(m_hwndTT, TTM_SETMARGIN, 0, r)
  m_Bottom = NewValue / i
  Call PropertyChanged("MarginBottom")
  
End Property

Public Property Get TooltipWidth() As Long
Attribute TooltipWidth.VB_Description = "Set the maximum tooltip window width (in twips)"
  Dim i As Long
  
  If (m_hwndTT = 0) Then Exit Property
  
  i = Screen.TwipsPerPixelX
  If m_Width = 0 Then
    TooltipWidth = (LOWORD(SendMessage(m_hwndTT, TTM_GETMAXTIPWIDTH, 0, 0))) * i
  Else
    TooltipWidth = m_Width
  End If
End Property

Public Property Let TooltipWidth(ByVal NewValue As Long)
  Dim i As Long
  
  If (m_hwndTT = 0) Then Exit Property
  If (NewValue < 1) Then NewValue = -1
  
  i = Screen.TwipsPerPixelX
  Call SendMessage(m_hwndTT, TTM_SETMAXTIPWIDTH, 0, NewValue / i)
  m_Width = NewValue / i
  Call PropertyChanged("TooltipWidth")
  
End Property

Public Property Get ToolCount() As Integer
Attribute ToolCount.VB_Description = "Count of tools maintained by tooltip control"
  If (m_hwndTT = 0) Then Exit Property
  
  ToolCount = SendMessage(m_hwndTT, TTM_GETTOOLCOUNT, 0, 0)
End Property

Public Property Get ToolText(ByVal ctrl As Object) As String
Attribute ToolText.VB_Description = "Sets the tooltip text - can use vbCrLf to create new line"
  Dim ti As TOOLINFO
  Dim ctl As Control
  
  Set ctl = ctrl
  If (m_hwndTT = 0) Then Exit Property
  
  If GetToolInfo(ctl.hWnd, ti, True) Then
    ToolText = GetStrFromBufferA(ti.lpszText)
  End If

End Property

Public Property Let ToolText(ByVal ctrl As Object, ByVal NewValue As String)
  Dim ti As TOOLINFO
  Dim ctl As Control
  
  Set ctl = ctrl
  If (m_hwndTT = 0) Then Exit Property
  
  If GetToolInfo(ctl.hWnd, ti) Then
    ti.lpszText = NewValue
    ' Set the buffer size to the length of text
    m_lngMaxTip = Max(m_lngMaxTip, Len(NewValue) + 1)
    ' The tooltip won't appear for the control if lpszText is an empty string
    Call SendMessage(m_hwndTT, TTM_UPDATETIPTEXT, 0, ti)
  End If
  
End Property


' Public functions
' ----------------

Public Sub Create(ByVal frm As Object)
Attribute Create.VB_Description = "Creates tooltip window (if the computer has comctl32.dll version 5.80 or higher, will be balloon tooltips)"
  Dim f As Form
  Dim lngStyle As Long
  Dim dvi As DLLVersionInfo, DLLMajVer As Long, DLLMinVer As Long, DLLBuild As Long, DLLPlatID As Long
  
  On Error GoTo CreateErr
  
  Set f = frm
  If (m_hwndTT = 0) Then
    Call InitCommonControls
    
    If Len(m_hwndTV) <> 0 Then
      lngStyle = GetWindowLong(m_hwndTV, GWL_STYLE)
      Call SetWindowLong(m_hwndTV, GWL_STYLE, lngStyle Or TVS_NOTOOLTIPS)
    End If

    ' Get DLL version number
    dvi.cbSize = Len(dvi)
    Call GetComctl32Version(dvi)
    DLLMajVer = dvi.dwMajorVersion
    DLLMinVer = dvi.dwMinorVersion
    DLLPlatID = dvi.dwPlatformID
    DLLBuild = dvi.dwBuildNumber

    If DLLMajVer >= 5 And DLLMinVer >= 80 And m_Style = Balloon Then
      m_hwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP Or TTS_NOPREFIX Or TTS_BALLOON, 0, 0, 0, 0, f.hWnd, 0, App.hInstance, ByVal 0)
    Else
      m_hwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, f.hWnd, 0, App.hInstance, ByVal 0)
    End If
    
  End If

CreateEnd:
  Exit Sub

CreateErr:
  Call Err.Raise(ERR_CREATETOOLTIPWINDOW, "Create", "Error creating the tooltip window", Err.HelpFile, Err.HelpContext)
  Resume CreateEnd
End Sub

Public Sub AddTool(ByVal ctrl As Object, Optional ByVal strText As String)
Attribute AddTool.VB_Description = "Registers tool with tooltip control"
  Dim ti As TOOLINFO
  Dim ctl As Control
  Dim hwndCtrl As Long
  
  On Error GoTo AddToolErr
  
  Set ctl = ctrl
  If (m_hwndTT = 0) Then Exit Sub
  
  hwndCtrl = ctl.hWnd
  If (GetToolInfo(hwndCtrl, ti) = False) Then
    With ti
      .cbSize = Len(ti)
      ' TTF_IDISHWND must be specified to tell the tooltip control
      ' to retrieve the control's rect from it's hWnd specified in uId.
      .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
      .hWnd = ctl.Container.hWnd
      .uId = hwndCtrl
      
      If Len(strText) Then
        .lpszText = strText
      Else
        '.lpszText = "Tool" & ToolCount + 1 'pc 16/11
        .lpszText = ""
      End If
      
      ' Maintain the maximun tip text length for GetToolInfo
      m_lngMaxTip = Max(m_lngMaxTip, Len(.lpszText) + 1)
    
    End With
    
    Call SendMessage(m_hwndTT, TTM_ADDTOOL, 0, ti)
  End If
  
AddToolEnd:
  Exit Sub

AddToolErr:
  Call Err.Raise(ERR_ADDTOOL, "AddTool", "Error adding tool to the tooltip control", Err.HelpFile, Err.HelpContext)
  Resume AddToolEnd

End Sub

Public Sub RemoveTool(ByVal ctrl As Object)
Attribute RemoveTool.VB_Description = "Removes tool from tooltip control"
  Dim ti As TOOLINFO
  Dim ctl As Control
  
  On Error GoTo RemoveToolErr
  
  Set ctl = ctrl
  If (m_hwndTT = 0) Then Exit Sub
  
  If GetToolInfo(ctl.hWnd, ti) Then
    Call SendMessage(m_hwndTT, TTM_DELTOOL, 0, ti)
  End If

RemoveToolEnd:
  Exit Sub

RemoveToolErr:
  Call Err.Raise(ERR_REMOVETOOL, "RemoveTool", "Error removing the tool from the tooltip control", Err.HelpFile, Err.HelpContext)
  Resume RemoveToolEnd
  
End Sub


' Private functions
' -----------------
Private Function GetToolInfo(ByVal hwndTool As Long, ti As TOOLINFO, Optional ByVal blnGetText As Boolean = False) As Boolean
  Dim nItems As Integer
  Dim i As Integer
  
  ' Must set the ti size before sending the message below
  ti.cbSize = Len(ti)
  ' Fill the buffer with 0s
  If blnGetText Then ti.lpszText = String$(m_lngMaxTip, 0)
    
  nItems = ToolCount
  
  For i = 0 To nItems - 1
    ' Returns 1 on success, 0 on failure
    ' i is zero-based index for the enumerated tools
    If SendMessage(m_hwndTT, TTM_ENUMTOOLS, (i), ti) Then
      If (hwndTool = ti.uId) Then
        GetToolInfo = True
        Exit Function
      End If
    End If
  Next

End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    m_bkColor = .ReadProperty("bkColor", RGB(&HFF, &HFF, &HE1))
    m_Bottom = .ReadProperty("Bottom", 0)
    m_Left = .ReadProperty("Left", 0)
    m_Right = .ReadProperty("Right", 0)
    m_Style = .ReadProperty("Style", 0)
    m_Top = .ReadProperty("Top", 0)
    m_txtColor = .ReadProperty("txtColor", RGB(0, 0, 0))
    m_Width = .ReadProperty("Width", -1)
  End With
End Sub

Private Sub UserControl_Terminate()
  If m_hwndTT Then Call DestroyWindow(m_hwndTT)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("bkColor", m_bkColor, RGB(&HFF, &HFF, &HE1))
    Call .WriteProperty("Bottom", m_Bottom, 0)
    Call .WriteProperty("Left", m_Left, 0)
    Call .WriteProperty("Right", m_Right, 0)
    Call .WriteProperty("Style", m_Style, 0)
    Call .WriteProperty("Top", m_Top, 0)
    Call .WriteProperty("txtColor", m_txtColor, RGB(0, 0, 0))
    Call .WriteProperty("Width", m_Width, -1)
  End With
End Sub

