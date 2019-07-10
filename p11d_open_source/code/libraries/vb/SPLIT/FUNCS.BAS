Attribute VB_Name = "Funcs"
Option Explicit

Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As Long, lpInitData As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal pTx As Long, ByVal pTy As Long) As Long


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const SRCINVERT = &H660046
Public Const GW_CHILD = 5
Public Const HS_DIAGCROSS = 5
Public Const PS_NULL = 5
Public Const GWL_WNDPROC = (-4)
Public Const WM_SIZE = &H5

Public Enum SPLIT_MOUSE_MODE
  SMM_DOWN
  SMM_MOVE
  SMM_UP
End Enum


Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type SIZEL
  cx As Long
  cy As Long
End Type

Private Type SPLIT_SUBCLASS
  OldWindLong As Long
  hWndParent As Long
  hWnd As Long
  SPLIT As SPLIT
End Type

Public Type SPLIT_INFO
  hWndClient As Long
  hWndParent As Long
  pTStartPosRelSplit As POINTAPI
  SplitRect As RECT
  StartSplitRect As RECT
  SplitDimensions As SIZEL
  LastSplitRect As RECT
  MousePosition As POINTAPI
  XMouseMin As Long
  XMouseMax As Long
  XMouseMinUser As Long
  XMouseMaxUser As Long
  hDCScreen As Long
End Type

Private m_SplitSubClass() As SPLIT_SUBCLASS
Private m_NoOfControlsUbound As Long

Public Sub AddControl(TSP As Object)
  Dim i As Long
  Dim bFound As Boolean
  Dim lIndex As Long
  
  If m_NoOfControlsUbound > 0 Then
  
    For i = 1 To m_NoOfControlsUbound
      If m_SplitSubClass(i).SPLIT Is TSP Then
        bFound = True
        Exit For
      End If
    Next
    
    For i = 1 To m_NoOfControlsUbound
      If Not bFound Then
        If m_SplitSubClass(i).SPLIT Is Nothing Then
          Set m_SplitSubClass(i).SPLIT = TSP
          lIndex = i
          bFound = True
          Exit For
        End If
      End If
    Next
  End If
  
  If Not bFound Then
    m_NoOfControlsUbound = m_NoOfControlsUbound + 1
    ReDim Preserve m_SplitSubClass(1 To m_NoOfControlsUbound)
    Set m_SplitSubClass(m_NoOfControlsUbound).SPLIT = TSP
    lIndex = m_NoOfControlsUbound
  End If
  
  If lIndex > 0 Then
    m_SplitSubClass(lIndex).hWndParent = m_SplitSubClass(lIndex).SPLIT.hWndParent
    m_SplitSubClass(lIndex).OldWindLong = SetWindowLong(m_SplitSubClass(lIndex).hWndParent, GWL_WNDPROC, AddressOf MyWindowProc)
  End If
  
End Sub

Public Sub RemoveControl(TSP As SPLIT)
  Dim i As Long
  
  For i = 1 To m_NoOfControlsUbound
    If m_SplitSubClass(i).SPLIT Is TSP Then
      Set m_SplitSubClass(i).SPLIT = Nothing
      Call SetWindowLong(m_SplitSubClass(i).hWndParent, GWL_WNDPROC, m_SplitSubClass(i).OldWindLong)
      Exit For
    End If
  Next
  
End Sub

Public Function MyWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim i As Long
  Dim j As Long
  For i = 1 To m_NoOfControlsUbound
    If uMsg = WM_SIZE Then
      Debug.Print "hwnd = " & m_SplitSubClass(i).hWndParent
      For j = 1 To m_NoOfControlsUbound
        Debug.Print m_SplitSubClass(i).hWndParent
      Next
    End If
    If hWnd = m_SplitSubClass(i).hWndParent Then
      MyWindowProc = CallWindowProc(m_SplitSubClass(i).OldWindLong, hWnd, uMsg, wParam, lParam)
      Select Case uMsg
        Case WM_SIZE
          m_SplitSubClass(i).SPLIT.SizeMeToParent
      End Select
      Exit For
    End If
  Next
  
End Function
Public Function ConvertCoords(x As Single, y As Single) As POINTAPI
  ConvertCoords.x = x / Screen.TwipsPerPixelX
  ConvertCoords.y = y / Screen.TwipsPerPixelY
End Function
Public Function TranslateCoordsEx(xTwips As Single, yTwips As Single, hWndFrom As Long, hWndTo As Long) As POINTAPI
  Dim pT As POINTAPI
  Dim lRet As Long
  
  pT = ConvertCoords(xTwips, yTwips)
  lRet = MapWindowPoints(hWndFrom, hWndTo, pT.x, 1)
  TranslateCoordsEx = pT
End Function

Public Function RectDimensions(r As RECT) As SIZEL
  RectDimensions.cx = r.Right - r.Left
  RectDimensions.cy = r.Bottom - r.Top
End Function

