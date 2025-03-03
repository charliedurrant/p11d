VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D08C90A4-2337-4BE1-8137-EB1A093571A4}#1.0#0"; "atc2dmenu.ocx"
Begin VB.UserControl CodeEditor 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin atc2dmenu.DMenu dmenu 
      Left            =   960
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin RichTextLib.RichTextBox rtMultiLine 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"CodeEditor.ctx":0000
   End
   Begin RichTextLib.RichTextBox rtSingleLine 
      Height          =   3615
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"CodeEditor.ctx":0081
   End
End
Attribute VB_Name = "CodeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum CURSOR_ACTION
  CA_SELECT_LEFT
  CA_SELECT_RIGHT
  CA_MOVE_LEFT
  CA_MOVE_RIGHT
  CA_NONE
End Enum

Private Type BLOCK_DEL
  LineTextCurrent As String
  StartSelPos As Long
  InShift As Boolean
  EraseKeyPressed As Boolean
  SelLength As Long
End Type

Private Enum SEL_COLOR_SET
  SCS_BLUE
  SCS_NORMAL
End Enum

Private BD As BLOCK_DEL
Private m_dirty  As Boolean
Private m_codes As StringList
Private m_vbm As VBMenu
Private m_MultiLine As Boolean

Private Declare Function EbMode Lib "vba6" () As Long ' 0=Design, 1=Run, 2=Break.

Event MenuClick(ByVal vbm As VBMenu, ByVal vbmi As VBMenuItem)

Private Sub UserControl_Initialize()
  Set m_codes = New StringList
  Set m_vbm = dmenu.Add("Menu")
End Sub
Friend Sub RaiseMenuClick(ByVal vbm As VBMenu, ByVal vbmi As VBMenuItem)
  Dim codeAdded As String
  codeAdded = ""
  RaiseEvent MenuClick(vbm, vbmi)
End Sub
Public Sub InsertCodeIntoText(ByVal Code As String)
  rt.SelText = Code
  Dirty = True
  Call ColorCode(Code, True, "")
End Sub
Private Sub dmenu_MenuClick(ByVal vbm As atc2dmenu.VBMenu, ByVal vbmi As atc2dmenu.VBMenuItem)
  Call RaiseMenuClick(vbm, vbmi)
End Sub
Public Property Let FormHWnd(ByVal value As Long)
  dmenu.hwnd = value
  
End Property
Public Sub AddCode(ByVal value As String)
  m_codes.Add (value)
End Sub
Public Property Get Menu() As VBMenu
  Set Menu = m_vbm
End Property
Private Sub ColorCodes()
  Dim i As Long
  Dim txt As String
  
  On Error GoTo ColorCodes_ERR
  
  Call xSet("ColorCodes")
  txt = Me.Text
  
  For i = 1 To m_codes.count
    Call ColorCode(m_codes.Item(i), False, txt)
  Next
  rt.SelStart = 0
  Call SetSelTextProperties(SCS_NORMAL)
  
ColorCodes_END:
  Call xReturn("ColorCodes")
  Exit Sub
ColorCodes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ColorCodes", "Color Codes", "Error setting the color of the control codes in the current employee letter file.")
  Resume ColorCodes_END
End Sub
Private Sub ColorCode(sCode As String, bFromMenu As Boolean, richTextBoxText As String)
  Dim l As Long
  
  On Error GoTo ColorCode_ERR
  
  Call xSet("ColorCode")
    
  If bFromMenu Then
    rt.SelStart = rt.SelStart - Len(sCode)
    rt.SelLength = Len(sCode)
    Call SetSelTextProperties(SCS_BLUE)
    rt.SelStart = rt.SelStart + rt.SelLength
  Else
    'bug with RT does not do ignore case for loop twice?
    Call ColorCodeEx(sCode, richTextBoxText)
  End If
    
  
ColorCode_END:
  Call SetSelTextProperties(SCS_NORMAL)
  Call xReturn("ColorCode")
  Exit Sub
ColorCode_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ColorCode", "Color Code", "Error setting the color of the code " & sCode & " in the employee letter file.")
  Resume ColorCode_END
  
  
End Sub
Private Sub ColorCodeEx(sCode As String, richTextBoxText As String)
  Dim l As Long
  
  On Error GoTo ColorCodeEx_ERR
  
  Call xSet("ColorCodeEx")
  
  
  l = 1
  Do
    l = InStr(l, richTextBoxText, sCode, vbTextCompare)
    If l <> 0 Then
      rt.SelStart = l - 1
      rt.SelLength = Len(sCode)
      Call SetSelTextProperties(SCS_BLUE)
      rt.SelStart = rt.SelStart + Len(sCode)
      rt.SelLength = 0
      Call SetSelTextProperties(SCS_NORMAL)
      l = l + Len(sCode)
    Else
      Exit Do
    End If
  Loop While True
  
  GoTo ColorCodeEx_END
  Do
    l = rt.Find(sCode, l, , rtfWholeWord)
    
    If l <> -1 Then
      rt.SelStart = l
      rt.SelLength = Len(sCode)
      Call SetSelTextProperties(SCS_BLUE)
      rt.SelStart = rt.SelStart + Len(sCode)
      rt.SelLength = 0
      Call SetSelTextProperties(SCS_NORMAL)
      l = l + 1
    Else
      Exit Do
    End If
  Loop While True
    
  
ColorCodeEx_END:
  Call SetSelTextProperties(SCS_NORMAL)
  Call xReturn("ColorCodeEx")
  Exit Sub
ColorCodeEx_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ColorCodeEx", "Color Code Ex", "Error setting the color of the code " & sCode & " in the employee letter file.")
  Resume ColorCodeEx_END
End Sub

Private Function GetCursorAction(KeyCode As Integer) As CURSOR_ACTION
  On Error GoTo GetCursorAction_ERR
  
  Call xSet("GetCursorAction")

  If BD.EraseKeyPressed And BD.SelLength > 0 Then
    GetCursorAction = CA_NONE
    Exit Function
  End If
  
  If Not BD.InShift Then
    Select Case KeyCode
      Case vbKeyLeft, vbKeyUp, vbKeyBack
        GetCursorAction = CA_MOVE_LEFT
      Case vbKeyRight, vbKeyDown, vbKeyDelete
        GetCursorAction = CA_MOVE_RIGHT
    End Select
  Else
    Select Case rt.SelStart
      Case BD.StartSelPos
        If rt.SelLength > 0 Then
          GetCursorAction = CA_SELECT_RIGHT
        End If
      Case Is < BD.StartSelPos
        GetCursorAction = CA_SELECT_LEFT
    End Select
  End If
  
GetCursorAction_END:
  Call xReturn("GetCursorAction")
  Exit Function
GetCursorAction_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Get Cursor Action", "Get Cursor Action", "Error getting the cursor action")
  Resume GetCursorAction_END
End Function

Private Sub SetSelTextProperties(SCS As SEL_COLOR_SET)
  On Error GoTo SetSelTextProperties_ERR
  
  Call xSet("SetSelTextProperties")

  Select Case SCS
    Case SCS_BLUE
      rt.SelBold = True
      rt.SelColor = vbBlue
    Case SCS_NORMAL
      rt.SelItalic = False
      rt.SelBold = False
      rt.SelColor = vbBlack
  End Select
  
SetSelTextProperties_END:
  Call xReturn("SetSelTextProperties")
  Exit Sub
SetSelTextProperties_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetSelTextProperties", "Set SelText Properties", "Error setting the seltext font properties.")
  Resume SetSelTextProperties_END
End Sub

Private Sub UserControl_Resize()
  rtMultiLine.Top = 0
  rtMultiLine.Left = 0
  rtMultiLine.width = UserControl.width
  rtMultiLine.height = UserControl.height
  
  rtSingleLine.Top = 0
  rtSingleLine.Left = 0
  rtSingleLine.width = UserControl.width
  rtSingleLine.height = UserControl.height
End Sub


Private Sub rtSingleLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call RichTextBox_MouseUp
End Sub

Private Sub rtMultiLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call RichTextBox_MouseUp
End Sub
Private Sub RichTextBox_MouseUp()
   Dim lStartBrace As Long
  
  If rt.SelLength = 0 Then
    Call SelCodes(-1, CA_MOVE_LEFT)
  Else
    If InsideBrace(lStartBrace, 0, BD.LineTextCurrent, rt.SelStart + 1) Then
      rt.SelLength = 0
      rt.SelStart = lStartBrace - 1
    ElseIf InsideBrace(lStartBrace, 0, BD.LineTextCurrent, (rt.SelStart + rt.SelLength + 1)) Then
      rt.SelLength = 0
      rt.SelStart = lStartBrace - 1
    End If
  End If
End Sub
Private Sub rtSingleLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RichTextBox_MouseDown(Button, X, Y)
End Sub
Private Sub rtMultiLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RichTextBox_MouseDown(Button, X, Y)
End Sub
Private Sub RichTextBox_MouseDown(Button As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    Call m_vbm.Popup(S_ELMC_MASTER, X, Y)
  Else
    Call RecordKeyDown(-1, 0)
  End If
End Sub

Private Sub rtMultiLine_KeyUp(KeyCode As Integer, Shift As Integer)
 Call RichTextBox_KeyUp(KeyCode)
End Sub
Private Sub rtSingleLine_KeyUp(KeyCode As Integer, Shift As Integer)
 Call RichTextBox_KeyUp(KeyCode)
End Sub
Private Sub RichTextBox_KeyUp(KeyCode As Integer)
  Call SelCodes(KeyCode, GetCursorAction(KeyCode))
End Sub
Private Sub rtMultiLine_KeyDown(KeyCode As Integer, Shift As Integer)
  Call RichTextBox_KeyDown(KeyCode, Shift)
End Sub
Private Sub rtSingleLine_KeyDown(KeyCode As Integer, Shift As Integer)
  Call RichTextBox_KeyDown(KeyCode, Shift)
End Sub

Private Sub RichTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
  Dirty = True
  If KeyCode = 221 Or KeyCode = 219 Then '{}
    KeyCode = 0
  End If
  Call RecordKeyDown(KeyCode, Shift)
  If rt.SelLength = 0 Then Call SetSelTextProperties(SCS_NORMAL)
End Sub
Private Sub SelCodes(KeyCode As Integer, CA As CURSOR_ACTION)
  Dim lStartBrace As Long, lEndBrace As Long
  
  On Error GoTo SelCodes_ERR
  
  Call xSet("SelCodes")
  
  Select Case CA
    Case CA_NONE
    Case CA_MOVE_RIGHT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rt.SelStart + 1 + Abs(BD.EraseKeyPressed)) Then
        If BD.EraseKeyPressed Then
          rt.SelStart = lStartBrace - 1
          rt.SelLength = lEndBrace - lStartBrace
          rt.SelText = ""
        Else
          rt.SelStart = lEndBrace
        End If
      End If
    Case CA_MOVE_LEFT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rt.SelStart + 1) Then
          If BD.EraseKeyPressed Then
            rt.SelStart = lStartBrace - 1
            rt.SelLength = lEndBrace - lStartBrace
            rt.SelText = ""
          Else
            rt.SelStart = lStartBrace - 1
          End If
      End If
    Case CA_SELECT_RIGHT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rt.SelStart + 1 + rt.SelLength) Then
        Select Case KeyCode
          Case vbKeyLeft, vbKeyUp
            rt.SelLength = (lStartBrace - 1) - rt.SelStart
          Case vbKeyRight, vbKeyDown
            rt.SelLength = lEndBrace - rt.SelStart
        End Select
      End If
    Case CA_SELECT_LEFT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rt.SelStart + 1) Then
        Select Case KeyCode
          Case vbKeyLeft, vbKeyUp
              rt.SelStart = lStartBrace - 1
              rt.SelLength = BD.StartSelPos - (lStartBrace - 1)
              BD.InShift = False
          Case vbKeyRight, vbKeyDown
              rt.SelStart = lEndBrace - 1
              rt.SelLength = 0
        End Select
      End If
  End Select

SelCodes_END:
  Call xReturn("SelCodes")
  Exit Sub
SelCodes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SelCodes", "Sel Codes", "Error selecting an employee letter control code.")
  Resume SelCodes_END
  
End Sub

Private Function InsideBrace(lStartBrace As Long, lEndBrace As Long, sTextToSearch As String, lStartPos As Long) As Boolean
  Dim l As Long, m As Long, n As Long, o As Long
  
  On Error GoTo InsideBrace_ERR
  
  Call xSet("InsideBrace")
  
  lStartBrace = 0
  lEndBrace = 0
  
  If Len(sTextToSearch) Then
    l = InStr(lStartPos, sTextToSearch, "}", vbTextCompare)
    If l > 0 Then
      m = InStr(lStartPos, sTextToSearch, "{", vbTextCompare)
      If m = 0 Or m > l Then
        n = InStrRev(sTextToSearch, "{", lStartPos, vbTextCompare)
        If n > 0 Then
          o = InStrRev(sTextToSearch, "}", lStartPos - 1, vbTextCompare)
          If (o = 0) Or o > 0 And o < n Then
            lStartBrace = n
            lEndBrace = l
            InsideBrace = True
          End If
        End If
      End If
    Else
      lStartBrace = 0
      lEndBrace = 0
    End If
  End If

InsideBrace_END:
  Call xReturn("InsideBrace")
  Exit Function
InsideBrace_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "InsideBrace", "Inside Brace", "Error determining whether the caret is inside a set of braces.")
  Resume InsideBrace_END
End Function
Private Function rt() As RichTextBox
  If (Me.MultiLine) Then
    Set rt = rtMultiLine
  Else
    Set rt = rtSingleLine
  End If
End Function
Private Sub RecordKeyDown(KeyCode As Integer, Shift As Integer)

  On Error GoTo RecordKeyDown_ERR
  
  Call xSet("RecordKeyDown")
     
  BD.LineTextCurrent = rt.Text
  BD.SelLength = rt.SelLength
  
  If (Shift And vbShiftMask) Then
    If Not BD.InShift Then
      BD.InShift = True
      BD.StartSelPos = rt.SelStart
    End If
  Else
    BD.InShift = False
    BD.StartSelPos = -1
  End If
  If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    BD.EraseKeyPressed = True
  Else
    BD.EraseKeyPressed = False
  End If
  
RecordKeyDown_END:
  Call xReturn("RecordKeyDown")
  Exit Sub
RecordKeyDown_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RecordKeyDown", "Record Key Down", "Error recording the keydown in the employee letter.")
  Resume RecordKeyDown_END
End Sub
Public Property Get SelLength() As Long
  SelLength = rt.SelLength
End Property
Public Property Let SelLength(ByVal value As Long)
  rt.SelLength = value
End Property
Public Property Get SelText() As String
  SelText = rt.SelText
End Property
Public Property Let SelText(ByVal value As String)
  rt.SelText = value
End Property
Public Property Get SelStart() As Long
  SelLength = rt.SelStart
End Property
Public Property Let SelStart(ByVal value As Long)
  rt.SelStart = value
End Property
Public Property Set Font(ByVal value As StdFont)
  Set rt.Font = value
End Property
Public Property Get Font() As StdFont
  Set Font = rt.Font
End Property
Public Property Let Text(ByVal value As String)
  rt.Text = value
  m_dirty = True
  Call ColorCodes
  
End Property
Public Property Get Text() As String
  Text = rt.Text
End Property
Public Property Get Dirty() As Boolean
  Dirty = m_dirty
End Property
Public Property Let Dirty(ByVal value As Boolean)
  m_dirty = value
End Property
Public Property Get MultiLine() As Boolean
  MultiLine = m_MultiLine
End Property
Public Property Let MultiLine(ByVal value As Boolean)
  m_MultiLine = value
  rtSingleLine.Visible = Not Me.MultiLine
  rtMultiLine.Visible = Me.MultiLine
  
  Call PropertyChanged("MultiLine")
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Me.MultiLine = PropBag.ReadProperty("MultiLine", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("MultiLine", Me.MultiLine, True)
End Sub

