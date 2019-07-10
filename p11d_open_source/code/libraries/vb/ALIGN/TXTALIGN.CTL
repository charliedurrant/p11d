VERSION 5.00
Begin VB.UserControl TXTAlign 
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ScaleHeight     =   2085
   ScaleWidth      =   4155
   ToolboxBitmap   =   "txtAlign.ctx":0000
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "TXTAlign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:

'Property Variables:
Private m_TXT As TextBox
Attribute m_TXT.VB_VarHelpID = -1
'Event Declarations:
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=m_TXT,txt,0,OLEStartDrag
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=m_TXT,txt,0,OLESetData
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=m_TXT,txt,0,OLEGiveFeedback
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=m_TXT,txt,0,OLEDragOver
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=m_TXT,txt,0,OLEDragDrop
Event OLECompleteDrag(Effect As Long) 'MappingInfo=m_TXT,txt,0,OLECompleteDrag
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=m_TXT,txt,0,MouseUp
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=m_TXT,txt,0,MouseMove
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=m_TXT,txt,0,MouseDown
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=m_TXT,txt,0,KeyUp
Event KeyPress(KeyAscii As Integer) 'MappingInfo=m_TXT,txt,0,KeyPress
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=m_TXT,txt,0,KeyDown
Event Click() 'MappingInfo=m_TXT,txt,0,Click
Event Change() 'MappingInfo=m_TXT,txt,0,Change
Event DblClick() 'MappingInfo=m_TXT,txt,0,DblClick

Public Enum txtAlignment
  TXT_LEFT = 0
  TXT_CENTRE = 1
  TXT_RIGHT = 2
End Enum

Private Const L_NO_TXT_ALIGNS = 3

Dim m_txtAlign As txtAlignment
Const m_def_txtAlign = txtAlignment.TXT_LEFT



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
  WhatsThisHelpID = m_TXT.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
  m_TXT.WhatsThisHelpID() = New_WhatsThisHelpID
  PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
  ToolTipText = m_TXT.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
  m_TXT.ToolTipText() = New_ToolTipText
  PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
  Text = m_TXT.Text
  
End Property

Public Property Let Text(ByVal New_Text As String)
  m_TXT.Text() = New_Text
  PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
  SelText = m_TXT.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
  m_TXT.SelText() = New_SelText
  PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
  SelStart = m_TXT.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
  m_TXT.SelStart() = New_SelStart
  PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
  SelLength = m_TXT.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
  m_TXT.SelLength() = New_SelLength
  PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
  ScrollBars = m_TXT.ScrollBars
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
  RightToLeft = m_TXT.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
  m_TXT.RightToLeft() = New_RightToLeft
  PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  m_TXT.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
  PasswordChar = m_TXT.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
  m_TXT.PasswordChar() = New_PasswordChar
  PropertyChanged "PasswordChar"
End Property

Private Sub txt_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
  RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub txt_OLESetData(Index As Integer, Data As DataObject, DataFormat As Integer)
  RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub txt_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
  RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
  OLEDropMode = m_TXT.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
  m_TXT.OLEDropMode() = New_OLEDropMode
  PropertyChanged "OLEDropMode"
End Property

Private Sub txt_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
  RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,OLEDragMode
Public Property Get OLEDragMode() As Integer
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
  OLEDragMode = m_TXT.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
  m_TXT.OLEDragMode() = New_OLEDragMode
  PropertyChanged "OLEDragMode"
End Property

Private Sub txt_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
  m_TXT.OLEDrag
End Sub

Private Sub txt_OLECompleteDrag(Index As Integer, Effect As Long)
  RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
  MultiLine = m_TXT.MultiLine
End Property

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
  MousePointer = m_TXT.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
  m_TXT.MousePointer() = New_MousePointer
  PropertyChanged "MousePointer"
End Property

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
  Set MouseIcon = m_TXT.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
  Set m_TXT.MouseIcon = New_MouseIcon
  PropertyChanged "MouseIcon"
End Property

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
  MaxLength = m_TXT.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
  m_TXT.MaxLength() = New_MaxLength
  PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
  Locked = m_TXT.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  m_TXT.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "Returns/sets the source application and topic for a destination control."
  LinkTopic = m_TXT.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
  m_TXT.LinkTopic() = New_LinkTopic
  PropertyChanged "LinkTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "Returns/sets the amount of time a control waits for a response to a DDE message."
  LinkTimeout = m_TXT.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
  m_TXT.LinkTimeout() = New_LinkTimeout
  PropertyChanged "LinkTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkSend
Public Sub LinkSend()
Attribute LinkSend.VB_Description = "Transfers contents of PictureBox to destination application in DDE conversation."
  m_TXT.LinkSend
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkRequest
Public Sub LinkRequest()
Attribute LinkRequest.VB_Description = "Asks the source DDE application to update the contents of a Label, PictureBox, or Textbox control."
  m_TXT.LinkRequest
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkPoke
Public Sub LinkPoke()
Attribute LinkPoke.VB_Description = "Transfers contents of Label, PictureBox, or TextBox to source application in DDE conversation."
  m_TXT.LinkPoke
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkMode
Public Property Get LinkMode() As Integer
Attribute LinkMode.VB_Description = "Returns/sets the type of link used for a DDE conversation and activates the connection."
  LinkMode = m_TXT.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As Integer)
  m_TXT.LinkMode() = New_LinkMode
  PropertyChanged "LinkMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "Returns/sets the data passed to a destination control in a DDE conversation with another application."
  LinkItem = m_TXT.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
  m_TXT.LinkItem() = New_LinkItem
  PropertyChanged "LinkItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,LinkExecute
Public Sub LinkExecute(Command As String)
Attribute LinkExecute.VB_Description = "Sends a command string to the source application in a DDE conversation."
  m_TXT.LinkExecute Command
End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
  hWnd = m_TXT.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Specifies whether the selection in a Masked edit control is hidden when the control loses focus."
  HideSelection = m_TXT.HideSelection
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = m_TXT.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_TXT.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
  FontUnderline = m_TXT.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
  m_TXT.FontUnderline() = New_FontUnderline
  PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
  FontStrikethru = m_TXT.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
  m_TXT.FontStrikethru() = New_FontStrikethru
  PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
  FontSize = m_TXT.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
  m_TXT.FontSize() = New_FontSize
  PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
  FontName = m_TXT.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
  m_TXT.FontName() = New_FontName
  PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
  FontItalic = m_TXT.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
  m_TXT.FontItalic() = New_FontItalic
  PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
  FontBold = m_TXT.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
  m_TXT.FontBold() = New_FontBold
  PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = m_TXT.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set m_TXT.Font = New_Font
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

Private Sub txt_Click(Index As Integer)
  RaiseEvent Click
End Sub

Private Sub txt_Change(Index As Integer)
  RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = m_TXT.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  m_TXT.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
  Appearance = m_TXT.Appearance
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=m_TXT,txt,0,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
  Alignment = m_TXT.Alignment
End Property

Public Property Get TXTAlign() As txtAlignment
  TXTAlign = m_txtAlign
End Property

Public Property Let TXTAlign(ByVal New_txtAlign As txtAlignment)
  m_txtAlign = New_txtAlign
  PropertyChanged "txtAlign"
  Call lSetAlign
End Property

Private Sub txt_DblClick(Index As Integer)
  RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
  m_txtAlign = m_def_txtAlign
  Set m_TXT = UserControl.txt(L_NO_TXT_ALIGNS - 1)
  Call lSetAlign
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

  
  
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_TXT.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
  m_TXT.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
  m_TXT.Text = PropBag.ReadProperty("Text", "0")
  m_TXT.SelText = PropBag.ReadProperty("SelText", "")
  m_TXT.SelStart = PropBag.ReadProperty("SelStart", 0)
  m_TXT.SelLength = PropBag.ReadProperty("SelLength", 0)
  m_TXT.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
  m_TXT.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
  m_TXT.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
  m_TXT.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
  m_TXT.MousePointer = PropBag.ReadProperty("MousePointer", 0)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  m_TXT.MaxLength = PropBag.ReadProperty("MaxLength", 0)
  m_TXT.Locked = PropBag.ReadProperty("Locked", False)
  m_TXT.LinkTopic = PropBag.ReadProperty("LinkTopic", "")
  m_TXT.LinkTimeout = PropBag.ReadProperty("LinkTimeout", 50)
  m_TXT.LinkMode = PropBag.ReadProperty("LinkMode", 0)
  m_TXT.LinkItem = PropBag.ReadProperty("LinkItem", "")
  m_TXT.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
  m_TXT.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
  m_TXT.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
  m_TXT.FontSize = PropBag.ReadProperty("FontSize", 0)
  m_TXT.FontName = PropBag.ReadProperty("FontName", "")
  m_TXT.FontItalic = PropBag.ReadProperty("FontItalic", 0)
  m_TXT.FontBold = PropBag.ReadProperty("FontBold", 0)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  m_TXT.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  m_txtAlign = PropBag.ReadProperty("txtAlign", m_def_txtAlign)
  Call lSetAlign
End Sub

Private Sub UserControl_Resize()
  Dim l As Long
  
  On Error Resume Next
  With UserControl
    For l = 0 To L_NO_TXT_ALIGNS - 1
       txt(l).Width = UserControl.Width
       txt(l).Height = UserControl.Height
    Next l
  End With
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("WhatsThisHelpID", m_TXT.WhatsThisHelpID, 0)
  Call PropBag.WriteProperty("ToolTipText", m_TXT.ToolTipText, "")
  Call PropBag.WriteProperty("Text", m_TXT.Text, "0")
  Call PropBag.WriteProperty("SelText", m_TXT.SelText, "")
  Call PropBag.WriteProperty("SelStart", m_TXT.SelStart, 0)
  Call PropBag.WriteProperty("SelLength", m_TXT.SelLength, 0)
  Call PropBag.WriteProperty("RightToLeft", m_TXT.RightToLeft, False)
  Call PropBag.WriteProperty("PasswordChar", m_TXT.PasswordChar, "")
  Call PropBag.WriteProperty("OLEDropMode", m_TXT.OLEDropMode, 0)
  Call PropBag.WriteProperty("OLEDragMode", m_TXT.OLEDragMode, 0)
  Call PropBag.WriteProperty("MousePointer", m_TXT.MousePointer, 0)
  Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call PropBag.WriteProperty("MaxLength", m_TXT.MaxLength, 0)
  Call PropBag.WriteProperty("Locked", m_TXT.Locked, False)
  Call PropBag.WriteProperty("LinkTopic", m_TXT.LinkTopic, "")
  Call PropBag.WriteProperty("LinkTimeout", m_TXT.LinkTimeout, 50)
  Call PropBag.WriteProperty("LinkMode", m_TXT.LinkMode, 0)
  Call PropBag.WriteProperty("LinkItem", m_TXT.LinkItem, "")
  Call PropBag.WriteProperty("ForeColor", m_TXT.ForeColor, &H80000008)
  Call PropBag.WriteProperty("FontUnderline", m_TXT.FontUnderline, 0)
  Call PropBag.WriteProperty("FontStrikethru", m_TXT.FontStrikethru, 0)
  Call PropBag.WriteProperty("FontSize", m_TXT.FontSize, 0)
  Call PropBag.WriteProperty("FontName", m_TXT.FontName, "")
  Call PropBag.WriteProperty("FontItalic", m_TXT.FontItalic, 0)
  Call PropBag.WriteProperty("FontBold", m_TXT.FontBold, 0)
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("BackColor", m_TXT.BackColor, &H80000005)
  Call PropBag.WriteProperty("txtAlign", m_txtAlign, m_def_txtAlign)
End Sub

Private Function lSetAlign()
  Dim p As PropertyBag
  Dim l&, sTxt$
  
On Error GoTo lSetAlign_ERR
  'Call UserControl_ReadProperties(p)
  '
  
  With UserControl
    Call lMapProperties
    Set m_TXT = .txt(m_txtAlign)
    'Call UserControl_WriteProperties(p)
    '
    m_TXT.Visible = True
    For l = 0 To L_NO_TXT_ALIGNS - 1
     If l <> m_txtAlign Then .txt(l).Visible = False
    Next l
  End With
  
lSetAlign_ERR:
  Set p = Nothing
  Exit Function
End Function
Private Function lMapProperties&()
  'maps the properties that can be changed at run time
  With UserControl.txt(m_txtAlign)
    .SelLength = m_TXT.SelLength
    .RightToLeft = m_TXT.RightToLeft
    .PasswordChar = m_TXT.PasswordChar
    .OLEDropMode = m_TXT.OLEDropMode
    .OLEDragMode = m_TXT.OLEDragMode
    .MousePointer = m_TXT.MousePointer
    Set .MouseIcon = m_TXT.MouseIcon
    .MaxLength = m_TXT.MaxLength
    .Locked = m_TXT.Locked
    .LinkTopic = m_TXT.LinkTopic
    .LinkTimeout = m_TXT.LinkTimeout
    .LinkMode = m_TXT.LinkMode
    .LinkItem = m_TXT.LinkItem
    .FontUnderline = m_TXT.FontUnderline
    .FontStrikethru = m_TXT.FontStrikethru
    .FontSize = m_TXT.FontSize
    .FontName = m_TXT.FontName
    .FontItalic = m_TXT.FontItalic
    .FontBold = m_TXT.FontBold
    Set .Font = m_TXT.Font
    '.Enabled = m_TXT.Enabled
    .BackColor = m_TXT.BackColor
    .Text = m_TXT.T