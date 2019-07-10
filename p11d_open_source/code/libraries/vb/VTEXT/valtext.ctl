VERSION 5.00
Object = "{C8A8E78F-1AF8-4AD9-A6D0-E3456B7DA96B}#1.0#0"; "atc2align.OCX"
Begin VB.UserControl ValText 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "valtext.ctx":0000
   Begin atc2TextAlign.TXTAlign TXTControl 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   661
      Text            =   ""
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ValText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum datatypes
  VT_LONG
  VT_DOUBLE
  VT_DATE
  VT_STRING
  VT_USER
End Enum

'Default Property Texts:
Const m_def_ReturnInvokesTab As Boolean = True
Const m_def_EscapeInvokesOriginal As Boolean = True
Const m_def_ValidColor = &H80000005
Const m_def_InvalidColor = &HFF&
Const m_def_AllowEmpty = -1
Const m_def_Type = 0
Const m_def_ValidationText = ""
Const m_def_Maximum = Empty
Const m_def_Minimum = Empty
Const m_def_txtAlign = txtAlignment.TXT_LEFT
Const m_def_AutoSwitchAlign = True
Const m_def_ItemsKeyCode = vbKeyInsert
Const m_def_AllowUserAddItems = False
Const m_def_SelectItemsCaption$ = "Items:"
Const m_def_ItemsKeyString$ = "Insert"
Const m_def_AllowUserDeleteItems As Boolean = False
Const m_def_Validate As Boolean = True
Const m_def_AutoSelect As Boolean = True
Const m_def_FullDate As Boolean = True
'my constants
Private Const L_LONGMAX As Long = 2147483647
Private Const L_LONGMIN As Long = -2147483647

'Property Variables:
Dim m_ValidColor As OLE_COLOR
Dim m_InvalidColor As OLE_COLOR
Dim m_AllowEmpty As Boolean
Dim m_TypeOfData As Long
Dim m_Maximum As Variant
Dim m_Minimum As Variant
Dim m_ValidationText As String
Dim m_FieldInvalid As Boolean
Dim m_txtAlign As Long
Dim m_AutoSwitchAlign As Boolean
Dim m_ItemsKeyCode As Integer
Dim m_AllowUserAddItems As Boolean
Dim m_SelectItemsCaption As String
Dim m_ItemsKeyString As String
Dim m_AllowUserDeleteItems As Boolean
Dim m_Validate As Boolean, m_ReturnInvokesTab As Boolean
Dim m_OriginalString As String
Dim m_EscapeInvokesOriginal As Boolean
Dim m_AutoSelect As Long
Dim m_FullDate As Boolean

'Event Declarations:
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Event FieldInvalid(Valid As Boolean, Message As String)
Event Click() 'MappingInfo=txtControl,txtControl,-1,Click
Event DblClick() 'MappingInfo=txtControl,txtControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtControl,txtControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtControl,txtControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtControl,txtControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtControl,txtControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtControl,txtControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtControl,txtControl,-1,MouseUp
Event UserValidate(Valid As Boolean, Message As String, sTextEntered As String)

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,BackColor
Public Property Get AutoSelect() As Boolean
  AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(b As Boolean)
  m_AutoSelect = b
  PropertyChanged ("AutoSelect")
End Property

Public Property Let ReturnInvokesTab(b As Boolean)
  m_ReturnInvokesTab = b
  PropertyChanged "ReturnInvokesTab"
End Property
Public Property Get ReturnInvokesTab() As Boolean
  ReturnInvokesTab = m_ReturnInvokesTab
End Property
Public Property Let EscapeInvokesOriginal(b As Boolean)
  m_EscapeInvokesOriginal = b
  PropertyChanged "EscapeInvokesOriginal"
End Property
Public Property Get EscapeInvokesOriginal() As Boolean
  EscapeInvokesOriginal = m_EscapeInvokesOriginal
End Property
Public Property Let AllowUserDeleteItems(b As Boolean)
  m_AllowUserDeleteItems = b
  PropertyChanged "AllowUserDeleteItems"
End Property
Public Property Get AllowUserDeleteItems() As Boolean
  AllowUserDeleteItems = m_AllowUserDeleteItems
End Property
Property Get ItemsKeyString$()
  ItemsKeyString = m_ItemsKeyString
End Property
Property Let ItemsKeyString(s$)
  m_ItemsKeyString = s
  PropertyChanged "ItemsKeyString"
End Property

Public Property Get SelectItemsCaption$()
  SelectItemsCaption$ = m_SelectItemsCaption
End Property

Public Property Let SelectItemsCaption(s$)
  m_SelectItemsCaption$ = s$
  PropertyChanged "SelectItemsCaption"
End Property

Public Property Get AllowUserAddItems() As Boolean
  AllowUserAddItems = m_AllowUserAddItems
End Property
Public Property Let AllowUserAddItems(b As Boolean)
  m_AllowUserAddItems = b
  PropertyChanged "AllowUserAddItems"
End Property

Public Property Let ItemsKeyCode(l As Long)
  m_ItemsKeyCode = l
  PropertyChanged "ItemsKeyCode"
End Property
Public Property Get ItemsKeyCode&()
  ItemsKeyCode = m_ItemsKeyCode
End Property

Public Property Get AutoSwitchAlign() As Boolean
  AutoSwitchAlign = m_AutoSwitchAlign
End Property

Public Property Let AutoSwitchAlign(b As Boolean)
  m_AutoSwitchAlign = b
  Call lAutoSwitchAlign&
  PropertyChanged "AutoSwitchAlign"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = TXTControl.BackColor
End Property

Public Property Get FieldInvalid() As Boolean
  FieldInvalid = m_FieldInvalid
End Property

Public Property Let TXTAlign(l As txtAlignment)
  With UserControl.TXTControl
    .TXTAlign = l
    Call lValidate
    Call lResize&
    PropertyChanged "TXTAlign"
  End With
End Property

Public Property Get TXTAlign() As txtAlignment
  TXTAlign = m_txtAlign
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  TXTControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = TXTControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  TXTControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = TXTControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set TXTControl.Font = New_Font
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'  BorderStyle = txtControl.BorderStyle
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  TXTControl.Refresh
End Sub

Private Sub txtControl_Change()
  Call lValidate
  RaiseEvent Change
End Sub

Private Sub txtControl_Click()
  RaiseEvent Click
End Sub

Private Sub txtControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub TxtControl_GotFocus()
  Dim lLen&
  
  With TXTControl
    m_OriginalString = .Text
    If m_AutoSelect Then
      lLen = Len(.Text)
      If lLen Then
        .SelStart = 0
        .SelLength = lLen
      End If
    End If
  End With
  'MsgBox m_OriginalString
End Sub

Private Sub txtControl_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim sNewString$, lLen&
  Select Case KeyCode
    Case 13
      If m_ReturnInvokesTab Then
        KeyCode = 0
        SendKeys "{TAB}"
      End If
    Case vbKeyEscape
      If m_EscapeInvokesOriginal Then
        With TXTControl
          sNewString = .Text
          ' if the not the same then replace with old
          If StrComp(sNewString, m_OriginalString) Then
            .Text = m_OriginalString
          End If
          'select if not 0 length
          lLen = Len(m_OriginalString)
          If lLen Then
            .SelStart = 0
            .SelLength = lLen
          End If
        End With
      End If
  End Select
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtControl_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And m_ReturnInvokesTab Then
    KeyAscii = 0
  ElseIf StrComp(Chr$(KeyAscii), ".") = 0 Then
    If m_TypeOfData = VT_LONG Then KeyAscii = 0
  End If
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
  Alignment = TXTControl.Alignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
  Appearance = TXTControl.Appearance
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
  Locked = TXTControl.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  TXTControl.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
  MaxLength = TXTControl.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
  If New_MaxLength < 0 Then New_MaxLength = 0
  TXTControl.MaxLength = New_MaxLength
  Call lValidate
  PropertyChanged "MaxLength"
End Property

Public Property Get ValidationText() As String
  ValidationText = m_ValidationText
End Property

Public Property Let ValidationText(ByVal New_ValidationText As String)
  m_ValidationText = New_ValidationText
  Call lValidate
  PropertyChanged " ValidationText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
  Set MouseIcon = TXTControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
  Set TXTControl.MouseIcon = New_MouseIcon
  PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
  MousePointer = TXTControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
  TXTControl.MousePointer() = New_MousePointer
  PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
  MultiLine = TXTControl.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
  PasswordChar = TXTControl.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
  TXTControl.PasswordChar() = New_PasswordChar
  PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
  SelLength = TXTControl.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
  TXTControl.SelLength() = New_SelLength
  PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
  SelStart = TXTControl.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
  TXTControl.SelStart() = New_SelStart
  PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
  SelText = TXTControl.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
  TXTControl.SelText() = New_SelText
  PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
  Text = TXTControl.Text
End Property

Public Property Let Text(ByVal New_Text As String)
  TXTControl.Text = New_Text
  Call lValidate
  PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
  WhatsThisHelpID = TXTControl.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
  TXTControl.WhatsThisHelpID() = New_WhatsThisHelpID
  PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtControl,txtControl,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
  ToolTipText = TXTControl.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
  TXTControl.ToolTipText() = New_ToolTipText
  PropertyChanged "ToolTipText"
End Property

Public Property Get TypeOfData() As datatypes
  TypeOfData = m_TypeOfData
End Property

Public Property Let TypeOfData(ByVal New_TypeOfData As datatypes)
  m_TypeOfData = New_TypeOfData
  m_Minimum = Empty
  m_Maximum = Empty
  Call lResize
  Call lAutoSwitchAlign
  Call lValidate
  PropertyChanged "TypeOfData"
End Property

Public Function lValidate() As Long
Dim s As String, sErrMsg As String, d0 As Date
Dim bValid As Boolean, EmptyString As Boolean, bYearError As Boolean

On Error GoTo Validate_ERR
  
  bValid = True
  sErrMsg = ""
  s = TXTControl.Text
  EmptyString = IsStringEmpty(s)
  If (m_AllowEmpty And EmptyString) Or Not m_Validate Then GoTo change_exit
  Select Case m_TypeOfData
    Case VT_LONG
      Call lCheckLong(s, bValid, sErrMsg)
    Case VT_DOUBLE
      If IsNumeric(s) Then
        If IsEmpty(m_Maximum) Or (CDbl(s) <= CDbl(m_Maximum)) Then
          If IsEmpty(m_Minimum) Or (CDbl(s) >= CDbl(m_Minimum)) Then
            RaiseEvent FieldInvalid(True, "")
            TXTControl.BackColor = m_ValidColor
          Else
            bValid = False
            sErrMsg = "Invalid decimal number entered. The Text should be greater than or equal to " & m_Minimum
          End If
        Else
          bValid = False
          sErrMsg = "Invalid decimal number entered. The Text should be less than or equal to " & m_Maximum
        End If
      Else
        bValid = False
        sErrMsg = "Unable to convert your entry into a number."
      End If
    Case VT_STRING
       If ((Not m_AllowEmpty) And EmptyString) Then
        bValid = False
        sErrMsg = "The field is not allowed to be empty."
      ElseIf Not ((Len(m_ValidationText) = 0) Or s Like m_ValidationText) Then
        bValid = False
        sErrMsg = "Unable to format entry into the format" & m_ValidationText
      End If
    Case VT_DATE
      bValid = False
      d0 = TryConvertDateDMY(s, MIN_DATE)
      If d0 <> MIN_DATE Then
        If d0 <= TryConvertDateDMY(CStr(m_Maximum), MAX_DATE) Then
          If d0 >= TryConvertDateDMY(CStr(m_Minimum), MIN_DATE) Then
            If m_FullDate Then
              If Len(s) > 4 Then
                If InStr(Right$(s, 4), "/") = 0 Then bValid = True
              End If
            Else
              bValid = True
            End If
          Else
            sErrMsg = "Invalid date Text entered. The Text should be greater than " & Format$(m_Minimum, "dd/mm/yyyy")
          End If
        Else
          sErrMsg = "Invalid date Text entered. The Text should be less than " & Format$(m_Maximum, "dd/mm/yyyy")
        End If
      Else
        sErrMsg = "Unable to convert your entry into a date."
      End If
    Case VT_USER
      RaiseEvent UserValidate(bValid, sErrMsg, s)
  End Select
  
change_exit:
  If bValid Then
    TXTControl.BackColor = m_ValidColor
  Else
    TXTControl.BackColor = m_InvalidColor
  End If
  m_FieldInvalid = Not (bValid)
  RaiseEvent FieldInvalid(bValid, sErrMsg)
  Exit Function
  
Validate_ERR:
End Function
Private Function lCheckLong&(s$, bValid As Boolean, sErrMsg$)
 Dim d#, e#, l&
 
 If IsNumeric(s) Then
    d = CDbl(s)
    e = d
    d = d - Int(d)
    If d = 0 Then
      If (L_LONGMIN < e) And (e < L_LONGMAX) Then
        l = CLng(s)
        If IsEmpty(m_Maximum) Or (l <= CLng(m_Maximum)) Then
          If Not (IsEmpty(m_Minimum) Or (l >= CLng(m_Minimum))) Then
            bValid = False
            sErrMsg = "Invalid integer number entered. The Text should be greater than or equal to " & m_Minimum
          End If
        Else
          bValid = False
          sErrMsg = "Invalid integer number entered. The Text should be less than or equal to " & m_Maximum
        End If
      Else
        bValid = False
        sErrMsg = "Number is outside " & CStr(L_LONGMIN) & " to " & CStr(L_LONGMAX) & "."
      End If
    Else
      bValid = False
      sErrMsg = "Invalid integer number entered. No decimal values allowed."
    End If
  Else
    bValid = False
    sErrMsg = "Unable to convert your entry into a number."
  End If
End Function

Private Function lReturnChoice&(Optional sText$ = "")
  On Error Resume Next 'Dodgy Hack to stop the control GPFing when clicking away from the expanded combo-box style drop down
  With TXTControl
    If Len(sText) Then .Text = sText
    'highlight the text in the user control
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
  End With
End Function

Public Property Get Maximum() As Variant
Attribute Maximum.VB_Description = "The maximum value to be entered into this box"
  Maximum = m_Maximum
End Property

Public Property Let Maximum(ByVal New_Maximum As Variant)
  m_Maximum = New_Maximum
  Call lValidate
  PropertyChanged "Maximum"
End Property

Public Property Get Minimum() As Variant
Attribute Minimum.VB_Description = "The minimun value to be entered into this box"
  Minimum = m_Minimum
End Property

Public Property Let Minimum(ByVal New_Minimum As Variant)
  m_Minimum = New_Minimum
  Call lValidate
  PropertyChanged "Minimum"
End Property

'Initialize Properties for User Control
Private Sub usercontrol_InitProperties()
  m_TypeOfData = m_def_Type
  m_Maximum = m_def_Maximum
  m_Minimum = m_def_Minimum
  m_AllowEmpty = m_def_AllowEmpty
  m_ValidationText = m_def_ValidationText
  If Ambient.UserMode Then Call lValidate
  m_ValidColor = m_def_ValidColor
  m_InvalidColor = m_def_InvalidColor
  m_AutoSwitchAlign = m_def_AutoSwitchAlign
  m_txtAlign = m_def_txtAlign
  m_AllowUserAddItems = m_def_AllowUserAddItems
  m_SelectItemsCaption = m_def_SelectItemsCaption
  m_ItemsKeyString = m_def_ItemsKeyString
  m_AllowUserDeleteItems = m_def_AllowUserDeleteItems
  m_ItemsKeyCode = m_def_ItemsKeyCode
  m_Validate = m_def_Validate
  m_ReturnInvokesTab = m_def_ReturnInvokesTab
  m_EscapeInvokesOriginal = m_def_EscapeInvokesOriginal
  m_FullDate = m_def_FullDate
End Sub

'Load property Texts from storage
Private Sub usercontrol_ReadProperties(PropBag As PropertyBag)
  With PropBag
    TXTControl.BackColor = .ReadProperty("BackColor", &H80000005)
    TXTControl.ForeColor = .ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = .ReadProperty("Enabled", True)
    Set Font = .ReadProperty("Font", Ambient.Font)
    TXTControl.Locked = .ReadProperty("Locked", False)
    TXTControl.MaxLength = .ReadProperty("MaxLength", 0)
    Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
    TXTControl.MousePointer = .ReadProperty("MousePointer", 0)
    TXTControl.PasswordChar = .ReadProperty("PasswordChar", "")
    TXTControl.SelLength = .ReadProperty("SelLength", 0)
    TXTControl.SelStart = .ReadProperty("SelStart", 0)
    TXTControl.SelText = .ReadProperty("SelText", "")
    TXTControl.Text = .ReadProperty("Text", "txtControl")
    TXTControl.WhatsThisHelpID = .ReadProperty("WhatsThisHelpID", 0)
    TXTControl.ToolTipText = .ReadProperty("ToolTipText", "")
    m_TypeOfData = .ReadProperty("TypeOfData", m_def_Type)
    m_Maximum = .ReadProperty("Maximum", m_def_Maximum)
    m_Minimum = .ReadProperty("Minimum", m_def_Minimum)
    m_AllowEmpty = .ReadProperty("AllowEmpty", m_def_AllowEmpty)
    m_ValidationText = .ReadProperty("ValidationText", m_def_ValidationText)
    m_ValidColor = .ReadProperty("ValidColor", m_def_ValidColor)
    m_InvalidColor = .ReadProperty("InvalidColor", m_def_InvalidColor)
    m_txtAlign = .ReadProperty("TXTAlign", m_def_txtAlign)
    m_AutoSwitchAlign = .ReadProperty("AutoSwitchAlign", m_def_AutoSwitchAlign)
    m_ItemsKeyCode = .ReadProperty("ItemsKeyCode", m_def_ItemsKeyCode)
    m_AllowUserAddItems = .ReadProperty("AllowUserAddItems", m_def_AllowUserAddItems)
    m_SelectItemsCaption = .ReadProperty("SelectItemsCaption", m_def_SelectItemsCaption)
    m_ItemsKeyString = .ReadProperty("ItemsKeyString", m_def_ItemsKeyString)
    m_AllowUserDeleteItems = .ReadProperty("AllowUserDeleteItems", m_def_AllowUserDeleteItems)
    m_Validate = .ReadProperty("Validate", m_def_Validate)
    m_ReturnInvokesTab = .ReadProperty("ReturnInvokesTab", m_def_ReturnInvokesTab)
    m_EscapeInvokesOriginal = .ReadProperty("EscapeInvokesOriginal", m_def_EscapeInvokesOriginal)
    m_AutoSelect = .ReadProperty("AutoSelect", m_def_AutoSelect)
    m_FullDate = .ReadProperty("FullDate", m_def_FullDate)
    Call lAutoSwitchAlign
  End With
End Sub

Private Sub usercontrol_Resize()
  Call lResize
End Sub

Private Function lResize&()
  Call TXTControl.Move(0, 0, UserControl.Width, UserControl.Height)
End Function

Private Sub usercontrol_Show()
  If Ambient.UserMode Then Call lValidate
End Sub

'Write property Texts to storage
Private Sub usercontrol_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("BackColor", TXTControl.BackColor, &H80000005)
    Call .WriteProperty("ForeColor", TXTControl.ForeColor, &H80000008)
    Call .WriteProperty("Enabled", UserControl.Enabled, True)
    Call .WriteProperty("Font", Font, Ambient.Font)
    Call .WriteProperty("Locked", TXTControl.Locked, False)
    Call .WriteProperty("MaxLength", TXTControl.MaxLength, 0)
    Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call .WriteProperty("MousePointer", TXTControl.MousePointer, 0)
    Call .WriteProperty("PasswordChar", TXTControl.PasswordChar, "")
    Call .WriteProperty("SelLength", TXTControl.SelLength, 0)
    Call .WriteProperty("SelStart", TXTControl.SelStart, 0)
    Call .WriteProperty("SelText", TXTControl.SelText, "")
    Call .WriteProperty("Text", TXTControl.Text, "txtControl")
    Call .WriteProperty("WhatsThisHelpID", TXTControl.WhatsThisHelpID, 0)
    Call .WriteProperty("ToolTipText", TXTControl.ToolTipText, "")
    Call .WriteProperty("TypeOfData", m_TypeOfData, m_def_Type)
    Call .WriteProperty("Maximum", m_Maximum, m_def_Maximum)
    Call .WriteProperty("Minimum", m_Minimum, m_def_Minimum)
    Call .WriteProperty("AllowEmpty", m_AllowEmpty, m_def_AllowEmpty)
    Call .WriteProperty("ValidationText", m_ValidationText, m_def_ValidationText)
    Call .WriteProperty("ValidColor", m_ValidColor, m_def_ValidColor)
    Call .WriteProperty("InvalidColor", m_InvalidColor, m_def_InvalidColor)
    Call .WriteProperty("TXTAlign", m_txtAlign, m_def_txtAlign)
    Call .WriteProperty("AutoSwitchAlign", m_AutoSwitchAlign, m_def_AutoSwitchAlign)
    Call .WriteProperty("ItemsKeyCode", m_ItemsKeyCode, m_def_ItemsKeyCode)
    Call .WriteProperty("AllowUserAddItems", m_AllowUserAddItems, m_def_AllowUserAddItems)
    Call .WriteProperty("SelectItemsCaption", m_SelectItemsCaption, m_def_SelectItemsCaption)
    Call .WriteProperty("ItemsKeyString", m_ItemsKeyString, m_def_ItemsKeyString)
    Call .WriteProperty("AllowUserDeleteItems", m_AllowUserDeleteItems, m_def_AllowUserDeleteItems)
    Call .WriteProperty("Validate", m_Validate, m_def_Validate)
    Call .WriteProperty("ReturnInvokesTab", m_ReturnInvokesTab, m_def_ReturnInvokesTab)
    Call .WriteProperty("EscapeInvokesOriginal", m_EscapeInvokesOriginal, m_def_EscapeInvokesOriginal)
    Call .WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
    Call .WriteProperty("FullDate", m_FullDate, m_def_FullDate)
  End With
End Sub

Public Property Get Validate() As Boolean
  Validate = m_Validate
End Property

Public Property Let Validate(ByVal New_Validate As Boolean)
  m_Validate = New_Validate
  Call lValidate
  PropertyChanged "Validate"
End Property

Public Property Get AllowEmpty() As Boolean
  AllowEmpty = m_AllowEmpty
End Property

Public Property Let AllowEmpty(ByVal New_AllowEmpty As Boolean)
  m_AllowEmpty = New_AllowEmpty
  Call lValidate
  PropertyChanged "AllowEmpty"
End Property
Public Property Let FullDate(ByVal NewValue As Boolean)
  m_FullDate = NewValue
  Call lValidate
  PropertyChanged "FullDate"
End Property
Public Property Get FullDate() As Boolean
  FullDate = m_FullDate
End Property

Public Property Get ValidColor() As OLE_COLOR
  ValidColor = m_ValidColor
End Property

Public Property Let ValidColor(ByVal New_ValidColor As OLE_COLOR)
  m_ValidColor = New_ValidColor
  PropertyChanged "ValidColor"
End Property

Public Property Get InvalidColor() As OLE_COLOR
  InvalidColor = m_InvalidColor
End Property

Public Property Let InvalidColor(ByVal New_InvalidColor As OLE_COLOR)
  m_InvalidColor = New_InvalidColor
  PropertyChanged "InvalidColor"
End Property

Private Function lAutoSwitchAlign&()
  Dim ltxtAlign As txtAlignment
  'refer to enum datatypes
  If m_AutoSwitchAlign Then
    Select Case m_TypeOfData
      Case VT_LONG, VT_DOUBLE
        ltxtAlign = txtAlignment.TXT_RIGHT
      Case Else
        ltxtAlign = txtAlignment.TXT_LEFT
    End Select
      m_txtAlign = ltxtAlign
      TXTAlign = ltxtAlign
  End If
End Function

Private Function IsStringEmpty(s As String) As Boolean
  Dim i As Long, s2 As String
  
  IsStringEmpty = True
  If Len(s) > 0 Then
    For i = 1 To Len(s)
      s2 = Mid$(s, i, 1)
      If Not ((StrComp(s2, " ") = 0) Or (StrComp(s2, Chr$(9)) = 0)) Then GoTo IsStringEmpty_end
    Next i
  End If
  Exit Function
  
IsStringEmpty_end:
  IsStringEmpty = False
End Function

  
