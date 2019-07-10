VERSION 5.00
Begin VB.UserControl QBEObj 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   ScaleHeight     =   1575
   ScaleWidth      =   1815
   Begin VB.TextBox txtFilter3 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   1260
      Width           =   1815
   End
   Begin VB.TextBox txtFilter2 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   945
      Width           =   1815
   End
   Begin VB.TextBox txtFilter1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   630
      Width           =   1815
   End
   Begin VB.ComboBox cmbSort 
      Height          =   315
      ItemData        =   "QBEObj.ctx":0000
      Left            =   0
      List            =   "QBEObj.ctx":000D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   315
      Width           =   1815
   End
   Begin VB.ComboBox cmbField 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "QBEObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Consts
Const QBE_HEIGHT As Long = 1615
Const QBE_WIDTH As Long = 1855

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
  BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
  UserControl.BackStyle() = New_BackStyle
  PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  UserControl.Refresh
End Sub

Private Sub txtFilter1_LostFocus()
  txtFilter1 = Validate(txtFilter1)
End Sub

Private Sub txtFilter2_LostFocus()
  txtFilter2 = Validate(txtFilter2)
End Sub

Private Sub txtFilter3_LostFocus()
  txtFilter3 = Validate(txtFilter3)
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbField,cmbField,-1,AddItem
Public Sub AddItem(Item As String, Optional Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
  cmbField.AddItem Item, Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFilter1,txtFilter1,-1,Text
Public Property Get Filter1() As String
Attribute Filter1.VB_Description = "Returns/sets the text contained in the control."
  Filter1 = txtFilter1.Text
End Property

Public Property Let Filter1(ByVal New_Filter1 As String)
  txtFilter1.Text() = New_Filter1
  PropertyChanged "Filter1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFilter2,txtFilter2,-1,Text
Public Property Get Filter2() As String
Attribute Filter2.VB_Description = "Returns/sets the text contained in the control."
  Filter2 = txtFilter2.Text
End Property

Public Property Let Filter2(ByVal New_Filter2 As String)
  txtFilter2.Text() = New_Filter2
  PropertyChanged "Filter2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFilter3,txtFilter3,-1,Text
Public Property Get Filter3() As String
Attribute Filter3.VB_Description = "Returns/sets the text contained in the control."
  Filter3 = txtFilter3.Text
End Property

Public Property Let Filter3(ByVal New_Filter3 As String)
  txtFilter3.Text() = New_Filter3
  PropertyChanged "Filter3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbSort,cmbSort,-1,ListIndex
Public Property Get Sort() As Integer
Attribute Sort.VB_Description = "Returns/sets the index of the currently selected item in the control."
  Sort = cmbSort.ListIndex
End Property

Public Property Let Sort(ByVal New_Sort As Integer)
  cmbSort.ListIndex() = New_Sort
  PropertyChanged "Sort"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  Set Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  txtFilter1.Text = PropBag.ReadProperty("Filter1", "")
  txtFilter2.Text = PropBag.ReadProperty("Filter2", "")
  txtFilter3.Text = PropBag.ReadProperty("Filter3", "")
  cmbSort.ListIndex = PropBag.ReadProperty("Sort", 0)
  'cmbField.ListIndex = PropBag.ReadProperty("ListIndex", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
  'cmbField.List(Index) = PropBag.ReadProperty("List" & Index, "")
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = QBE_HEIGHT
  UserControl.Width = QBE_WIDTH
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
  Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
  Call PropBag.WriteProperty("Filter1", txtFilter1.Text, "")
  Call PropBag.WriteProperty("Filter2", txtFilter2.Text, "")
  Call PropBag.WriteProperty("Filter3", txtFilter3.Text, "")
  Call PropBag.WriteProperty("Sort", cmbSort.ListIndex, 0)
  Call PropBag.WriteProperty("ListIndex", cmbField.ListIndex, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
  Call PropBag.WriteProperty("List" & Index, cmbField.List(Index), "")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbField,cmbField,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
  ListIndex = cmbField.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
  cmbField.ListIndex() = New_ListIndex
  PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbField,cmbField,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
  ListCount = cmbField.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbField,cmbField,-1,List
Public Property Get List(Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
  List = cmbField.List(Index)
End Property

Public Property Let List(Index As Integer, ByVal New_List As String)
  cmbField.List(Index) = New_List
  PropertyChanged "List"
End Property

Private Function Validate(txt As String) As String
  If IsDate(txt) Then
    Validate = "=#" & txt & "#"
  ElseIf IsNumeric(txt) Then
    Validate = "=" & txt
  ElseIf Not Key(txt) Then
    Validate = "=" & Chr$(34) & txt & Chr$(34)
  Else
    Validate = txt
  End If
End Function

Private Function Key(txt As String) As Boolean
  Key = False
  If InStr(txt, "=") Then
    Key = True
    Exit Function
  ElseIf InStr(txt, "LIKE") Then
    Key = True
    Exit Function
  ElseIf InStr(txt, ">") Then
    Key = True
    Exit Function
  ElseIf InStr(txt, "<") Then
    Key = True
    Exit Function
  ElseIf txt = "" Then
    Key = True
    Exit Function
  End If
End Function
