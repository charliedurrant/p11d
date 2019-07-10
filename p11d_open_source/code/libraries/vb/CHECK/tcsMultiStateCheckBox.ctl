VERSION 5.00
Begin VB.UserControl tcsCheck 
   AccessKeys      =   " "
   BackColor       =   &H8000000A&
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   PropertyPages   =   "tcsMultiStateCheckBox.ctx":0000
   ScaleHeight     =   13
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13
End
Attribute VB_Name = "tcsCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Public Event Change()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Public Enum CHECK_STATE
  'Min must be 2
  [_INVALID_STATE] = 0
  CHECK_EMPTY = 2
  CHECK_TICK = 4
  CHECK_CROSS = 8
  CHECK_QUESTION = 16
  [_CHECK_STATE_N] = 4 'Number of valid states
  [_STATE_ALL] = CHECK_EMPTY + CHECK_TICK + CHECK_CROSS + CHECK_QUESTION
End Enum

Private m_State As CHECK_STATE
Private m_ValidStates As Long
Private m_AllowClick As Boolean

Private Function ValidState(ByVal NewState As CHECK_STATE) As Boolean
  If (NewState = CHECK_EMPTY) Or (NewState = CHECK_TICK) Or (NewState = CHECK_CROSS) Or (NewState = CHECK_QUESTION) Then
    ValidState = ((NewState And m_ValidStates) = NewState)
  End If
End Function

Private Sub ChangeCheck(ByVal CurState As CHECK_STATE)
  Dim NewState As Long
  Dim i As Long
  Dim State_N As Long
  
  CurState = CurState * 2
  If Not ValidState(CurState) Then CurState = CHECK_EMPTY
  i = 0
  Do While Not ValidState(CurState)
    CurState = CurState * 2
    i = i + 1
    If i = [_CHECK_STATE_N] Then Call Err.Raise(380, "ChangeCheck", "Unable to set a valid state")
  Loop
  'state set
  UserControl.Picture = LoadResPicture(CurState + 100, vbResBitmap)
  UserControl.Refresh
  m_State = CurState
  RaiseEvent Change
End Sub
Public Sub CycleCheck()
  If m_AllowClick Then
    Call ChangeCheck(m_State)
   End If
End Sub
Private Sub UserControl_Click()
  If m_AllowClick Then
    Call ChangeCheck(m_State)
   End If
End Sub

Private Sub UserControl_DblClick()
  If m_AllowClick Then
    Call ChangeCheck(m_State)
  End If
End Sub

Private Sub UserControl_EnterFocus()
  Debug.Print "EnterFocus"
End Sub
Private Sub UserControl_ExitFocus()
  Debug.Print "ExitFocus"
End Sub

Private Sub UserControl_Initialize()
  Dim i As Long
  If m_ValidStates = 0 Then
    m_ValidStates = [_STATE_ALL]
  End If
  m_State = CHECK_EMPTY
  UserControl.Picture = LoadResPicture(m_State + 100, vbResBitmap)
  UserControl.AccessKeys = " z"
End Sub

Public Property Let State(ByVal NewValue As CHECK_STATE)
  If Not ValidState(NewValue) Then Call Err.Raise(380, "LetState", "Unable to set a valid state")
  Call ChangeCheck(NewValue / 2)
  Call PropertyChanged("State")
End Property

Public Property Get State() As CHECK_STATE
  State = m_State
End Property

Public Property Let AllowClick(ByVal NewValue As Boolean)
  m_AllowClick = NewValue
  Call PropertyChanged("m_AllowClick")
End Property

Public Property Get AllowClick() As Boolean
  AllowClick = m_AllowClick
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(NewVal As Boolean)
  UserControl.Enabled = NewVal
End Property

Private Sub UserControl_InitProperties()
  m_AllowClick = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeySpace) Then
    Call CycleCheck
  End If
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_AllowClick = PropBag.ReadProperty("m_AllowClick", True)
  m_ValidStates = PropBag.ReadProperty("ValidStates", [_STATE_ALL])
  m_State = PropBag.ReadProperty("State", CHECK_EMPTY)
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = 200
  UserControl.Width = 200
End Sub

Public Property Get ValidStates() As CHECK_STATE
Attribute ValidStates.VB_ProcData.VB_Invoke_Property = "Valid_States"
  ValidStates = m_ValidStates
End Property

Public Property Let ValidStates(ByVal NewValue As CHECK_STATE)
  m_ValidStates = NewValue
  Call PropertyChanged("ValidStates")
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("m_AllowClick", m_AllowClick)
  Call PropBag.WriteProperty("ValidStates", m_ValidStates)
  Call PropBag.WriteProperty("State", m_State)
End Sub

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

