VERSION 5.00
Begin VB.UserControl ValCombo 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "ValCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Change()
Public Event Click()
Public Event FieldInvalid(Valid As Boolean, Message As String)
Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Private m_InvalidValue As String
Private Sub cbo_Change()
  Validate
  RaiseEvent Change
End Sub
Private Sub cbo_Click()
  Validate
  RaiseEvent Click
End Sub
Private Sub Validate()
  Dim b As Boolean
  b = FieldInvalid
End Sub


Private Sub UserControl_Initialize()
  Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
  cbo.Width = UserControl.Width
  cbo.Top = 0
  cbo.Left = 0
  UserControl.Height = cbo.Height
End Sub

Public Sub Clear()
  Call cbo.Clear
  Validate
End Sub
Public Sub AddItem(Item As String)
   Call cbo.AddItem(Item)
   Validate
End Sub
Public Property Get FieldInvalid() As Boolean
   cbo.BackColor = vbWhite
   If (UserControl.Enabled) Then
    If (StrComp(cbo.Text, m_InvalidValue) = 0) Then
      FieldInvalid = True
      cbo.BackColor = vbRed
      Exit Property
    End If
   End If
End Property
Public Property Let InvalidValue(ByVal NewValue As String)
    m_InvalidValue = NewValue
    Validate
End Property
Public Property Get ComboBox() As ComboBox
  Set ComboBox = cbo
End Property
Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  UserControl.Enabled = NewValue
  cbo.Enabled = NewValue
  PropertyChanged "Enabled"
End Property


Private Sub UserControl_InitProperties()
  
  Set m_Font = Ambient.Font
  Set UserControl.Font = m_Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
  Set UserControl.Font = m_Font
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Font", m_Font, Ambient.Font
End Sub
Public Property Get Font() As StdFont
  Set Font = m_Font
End Property

Public Property Set Font(mnewFont As StdFont)
  With m_Font
    .Bold = mnewFont.Bold
    .Italic = mnewFont.Italic
    .Name = mnewFont.Name
    .size = mnewFont.size
    .Strikethrough = mnewFont.Strikethrough
    .Charset = mnewFont.Charset
    .Underline = mnewFont.Underline
    .Weight = mnewFont.Weight
  End With
  PropertyChanged "Font"
End Property

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
  Set UserControl.Font = m_Font
  Set cbo.Font = m_Font
  Call UserControl_Resize
  Refresh
End Sub

