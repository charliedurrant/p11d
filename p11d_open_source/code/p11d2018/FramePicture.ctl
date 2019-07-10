VERSION 5.00
Begin VB.UserControl FramePicture 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlContainer=   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   4770
   Begin VB.Frame fra 
      Caption         =   "Frame1"
      Height          =   3480
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4695
      Begin VB.PictureBox pct 
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   90
         ScaleHeight     =   2715
         ScaleWidth      =   4515
         TabIndex        =   1
         Top             =   675
         Width           =   4515
      End
   End
End
Attribute VB_Name = "FramePicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Font As StdFont
Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  UserControl.Enabled = NewValue
  fra.Enabled = NewValue
  pct.Enabled = NewValue
  PropertyChanged "Enabled"
End Property

Private Sub UserControl_InitProperties()
  Set m_Font = Ambient.Font
  Set UserControl.Font = m_Font
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
  Set UserControl.Font = m_Font
  Me.Caption = PropBag.ReadProperty("Caption", UserControl.Name)
End Sub

Private Sub UserControl_Show()
   Call fra.ZOrder(0)
   Call pct.ZOrder(1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Font", m_Font, Ambient.Font
  PropBag.WriteProperty "Caption", Caption, UserControl.Name
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
  fra.Font = m_Font
  PropertyChanged "Font"
End Property

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
  Set UserControl.Font = m_Font
  
  Set fra.Font = m_Font
  Call UserControl_Resize
  Refresh
End Sub


Public Property Get Caption() As String
 Caption = fra.Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
  fra.Caption = NewValue
  PropertyChanged "Caption"

End Property
Public Property Get Controls() As Collection
  Set Controls = UserControl.Controls
End Property


Private Sub UserControl_Initialize()
  Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
  Dim i As Long
  
  i = (UserControl.Font.size + 5) * Screen.TwipsPerPixelY
  Call fra.Move(0, 0, UserControl.Width, UserControl.Height)
  Call pct.Move(1, i, fra.Width - 2, fra.Height - i)
End Sub
