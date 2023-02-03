VERSION 5.00
Begin VB.UserControl MyFolderBrowser 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ScaleHeight     =   375
   ScaleWidth      =   4785
   Begin VB.CommandButton cmdBrowse 
      Height          =   330
      Left            =   0
      Picture         =   "FolderBrowser.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   35
      Width           =   375
   End
   Begin VB.Label lbl 
      Height          =   375
      Left            =   405
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "MyFolderBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Folder As String
Private m_Title As String
Private m_ReadOnly As Boolean

Public Event Started()
Public Event Ended()
Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Let Title(ByVal NewValue As String)
  m_Title = NewValue
End Property
Public Property Get Directory() As String
  Directory = lbl.Caption
End Property
Public Property Let Directory(ByVal NewValue As String)
  lbl.Caption = FullPath(NewValue)
  lbl.ToolTipText = lbl.Caption
End Property

Private Sub cmdBrowse_Click()
 Dim sIntialDirectory As String
 Dim s As String
 RaiseEvent Started
 If (Len(Directory)) = 0 Then
  sIntialDirectory = CurDir
 Else
  sIntialDirectory = Directory
 End If
 sIntialDirectory = FullPath(sIntialDirectory)
 s = FullPath(BrowseForFolderEx(UserControl.hwnd, sIntialDirectory, m_Title))
 If (Len(s) > 0) Then
  Directory = s
  RaiseEvent Ended
 End If
End Sub

Private Sub UserControl_Initialize()
  m_Title = "Choose a folder"
  cmdBrowse.ToolTipText = "Click to select a folder"
  Directory = CurDir
  lbl.ForeColor = UserControl.ForeColor
End Sub

Private Sub UserControl_Resize()
  cmdBrowse.Visible = Not Me.ReadOnly
    
  If (Me.ReadOnly) Then
    lbl.height = UserControl.height
    lbl.Top = 0
    lbl.Left = 0
    lbl.width = UserControl.width
  Else
    cmdBrowse.Left = 0
    cmdBrowse.Top = 0
    lbl.height = UserControl.height
    lbl.Top = 0
    lbl.Left = cmdBrowse.width + cmdBrowse.Left + 60
    lbl.width = UserControl.width - lbl.Left
  End If
  
End Sub
Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
 UserControl.Enabled = NewValue
 lbl.Enabled = NewValue
 cmdBrowse.Enabled = NewValue
End Property
Public Property Get Font() As StdFont
   Set Font = lbl.Font
End Property
Public Property Set Font(po_Font As StdFont)
   Set lbl.Font = po_Font
   PropertyChanged "Font"  ' Signal that the property should be saved in the WriteProperties event
   lbl.Refresh  ' Redraw the usercontrol using the new font
End Property
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = lbl.ForeColor
End Property
Public Property Let ForeColor(value As OLE_COLOR)
   lbl.ForeColor = value
   PropertyChanged "ForeColor"  ' Signal that the property should be saved in the WriteProperties event
End Property
Private Sub UserControl_InitProperties()
   Set lbl.Font = UserControl.Font   ' Initialize to the current font
   lbl.ForeColor = UserControl.ForeColor
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   On Error Resume Next
   Set lbl.Font = PropBag.ReadProperty("Font")
   lbl.ForeColor = PropBag.ReadProperty("ForeColor")
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "Font", Me.Font  ' Persist the font
   PropBag.WriteProperty "ForeColor", Me.ForeColor
End Sub
Public Property Let ReadOnly(ByVal value As Boolean)
  m_ReadOnly = True
  Call UserControl_Resize
End Property
Public Property Get ReadOnly() As Boolean
  ReadOnly = m_ReadOnly
End Property

