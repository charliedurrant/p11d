VERSION 5.00
Begin VB.UserControl FolderBrowser 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ScaleHeight     =   555
   ScaleWidth      =   4785
   Begin VB.CommandButton cmdBrowse 
      Height          =   330
      Left            =   0
      Picture         =   "FolderBrowser.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lbl 
      Height          =   285
      Left            =   405
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "FolderBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Folder As String
Private m_Title As String
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
End Sub

Private Sub UserControl_Resize()
  lbl.Height = UserControl.Height
  lbl.Width = UserControl.Width - (cmdBrowse.Width - 3)
End Sub
Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
 UserControl.Enabled = NewValue
 lbl.Enabled = NewValue
 cmdBrowse.Enabled = NewValue
End Property
