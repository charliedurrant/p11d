VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl FolderBrowser 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ScaleHeight     =   555
   ScaleWidth      =   4785
   Begin MSComctlLib.ImageList iml 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderBrowser.ctx":0000
            Key             =   "ICON_ENABLED"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FolderBrowser.ctx":0112
            Key             =   "ICON_DISABLED"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   330
      Left            =   0
      Picture         =   "FolderBrowser.ctx":0224
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
Attribute VB_Exposed = True
Option Explicit

Private m_Folder As String
Private m_Title As String
Private m_Style As FolderBrowserStyle

Private m_DefaultStyle As FolderBrowserStyle
Private m_DefaultTitle As String

Public Event Started()
Public Event Ended()

'RK Addition of new enum for style property
Public Enum FolderBrowserStyle
  FolderBrowser = 0
  FileBrowser
End Enum

'RK Addition of filefilters
Private m_FileOpenFilter As Long
Private m_FileOpenExtensions As String

'RK Addition of Style property
Public Property Let Style(ByVal NewValue As FolderBrowserStyle)
  If Style = FolderBrowser Or Style = FileBrowser Then
    m_Style = NewValue
    m_Title = GetTitleText
    cmdBrowse.ToolTipText = GetBrowserText
    PropertyChanged "Style"
  End If
End Property

Public Property Get Style() As FolderBrowserStyle
  Style = m_Style
End Property

Public Property Get Title() As String
  Title = m_Title
End Property
Public Property Let Title(ByVal NewValue As String)
  m_Title = NewValue
  PropertyChanged "Title"
End Property
'RK QUERY: Rename property (if possible) as now serves dual purpose of file and directory path
Public Property Get Directory() As String
  Directory = lbl.Caption
End Property
Public Property Let Directory(ByVal NewValue As String)
  'RK Apply fullpath based on style property
  Select Case m_Style
    Case FileBrowser
      lbl.Caption = NewValue
    Case FolderBrowser
      lbl.Caption = FullPath(NewValue)
  End Select
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
 
 Select Case m_Style
   Case FolderBrowser
    s = FullPath(BrowseForFolderEx(UserControl.hwnd, sIntialDirectory, GetBrowserText))
   Case FileBrowser
    s = FileOpenDlgFilter(m_FileOpenFilter, GetTitleText, m_FileOpenExtensions, sIntialDirectory, False)
 End Select
 
 If (Len(s) > 0) Then
  Directory = s
  RaiseEvent Ended
 End If
End Sub
'should be added to core library
Private Function BrowseForFolderEx(ByVal hWndOwener As Long, Optional ByVal Initialdirectory As String, Optional ByVal Title As String) As String

  Dim sRet As String
  Dim c As BrowseForFolderClass
  
On Error GoTo err_Err
  Set c = New BrowseForFolderClass
  
   c.hwndOwner = hWndOwener
   If (Len(Initialdirectory) = 0) Then
    Initialdirectory = CurDir
   End If
   
   c.InitialDir = Initialdirectory
   c.FileSystemOnly = True
   c.StatusText = True
   c.Title = Title
   c.EditBox = True
   c.UseNewUI = True
   sRet = c.BrowseForFolder
   
err_End:
  BrowseForFolderEx = sRet
  Exit Function
err_Err:
  
  sRet = BrowseForFolder(hWndOwener, Initialdirectory, Title)
  Resume err_End
End Function


Private Sub UserControl_Initialize()
  m_DefaultStyle = FolderBrowser
  m_DefaultTitle = GetTitleText
  'RK TODO Interaction between Style and Title properties - This needs moved to after ReadProperties
  Select Case m_Style
    Case FolderBrowser
     m_Title = GetTitleText
     cmdBrowse.ToolTipText = GetBrowserText
    Case FileBrowser
     m_Title = GetTitleText
     cmdBrowse.ToolTipText = GetBrowserText
     m_FileOpenFilter = 1
     m_FileOpenExtensions = "Text Files (*.txt)|*.txt|Comma Separated Variable Files (*.csv)|*.csv|All Files (*.*)|*.*"
  End Select
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
 If NewValue Then
   cmdBrowse.Picture = iml.ListImages("ICON_ENABLED").Picture
 Else
   cmdBrowse.Picture = iml.ListImages("ICON_DISABLED").Picture
 End If
End Property


'RK Addition of file filters
Public Property Let FileOpenFilter(ByVal NewValue As Long)
  m_FileOpenFilter = NewValue
End Property

Public Sub AddFileExtension(ByVal DisplayName As String, ByVal FileExtension As String, Optional ByVal DefaultChoice As Boolean = False, Optional ByVal ClearExisting As Boolean = False)
  If ClearExisting Then
    m_FileOpenExtensions = ""
    m_FileOpenFilter = 1
  End If
  
  If Len(FileExtension) Then
    If DefaultChoice Then
      m_FileOpenFilter = 1
      m_FileOpenExtensions = DisplayName & "|" & FileExtension & IIf(Len(m_FileOpenExtensions), "|", "") & m_FileOpenExtensions
    Else
      m_FileOpenExtensions = m_FileOpenExtensions & IIf(Len(m_FileOpenExtensions), "|", "") & DisplayName & "|" & FileExtension
    End If
  End If
End Sub

'RK Added in persistence for Style & Title properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Style", m_Style, m_DefaultStyle)
  Call PropBag.WriteProperty("Title", m_Title, m_DefaultTitle)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Style = PropBag.ReadProperty("Style", m_DefaultStyle)
  m_Title = PropBag.ReadProperty("Title", m_DefaultTitle)
End Sub

Private Function GetBrowserText() As String
  Select Case m_Style
    Case FolderBrowser
     GetBrowserText = "Click to select a folder"
    Case FileBrowser
     GetBrowserText = "Click to select a file"
  End Select
End Function

Private Function GetTitleText() As String
  Select Case m_Style
    Case FolderBrowser
     GetTitleText = "Choose a folder"
    Case FileBrowser
     GetTitleText = "Choose a file"
  End Select
End Function

