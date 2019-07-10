VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderBrowserControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If AbacusReporter Then
Private WithEvents m_FolderBrowser As FolderBrowser
Attribute m_FolderBrowser.VB_VarHelpID = -1
Private m_fbce As IFolderBrowserControlEvents
Private m_id As Variant
Public Function Setup(ByVal fbce As IFolderBrowserControlEvents, ByVal lblPlaceHolder As Label, ByVal id As Variant) As Object
  On Error GoTo Err_Err
    
  Set m_fbce = fbce
  Set m_FolderBrowser = lblPlaceHolder.Parent.Controls.Add("atc3FolderBrowser.FolderBrowser", id)
  m_id = id
  With m_FolderBrowser
    If Not lblPlaceHolder.Container Is Nothing Then
      Set .Container = lblPlaceHolder.Container
    End If
    .Visible = True
    .top = lblPlaceHolder.top
    .left = lblPlaceHolder.left
    .Height = lblPlaceHolder.Height
    .Width = lblPlaceHolder.Width
  End With
  
  Set Setup = m_FolderBrowser
Err_End:
  Exit Function
Err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Setup", "Setup FolderBrowser Control", Err.Description)
  Resume Err_End
  Resume
End Function
#If AbacusReporter Then
Public Property Get Style() As FolderBrowserStyle
  Style = m_FolderBrowser.Style
End Property

Public Property Let Style(NewVal As FolderBrowserStyle)
  m_FolderBrowser.Style = NewVal
End Property
#End If
Public Property Get Directory() As String
  Directory = m_FolderBrowser.Directory
End Property

Public Property Let Directory(NewVal As String)
  m_FolderBrowser.Directory = NewVal
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_FolderBrowser.Enabled
End Property

Public Property Let Enabled(NewVal As Boolean)
  m_FolderBrowser.Enabled = NewVal
End Property

Public Property Get Visible() As Boolean
  Visible = m_FolderBrowser.Visible
End Property

Public Property Let Visible(NewVal As Boolean)
  m_FolderBrowser.Visible = NewVal
End Property

Public Sub AddFileExtension(ByVal DisplayName As String, ByVal FileExtension As String, Optional ByVal DefaultChoice As Boolean = False, Optional ByVal ClearExisting As Boolean = False)
   Call m_FolderBrowser.AddFileExtension(DisplayName, FileExtension, DefaultChoice, ClearExisting)
End Sub

Private Sub m_FolderBrowser_Ended()
  Call m_fbce.Ended(m_id)
End Sub

Private Sub m_FolderBrowser_Started()
  Call m_fbce.Started(m_id)
End Sub

Private Sub m_FolderBrowser_Validate(Cancel As Boolean)
  Call m_fbce.Validate(m_id, Cancel)
End Sub
#End If

