VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLDebugADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_sFileFilter As String
Private m_sFileCaption As String
Private m_sFileStartDirectory As String
Private m_FileDirectory As String

Private m_sFileFilterSystem As String
Private m_sFileCaptionSystem As String

Private m_OpenExclusive As Boolean
Private m_OpenExclusiveMessage As String
Private m_sConnectionString As String
Private Sub Class_Initialize()
  On Error Resume Next
  InitCount = InitCount + 1
  m_sFileFilter = "Database files (*.mdb)|*.mdb"
  m_sFileCaption = "Please select a database file..."
  m_OpenExclusiveMessage = "Cannot Open SQL Explorer in exclusive mode if you are in a database"
  m_sFileStartDirectory = CurDir
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  InitCount = InitCount - 1
End Sub

Public Property Let OpenExclusive(ByVal NewValue As Boolean)
  m_OpenExclusive = NewValue
End Property

Public Property Get OpenExclusive() As Boolean
  OpenExclusive = m_OpenExclusive
End Property

Public Property Let OpenExclusiveMessage(ByVal NewValue As String)
  m_OpenExclusiveMessage = NewValue
End Property

Public Property Get OpenExclusiveMessage() As String
  OpenExclusiveMessage = m_OpenExclusiveMessage
End Property

Public Property Let FileFilter(ByVal NewValue As String)
  m_sFileFilter = NewValue
End Property

Public Property Get FileFilter() As String
  FileFilter = m_sFileFilter
End Property
Public Property Let FileFilterSystem(ByVal NewValue As String)
  m_sFileFilterSystem = NewValue
End Property

Public Property Get FileFilterSystem() As String
  FileFilterSystem = m_sFileFilterSystem
End Property

Public Property Let FileStartDirectory(ByVal NewValue As String)
  m_sFileStartDirectory = NewValue
  m_FileDirectory = NewValue
End Property

Public Property Get FileStartDirectory() As String
  FileStartDirectory = m_sFileStartDirectory
End Property


Public Property Let CurrentDirectory(ByVal NewValue As String)
  m_FileDirectory = NewValue
End Property

Public Property Get CurrentDirectory() As String
  CurrentDirectory = m_FileDirectory
End Property

Public Property Let FileCaption(ByVal NewValue As String)
  m_sFileCaption = NewValue
End Property

Public Property Get FileCaption() As String
  FileCaption = m_sFileCaption
End Property
Public Property Let FileCaptionSystem(ByVal NewValue As String)
  m_sFileCaptionSystem = NewValue
End Property

Public Property Get FileCaptionSystem() As String
  FileCaptionSystem = m_sFileCaptionSystem
End Property

Public Sub Show(Optional ByVal vConnection As Variant, Optional ByVal RequirePassword As Boolean = True, Optional ByVal Target As DATABASE_TARGET = DB_TARGET_JET, Optional ByVal ParentForm As Object)
  If RequirePassword Then
    If Not GetTCSPasswordAlways2("SQL Debug View") Then Exit Sub
  End If
  
  m_sConnectionString = frmSQL.InitialiseForm(vConnection, Me, Target)
  If (Not ParentForm Is Nothing) Then
    If (TypeOf ParentForm Is Form) And Not IsRunningInIDE Then
      frmSQL.Show , ParentForm 'no modal as f12 click fails
    End If
  Else
    frmSQL.Show
  End If
End Sub

