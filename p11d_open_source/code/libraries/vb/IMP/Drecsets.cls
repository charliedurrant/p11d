VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DestRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public rs As Recordset
Public Name As String
Public DisplayName  As String
Public Tag As Variant
Public HelpText As String

Public UpdateType As IMPORT_UPDATE_TYPE
Public LockUpdate As IMPORT_UPDATE_TYPE

Public PrimaryKeys As New StringList      '
Public RequiredFields As New StringList   '
Public ProcessImport As IProcessImport
Public SQLString As String
Public dB As Database

Public DefaultStatics As Collection
Public HiddenFields As StringList
Public Constraints As Collection
Public DefaultValues As Collection
Public FieldDisplayNames As Collection  'cad p11d
Private Sub Class_Initialize()
  
  On Error Resume Next
  Set DefaultStatics = New Collection
  Set HiddenFields = New StringList
  Set Constraints = New Collection
  Set DefaultValues = New Collection
  Set FieldDisplayNames = New Collection
  LockUpdate = [_NO_LOCK_UPDATES]
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  If Not dB Is Nothing Then Set dB = Nothing
  Set DefaultStatics = Nothing
  Set HiddenFields = Nothing
  Set Constraints = Nothing
  Set DefaultValues = Nothing
  Set FieldDisplayNames = Nothing
End Sub
