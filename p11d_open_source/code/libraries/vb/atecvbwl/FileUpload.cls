VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiPartFormElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cadxx completely replaced
Option Explicit
Public Enum MULTU_PART_FORM_DATA_TYPE
  MPFDT_BINARY
  MPFDT_TEXT
End Enum
Public ContentType As String
Public FileName As String
Public Path As String
Public Data As Variant
Public Name As String
Public DataType As MULTU_PART_FORM_DATA_TYPE
Public Sub SaveToDisk(ByVal PathAndFile As String)
  Dim i As Integer
  Dim bOpened As Boolean
  Dim b() As Byte
  
  On Error GoTo SaveToDisk_ERR
  i = FreeFile
  Open PathAndFile For Binary Access Write As #i
  bOpened = True
  'CAD todo copy b() ptr to point to variant data then put
  b = Data
  Put #i, , b
  Close #i
  
SaveToDisk_END:
  Exit Sub
SaveToDisk_ERR:
  If bOpened Then
    Close #i
    Call xKill(PathAndFile)
  End If
  Call Err.Raise(ERR_MULTI_PART_FORM, ErrorSourceComponentEx(Err, "SaveToDisk", COMPONENT_NAME), Err.Description)
  Resume
End Sub

