VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VersionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public DLLName As String
Public DLLPathSystem As String
Public DLLPathLoader As String
Public CurrentVersion As String
Public LoaderVersion As String
Public CompareVersions As Double
Public CurrentDateTime As String
Public LoaderDateTime As String
Public LoaderFileExists As Boolean
Public SystemFileExists As Boolean
