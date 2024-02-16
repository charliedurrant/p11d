VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestingSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sheet As String
Public ImportSpec As String
Public ReportSpec As String
Public ReportResultsSheet As String

Public Property Get ReportEnabled() As Boolean
  ReportEnabled = Len(ReportSpec) > 0 And Len(ReportResultsSheet)
End Property

