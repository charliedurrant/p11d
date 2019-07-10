VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReportPreProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' datasetpath is from root to leaf ( 1 .. N )
Public Function BeforeReportPrepare(ByVal SessionName As String) As Boolean
  ' return true to abort report
End Function

Public Function BeforeReportOutput(ByVal SessionName As String, DataSetPath() As ReportDataSet) As Boolean
  ' return true to abort report
End Function

Public Sub AfterReportOutput(ByVal SessionName As String, DataSetPath() As ReportDataSet)

End Sub


