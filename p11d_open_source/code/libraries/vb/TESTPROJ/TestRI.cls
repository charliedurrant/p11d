VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestRepInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Value As String
Private frm As frmRepInterface
Private Const SHOW_FORM As Boolean = False
Implements IReportPreProcess

Private Function IReportPreProcess_BeforeReportPrepare(ByVal SessionName As String) As Boolean
  Set frm = New frmRepInterface
  
  frm.lblSession = "Session: " & SessionName
  frm.sbrTest.SimpleText = "BeforeReportPrepare"
  If SHOW_FORM Then frm.Show vbModal
  Value = frm.txtValue
End Function

Private Function IReportPreProcess_BeforeReportOutput(ByVal SessionName As String, DataSetPath() As ReportDataSet) As Boolean
  frm.lblSession = "Session: " & SessionName
  frm.lblDataset = "Dataset: " & DataSetPath(1).Name
  frm.lblCriteria = "CriteriaSQL: " & DataSetPath(1).CriteriaSQL
  frm.sbrTest.SimpleText = "BeforeReportOutput"
  If SHOW_FORM Then frm.Show vbModal
End Function

Private Sub IReportPreProcess_AfterReportOutput(ByVal SessionName As String, DataSetPath() As ReportDataSet)
  frm.lblSession = "Session: " & SessionName
  frm.lblDataset = "Dataset: " & DataSetPath(1).Name
  frm.sbrTest.SimpleText = "AfterReportOutput"
  If SHOW_FORM Then frm.Show vbModal
End Sub

