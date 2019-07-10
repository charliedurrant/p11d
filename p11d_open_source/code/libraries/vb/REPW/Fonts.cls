VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontHeights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Fonts As ObjectHash
Public Property Get FontHeightCurrentReportStyle() As Single
  FontHeightCurrentReportStyle = FontHeight(ReportControl.fStyle.Name, ReportControl.fStyle.Size, ReportControl.rTarget)
End Property

Public Function FontHeight(ByVal Name As String, ByVal Size As Long, prt As PRIVATE_REPORT_TARGET) As Single
  Dim fh As FontHeight
  Dim sKey As String
  
  sKey = Name & ":" & prt
  Set fh = m_Fonts.Item(sKey, False)
  If (fh Is Nothing) Then
    Set fh = New FontHeight
    Call fh.Init(Name, prt)
    Call m_Fonts.Add(fh, sKey)
  End If
  FontHeight = fh.Height(Size)
End Function
Private Sub Class_Initialize()
  Set m_Fonts = New ObjectHash
End Sub
