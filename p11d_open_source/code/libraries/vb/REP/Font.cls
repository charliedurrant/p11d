VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontHeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Heights(1 To 50) As Single
Public Function Init(ByVal FontName As String, prt As PRIVATE_REPORT_TARGET)
  Dim i As Long
  Dim sOldFontName As String
  Dim iOldFontSize As Long
  
  sOldFontName = ReportControl.Preview.font.Name
  iOldFontSize = ReportControl.Preview.font.Size
  ReportControl.Preview.FontName = FontName
  For i = 1 To MAX_POINT_SIZE
    If (ReportControl.UseExactFontHeight Or True) Then
      m_Heights(i) = GetTextHeightFromFontStyle(i, "Ay")
    Else
      m_Heights(i) = FontHeights(i)
    End If
  Next
  ReportControl.Preview.font.Name = sOldFontName
  ReportControl.Preview.font.Size = iOldFontSize
End Function
Public Function Height(ByVal PointSize As Long) As Single
  Height = m_Heights(PointSize)
End Function
