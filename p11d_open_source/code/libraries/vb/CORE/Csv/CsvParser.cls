VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CsvParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_csv As tCSVResult
Private m_FirstRowIsHeader As Boolean
Public Sub Init(csv As String, firstRowIsHeader As Boolean)
  m_csv = ParseCsv(csv)
  m_FirstRowIsHeader = firstRowIsHeader
End Sub
Public Function RowCount() As Long
  If (m_FirstRowIsHeader) Then
      RowCount = m_csv.RowCount - 1
  Else
    RowCount = m_csv.RowCount
  End If
End Function
Public Function Value(row As Long, field As String) As String
  If (Not m_FirstRowIsHeader) Then
    Call Err.Raise(ERR_ERROR, "Value", "FirstrowAsHeader = false so can not access by field name")
  End If
  Value = m_csv.Rows(row + 1).Columns(FieldIndex(field))
End Function
Public Function ValueByIndex(row As Long, field As Long) As String
  ValueByIndex = m_csv.Rows(row).Columns(field)
End Function
Private Function FieldIndex(field As String)
  FieldIndex = -1
  Dim i As Long
  For i = 0 To m_csv.ColumnCount - 1
    If (field = m_csv.Rows(0).Columns(i)) Then
      FieldIndex = i
      Exit Function
    End If
  Next
End Function
