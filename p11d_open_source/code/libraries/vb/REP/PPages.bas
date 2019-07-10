Attribute VB_Name = "PreviewPages"
Option Explicit

Public Type PreviewPage
  data As QString
  complete As Boolean
  'CurrentY As Single
  PageNumber As Long
  PrePageNumber As String
  PostPageNumber  As String
  ExportOnlyFooter As String
  statics(0 To (REPORT_CONSTANTS_N - 1)) As String
End Type

'* 1 .. MAX_PAGE
Public Pages() As PreviewPage
Public PageStatics(0 To (REPORT_CONSTANTS_N - 1)) As Boolean
Public PageStaticsDefault(0 To (REPORT_CONSTANTS_N - 1)) As String

Private Const PAGE_INCREMENT As Long = 100
Private CurPage As Long
Private MAX_PAGE As Long

Public Sub AddPage()
  CurPage = CurPage + 1
  If CurPage > MAX_PAGE Then
    MAX_PAGE = MAX_PAGE + PAGE_INCREMENT
    ReDim Preserve Pages(1 To MAX_PAGE) As PreviewPage
  End If
  Set Pages(CurPage).data = New QString
End Sub

Public Sub ClearPages()
  CurPage = 0
End Sub



