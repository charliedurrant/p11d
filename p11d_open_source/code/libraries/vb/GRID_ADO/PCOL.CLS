VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'All

Public RealName As String
Public DisplayName As String
Public lID As Long
Public FieldCodes As String
Public Width As Long            'W
Public Precision As Long        'P
Public Significant As Long      'S
Public DataType As String       'T
Public DateFormat As String     'D
Public NumberFormat As String   'Combination of Precision/Significant/Width
Public Alignment As Long        'A
Public Hide As Boolean          'H

'Print
Public SubTotal As Boolean      'S
Public PageBreak As Boolean     'N
Public Header As Boolean        'H
Public Sum As Boolean           'T

'Screen
Public QueryName As String      'v#{n}sqlname{p1}{p2}{n}#
Public Drop As Integer          'D#{p1}{p2}{p3}#
Public Items As Integer         'I
Public QueryStyle As Integer    'v#{n}sqlname{p1}{p2}{n}# n=0 List n=1 combo - Required!!
Private Parameters() As String  'Array of Strings containing all parameters
Public NoEdit As Boolean        'E
Public NewRecord As String      'N#fldName#

'Dialog
Public NewColumn As Boolean     'C
Public TabTitle As String       't#tabtitle

Private Sub Class_Initialize()
  Alignment = vbLeftJustify
  ReDim Parameters(0)
  Items = 6
End Sub

Public Sub AddParam(sParam As String)
  ReDim Preserve Parameters(UBound(Parameters()) + 1)
  Parameters(UBound(Parameters)) = sParam
End Sub

Public Function GetParam(l As Long) As String
  If (l > 0) And (l <= UBound(Parameters())) Then
    GetParam = Parameters(l)
  Else
    GetParam = ""
  End If
End Function

Public Function NoParams() As Long
  NoParams = UBound(Parameters())
End Function
