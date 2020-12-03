Attribute VB_Name = "stringF"
Option Explicit
Public Function ServicePack() As String
  ServicePack = "" '" (Service Pack 2)"
End Function
Public Function FormatWN(ByVal v As Variant, Optional ByVal sCurrency As String = S_CURRENCY, Optional ByVal bNegative As Boolean = False, Optional b2DP As Boolean = False) As String
  Dim sFormatString As String
On Error GoTo FormatWN_err
  
  If bNegative And IsNumeric(v) Then
    If v > 0 Then v = v * -1
  End If
  If b2DP Then
    sFormatString = sCurrency & "#,##0.00;" & sCurrency & "(#,##0.00);" & sCurrency & "0"
  Else
    sFormatString = sCurrency & "#,##0;" & sCurrency & "(#,##0);" & sCurrency & "0"
  End If
  FormatWN = Format$(v, sFormatString)
  
FormatWN_end:
  Exit Function
FormatWN_err:
  FormatWN = sCurrency & v
  Resume FormatWN_end
End Function
Public Function FormatWNRPT(v As Variant, Optional sCurrency As String = S_CURRENCY, Optional bNegative As Boolean = False, Optional b2DP As Boolean = False) As String
  FormatWNRPT = Chr$(34) & FormatWN(v, sCurrency, bNegative, b2DP) & Chr$(34)
End Function
Public Function ValueOfMaxStatus(sLeadingCaption, lValue As Long, lMax As Long) As String
  ValueOfMaxStatus = sLeadingCaption & " " & lValue & " of " & lMax
End Function

