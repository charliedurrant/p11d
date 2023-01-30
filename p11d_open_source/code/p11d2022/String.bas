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
Public Function CsvField(ByVal value As String) As String
  Dim needsQuotes As Boolean
  Dim p0 As Long
  
  If (Len(value) = 0) Then
    CsvField = value
    Exit Function
  End If
  
  
  p0 = InStr(1, value, """")
  If (p0 > 0) Then
    value = Replace(value, """", """""")
    needsQuotes = True
  End If
  
  p0 = InStr(1, value, vbCrLf)
  If (p0 > 0) Then
    needsQuotes = True
  End If
  
  p0 = InStr(1, value, ",")
  If (p0 > 0) Then
    needsQuotes = True
  End If
  
  If (needsQuotes) Then
    value = """" & value & """"
  End If
  CsvField = value
End Function
