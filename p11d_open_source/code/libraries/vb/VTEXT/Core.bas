Attribute VB_Name = "Core"
Option Explicit


'Public Function TryConvertDateDMY(sDate As String, Optional ByVal DefaultDate As Variant) As Variant
'  On Error GoTo TryConvertDateDMY_err
'
'  TryConvertDateDMY = ConvertDate(sDate, CONVERT_DELIMITED, "DMY")
'
'TryConvertDateDMY_end:
'  Exit Function
'
'TryConvertDateDMY_err:
'  If IsMissing(DefaultDate) Then
'    TryConvertDateDMY = sDate
'  Else
'    TryConvertDateDMY = DefaultDate
'  End If
'  Resume TryConvertDateDMY_end
'End Function
'
'Private Function ConvertDate(sDate As String, ByVal ConvType As DATECONVERT_TYPE, ConvStr As String, Optional ByVal DateDelimit As String = "/", Optional ByVal TimeDelimit As String = ":") As Date
'  Dim i As Long, j As Long, p0 As Long, p1 As Long, ch As String
'  Dim nyear As Integer, nmonth As Integer, nday As Integer
'  Dim nhour As Integer, nminute As Integer, nsecond As Integer
'  Dim d0 As Date
'
'  If (ConvType = CONVERT_FIXEDDATE) Or (ConvType = CONVERT_FIXEDDATETIME) Then
'    i = InStr(1, ConvStr, "Y", vbTextCompare)
'    j = InStrBackEx(ConvStr, "Y", 1, vbTextCompare)
'    If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Year value must allow at least 2 digits"
'
'    nyear = CLng(Mid$(sDate, i, j - i + 1))
'    nyear = GetFullYear_CD(nyear, ConvStr)
'
'    i = InStr(1, ConvStr, "D", vbTextCompare)
'    j = InStrBackEx(ConvStr, "D", 1, vbTextCompare)
'    If (i = 0) And (j = 0) Then
'      nday = 1&
'    Else
'      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Day value must allow 2 digits"
'      nday = CLng(Mid$(sDate, i, j - i + 1))
'    End If
'
'    i = InStr(1, ConvStr, "M", vbTextCompare)
'    j = InStrBackEx(ConvStr, "M", 1, vbTextCompare)
'    If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Month value must allow 2 digits"
'    nmonth = CLng(Mid$(sDate, i, j - i + 1))
'    If ConvType = CONVERT_FIXEDDATETIME Then
'      i = InStr(1, ConvStr, "H", vbTextCompare)
'      j = InStrBackEx(ConvStr, "H", 1, vbTextCompare)
'      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Hour value must allow 2 digits"
'      nhour = CLng(Mid$(sDate, i, j - i + 1))
'
'      i = InStr(1, ConvStr, "N", vbTextCompare)
'      j = InStrBackEx(ConvStr, "N", 1, vbTextCompare)
'      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Minute value must allow 2 digits"
'      nminute = CLng(Mid$(sDate, i, j - i + 1))
'
'      i = InStr(1, ConvStr, "S", vbTextCompare)
'      j = InStrBackEx(ConvStr, "S", 1, vbTextCompare)
'      If (j - i) >= 1 Then nsecond = CLng(Mid$(sDate, i, j - i + 1))
'    End If
'  ElseIf ConvType = CONVERT_DELIMITED Then
'    i = 1: p0 = 1: p1 = 1
'    nday = 1
'    For i = 1 To Len(ConvStr)
'      ch = UCase$(Mid$(ConvStr, i, 1))
'      If InStr("HNS", ch) = 0 Then
'        p1 = InStr(p0, sDate, DateDelimit, vbTextCompare)
'      Else
'        p1 = InStr(p0, sDate, TimeDelimit, vbTextCompare)
'      End If
'      If p1 = 0 Then
'        If i = Len(ConvStr) Then p1 = Len(sDate) + 1
'        If (p1 - p0) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Date not complete: " & sDate & vbCr & "Format: " & ConvStr
'      End If
'      Select Case ch
'        Case "D"
'            nday = CLng(Mid$(sDate, p0, p1 - p0))
'        Case "M"
'            nmonth = CLng(Mid$(sDate, p0, p1 - p0))
'        Case "Y"
'            nyear = CLng(Mid$(sDate, p0, p1 - p0))
'            nyear = GetFullYear_CD(nyear, ConvStr)
'        Case "H"
'            nhour = CLng(Mid$(sDate, p0, p1 - p0))
'        Case "N"
'            nminute = CLng(Mid$(sDate, p0, p1 - p0))
'        Case "S"
'            nsecond = CLng(Mid$(sDate, p0, p1 - p0))
'        Case Else
'            Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format contains invalid charactere: " & ConvStr & vbCr & "Date must have a Year, Month and Day order - DMY and HNS for Hours, minutes and seconds"
'      End Select
'      p0 = p1 + 1
'    Next i
'  Else
'    If Not IsDate(sDate) Then
'      Err.Raise ERR_CONVERTDATE, "ConvertDate", "Unknown date format: " & sDate
'    End If
'    d0 = CDate(sDate)
'    nday = DatePart("d", d0)
'    nmonth = DatePart("m", d0)
'    nyear = DatePart("yyyy", d0)
'    nhour = DatePart("h", d0)
'    nminute = DatePart("n", d0)
'    nsecond = DatePart("s", d0)
'  End If
'  d0 = DateSerial(nyear, nmonth, nday) + TimeSerial(nhour, nminute, nsecond)
'  If (Day(d0) <> nday) Or (Month(d0) <> nmonth) Or (Year(d0) <> nyear) Then
'    Err.Raise ERR_CONVERTDATE, "ConvertDate", "Failed to convert date: " & sDate & "Converted to " & Format$(d0, "DD/MM/YYYY") & " (DD/MM/YYYY)"
'  End If
'  ConvertDate = d0
'End Function
'
'Private Function InStrBackEx(String1 As String, String2 As String, Start As Long, Compare As VbCompareMethod) As Long
'  Dim pos As Long
'  Dim lastpos As Long
'
'  pos = InStr(Start, String1, String2, Compare)
'  If pos > 0 Then
'    Do
'      lastpos = pos
'      pos = InStr(pos + 1, String1, String2, Compare)
'    Loop Until pos = 0
'    pos = lastpos
'  End If
'  InStrBackEx = pos
'End Function
'
'
'Private Function GetFullYear_CD(ByVal nyear As Integer, ConvStr As String) As Integer
'  If nyear < 100 Then
'    If nyear > YEAR1900CONV Then
'      nyear = nyear + 1900
'    ElseIf nyear < YEAR2000CONV Then
'      nyear = nyear + 2000
'    Else
'      Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Year value must be in format YYYY"
'    End If
'  End If
'  GetFullYear_CD = nyear
'End Function
'
