VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataConversion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum DATECONVERT_TYPE ' note: see VTEXT if changed
  CONVERT_DATETYPEUNKNOWN = 0
  CONVERT_FIXEDDATE = 256
  CONVERT_DELIMITED
  CONVERT_FIXEDDATETIME
End Enum

Public Function DataTypeName(ByVal dType As DATABASE_FIELD_TYPES) As String
  Select Case dType
    Case [_TYPE_INVALID]
      DataTypeName = "Invalid"
    Case TYPE_LONG
      DataTypeName = "Long"
    Case TYPE_STR
      DataTypeName = "String"
    Case TYPE_DOUBLE
      DataTypeName = "Double"
    Case TYPE_DATE
      DataTypeName = "Date"
    Case TYPE_BOOL
      DataTypeName = "Boolean"
    Case Else
      Err.Raise ERR_DATACONV, "DataTypeName", "Unrecognised data type: " & CStr(dType)
  End Select
End Function

Public Function VarTypetoDatatype(ByVal vbType As VbVarType) As DATABASE_FIELD_TYPES
  Select Case vbType
    Case vbCurrency, vbDecimal, vbDouble, vbSingle
      VarTypetoDatatype = TYPE_DOUBLE
    Case vbInteger, vbLong, vbByte
      VarTypetoDatatype = TYPE_LONG
    Case vbDate
      VarTypetoDatatype = TYPE_DATE
    Case vbString
      VarTypetoDatatype = TYPE_STR
    Case vbBoolean
      VarTypetoDatatype = TYPE_BOOL
    Case Else
      Err.Raise ERR_DATACONV, "VarTypetoDatatype", "Unrecognised data type: " & CStr(vbType)
    End Select
End Function

Public Function IsNumberField(ByVal dbDataType As DATABASE_FIELD_TYPES) As Boolean
  IsNumberField = (dbDataType = TYPE_DOUBLE) Or (dbDataType = TYPE_LONG)
End Function


'##D Modifies a date by changing any specific field
'##V ChangeDate           Date to be modified
'##V Day               Overide day (Default -1)
'##V Month              Overide month (Default -1)
'##V Year               Overide year (Default -1)
'##V Hour              Overide hours (Default -1)
'##V Minute             Overide minutes (Default -1)
'##V Seconds            Overide seconds (Default -1)
'##LD Passing a -1 (the default value) means leaves that portion of the date unchanged
'##RV The changed date
'##C DTF
Public Function SetDateTime(ByVal ChangeDate As Date, Optional ByVal Day As Integer = -1, Optional ByVal Month As Integer = -1, Optional ByVal Year As Integer = -1, Optional ByVal Hour As Integer = 0, Optional ByVal Minute As Integer = 0, Optional ByVal Seconds As Integer = 0) As Date
  Dim nday0 As Integer, nmonth0 As Integer, nyear0 As Integer
  Dim nhour0 As Integer, nminute0 As Integer, nsecond0 As Integer
  
  On Error GoTo SetDateTime_err
  ' date
  If Day <> -1 Then
    nday0 = Day
  Else
    nday0 = DatePart("d", ChangeDate)
  End If
  If Month <> -1 Then
    nmonth0 = Month
  Else
    nmonth0 = DatePart("m", ChangeDate)
  End If
  If Year <> -1 Then
    nyear0 = Year
  Else
    nyear0 = DatePart("yyyy", ChangeDate)
  End If
  
  ' time
  If Hour <> -1 Then
    nhour0 = Hour
  Else
    nhour0 = DatePart("h", ChangeDate)
  End If
  If Minute <> -1 Then
    nminute0 = Minute
  Else
    nminute0 = DatePart("n", ChangeDate)
  End If
  If Seconds <> -1 Then
    nsecond0 = Seconds
  Else
    nsecond0 = DatePart("s", ChangeDate)
  End If
  SetDateTime = DateSerialEx(nyear0, nmonth0, nday0) + TimeSerial(nhour0, nminute0, nsecond0)
  Exit Function
  
SetDateTime_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "SetDateTime"), Err.Description
End Function

'##D Attempt to convert a string of the form DD/MM/YY HH:NN:SS to a date
'##LD If the string is not a valid date then returns DefaultDate
'##C DTF
Public Function TryConvertDate(ByVal DateString As Variant, Optional ByVal DefaultDate As Date = UNDATED) As Date
  On Error GoTo TryConvertDate_err
  
  TryConvertDate = ConvertDate(DateString, CONVERT_DELIMITED, "DMY", "/", ":")
  
TryConvertDate_end:
  Exit Function
  
TryConvertDate_err:
  TryConvertDate = DefaultDate
  Resume TryConvertDate_end
End Function

'##D Attempt to convert a string of the form DD/MM/YY HH:NN:SS to a date
'##LD If the string is not a valid date then returns DefaultDate if present, or DateString if DefaultDate is not specified
'##C DTF
' note: see VTEXT if changed
Public Function TryConvertDateDMY(ByVal DateString As Variant, Optional ByVal DefaultDate As Variant) As Variant
  On Error GoTo TryConvertDateDMY_err
  
  TryConvertDateDMY = ConvertDate(DateString, CONVERT_DELIMITED, "DMY", "/", ":")
  
TryConvertDateDMY_end:
  Exit Function
  
TryConvertDateDMY_err:
  If IsMissing(DefaultDate) Then
    TryConvertDateDMY = DateString
  Else
    TryConvertDateDMY = DefaultDate
  End If
  Resume TryConvertDateDMY_end
End Function

'##D Converts Fixed/Delimited date string to dates
'##V ConvStr order DMY or MDY (or DMYHNS) for delimited dates, mask for fixed dates
'##C DTF
Public Function ConvertDate(ByVal DateString As String, ByVal ConvType As DATECONVERT_TYPE, ByVal ConvStr As String, Optional ByVal DateDelimit As String = "/", Optional ByVal TimeDelimit As String = ":") As Date
  Dim i As Long, j As Long, p0 As Long, p1 As Long, ch As String
  Dim nyear As Integer, nmonth As Integer, nday As Integer
  Dim nhour As Integer, nminute As Integer, nsecond As Integer
  Dim d0 As Date

  On Error GoTo ConvertDate_err
  If (ConvType = CONVERT_FIXEDDATE) Or (ConvType = CONVERT_FIXEDDATETIME) Then
    i = InStr(1, ConvStr, "Y", vbTextCompare)
    j = InStrRev(ConvStr, "Y", , vbTextCompare)
    If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Year value must allow at least 2 digits"
    
    nyear = CLng(Mid$(DateString, i, j - i + 1))
    nyear = GetFullYear_CD(nyear, ConvStr)
        
    i = InStr(1, ConvStr, "D", vbTextCompare)
    j = InStrRev(ConvStr, "D", , vbTextCompare)
    If (i = 0) And (j = 0) Then
      nday = 1&
    Else
      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Day value must allow 2 digits"
      nday = CLng(Mid$(DateString, i, j - i + 1))
    End If
  
    i = InStr(1, ConvStr, "M", vbTextCompare)
    j = InStrRev(ConvStr, "M", , vbTextCompare)
    If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Month value must allow 2 digits"
    nmonth = CLng(Mid$(DateString, i, j - i + 1))
    If ConvType = CONVERT_FIXEDDATETIME Then
      i = InStr(1, ConvStr, "H", vbTextCompare)
      j = InStrRev(ConvStr, "H", , vbTextCompare)
      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Hour value must allow 2 digits"
      nhour = CLng(Mid$(DateString, i, j - i + 1))
     
      i = InStr(1, ConvStr, "N", vbTextCompare)
      j = InStrRev(ConvStr, "N", , vbTextCompare)
      If (j - i) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCr & "Minute value must allow 2 digits"
      nminute = CLng(Mid$(DateString, i, j - i + 1))
           
      i = InStr(1, ConvStr, "S", vbTextCompare)
      j = InStrRev(ConvStr, "S", , vbTextCompare)
      If (j - i) >= 1 Then nsecond = CLng(Mid$(DateString, i, j - i + 1))
    End If
  ElseIf ConvType = CONVERT_DELIMITED Then
    i = 1: p0 = 1: p1 = 1
    nday = 1
    For i = 1 To Len(ConvStr)
      ch = UCase$(Mid$(ConvStr, i, 1))
      If InStr("HNS", ch) = 0 Then
        p1 = InStr(p0, DateString, DateDelimit, vbTextCompare)
      Else
        p1 = InStr(p0, DateString, TimeDelimit, vbTextCompare)
      End If
      If p1 = 0 Then p1 = InStr(p0, DateString, " ", vbTextCompare)
      If p1 = 0 Then
        If i = Len(ConvStr) Then p1 = Len(DateString) + 1
        If (p1 - p0) < 1 Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr
      End If
      Select Case ch
        Case "D"
            nday = CLng(Mid$(DateString, p0, p1 - p0))
        Case "M"
            nmonth = CLng(Mid$(DateString, p0, p1 - p0))
        Case "Y"
            nyear = CLng(Mid$(DateString, p0, p1 - p0))
            nyear = GetFullYear_CD(nyear, ConvStr)
        Case "H"
            nhour = CLng(Mid$(DateString, p0, p1 - p0))
        Case "N"
            nminute = CLng(Mid$(DateString, p0, p1 - p0))
        Case "S"
            nsecond = CLng(Mid$(DateString, p0, p1 - p0))
        Case Else
            Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format contains invalid charactere: " & ConvStr & vbCr & "Date must have a Year, Month and Day order - DMY and HNS for Hours, minutes and seconds"
      End Select
      p0 = p1 + 1
    Next i
  Else
    If Not IsDate(DateString) Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Unknown Date format"
    d0 = CDate(DateString)
    nday = DatePart("d", d0)
    nmonth = DatePart("m", d0)
    nyear = DatePart("yyyy", d0)
    nhour = DatePart("h", d0)
    nminute = DatePart("n", d0)
    nsecond = DatePart("s", d0)
  End If
  d0 = DateSerialEx(nyear, nmonth, nday) + TimeSerial(nhour, nminute, nsecond)
  If (Day(d0) <> nday) Or (Month(d0) <> nmonth) Or (Year(d0) <> nyear) Then Err.Raise ERR_CONVERTDATE, "ConvertDate", "Failed to convert date. Converted to " & Format$(d0, "DD/MM/YYYY") & " (DD/MM/YYYY)"
  ConvertDate = d0
  Exit Function
  
ConvertDate_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "ConvertDate"), "Unable to convert string " & DateString & " to a date" & vbCrLf & Err.Description
End Function

'##D Takes a Variant converts into either a Date String of the format specified or the DefaultValue if is is not a date
'##C DTF
Public Function DateString2(ByVal v As Variant, Optional ByVal DefaultValue As String = "", Optional ByVal FormatString As String = "DD/MM/YYYY") As String
  
  On Error GoTo DateString2_err
  If VarType(v) <> vbDate Then
    If Month(CDate("01/02/2010")) = 2 Then
      v = ConvertDate(v, CONVERT_DELIMITED, "DMY", "/", ":")
    Else
      v = ConvertDate(v, CONVERT_DELIMITED, "MDY", "/", ":")
    End If
  End If
  
  If v = UNDATED Then
    DateString2 = DefaultValue
  Else
    DateString2 = Format$(v, FormatString)
  End If
  
DateString2_end:
  Exit Function
  
DateString2_err:
  DateString2 = DefaultValue
  Resume DateString2_end
End Function

'##D Takes a Variant converts into either a Date String of the form DD/MM/YYYY or the Empty String if is is not a date
'##C DTF
Public Function DateString(ByVal v As Variant) As String
  DateString = DateString2(v, "", "DD/MM/YYYY")
End Function

'##D Takes a variant converts into either a Date String of the form DD/MM/YYYY or the DefaultValue if is is not a date
'##C DTF
Public Function DateStringEx(ByVal v As Variant, ByVal DefaultValue As String) As String
  DateStringEx = DateString2(v, DefaultValue, "DD/MM/YYYY")
End Function


Public Function IsNullEx(ByVal v As Variant, ByVal NullValue As Variant) As Variant
  If IsNull(v) Then
    IsNullEx = NullValue
  Else
    IsNullEx = v
  End If
End Function

'##D Converts a string to a Value of a specific data type
'##LD This is normally used for converting from saved representations ( web forms etc. ) to typed values
'##LD Returns null if type cannot be converted or it is the empty string
'##LD GetTypedValueAsString is used for converting back
'##C GCF
Public Function GetTypedValueFromString(ByVal sValue As String, ByVal dType As DATABASE_FIELD_TYPES) As Variant
  On Error GoTo GetTypedValueFromString_err
  sValue = Trim$(sValue)
  If Len(sValue) = 0 Then
    GetTypedValueFromString = Null
  Else
    GetTypedValueFromString = GetTypedValue(sValue, dType)
  End If
GetTypedValueFromString_end:
  Exit Function
  
GetTypedValueFromString_err:
  GetTypedValueFromString = Null
  Resume GetTypedValueFromString_end
End Function

'##D Converts a variant to a Value of a specific data type
'##LD If conversion fails then raises an error
'##C GCF
Public Function GetTypedValue(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As Variant
    Select Case dType
    Case TYPE_STR
      GetTypedValue = CStr(v)
    Case TYPE_LONG
      GetTypedValue = CLng(v)
    Case TYPE_DOUBLE
      GetTypedValue = CDbl(v)
    Case TYPE_DATE
      GetTypedValue = Null
      If VarType(v) = vbDate Then
        GetTypedValue = v
      ElseIf VarType(v) = vbString Then
        If InStr(1, v, "/", vbBinaryCompare) > 0 Then
          If InStr(v, ":") > 0 Then
            GetTypedValue = ConvertDate(v, CONVERT_DELIMITED, "DMYHNS", "/", ":")
          Else
            GetTypedValue = ConvertDate(v, CONVERT_DELIMITED, "DMY", "/", ":")
          End If
        ElseIf InStr(1, v, "-", vbBinaryCompare) > 0 Then
          If InStr(v, ":") > 0 Then
            GetTypedValue = ConvertDate(v, CONVERT_DELIMITED, "DMYHNS", "-", ":")
          Else
            GetTypedValue = ConvertDate(v, CONVERT_DELIMITED, "YMD", "-", ":")
          End If
        End If
      End If
      If IsNull(GetTypedValue) Then GetTypedValue = CDate(v)
      If GetTypedValue = UNDATED Then Err.Raise ERR_INVALID_TYPE, "GetTypedValue", "Unable to convert " & CStr(v) & " to a date."
    Case TYPE_BOOL
      If Not IsBoolean(v) Then Err.Raise ERR_INVALID_TYPE, "GetTypedValue", "Unable to convert " & CStr(v) & " to a boolean value."
      GetTypedValue = CBoolean(v)
    Case Else
      Err.Raise ERR_INVALID_TYPE, "GetTypedValue", "Unrecognised type: " & CStr(dType)
  End Select
End Function

'##D Converts a variant to a Value of a specific data type
'##LD If conversion fails then returns DefaultValue
'##C GCF
Public Function GetTypedValueDefault(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES, ByVal DefaultValue As Variant) As Variant
  On Error GoTo GetTypedValueDefault_err
  GetTypedValueDefault = GetTypedValue(v, dType)
  
GetTypedValueDefault_end:
  Exit Function
  
GetTypedValueDefault_err:
  GetTypedValueDefault = DefaultValue
  Resume GetTypedValueDefault_end
End Function

'##D Converts a variant to a Value of a specific data type
'##LD Deals with Null values, GetTypedValueFromString will convert back to a variant
'##C GCF
Public Function GetTypedValueAsString(ByVal vValue As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  Select Case dType
    Case TYPE_BOOL
      vValue = CBoolean(IsNullEx(vValue, False))
      vValue = CLng(vValue)
    Case TYPE_DOUBLE, TYPE_LONG
      vValue = IsNullEx(vValue, "")
    Case TYPE_DATE
      vValue = IsNullEx(vValue, UNDATED)
      If vValue = UNDATED Then
        vValue = ""
      Else
        vValue = DateString(vValue)
      End If
    Case Else
      vValue = IsNullEx(vValue, "")
  End Select
  GetTypedValueAsString = vValue
End Function

'##D Converts Value to a string
'##LD Does not deal with Null values, GetTypedValue will convert back to a variant
'##C GCF
Public Function GetValueAsString(ByVal v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As String
  If IsNull(v) Then Err.Raise ERR_INVALID_TYPE, "GetValueAsString", "Unable to convert a Null Value to a string"
  v = GetTypedValue(v, dType)
  Select Case dType
    Case TYPE_STR, TYPE_LONG, TYPE_DOUBLE
      GetValueAsString = CStr(v)
    Case TYPE_DATE
      GetValueAsString = DateString(v)
    Case TYPE_BOOL
      GetValueAsString = IIf(v, "True", "False")
    Case Else
      Err.Raise ERR_INVALID_TYPE, "GetValueAsString", "Unrecognised type: " & CStr(dType)
  End Select
End Function

'##D Rounds Given Number to nearest Integer
'##C NF
Public Function RoundN(ByVal Number As Double, Optional ByVal DecimalPlaces As Long = 2) As Double
  RoundN = RoundDouble(Number, DecimalPlaces, R_NORMAL)
End Function
'##D Rounds Up Given Number (Must not be negative)
'##C NF
Public Function RoundUp(ByVal Number As Double, Optional ByVal DecimalPlaces As Long = 2) As Double
  If Number < 0 Then Err.Raise ERR_DATACONV, "RoundUp", "Unable to round negative numbers up"
  RoundUp = RoundDouble(Number, DecimalPlaces, R_UP)
End Function

'##D Rounds Down Given Number (must not be negative)
'##C NF
Public Function RoundDown(ByVal Number As Double, Optional ByVal DecimalPlaces As Long = 2) As Double
  If Number < 0 Then Err.Raise ERR_DATACONV, "RoundDown", "Unable to round negative numbers down"
  RoundDown = RoundDouble(Number, DecimalPlaces, R_DOWN)
End Function

Public Function IsVarNumeric(v As Variant) As Boolean
  IsVarNumeric = (VarType(v) = vbInteger) Or (VarType(v) = vbLong) Or (VarType(v) = vbSingle) Or (VarType(v) = vbDouble) Or (VarType(v) = vbByte) Or (VarType(v) = vbDecimal) Or (VarType(v) = vbCurrency)
End Function

Public Function IsDateTime(ByVal d0 As Date) As Boolean
  IsDateTime = Not (Fix(d0) = d0)
End Function

'##D Determines whether a variable is Boolean
'##LD Takes account of numbers and text true/false, on/off, yes/no
'##C GCF
Public Function IsBoolean(v As Variant) As Boolean
  If IsVarNumeric(v) Or (VarType(v) = vbBoolean) Then
    IsBoolean = True
  ElseIf VarType(v) = vbString Then
    IsBoolean = (StrComp(v, "True", vbTextCompare) = 0) Or _
                (StrComp(v, "False", vbTextCompare) = 0) Or _
                (StrComp(v, "on", vbTextCompare) = 0) Or _
                (StrComp(v, "off", vbTextCompare) = 0) Or _
                (StrComp(v, "yes", vbTextCompare) = 0) Or _
                (StrComp(v, "no", vbTextCompare) = 0) Or _
                (StrComp(v, "-1", vbTextCompare) = 0) Or _
                (StrComp(v, "0", vbTextCompare) = 0)
  End If
End Function

'##D Converts Variable to Boolean
'##LD Takes account of numbers and text true/false, on/off, yes/no
'##C GCF
Public Function CBoolean(v As Variant) As Boolean
 If VarType(v) = vbBoolean Then
    CBoolean = v
  ElseIf IsVarNumeric(v) Then
    CBoolean = Not (v = 0)
  ElseIf VarType(v) = vbString Then
    CBoolean = (StrComp(v, "True", vbTextCompare) = 0) Or (StrComp(v, "on", vbTextCompare) = 0) Or (StrComp(v, "yes", vbTextCompare) = 0) Or (StrComp(v, "-1", vbTextCompare) = 0)
  End If
End Function

Public Function GetDecimalPrecision(ByVal sFormat As String, Optional ByVal DecimalPlaceholder As String = ".") As Long
  Dim p0 As Long, p1 As Long
  
  GetDecimalPrecision = 0
  p0 = InStr(1, sFormat, DecimalPlaceholder, vbBinaryCompare)
  If p0 > 0 Then
    p1 = NotInStrAny(sFormat & ";", "#0123456789", p0 + 1, vbBinaryCompare) - 1
    GetDecimalPrecision = p1 - p0
  End If
  If GetDecimalPrecision > 14 Then Err.Raise ERR_DATACONV, "GetDecimalPrecision", "Numbers (Doubles) are limited to 14 significant digits after the decimal point"
End Function



