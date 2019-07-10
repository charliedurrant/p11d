VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocaleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetSystemDefaultLcidEx Lib "kernel32" Alias "GetSystemDefaultLCID" () As Long
Private Declare Function GetUserDefaultLcidEx Lib "kernel32" Alias "GetUserDefaultLCID" () As Long
Private Declare Function GetLocaleInfoEx Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Enum LocalInfo
  LOCALE_S1159 = &H28        '  AM designator
  LOCALE_S2359 = &H29        '  PM designator
  LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
  LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
  LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
  LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
  LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
  LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
  LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
  LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
  LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
  LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
  LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
  LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
  LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
  LOCALE_SABBREVMONTHNAME13 = &H100F
  LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
  LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
  LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
  LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
  LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
  LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
  LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
  LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
  LOCALE_SCOUNTRY = &H6         '  localized name of country
  LOCALE_SCURRENCY = &H14        '  local monetary symbol
  LOCALE_SDATE = &H1D        '  date separator
  LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
  LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
  LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
  LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
  LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
  LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
  LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
  LOCALE_SDECIMAL = &HE         '  decimal separator
  LOCALE_SENGCOUNTRY = &H1002      '  English name of country
  LOCALE_SENGLANGUAGE = &H1001      '  English name of language
  LOCALE_SGROUPING = &H10        '  digit grouping
  LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
  LOCALE_SLANGUAGE = &H2         '  localized name of language
  LOCALE_SLIST = &HC         '  list item separator
  LOCALE_SLONGDATE = &H20        '  long date format string
  LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
  LOCALE_SMONGROUPING = &H18        '  monetary grouping
  LOCALE_SMONTHNAME1 = &H38        '  long name for January
  LOCALE_SMONTHNAME10 = &H41        '  long name for October
  LOCALE_SMONTHNAME11 = &H42        '  long name for November
  LOCALE_SMONTHNAME12 = &H43        '  long name for December
  LOCALE_SMONTHNAME2 = &H39        '  long name for February
  LOCALE_SMONTHNAME3 = &H3A        '  long name for March
  LOCALE_SMONTHNAME4 = &H3B        '  long name for April
  LOCALE_SMONTHNAME5 = &H3C        '  long name for May
  LOCALE_SMONTHNAME6 = &H3D        '  long name for June
  LOCALE_SMONTHNAME7 = &H3E        '  long name for July
  LOCALE_SMONTHNAME8 = &H3F        '  long name for August
  LOCALE_SMONTHNAME9 = &H40        '  long name for September
  LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
  LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
  LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9
  LOCALE_SNATIVELANGNAME = &H4         '  native name of language
  LOCALE_SNEGATIVESIGN = &H51        '  negative sign
  LOCALE_SPOSITIVESIGN = &H50        '  positive sign
  LOCALE_SSHORTDATE = &H1F        '  short date format string
  LOCALE_STHOUSAND = &HF         '  thousand separator
  LOCALE_STIME = &H1E        '  time separator
  LOCALE_STIMEFORMAT = &H1003      '  time format string
End Enum

Public Enum LocaleScope
  LOCALE_SYSTEM_DEFAULT
  LOCALE_USER_DEFAULT
End Enum

Public Function GetSystemDefaultLcid() As Long
  GetSystemDefaultLcid = GetSystemDefaultLcidEx()
End Function

Public Function GetUserDefaultLcid() As Long
  GetUserDefaultLcid = GetUserDefaultLcidEx()
End Function

Public Function GetLocaleValue(ByVal Lcid As LocaleScope, Linfo As LocalInfo) As String
  Dim buffer As String
  Dim retval As Long
    
  If Lcid = LOCALE_SYSTEM_DEFAULT Then
    Lcid = GetSystemDefaultLcidEx()
  Else
    Lcid = GetUserDefaultLcidEx()
  End If
  retval = GetLocaleInfoEx(Lcid, Linfo, vbNullString, 0)
  If retval = 0 Then Err.Raise ERR_GETVALUE, "GetLocaleValue", "Error retrieving locale value size Linfo (0x" & right$("0000" & Hex(Linfo), 4) & ") LastError (" & Err.LastDllError & ")"
  buffer = String$(retval, vbNullChar)
  retval = GetLocaleInfoEx(Lcid, Linfo, buffer, retval)
  buffer = RTrimChar(buffer, vbNullChar)
  If retval = 0 Then Err.Raise ERR_GETVALUE, "GetLocaleValue", "Error retrieving locale value 0x" & right$("0000" & Hex$(Linfo), 4) & " LastError (" & Err.LastDllError & ")"
  GetLocaleValue = buffer
End Function

