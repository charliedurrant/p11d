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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemDefaultLcidEx Lib "kernel32" Alias "GetSystemDefaultLCID" () As Long
Private Declare Function GetUserDefaultLcidEx Lib "kernel32" Alias "GetUserDefaultLCID" () As Long
Private Declare Function GetLocaleInfoEx Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Enum LocalInfo
  LI_LOCALE_S1159 = &H28        '  AM designator
  LI_LOCALE_S2359 = &H29        '  PM designator
  LI_LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
  LI_LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
  LI_LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
  LI_LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
  LI_LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
  LI_LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
  LI_LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
  LI_LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
  LI_LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
  LI_LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
  LI_LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
  LI_LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
  LI_LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
  LI_LOCALE_SABBREVMONTHNAME13 = &H100F
  LI_LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
  LI_LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
  LI_LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
  LI_LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
  LI_LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
  LI_LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
  LI_LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
  LI_LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
  LI_LOCALE_SCOUNTRY = &H6         '  localized name of country
  LI_LOCALE_SCURRENCY = &H14        '  local monetary symbol
  LI_LOCALE_SDATE = &H1D        '  date separator
  LI_LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
  LI_LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
  LI_LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
  LI_LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
  LI_LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
  LI_LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
  LI_LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
  LI_LOCALE_SDECIMAL = &HE         '  decimal separator
  LI_LOCALE_SENGCOUNTRY = &H1002      '  English name of country
  LI_LOCALE_SENGLANGUAGE = &H1001      '  English name of language
  LI_LOCALE_SGROUPING = &H10        '  digit grouping
  LI_LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
  LI_LOCALE_SLANGUAGE = &H2         '  localized name of language
  LI_LOCALE_SLIST = &HC         '  list item separator
  LI_LOCALE_SLONGDATE = &H20        '  long date format string
  LI_LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
  LI_LOCALE_SMONGROUPING = &H18        '  monetary grouping
  LI_LOCALE_SMONTHNAME1 = &H38        '  long name for January
  LI_LOCALE_SMONTHNAME10 = &H41        '  long name for October
  LI_LOCALE_SMONTHNAME11 = &H42        '  long name for November
  LI_LOCALE_SMONTHNAME12 = &H43        '  long name for December
  LI_LOCALE_SMONTHNAME2 = &H39        '  long name for February
  LI_LOCALE_SMONTHNAME3 = &H3A        '  long name for March
  LI_LOCALE_SMONTHNAME4 = &H3B        '  long name for April
  LI_LOCALE_SMONTHNAME5 = &H3C        '  long name for May
  LI_LOCALE_SMONTHNAME6 = &H3D        '  long name for June
  LI_LOCALE_SMONTHNAME7 = &H3E        '  long name for July
  LI_LOCALE_SMONTHNAME8 = &H3F        '  long name for August
  LI_LOCALE_SMONTHNAME9 = &H40        '  long name for September
  LI_LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
  LI_LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
  LI_LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9
  LI_LOCALE_SNATIVELANGNAME = &H4         '  native name of language
  LI_LOCALE_SNEGATIVESIGN = &H51        '  negative sign
  LI_LOCALE_SPOSITIVESIGN = &H50        '  positive sign
  LI_LOCALE_SSHORTDATE = &H1F        '  short date format string
  LI_LOCALE_STHOUSAND = &HF         '  thousand separator
  LI_LOCALE_STIME = &H1E        '  time separator
  LI_LOCALE_STIMEFORMAT = &H1003      '  time format string
End Enum

Public Enum LocaleScope
  LI_LOCALE_SYSTEM_DEFAULT
  LI_LOCALE_USER_DEFAULT
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
  
  If Lcid = LI_LOCALE_SYSTEM_DEFAULT Then
    Lcid = GetSystemDefaultLcidEx()
  Else
    Lcid = GetUserDefaultLcidEx()
  End If
  retval = GetLocaleInfoEx(Lcid, Linfo, vbNullString, 0)
  If retval = 0 Then Err.Raise ERR_GETVALUE, "GetLocaleValue", "Error retrieving locale value size Linfo (0x" & Right$("0000" & Hex(Linfo), 4) & ") LastError (" & Err.LastDllError & ")"
  buffer = String$(retval, vbNullChar)
  retval = GetLocaleInfoEx(Lcid, Linfo, buffer, retval)
  buffer = RTrimChar(buffer, vbNullChar)
  If retval = 0 Then Err.Raise ERR_GETVALUE, "GetLocaleValue", "Error retrieving locale value 0x" & Right$("0000" & Hex$(Linfo), 4) & " LastError (" & Err.LastDllError & ")"
  GetLocaleValue = buffer
End Function
