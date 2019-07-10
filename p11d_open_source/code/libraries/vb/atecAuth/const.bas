Attribute VB_Name = "Const"
Option Explicit
Public Enum LDAP_ERRORS
  ERR_ENUMERATEPEOPLE = TCSCLIENT_ERROR 'apf change
  ERR_NOSERVERCONTEXT
  ERR_GETUSEROBJECT
End Enum
Public Const S_LDAP_PREFIX As String = "LDAP://"
Public gDBHelp As DBHelper

Public Sub main()
  Set gDBHelp = New DBHelper
  gDBHelp.DatabaseTarget = DB_TARGET_SQLSERVER
End Sub

Public Function IsServerAlive(ByVal ServerName As String) As Long
  Dim Root As IADs
  Dim t0 As Long
  
  On Error GoTo IsServerAlive_Err
  t0 = GetTicks
  Set Root = GetObject(S_LDAP_PREFIX & ServerName)
  IsServerAlive = GetTicks - t0
  Exit Function
  
IsServerAlive_Err:
  IsServerAlive = -1
End Function

Public Sub GetProperty(ByVal PE As PropertyEntry, ByVal LP As LDAPProperty)
  Dim DType As ADSTYPEENUM
  Dim PV As PropertyValue
  Dim P As IADsPropertyValue2
  Dim LPValues As Variant
  Dim PEValues As Variant

  Dim l As Long
  On Error GoTo GetProperty_Err
  PEValues = PE.Values
  If IsArray(PEValues) Then
    If LBound(PEValues) <> UBound(PEValues) Then
      LP.MultiValued = True
      ReDim LPValues(LBound(PEValues) To UBound(PEValues))
    End If
  Else
    'PC - Never reach here: Single values come back as arrays
  End If
  
  DType = PE.ADsType
  Select Case DType
    Case ADSTYPE_DN_STRING
      'cad new
      LP.MultiValued = False
      Set PV = PEValues(LBound(PEValues))
      LPValues = PV.DNString
    Case ADSTYPE_CASE_IGNORE_STRING
      LP.DType = TYPE_STR
      If LP.MultiValued Then
        For l = LBound(PEValues) To UBound(PEValues)
          Set PV = PEValues(l)
          LPValues(l) = GetTypedValue(PV.CaseIgnoreString, LP.DType)
        Next l
      Else
        Set PV = PEValues(LBound(PEValues))
        LPValues = GetTypedValue(PV.CaseIgnoreString, LP.DType)
      End If
    Case ADSTYPE_CASE_EXACT_STRING
      LP.DType = TYPE_STR
      If LP.MultiValued Then
        For l = LBound(PEValues) To UBound(PEValues)
          Set PV = PEValues(l)
          LPValues(l) = GetTypedValue(PV.CaseExactString, LP.DType)
        Next l
      Else
        Set PV = PEValues(LBound(PEValues))
        LPValues = GetTypedValue(PV.CaseExactString, LP.DType)
      End If
    Case ADSTYPE_BOOLEAN
      LP.DType = TYPE_BOOL
      If LP.MultiValued Then
        For l = LBound(PEValues) To UBound(PEValues)
          Set PV = PEValues(l)
          LPValues(l) = GetTypedValue(PV.Boolean, LP.DType)
        Next l
      Else
        Set PV = PEValues(LBound(PEValues))
        LPValues = GetTypedValue(PV.Boolean, LP.DType)
      End If
    Case ADSTYPE_INTEGER
      LP.DType = TYPE_LONG
      If LP.MultiValued Then
        For l = LBound(PEValues) To UBound(PEValues)
          Set PV = PEValues(l)
          LPValues(l) = GetTypedValue(PV.Integer, LP.DType)
        Next l
      Else
        Set PV = PEValues(LBound(PEValues))
        LPValues = GetTypedValue(PV.Integer, LP.DType)
      End If
    Case ADSTYPE_LARGE_INTEGER
      LP.DType = TYPE_LONG
      If LP.MultiValued Then
        For l = LBound(PEValues) To UBound(PEValues)
          Set PV = PEValues(l)
          LPValues(l) = GetTypedValue(PV.LargeInteger, LP.DType)
        Next l
      Else
        Set PV = PEValues(LBound(PEValues))
        LPValues = GetTypedValue(PV.LargeInteger, LP.DType)
      End If
    Case ADSTYPE_NUMERIC_STRING
      LP.DType = TYPE_DOUBLE
      If LP.MultiValued Then
        For l = LBound(PEValues) To UBound(PEValues)
          Set PV = PEValues(l)
          LPValues(l) = GetTypedValue(PV.NumericString, LP.DType)
        Next l
      Else
        Set PV = PEValues(LBound(PEValues))
        LPValues = GetTypedValue(PV.NumericString, LP.DType)
      End If
    Case ADSTYPE_UTC_TIME
      LP.DType = TYPE_DATE
      If LP.MultiValued Then
        For l = LBound(PEValues) To UBound(PEValues)
          Set PV = PEValues(l)
          LPValues(l) = GetTypedValue(PV.UTCTime, LP.DType)
        Next l
      Else
        Set PV = PEValues(LBound(PEValues))
        LPValues = GetTypedValue(PV.UTCTime, LP.DType)
      End If
    Case ADSTYPE_PRINTABLE_STRING
      'cad new
      LP.MultiValued = False
      Set PV = PEValues(LBound(PEValues))
      LPValues = PV.PrintableString
    Case ADSTYPE_PROV_SPECIFIC
      'cad new
      LP.MultiValued = False
      Set P = PEValues(LBound(PEValues))
      LPValues = abatecrt.CopyLPSTRtoString(StrPtr(P.GetObjectProperty(ADSTYPE_PROV_SPECIFIC)))
    Case Else
      LP.MultiValued = False
      LPValues = "(Unknown)"
  End Select
  LP.Values = LPValues
  Exit Sub
  
GetProperty_Err:
  Call LP.ClearValue
  'Err.Raise Err.Number, "atecAuth.GetProperty", Err.Description
  'Resume
End Sub

Public Function InList(ByVal Value As String, ByRef StrValues() As String) As Boolean
  Dim i As Long
  
  On Error GoTo InList_Err
  For i = LBound(StrValues) To UBound(StrValues)
    If StrComp(Value, StrValues(i), vbTextCompare) = 0 Then
      InList = True
      Exit Function
    End If
  Next i
  Exit Function
  
InList_Err:
  Err.Raise Err.Number, "atecAuth.InList", Err.Description
End Function


