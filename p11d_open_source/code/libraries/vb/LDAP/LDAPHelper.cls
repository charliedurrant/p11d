VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LDAPHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type LDAPServer
  Name As String
  nRequests As Long
  nSumRequestTime As Long
  bValid As Boolean
  nTotalRequests As Long
  nSwitches As Long
End Type

Private mLDAPServers() As LDAPServer
Private mLDAPServerCount As Long

Public Enum AUTHENTICATION_METHOD
  ADS_STANDARD = 0
  'ADS_SECURE_AUTHENTICATION = &H1
  ADS_USE_ENCRYPTION = &H2
  ADS_USE_SSL = &H2
End Enum
Private Const CONTEXT_STATS_DELIMITER As String = "|"
Private Const CONTEXT_SERVER_DELIMITER As String = ";"

Private Const ERROR_AUTHENTICATION0 As Long = &H80072030
Private Const ERROR_AUTHENTICATION1 As Long = &H8007052E
Private Const ERROR_DEADSERVER As Long = &H8007203A

Private Const SERVER_RANDOM_KEEPALIVE As Long = 100
Private Const SERVER_RUNNING_STATS As Long = 10

Private Sub Class_Initialize()
  ReDim mLDAPServers(0 To 0)
  mLDAPServerCount = 0
End Sub

Public Property Let ServerContext(ByVal NewValue As String)
  Dim server As String, s As String
  Dim p0 As Long, p1 As Long
  Dim q0 As Long, q1 As Long
  
  mLDAPServerCount = 0
  NewValue = Trim$(NewValue)
  If Left$(NewValue, 1) = "<" Then
    p0 = 2
    Do
      p1 = InStr(p0, NewValue, CONTEXT_SERVER_DELIMITER, vbBinaryCompare)
      If p1 = 0 Then p1 = InStr(p0, NewValue, ">", vbBinaryCompare)
      If p1 > 0 Then
        server = Trim$(Mid$(NewValue, p0, p1 - p0))
        'Debug.Print server
        If Len(server) > 0 Then
          ReDim Preserve mLDAPServers(0 To mLDAPServerCount)
          mLDAPServers(mLDAPServerCount).bValid = True
             
          q0 = 1
          q0 = GetDelimitedValue(s, server, q0, CONTEXT_STATS_DELIMITER)
          mLDAPServers(mLDAPServerCount).Name = s
          
          If q0 > Len(server) Then GoTo next_server
          q0 = GetDelimitedValue(s, server, q0, CONTEXT_STATS_DELIMITER)
          mLDAPServers(mLDAPServerCount).nSumRequestTime = CLng(s)
                   
          If q0 > Len(server) Then GoTo next_server
          q0 = GetDelimitedValue(s, server, q0, CONTEXT_STATS_DELIMITER)
          mLDAPServers(mLDAPServerCount).nRequests = CLng(s)
          
          If q0 > Len(server) Then GoTo next_server
          q0 = GetDelimitedValue(s, server, q0, CONTEXT_STATS_DELIMITER)
          mLDAPServers(mLDAPServerCount).nTotalRequests = CLng(s)
          
          If q0 > Len(server) Then GoTo next_server
          q0 = GetDelimitedValue(s, server, q0, CONTEXT_STATS_DELIMITER)
          mLDAPServers(mLDAPServerCount).nSwitches = CLng(s)
          
next_server:
          mLDAPServerCount = mLDAPServerCount + 1
        End If
        p0 = p1 + 1
      End If
    Loop Until p1 = 0
    If mLDAPServerCount = 0 Then Err.Raise 380, "ServerContext", "Invalid Server Context " & NewValue
  Else
    ReDim mLDAPServers(0 To 0)
  End If
End Property

Public Property Get DebugServerContext() As String
  Dim s As String, i As Long
  Dim server As String
  Dim nTotalRequests As Long
  
  s = ""
  For i = 0 To UBound(mLDAPServers)
    server = "Name  " & mLDAPServers(i).Name & vbCrLf & _
             "      Request Count " & CStr(mLDAPServers(i).nTotalRequests) & vbCrLf & _
             "      Switch Count " & CStr(mLDAPServers(i).nSwitches) & vbCrLf & _
             "      Valid " & CStr(mLDAPServers(i).bValid)
    nTotalRequests = nTotalRequests + mLDAPServers(i).nTotalRequests
    s = s & server & vbCrLf
  Next i
  DebugServerContext = s & "Total requests: " & nTotalRequests & vbCrLf
End Property

Public Property Get ServerContext() As String
  Dim sContext As String, i As Long
  Dim server As String
  
  sContext = "<"
  For i = 0 To UBound(mLDAPServers)
    server = mLDAPServers(i).Name & CONTEXT_STATS_DELIMITER & _
             CStr(mLDAPServers(i).nSumRequestTime) & CONTEXT_STATS_DELIMITER & _
             CStr(mLDAPServers(i).nRequests) & CONTEXT_STATS_DELIMITER & _
             CStr(mLDAPServers(i).nTotalRequests) & CONTEXT_STATS_DELIMITER & _
             CStr(mLDAPServers(i).nSwitches) & CONTEXT_STATS_DELIMITER & _
             "Valid=" & CStr(mLDAPServers(i).bValid)
    sContext = sContext & server & CONTEXT_SERVER_DELIMITER
  Next i
  ServerContext = sContext & ">"
End Property

Private Sub SortServers()
  Dim bestserver As Long, lstmp As LDAPServer, i As Long
  Dim avgrequest As Double, minavgrequest As Double
  
  'CONTEXT_SERVER_STATS_N
  bestserver = -1
  If Int((SERVER_RANDOM_KEEPALIVE + 1) * Rnd) = SERVER_RANDOM_KEEPALIVE Then
    bestserver = Int((UBound(mLDAPServers) + 1) * Rnd)
  Else
    For i = 0 To UBound(mLDAPServers)
      If mLDAPServers(i).bValid Then
        If mLDAPServers(i).nRequests > 0 Then
          avgrequest = (CDbl(mLDAPServers(i).nSumRequestTime) / mLDAPServers(i).nRequests)
        Else
          avgrequest = 0
        End If
        If (bestserver < 0) Or (avgrequest < minavgrequest) Then
          bestserver = i
          minavgrequest = avgrequest
        End If
      End If
    Next i
  End If
  If bestserver < 0 Then
    mLDAPServers(0).bValid = False
  ElseIf bestserver > 0 Then
    lstmp = mLDAPServers(0)
    mLDAPServers(0) = mLDAPServers(bestserver)
    mLDAPServers(0).bValid = True
    mLDAPServers(0).nSwitches = mLDAPServers(0).nSwitches + 1
    mLDAPServers(bestserver) = lstmp
  End If
End Sub

Private Sub AddStats(ByVal rTime As Long)
  Dim i As Long
      
  If rTime > 0 Then
    mLDAPServers(0).nRequests = mLDAPServers(0).nRequests + 1
    mLDAPServers(0).nTotalRequests = mLDAPServers(0).nTotalRequests + 1
  End If
  If rTime < 0 Then
    mLDAPServers(0).bValid = False
  Else
    mLDAPServers(0).nSumRequestTime = mLDAPServers(0).nSumRequestTime + rTime
  End If
  If mLDAPServers(0).nRequests > (2 * SERVER_RUNNING_STATS) Then
    mLDAPServers(0).nRequests = mLDAPServers(0).nRequests / 2
    mLDAPServers(0).nSumRequestTime = mLDAPServers(0).nSumRequestTime / 2
  End If
  Call SortServers
End Sub

Public Function IsServerAlive(ByVal ServerName As String) As Long
  Dim Root As IADs
  Dim t0 As Long
  
  On Error GoTo IsServerAlive_Err
  t0 = GetTicks
  Set Root = GetObject("LDAP://" & ServerName)
  IsServerAlive = GetTicks - t0
  Exit Function
  
IsServerAlive_Err:
  IsServerAlive = -1
End Function

Private Function GetUserObject(ByVal userid As String) As IADs
  If Not mLDAPServers(0).bValid Then Err.Raise ERR_NOSERVERCONTEXT, "GetUserObject", "Invalid server context."
  Set GetUserObject = GetObject("LDAP://" & mLDAPServers(0).Name & "/uid=" & userid & ",ou=people,o=arthurandersen.com")
End Function

Public Sub EnumeratePeople(ByVal iNotify As ILDAPNotify, Optional ByVal sGMU As String, Optional ByVal sLMU As String)
  Dim user As IADs, pCount As Long
  Dim cn As Connection, cmd As Command, rs As Recordset, vArray As Variant, sql As String
    
  On Error GoTo EnumeratePeople_err
  If Not mLDAPServers(0).bValid Then Err.Raise ERR_NOSERVERCONTEXT, "EnumeratePeople", "Invalid server context."
  Set cn = New Connection
  cn.Provider = "ADsDSOObject" 'this is the ADSI-OLEDB provider name
  cn.Open "Active Directory Provider"
  Set cmd = New Command
  Set cmd.ActiveConnection = cn
  sql = "select uid from 'LDAP://" & mLDAPServers(0).Name & "/ou=people,o=arthurandersen.com' " & _
        "WHERE objectClass = 'aaPerson'"
  If Len(sGMU) > 0 Then sql = sql & " AND destinationindicator=" & gDBHelp.StrSQL(sGMU)
  If Len(sLMU) > 0 Then sql = sql & " AND ou=" & gDBHelp.StrSQL(sLMU)

  cmd.CommandText = sql
  Set rs = cmd.Execute
  Do While Not rs.EOF
    vArray = rs.Fields("uid")
    Set user = GetUserObject(vArray(0))
    If Not iNotify.LDAPNotify(user) Then Exit Do
    Set user = Nothing
    pCount = pCount + 1
    rs.MoveNext
  Loop
  Set rs = Nothing
  If pCount = 2000 Then Err.Raise ERR_ENUMERATEPEOPLE, "LDAPHelper.EnumeratePeople", "Could not enumerate all people in the container please use filters"
  Exit Sub
  
EnumeratePeople_err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper.EnumeratePeople"), Err.Description
  Resume
End Sub

Public Function Authenticate(ByVal Username As String, ByVal Password As String, Optional AMethod As AUTHENTICATION_METHOD = ADS_STANDARD, Optional RaiseErrors As Boolean = False) As Boolean
  Dim DSO As IADsOpenDSObject
  Dim O As IADs
  Dim t0 As Long
  
  On Error GoTo Authenticate_Err
  If Len(Username) = 0 Or Len(Password) = 0 Then Exit Function
  
Authenticate_retry:
  If Not mLDAPServers(0).bValid Then Err.Raise ERR_NOSERVERCONTEXT, "Authenticate", "Invalid server context"
  t0 = GetTicks
  Set DSO = GetObject("LDAP:")
  Set O = DSO.OpenDSObject("LDAP://" & mLDAPServers(0).Name & "/o=arthurandersen.com", "uid=" & Username & ",ou=people,o=arthurandersen.com", Password, AMethod)
  t0 = GetTicks - t0
  Call AddStats(t0)
  Authenticate = True
  Exit Function

Authenticate_Err:
  'Debug.Print "User: " & Username & " Server: " & mLDAPServers(0).Name & " Error: 0x" & Hex$(Err.Number) & " " & Err.Description
  If Not ((Err.Number = ERROR_AUTHENTICATION0) Or (Err.Number = ERROR_AUTHENTICATION1) Or (Err.Number = ERR_NOSERVERCONTEXT)) Then
    Call AddStats(-1)
    Resume Authenticate_retry
  End If
  If RaiseErrors Then Err.Raise Err.Number, "LDAPHelper.Authenticate", Err.Description
  Authenticate = False
End Function

Private Function GetAllObjectPropertiesEx(ByVal ldapObject As IADs, ByRef PropertyNames() As String) As Collection
  Dim PropList As IADsPropertyList
  Dim PropEntry As PropertyEntry
  Dim Prop As LDAPProperty
  Dim i As Long, bPropertyList As Long
  Dim Values As Variant
  
  On Error GoTo GetAllObjectProperties_Err
  Set GetAllObjectPropertiesEx = New Collection
  bPropertyList = IsArrayEx(PropertyNames)
  Call ldapObject.GetInfo
  Set PropList = ldapObject
  For i = 0 To (PropList.PropertyCount - 1)
    Set PropEntry = PropList.Item(i)
    Set Prop = New LDAPProperty
    Prop.Name = PropEntry.Name
    If Not bPropertyList Then
      Call GetProperty(PropEntry, Prop)
      Call GetAllObjectPropertiesEx.Add(Prop, Prop.Name)
    Else
      If InList(Prop.Name, PropertyNames) Then
        Call GetProperty(PropEntry, Prop)
        Call GetAllObjectPropertiesEx.Add(Prop, Prop.Name)
      End If
    End If
  Next i
  Exit Function
  
GetAllObjectProperties_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper.GetAllObjectPropertiesEx"), Err.Description
  Resume
End Function

Public Function GetAllObjectProperties(ByVal ldapObject As IADs, ParamArray PropertyNames() As Variant) As Collection
  Dim i As Long
  Dim pCount As Long, props() As String
  
  On Error GoTo GetAllObjectProperties_Err
  pCount = UBound(PropertyNames) - LBound(PropertyNames)
  If pCount >= 0 Then
    ReDim props(0 To pCount)
    For i = 0 To pCount
      props(i) = PropertyNames(i)
    Next i
  End If
  Set GetAllObjectProperties = GetAllObjectPropertiesEx(ldapObject, props)
  Exit Function
  
GetAllObjectProperties_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper.GetAllObjectProperties"), Err.Description
End Function

Public Function GetAllProperties(ByVal Username As String, ParamArray PropertyNames() As Variant) As Collection
  Dim user As IADs, i As Long
  Dim pCount As Long, props() As String
  
  On Error GoTo GetAllProperties_Err
  pCount = UBound(PropertyNames) - LBound(PropertyNames)
  Set user = GetUserObject(Username)
  If pCount >= 0 Then
    ReDim props(0 To pCount)
    For i = 0 To pCount
      props(i) = PropertyNames(i)
    Next i
  End If
  Set GetAllProperties = GetAllObjectPropertiesEx(user, props)
  Exit Function
  
GetAllProperties_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper.GetAllProperties"), Err.Description
End Function

Public Function GetLDAPObjectProperty(ByVal user As IADs, ByVal PropertyName As String) As LDAPProperty
  Dim Prop As LDAPProperty
  Dim Values As Variant
  
  On Error GoTo GetLDAPObjectProperty_Err
  'Requesting a property with a zero length string will GPF ADSI - Number is ADSI Property not found
  If Len(PropertyName) = 0 Then Call Err.Raise(-2147463155, "", "Property Name not supplied")
  Values = user.GetEx(PropertyName)
  Set Prop = New LDAPProperty
  Prop.Name = PropertyName
  Prop.DType = -1
  Prop.MultiValued = (LBound(Values) <> UBound(Values))
  If Prop.MultiValued Then
    Prop.Values = Values
  Else
    Prop.Values = Values(LBound(Values))
  End If
  Set GetLDAPObjectProperty = Prop
  Exit Function
  
GetLDAPObjectProperty_Err:
  Err.Raise Err.Number, "LDAPHelper.GetLDAPObjectProperty", Err.Description
End Function

Public Function GetLDAPProperty(ByVal Username As String, ByVal PropertyName As String) As LDAPProperty
  Dim user As IADs
  
  On Error GoTo GetLDAPProperty_Err
  'Requesting a property with a zero length string will GPF ADSI - Number is ADSI Property not found
  Set user = GetUserObject(Username)
  Set GetLDAPProperty = GetLDAPObjectProperty(user, PropertyName)
  Exit Function
  
GetLDAPProperty_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper.GetLDAPProperty"), Err.Description
End Function

Private Sub GetProperty(ByVal PE As PropertyEntry, ByVal LP As LDAPProperty)
  Dim DType As ADSTYPEENUM
  Dim PV As PropertyValue
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
    Case Else
      LP.MultiValued = False
      LPValues = "(Unknown)"
  End Select
  LP.Values = LPValues
  Exit Sub
  
GetProperty_Err:
  Err.Raise Err.Number, "LDAPHelper.GetProperty", Err.Description
End Sub

Private Function InList(ByVal Value As String, ByRef StrValues() As String) As Boolean
  Dim i As Long
  
  On Error GoTo InList_Err
  For i = LBound(StrValues) To UBound(StrValues)
    If Value = StrValues(i) Then
      InList = True
      Exit Function
    End If
  Next i
  Exit Function
  
InList_Err:
  Err.Raise Err.Number, "LDAPHelper.InList", Err.Description
End Function

