VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LDAPHelper2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public S_USERNAME_DOMAIN As String
Public S_EMPLOYEEID_DOMAIN As String
Public S_AUTH_USER As String
Public S_AUTH_USER_SAM As String
Public S_AUTH_PASSWORD As String
'Private Const S_USERNAME_DOMAIN As String = "UK\"
'Private Const S_EMPLOYEEID_DOMAIN As String = "UK"
'Private Const S_AUTH_USER As String = "CN=Options Test,CN=Users,DC=uk,DC=deloitte,DC=com"
'Private Const S_AUTH_USER_SAM As String = "optionstest"
'Private Const S_AUTH_PASSWORD As String = "0pt10n5t3sT"

'Servers uklondc004.uk.deloitte.com;uklondc003.uk.deloitte.com;uklondc002.uk.deloitte.com;uklondc001.uk.deloitte.com
Private m_ServerList As String
Private m_ServerContext As String
Private m_NextServer As Long

Private m_ServerRoot As String
Private m_Server As String

Public Property Get ServerRoot() As String
  ServerRoot = m_ServerRoot
End Property

Private Sub ResetServerContext()
  Dim ServerCount As Long
  Dim Servers As Variant
  Dim TryServer As String
  Dim i As Long
  Dim t As Long

  On Error GoTo ResetServerContext_Err
  m_ServerRoot = ""
  m_Server = ""
  ServerCount = GetDelimitedValues(Servers, m_ServerList, , , ";")
  For i = m_NextServer To ServerCount
    TryServer = Servers(i)
    t = IsServerAlive(TryServer)
    If t > 0 Then
      m_ServerRoot = TryServer
      m_Server = m_ServerRoot & "/" & m_ServerContext
      m_NextServer = i + 1
      Exit Sub
    End If
  Next i
  For i = 1 To (m_NextServer - 1)
    TryServer = Servers(i)
    t = IsServerAlive(TryServer)
    If t > 0 Then
      m_ServerRoot = TryServer
      m_Server = m_ServerRoot & "/" & m_ServerContext
      m_NextServer = i + 1
      Exit Sub
    End If
  Next i
  Exit Sub

ResetServerContext_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper2.ResetServerContext"), Err.Description
  Resume
End Sub

Public Sub SetServerContext(ByVal ServerList As String, ByVal ServerContext As String)
  On Error GoTo ServerContext_Err
  m_ServerList = ServerList
  m_ServerContext = ServerContext
  m_NextServer = 1
  Call ResetServerContext
  Exit Sub

ServerContext_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper2.ServerContext"), Err.Description
  Resume
End Sub

Public Function UsernameNoDomain(ByVal username As String) As String
  If InStr(username, S_USERNAME_DOMAIN) = 1 Then
    UsernameNoDomain = Right$(username, Len(username) - Len(S_USERNAME_DOMAIN))
  Else
    UsernameNoDomain = username
  End If
End Function

Public Function EmployeeIDNoDomain(ByVal EmployeeID As String) As String
  If InStr(EmployeeID, S_EMPLOYEEID_DOMAIN) = 1 Then
    EmployeeIDNoDomain = Right$(EmployeeID, Len(EmployeeID) - Len(S_EMPLOYEEID_DOMAIN))
  Else
    EmployeeIDNoDomain = EmployeeID
  End If
End Function

Private Property Get LDAPAllUsersSQLFilter(ByVal Container As String, ByVal Fields As String, Optional ByVal fullyQualifiedName As Boolean = False, Optional ByVal Filter As String)
  LDAPAllUsersSQLFilter = "<" & S_LDAP_PREFIX & Me.ServerRoot & "/" & Container & ">;(&(objectCategory=person)(objectclass=user)" & Filter & "); " & Fields & ";"
  If fullyQualifiedName Then
    LDAPAllUsersSQLFilter = LDAPAllUsersSQLFilter & "base"
  Else
    LDAPAllUsersSQLFilter = LDAPAllUsersSQLFilter & "subtree"
  End If
End Property

Private Property Get LDAPSQLFilter(ByVal UsernameNoDomain As String, ByVal Fields As String)
  LDAPSQLFilter = "<" & S_LDAP_PREFIX & m_Server & ">;(&(objectCategory=person)(objectclass=user)(|(userPrincipalName=" + UsernameNoDomain + ")(samAccountName=" + UsernameNoDomain + "))); " & Fields & ";" & "subtree"
End Property

Public Function Authenticate(ByVal UsernameNoDomain As String, ByVal Password As String, Optional ByVal RaiseErrors As Boolean, Optional ByRef UserPath As String, Optional ByRef EmpRef As String)
  Dim rs As Recordset
      
  On Error GoTo ERR_ERR
  If Len(Me.ServerRoot) > 0 Then
    Set rs = UserProperties(UsernameNoDomain, "adsPath", "employeeID")
    If Not rs Is Nothing Then
      UserPath = Replace(rs.Fields("adsPath"), S_LDAP_PREFIX & Me.ServerRoot & "/", "", 1, -1, vbTextCompare)
      Set rs = UserProperties2(UsernameNoDomain, UserPath, Password, "employeeID")
      EmpRef = rs.Fields("employeeID").Value
      Authenticate = True
    End If
  End If
  Exit Function
  
ERR_ERR:
  Authenticate = False
  If RaiseErrors Then Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper2.Authenticate"), Err.Description
End Function

Private Function UserPropertiesEx(ByVal UsernameNoDomain As String, ByVal AuthUser As String, ByVal AuthPassword As String, Fields() As String) As Recordset
  Dim con As Connection
  Dim cmd As Command
  Dim rs As Recordset
  Dim sSQL As String, sFields As String
  Dim i As Long
  
  On Error GoTo ERR_ERR
  Set con = New Connection
  con.Provider = "ADsDSOObject"
  con.Properties("User ID") = AuthUser
  con.Properties("Password") = AuthPassword
  con.Properties("Encrypt Password") = False
  Call con.Open("ADs Provider")
  For i = LBound(Fields) To UBound(Fields)
    sFields = sFields & Fields(i)
    If i < UBound(Fields) Then
       sFields = sFields & ","
    End If
  Next
  Set cmd = New Command
  Set cmd.ActiveConnection = con
  cmd.CommandText = LDAPSQLFilter(UsernameNoDomain, sFields)
  Set rs = cmd.Execute()
  If Not (rs.BOF And rs.EOF) Then Set UserPropertiesEx = rs
  Exit Function

ERR_ERR:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper2.UserPropertiesEx"), Err.Description
  Resume
End Function

Public Function UserProperties(ByVal UsernameNoDomain As String, ParamArray Properties()) As Recordset
  Dim s() As String
  Dim i As Long

  On Error GoTo ERR_ERR
  ReDim s(LBound(Properties) To UBound(Properties))
  For i = LBound(Properties) To UBound(Properties)
    s(i) = Properties(i)
  Next
  Set UserProperties = UserPropertiesEx(UsernameNoDomain, S_AUTH_USER, S_AUTH_PASSWORD, s)
  Exit Function

ERR_ERR:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper2.UserProperties"), Err.Description
  Resume
End Function

Private Function UserProperties2(ByVal UsernameNoDomain As String, ByVal AuthUser As String, ByVal AuthPassword As String, ParamArray Properties()) As Recordset
  Dim s() As String
  Dim i As Long

  On Error GoTo ERR_ERR
  ReDim s(LBound(Properties) To UBound(Properties))
  For i = LBound(Properties) To UBound(Properties)
    s(i) = Properties(i)
  Next
  Set UserProperties2 = UserPropertiesEx(UsernameNoDomain, AuthUser, AuthPassword, s)
  Exit Function

ERR_ERR:
  Err.Raise Err.Number, ErrorSource(Err, "UserProperties2"), Err.Description
  Resume
End Function

Private Function GetUserObject(ByVal userid As String) As IADs
  Dim rs As Recordset
    
  Set rs = UserProperties(userid, "adsPath")
  If rs.EOF And rs.BOF Then Err.Raise ERR_GETUSEROBJECT, "GetUserObject", "Unable to retrieve adsPath for user [" & userid & "]"
  Set GetUserObject = GetAuthenticatedObject(rs.Fields(0).Value)
End Function

Private Function GetUserObjectDirect(ByVal distinguishedName As String) As IADs
  Set GetUserObjectDirect = GetAuthenticatedObject("LDAP://" & Me.ServerRoot & "/" & distinguishedName)
End Function

Private Function GetAuthenticatedObject(ByVal ldapObject As String) As IADs
  Dim np As IADsOpenDSObject
  
  Set np = GetObject("LDAP:")
  ' apf require power user rights to make this call !!
  Set GetAuthenticatedObject = np.OpenDSObject(ldapObject, S_AUTH_USER_SAM, S_AUTH_PASSWORD, ADS_SECURE_AUTHENTICATION + ADS_READONLY_SERVER + ADS_FAST_BIND)     ' ADS_SECURE_AUTHENTICATION
End Function

Public Sub EnumeratePeople(ByVal iNotify As ILDAPNotify, Optional ByVal sContainer As String, Optional ByVal sFilter As String)
  Dim user As IADs, pCount As Long
  Dim cn As Connection, cmd As Command, rs As Recordset, sql As String
    
  On Error GoTo EnumeratePeople_err
  If Len(m_Server) = 0 Then Err.Raise ERR_NOSERVERCONTEXT, "EnumeratePeople", "Invalid server context"
  Set cn = New Connection
  cn.Provider = "ADsDSOObject" 'this is the ADSI-OLEDB provider name
  cn.Properties("User ID") = S_AUTH_USER
  cn.Properties("Password") = S_AUTH_PASSWORD
  cn.Properties("Encrypt Password") = False
  Call cn.Open("ADs Provider")
  Set cmd = New Command
  Set cmd.ActiveConnection = cn
  cmd.CommandText = LDAPAllUsersSQLFilter(sContainer, "adsPath", , sFilter)
  cmd.Properties("Page Size") = 128
  Set rs = cmd.Execute()
  Do While Not rs.EOF
    Set user = GetAuthenticatedObject(rs.Fields("adsPath"))
    If Not iNotify.LDAPNotify(user) Then Exit Do
    Set user = Nothing
    pCount = pCount + 1
    rs.MoveNext
  Loop
  Set rs = Nothing
  'If pCount = 2000 Then Err.Raise ERR_ENUMERATEPEOPLE, "LDAPHelper.EnumeratePeople", "Could not enumerate all people in the container please use filters"
  Exit Sub
  
EnumeratePeople_err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper.EnumeratePeople"), Err.Description
  Resume
End Sub

Private Function GetPropListItem(ByVal PropList As IADsPropertyList, ByVal index As Long) As PropertyEntry
  On Error Resume Next
  Set GetPropListItem = PropList.Item(index)
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
    Set PropEntry = GetPropListItem(PropList, i)
    If PropEntry Is Nothing Then Debug.Print "PropEntry not valid: " & i
    If Not PropEntry Is Nothing Then
      Set Prop = New LDAPProperty
      Prop.Name = PropEntry.Name
      Debug.Print Prop.Name
      If Not bPropertyList Then
        Call GetProperty(PropEntry, Prop)
        Call GetAllObjectPropertiesEx.Add(Prop, Prop.Name)
      Else
        If InList(Prop.Name, PropertyNames) Then
          Call GetProperty(PropEntry, Prop)
          Call GetAllObjectPropertiesEx.Add(Prop, Prop.Name)
        End If
      End If
    End If
  Next i
  Exit Function
  
GetAllObjectProperties_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper2.GetAllObjectPropertiesEx"), Err.Description
  Resume
End Function

Public Function GetAllProperties2(ByVal user As IADs, ParamArray PropertyNames() As Variant) As Collection
  Dim i As Long
  Dim pCount As Long, props() As String
  
  On Error GoTo GetAllProperties2_Err
  pCount = UBound(PropertyNames) - LBound(PropertyNames)
  If pCount >= 0 Then
    ReDim props(0 To pCount)
    For i = 0 To pCount
      props(i) = PropertyNames(i)
    Next i
  End If
  Set GetAllProperties2 = GetAllObjectPropertiesEx(user, props)
  Exit Function
  
GetAllProperties2_Err:
  Err.Raise Err.Number, ErrorSource(Err, "LDAPHelper2.GetAllProperties2"), Err.Description
  Resume
End Function

Public Function GetAllProperties(ByVal username As String, ParamArray PropertyNames() As Variant) As Collection
  Dim user As IADs, i As Long
  Dim pCount As Long, props() As String
  
  On Error GoTo GetAllProperties_Err
  pCount = UBound(PropertyNames) - LBound(PropertyNames)
  Set user = GetUserObject(username)
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
  Resume
End Function

