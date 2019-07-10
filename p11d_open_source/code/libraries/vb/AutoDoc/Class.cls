VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PathAndFile As String
Private mName As String
Public GlobalNameSpace As Boolean
Public Creatable As Boolean
Public Exposed As Boolean
Public MulitUse As Boolean
Public Description As String
Public Parent As Project
Private mSortByCategory As Boolean
Private m_ScreenText As String

Private mFunctions As ObjectList
Implements ISortFunction

Public Property Get SortByCategory() As Boolean
  SortByCategory = mSortByCategory
End Property

Public Property Let SortByCategory(ByVal NewValue As Boolean)
  If Not GlobalNameSpace Then Err.Raise 380, "SortByCategory", "Cannot have categories in classes which are not global" & vbCrLf & "Class: " & mName
  mSortByCategory = NewValue
End Property

Public Property Get Name() As String
  If GlobalNameSpace Then
    Name = "Global"
  Else
    Name = mName
  End If
End Property

Public Property Let Name(ByVal NewValue As String)
  mName = NewValue
End Property

Private Sub Class_Initialize()
  Set mFunctions = New ObjectList
End Sub

Private Sub Class_Terminate()
  Set mFunctions = Nothing
End Sub

Public Sub kill()
  Dim fi As FunctionItem
  Dim i As Long
  
  Set Parent = Nothing
  For i = 1 To mFunctions.Count
    Set fi = mFunctions(i)
    Call fi.kill
  Next i
End Sub

Private Function GetItemIndex(vIndex As Variant) As Long
  Dim fi As FunctionItem
  Dim i As Long
  
  GetItemIndex = -1
  If (VarType(vIndex) = vbLong) Or (VarType(vIndex) = vbInteger) Then
    GetItemIndex = vIndex
  Else
    For i = 1 To mFunctions.Count
      Set fi = mFunctions.Item(i)
      If StrComp(fi.Name, vIndex) = 0 Then
        GetItemIndex = i
        Exit For
      End If
    Next i
  End If
End Function

Public Sub Remove(ByVal vIndex As Variant)
  Dim i As Long
  
  i = GetItemIndex(vIndex)
  If i < 0 Then Err.Raise 381, "Remove", "Cannot remove Functions by key: " & vIndex
  Call mFunctions.Remove(i)
  Call mFunctions.Compact
End Sub

Property Get Item(ByVal vIndex As Variant) As FunctionItem
  Dim i As Long
  
  i = GetItemIndex(vIndex)
  If i < 0 Then Err.Raise 381, "Item", "Cannot find Functions by key: " & vIndex
  Set Item = mFunctions.Item(i)
End Property

Property Get Count() As Long
  Count = mFunctions.Count
End Property

Public Function Add(sline As String, sToSearch As String, ByVal IsFunction As Boolean, ByVal IsProperty As Boolean) As FunctionItem
  Dim fi As FunctionItem
  
  Set fi = ParseFunction(sline, sToSearch, IsFunction, IsProperty)
  If Not fi Is Nothing Then
    Set fi.Parent = Me
    If Not fi.AddedToCollection Then
      Call mFunctions.Add(fi)
      Call gProjects.AllFunctions.Add(fi)
      fi.AddedToCollection = True
    End If
  End If
  Set Add = fi
End Function

Public Sub Sort()
  Call mFunctions.Sort(Me)
End Sub

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim i As Long
  Dim fi0 As FunctionItem, fi1 As FunctionItem
  Set fi0 = v0
  Set fi1 = v1
  
  If mSortByCategory Then
    i = StrComp(fi0.Category, fi1.Category)
    If i <> 0 Then
      ISortFunction_CompareItems = i
      Exit Function
    End If
  End If
  ISortFunction_CompareItems = StrComp(fi0.Name, fi1.Name)
End Function


Public Function Check(ByVal ForceAdd As Boolean) As Boolean
  Dim s As String, sRet As String
  Dim i As Long
  Dim ifr As TCSFileread
  
  Set ifr = New TCSFileread
  If ifr.OpenFile(PathAndFile) Then
    Do While ifr.GetLine(s)
      If ParseEquals(sRet, s, "VB_GlobalNameSpace ") Then GlobalNameSpace = StringToBool(sRet)
      If ParseEquals(sRet, s, "VB_Creatable ") Then Creatable = StringToBool(sRet)
      If ParseEquals(sRet, s, "VB_Exposed ") Then Exposed = StringToBool(sRet)
      If ParseEquals(sRet, s, "MultiUse ") Then MulitUse = StringToBool(sRet)
      i = i + 1
      If i > 20 Then Exit Do
    Loop
  End If
  Check = (Exposed And Creatable) Or ForceAdd
End Function


Private Function IsFunctionLine(sline As String, sToSearch As String, ByVal IsFunction As Boolean, ByVal IsProperty As Boolean) As FunctionItem
  If InStr(1, sline, sToSearch, vbBinaryCompare) = 1 Then
    Set IsFunctionLine = Add(sline, sToSearch, IsFunction, IsProperty)
  End If
End Function

Private Function FunctionFound(ByVal sline As String, fi As FunctionItem) As Boolean
  On Error GoTo FunctionFound_ERR
  
  If InStr(1, sline, "Private ") > 0 Then Exit Function
  If InStr(1, sline, "Enum ") > 0 Then Exit Function
  Set fi = Nothing
  If fi Is Nothing Then Set fi = IsFunctionLine(sline, S_FUNTION_SEARCH_PUBLIC_FUNCTION, True, False)
  If fi Is Nothing Then Set fi = IsFunctionLine(sline, S_FUNTION_SEARCH_PUBLIC_SUB, False, False)
  If fi Is Nothing Then Set fi = IsFunctionLine(sline, S_FUNTION_SEARCH_BARE_SUB, False, False)
  If fi Is Nothing Then Set fi = IsFunctionLine(sline, S_FUNTION_SEARCH_BARE_FUNCTION, True, False)
  If fi Is Nothing Then Set fi = IsFunctionLine(sline, S_FUNTION_PUBLIC_PROPERTY, False, True)
  If fi Is Nothing Then Set fi = IsFunctionLine(sline, S_FUNTION_PUBLIC_ANY, False, True)
  
   
FunctionFound_END:
  FunctionFound = Not (fi Is Nothing)
  Exit Function
  
FunctionFound_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FunctionFound", "Function Found", "Error finding function.")
  Resume FunctionFound_END
  Resume
End Function

Public Sub ReadClass()
  Dim ifr As New TCSFileread
  Dim p As Long
  Dim sline As String
  Dim fi As FunctionItem
  Dim IsAutoText As Boolean
  Dim LastPos As Long, AutoTextPos As Long
  Dim mode As CLASS_PARSEMODE
  
  On Error GoTo ReadFile_ERR
  Call xSet("ReadFile")
  
  mode = CP_SEARCH_ANY
  If ifr.OpenFile(PathAndFile) Then
    LastPos = ifr.CurrentPos
    AutoTextPos = 0
    Do While ifr.GetLine(sline)
      IsAutoText = False
      sline = Trim$(sline)
      If Len(sline) = 0 Then GoTo nextline
      
      'ignore commented lines
      If InStr(1, sline, S_AUTODOC_STUB, vbTextCompare) = 1 Then sline = Trim$(Mid$(sline, Len(S_AUTODOC_STUB)))
      IsAutoText = InStr(1, sline, S_AUTODOCSTART, vbTextCompare) > 0
      If Not IsAutoText Then
        If (mode = CP_PROCESS_AUTODOC) And (ifr.CurrentPos > LastPos) Then
          mode = CP_SEARCH_ANY
          GoTo nextline
        End If
        If (mode = CP_SEARCH_ANY) Or (mode = CP_SEARCH_FUNCTION) Then   ' check if this is a function
          If FunctionFound(sline, fi) Then
            fi.FileFunctionStartPos = LastPos
            If mode = CP_SEARCH_FUNCTION Then
              mode = CP_PROCESS_AUTODOC
              ifr.CurrentPos = AutoTextPos
              AutoTextPos = -1
            End If
          End If
        End If
      Else
        If mode = CP_PROCESS_AUTODOC Then
          Call ParseAutoDoc(sline, fi)
        ElseIf mode = CP_SEARCH_ANY Then
          ' check if this is a global autodoc setting otherwise ignore
          If InStr(1, sline, S_AUTODOC_CLASSDESCRIPTION, vbTextCompare) = 1 Then
            Me.Description = Me.Description & Trim$(Mid$(sline, Len(S_AUTODOC_CLASSDESCRIPTION) + 1)) & " "
          Else
            AutoTextPos = LastPos
            mode = CP_SEARCH_FUNCTION
          End If
        End If
      End If
      
nextline:
      If mode <> CP_PROCESS_AUTODOC Then LastPos = ifr.CurrentPos
    Loop
  Else
  End If
    
ReadFile_END:
  Call xReturn("ReadFile")
  Exit Sub
  
ReadFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReadFile", "ReadFile", "Error reading class file " & PathAndFile)
  Resume ReadFile_END
  Resume
End Sub

Private Function StringToBool(ByVal s As String) As Boolean
  s = Trim(s)
  
  If StrComp(s, "True") = 0 Then
    StringToBool = True
  ElseIf StrComp(s, "-1  'True") = 0 Then
    StringToBool = True
  End If
End Function


Public Property Get ScreenText() As String
  If Len(m_ScreenText) > 0 Then ScreenText = m_ScreenText: Exit Property
  
  m_ScreenText = "Class: " & Name & vbCrLf & Description & vbCrLf
  If Len(Description) > 0 Then m_ScreenText = m_ScreenText & vbCrLf
  ScreenText = m_ScreenText
End Property



Private Function ParseFunction(ByVal sToParse As String, sFunctionSearchString As String, ByVal IsFunction As Boolean, ByVal IsProperty As Boolean) As FunctionItem
  Dim param As Parameter
  Dim i As Long, j As Long, k As Long
  Dim fi As FunctionItem, fiIndex As Long
  Dim s As String
  Dim sParams() As String
  Dim p0 As Long
  Dim qCount As Long, pType As PROPERTY_TYPE
  
  On Error GoTo ParseFunction_ERR
  sToParse = Trim$(Right$(sToParse, Len(sToParse) - Len(sFunctionSearchString)))
  If IsProperty Then
    If ParseString(sToParse, S_PROPERTY_GET) Then
      pType = PROPERTY_GET
    ElseIf ParseString(sToParse, S_PROPERTY_SET) Then
      pType = PROPERTY_SET
    ElseIf ParseString(sToParse, S_PROPERTY_LET) Then
      pType = PROPERTY_LET
    Else
      pType = PROPERTY_LET + PROPERTY_GET
    End If
  End If
  
  i = InStr(1, sToParse, "(")
  If (i = 0) And IsProperty Then
    i = InStr(1, sToParse, " ")
  End If
  If i = 0 Then Err.Raise ERR_PARSEFUNCTION, "ParseFunction", "Unable to parse function." & vbCrLf & sToParse
  s = Left$(sToParse, i - 1)
  If StrComp(s, "Class_Terminate") = 0 Then GoTo ParseFunction_END
  If StrComp(s, "Class_Initialise") = 0 Then GoTo ParseFunction_END
  If IsProperty Then
    fiIndex = GetItemIndex(Trim$(s))
    If fiIndex >= 0 Then Set fi = mFunctions.Item(fiIndex)
  End If
  If fi Is Nothing Then Set fi = New FunctionItem
  Set ParseFunction = fi
  fi.Name = s
  sToParse = Right$(sToParse, Len(sToParse) - (i - 1))
  i = InStrRev(sToParse, ")")
  fi.IsFunction = IsFunction
  If IsProperty Then fi.PropertyType = fi.PropertyType Or pType
  If IsFunction Or (IsProperty And ((pType And PROPERTY_GET) = PROPERTY_GET)) Then
    'get the return value if a function
     If i < Len(sToParse) Then
      s = Right$(sToParse, Len(sToParse) - (i + 4)) '3 as ) as boolean , take out as also
      fi.ReturnValue = s
     End If
  End If
  
  If (i = 0) And IsProperty Then
    sToParse = "ByVal NewValue " & sToParse
  Else
    sToParse = Mid$(sToParse, 2, i - 2)
  End If
  'now left with params
  If Len(sToParse) > 0 Then
    qCount = 0
    p0 = 1
    For i = 1 To Len(sToParse)
      s = Mid$(sToParse, i, 1)
      If s = """" Then qCount = qCount + 1
      If ((s = ",") Or (i = Len(sToParse))) And ((qCount Mod 2) = 0) Then
        If i = Len(sToParse) Then i = i + 1
        Call ParseParam(Trim$(Mid$(sToParse, p0, i - p0)), param)
        If Not InCollection(fi.Params, param.Name) Then
          Call fi.Params.Add(param, param.Name)
        End If
        p0 = i + 1
      End If
    Next i
  End If
  
ParseFunction_END:
  Exit Function
ParseFunction_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ParseFunction", "Parse Function", "Error parsing a function = " & sToParse)
  Resume ParseFunction_END
  Resume
End Function

