VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Private mPath As String
Private mFile As String
Private mExt As String

Public Description As String
Public MajorVersion As String
Public MinorVersion As String
Public Revision As String
Private m_ScreenText As String
Private mClasses As ObjectList
Implements ISortFunction

Property Let PathAndFile(ByVal NewValue As String)
  Call SplitPath(NewValue, mPath, mFile, mExt)
End Property

Property Get PathAndFile() As String
  PathAndFile = mPath & mFile & mExt
End Property

Private Sub Class_Initialize()
  Set mClasses = New ObjectList
End Sub

Public Sub kill()
  Dim vc As Class
  Dim i As Long
  
  For i = 1 To mClasses.Count
    Set vc = mClasses(i)
    Call vc.kill
  Next i
  mClasses.RemoveAll
End Sub

Private Function GetItemIndex(vIndex As Variant) As Long
  Dim vc As Class
  Dim i As Long
  
  GetItemIndex = -1
  If (VarType(vIndex) = vbLong) Or (VarType(vIndex) = vbInteger) Then
    GetItemIndex = vIndex
  Else
    For i = 1 To mClasses.Count
      Set vc = mClasses.Item(i)
      If StrComp(vc.PathAndFile, vIndex) = 0 Then
        GetItemIndex = i
        Exit For
      End If
    Next i
  End If
End Function

Public Sub Remove(ByVal vIndex As Variant)
  Dim i As Long
  
  i = GetItemIndex(vIndex)
  If i < 0 Then Err.Raise 381, "Remove", "Cannot remove Class by key: " & vIndex
  Call mClasses.Remove(i)
  Call mClasses.Compact
End Sub

Property Get Item(ByVal vIndex As Variant) As Class
  Dim i As Long
  
  i = GetItemIndex(vIndex)
  If i < 0 Then Err.Raise 381, "Item", "Cannot get Class by key: " & vIndex
  Set Item = mClasses.Item(i)
End Property

Property Get Count() As Long
  Count = mClasses.Count
End Property

Public Function Add(ByVal PathAndFile As String, ByVal Name As String, ByVal ForceAdd As Boolean) As Class
  Dim vc As Class
  
  PathAndFile = Trim$(PathAndFile)
  If Len(PathAndFile) = 0 Then Err.Raise ERR_INVALIDNAME, "Add", "Invalid File Name"
  Set vc = New Class
  vc.PathAndFile = PathAndFile
  If vc.Check(ForceAdd) Then
    Call vc.ReadClass
    Set vc.Parent = Me
    vc.Name = Name
    Call mClasses.Add(vc)
  Else
    Set vc = Nothing
  End If
  Set Add = vc
End Function

Public Sub Sort()
  Call mClasses.Sort(Me)
End Sub

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim vc0 As Class, vc1 As Class
  Set vc0 = v0
  Set vc1 = v1
  
  If vc0.GlobalNameSpace And Not vc1.GlobalNameSpace Then
    ISortFunction_CompareItems = -1
  ElseIf Not vc0.GlobalNameSpace And vc1.GlobalNameSpace Then
    ISortFunction_CompareItems = 1
  Else
    ISortFunction_CompareItems = StrComp(vc0.Name, vc1.Name)
  End If
End Function

Private Function ParseVBPClass(sToParse As String, ByVal ClassList As String) As Boolean
  Dim sVals()
  Dim i As Long
  Dim s As String, sPathAndFile As String, sName As String

  On Error GoTo ParseVBPClass_ERR
  i = GetDelimitedValues(sVals, sToParse, , , ";")
  If i <> 2 Then Call Err.Raise(ERR_INVALID_VBP_CLASS, "ParseVBPClass", "Expected 2 parameters in vbp class definition.")
  sName = sVals(1)
  Call SplitPath(sVals(2), s)
  If Len(s) = 0 Then
    sPathAndFile = mPath & sVals(2)
  Else
    sPathAndFile = sVals(2)
  End If
  If Len(ClassList) > 0 Then
    If InList(ClassList, sName) Then Call Add(sPathAndFile, sName, True)
  Else
    Call Add(sPathAndFile, sName, False)
  End If
    
ParseVBPClass_END:
  Exit Function
ParseVBPClass_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ParseVBPClass", "Parse VBP Class", "Error pasing a VBP class line.")
  Resume ParseVBPClass_END
  Resume
End Function

Public Sub ReadProject(ByVal ClassList As String)
  Dim ifr As New TCSFileread
  Dim FileName As String, sRet As String
  Dim s As String
  Dim sClass As Class
  Dim i As Long
  
  On Error GoTo ReadFile_ERR
  If ifr.OpenFile(PathAndFile) Then
    Call frmMain.sts.StartPrg(ifr.LineCount, "Reading project ...", ValueOfMax)
    Do While ifr.GetLine(s)
      If Len(Name) = 0 Then
        If ParseEquals(sRet, s, "Name") Then
          Name = sRet
          Call frmMain.sts.StepCaption("Reading project " & Name)
          GoTo NEXT_LINE
        End If
      End If
      If Len(MajorVersion) = 0 Then If ParseEquals(sRet, s, "MajorVer") Then MajorVersion = sRet: GoTo NEXT_LINE
      If Len(MinorVersion) = 0 Then If ParseEquals(sRet, s, "MinorVer") Then MinorVersion = sRet: GoTo NEXT_LINE
      If Len(Revision) = 0 Then If ParseEquals(sRet, s, "RevisionVer") Then Revision = sRet: GoTo NEXT_LINE
      If Len(Description) = 0 Then If ParseEquals(sRet, s, "VersionFileDescription") Then Description = sRet: GoTo NEXT_LINE
      
      If ParseEquals(sRet, s, "Class") Then Call ParseVBPClass(sRet, ClassList)

NEXT_LINE:
      Call frmMain.sts.Step
    Loop
  Else
    Call SplitPath(PathAndFile, , FileName)
    Me.Name = FileName
    Me.Description = "Could not open project " & PathAndFile
  End If
  
  
ReadFile_END:
  Call frmMain.sts.StopPrg
  Exit Sub
  
ReadFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ReadFile", "Read File", "Error reading VBP file")
  Resume ReadFile_END
End Sub

Public Property Get ScreenText() As String
  
  If Len(m_ScreenText) <> 0 Then
    ScreenText = m_ScreenText
    Exit Property
  End If
  m_ScreenText = Name & " Version " & MajorVersion & "." & MinorVersion & "." & Revision & vbCrLf
  m_ScreenText = m_ScreenText & Description & vbCrLf & vbCrLf
  ScreenText = m_ScreenText
End Property

