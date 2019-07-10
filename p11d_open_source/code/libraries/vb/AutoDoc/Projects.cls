VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Projects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISortFunction
Private mProjects As ObjectList
Public AllFunctions As Collection

Private Sub Class_Initialize()
  Set mProjects = New ObjectList
  Set AllFunctions = New Collection
End Sub

Private Sub Class_Terminate()
  Set mProjects = Nothing
End Sub

Public Sub kill()
  Dim vp As Project
  Dim i As Long
  
  For i = 1 To mProjects.Count
    Set vp = mProjects(i)
    Call vp.kill
  Next i
  mProjects.RemoveAll
  Set AllFunctions = New Collection
End Sub

Private Function GetItemIndex(vIndex As Variant) As Long
  Dim vp As Project
  Dim i As Long
  
  GetItemIndex = -1
  If (VarType(vIndex) = vbLong) Or (VarType(vIndex) = vbInteger) Then
    GetItemIndex = vIndex
  Else
    For i = 1 To mProjects.Count
      Set vp = mProjects.Item(i)
      If StrComp(vp.PathAndFile, vIndex) = 0 Then
        GetItemIndex = i
        Exit For
      End If
    Next i
  End If
End Function

Public Sub Remove(ByVal vIndex As Variant)
  Dim i As Long
  
  i = GetItemIndex(vIndex)
  If i < 0 Then Err.Raise 381, "Remove", "Cannot remove Project by key: " & vIndex
  Call mProjects.Remove(i)
  Call mProjects.Compact
End Sub

Property Get Item(ByVal vIndex As Variant) As Project
  Dim i As Long
  
  i = GetItemIndex(vIndex)
  If i < 0 Then Err.Raise 381, "Item", "Cannot find Project by key: " & vIndex
  Set Item = mProjects.Item(i)
End Property

Property Get Count() As Long
  Count = mProjects.Count
End Property

Public Function Add(ByVal PathAndFile As String, ByVal ClassList As String) As Project
  Dim vp As Project
  
  PathAndFile = Trim$(PathAndFile)
  If Len(PathAndFile) = 0 Then Err.Raise ERR_INVALIDNAME, "Add", "Invalid Name"
  Set vp = New Project
  vp.PathAndFile = PathAndFile
  Call vp.ReadProject(ClassList)
  Call mProjects.Add(vp)
  Set Add = vp
End Function

Public Sub Sort()
  Call mProjects.Sort(Me)
End Sub

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  Dim vp0 As Project, vp1 As Project
  Set vp0 = v0
  Set vp1 = v1
  
  ISortFunction_CompareItems = StrComp(vp0.Name, vp1.Name)
End Function
