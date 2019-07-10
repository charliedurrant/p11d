VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CritColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mCriteria As Collection

Private Sub Class_Initialize()
  Set mCriteria = New Collection
End Sub

Public Function Item(Index As Variant) As Criterion
Attribute Item.VB_UserMemId = 0
  Set Item = mCriteria.Item(Index)
End Function

Public Property Get Count() As Long
  Count = mCriteria.Count
End Property

Public Sub Remove(Index As Variant)
  Call mCriteria.Remove(Index)
End Sub

Public Sub Add(Item As Criterion) ', Key As String)
  Call mCriteria.Add(Item) ', Key)
End Sub

Public Function NewEnum() As IUnknown
  Set NewEnum = mCriteria.[_NewEnum]
End Function

Public Property Get MaxCriterionNumber() As Long
  Dim Crit As Criterion, MaxNum As Long
  MaxNum = 0
  For Each Crit In mCriteria
    If Crit.Number > MaxNum Then MaxNum = Crit.Number
  Next Crit
  MaxCriterionNumber = MaxNum
End Property

Public Sub MeetsCriteria(CritMatches As Variant, ByVal TestValue As Variant)
  Dim Crit As Criterion
  
  For Each Crit In mCriteria
    If Crit.MeetsCriterion(TestValue) Then
      CritMatches(Crit.Number) = False
    End If
  Next Crit
  
End Sub

Private Sub Class_Terminate()
  Set mCriteria = Nothing
End Sub
