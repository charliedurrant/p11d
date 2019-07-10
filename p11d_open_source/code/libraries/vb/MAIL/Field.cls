VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum FIELD_DATEFORMAT
  FIELD_NUMERICAL = 0
End Enum

Public Name As String
Public DataSet As ReportDataSet
Public DataSetString As String ' fully qualified data set path (only on load)

Public Selected As Boolean
Public IndexUDM As Long


Public Prefix As String
Public Suffix As String
Public Width As Double
Public BooleanTrue As String
Public BooleanFalse As String
Public Alignment As ALIGNMENT_TYPE
Public DataType As DATABASE_FIELD_TYPES

Public Sort As SORT_TYPE
Public Hide As Boolean
Public SumType As SUM_TYPE
Public HeadingFont As FontDetails
Public DataFont As FontDetails
Public Order As Long
Public TextWidth As Long

Public Criteria As ObjectList
Public Group As Boolean
Public GroupCount  As Boolean
Public GroupHeader As Boolean
Public GroupTotal As Boolean
Public GroupPageBreak As Boolean

Public Format As String

Public FixedWidth As Long
Public Wrap As Boolean
Public NoSquash As Boolean

Private m_Description As String
Private m_FieldName As String

#If AbacusReporter Then
  Private m_Tag As Variant
  Friend Property Get Tag() As Variant
    Tag = m_Tag
  End Property
  Friend Property Let Tag(ByVal NewValue As Variant)
    m_Tag = NewValue
  End Property
#End If

Private Sub Class_Initialize()
  Set Criteria = New ObjectList
  Set HeadingFont = New FontDetails
  Set DataFont = New FontDetails
  Call Reinitialise
End Sub
Public Property Get KeyString() As String
  KeyString = UCase$(DataSetString & DATASET_SEP & Name)
End Property

' don't save non set values msphd
Public Function SaveDetails() As String
  Dim i As Long, s As String, s2 As String
  Dim Crit As Criterion
  
  s = s & "(DataSetName=" & AddEscapeChars(DataSet.CurrentDataSetString) & ")"
  s = s & "(Name=" & AddEscapeChars(Name) & ")"
  s = s & "(DataType=" & DataType & ")"
  s = s & "(FieldName=" & AddEscapeChars(FieldName) & ")"
  If Len(Prefix) > 0 Then s = s & "(Prefix=" & AddEscapeChars(Prefix) & ")"
  If Len(Suffix) > 0 Then s = s & "(Suffix=" & AddEscapeChars(Suffix) & ")"
  's = s & "(Width=" & Width & ")"
  's = s & "(FixedWidth=" & FixedWidth & ")"
  If DataType = TYPE_BOOL Then
    s = s & "(BooleanTrue=" & BooleanTrue & ")"
    s = s & "(BooleanFalse=" & BooleanFalse & ")"
  End If
  s = s & "(Alignment=" & Alignment & ")"
  If Sort <> SORT_NONE Then s = s & "(Sort=" & Sort & ")"
  If Hide Then s = s & "(Hide)"
  If NoSquash Then s = s & "(NoSquash)"
  If Wrap Then s = s & "(Wrap)"
  s = s & HeadingFont.SaveString("HeadingFont")
  s = s & DataFont.SaveString("DataFont")
  If (DataType = TYPE_STR) And (TextWidth > 0) Then s = s & "(Truncate=" & TextWidth & ")"
  If Len(Format) > 0 Then s = s & "(Format=" & AddEscapeChars(Format) & ")"
  If SumType <> TYPE_NOSUM Then s = s & "(SUMTYPE=" & SumType & ")"
  If Group Then
    s2 = ""
    If GroupHeader Then s2 = s2 & "H"
    If GroupTotal Then s2 = s2 & "T"
    If GroupCount Then s2 = s2 & "C"
    If GroupPageBreak Then s2 = s2 & "P"
    If Len(s2) > 0 Then
      s = s & "(GROUP=" & s2 & ")"
    Else
      s = s & "(GROUP)"
    End If
  End If
  For i = 1 To Criteria.Count
    Set Crit = Criteria(i)
    If Not Crit Is Nothing Then s = s & Crit.SaveString(i)
  Next i
  SaveDetails = s
End Function

Public Sub Reinitialise()
  'DataType = FIELD_DATATYPE_STRING
  'FieldName = Name
  FieldName = ""
  Hide = False
  'Set HeadingFont.Parent = Me
  Call HeadingFont.Reinitialise
  'Set DataFont.Parent = Me
  Call DataFont.Reinitialise
  Width = 0
  TextWidth = -1
  BooleanTrue = "True"
  BooleanFalse = "False"
  SumType = TYPE_NOSUM
  Group = False
  GroupCount = False
  GroupHeader = False
  GroupPageBreak = False
  GroupTotal = False
  Format = ""
  Criteria.RemoveAll
  SumType = TYPE_NOSUM
End Sub

' Set all the aCol Properties as appropriate
Public Sub SetAutoCol(aCol As AutoCol, ReportDataFontName As String, ReportDataFontSize As Long, HideGroupHeaderTypes As Boolean)
  Dim s As String
  
  If HideGroupHeaderTypes And aCol.bGroupHeader Then
    aCol.PrintCaption = ""
  Else
    'RK Change to allow Description property to flow through to Rep
    'RK: Previously aCol.PrintCaption = ReplaceChar(FieldName, "~", vbCrLf)
    'Check for Definer or User override of FieldName
    If Len(FieldName) = 0 Or (StrComp(FieldName, Description, vbTextCompare) = 0) Then
      s = Description
    Else
      s = FieldName
    End If
    aCol.PrintCaption = ReplaceChar(s, "~", vbCrLf)
  End If
  aCol.Alignment = Alignment
  aCol.SortType = Sort
  aCol.TrimField = Not NoSquash
  aCol.bGroup = Group
  aCol.bGroupCount = Group And GroupCount
  aCol.bGroupHeader = Group And GroupHeader
  aCol.bGroupPageBreak = Group And GroupPageBreak
  aCol.bGroupTotal = Group And GroupTotal
  If TextWidth > 0 Then aCol.Width = TextWidth
  aCol.Hide = Hide
  If IsNumberField(DataType) Then aCol.SumType = SumType
  aCol.Format = Format
  aCol.CaptionFormat = HeadingFont.FontStringRestricted(ReportDataFontName, ReportDataFontSize)
  aCol.DataFormat = DataFont.FontStringRestricted(ReportDataFontName, ReportDataFontSize)
  aCol.Wrap = Wrap
End Sub

Public Function Copy(rFldTo As ReportField, ByVal CopyUDMIndex As Boolean) As Boolean
  Dim i As Long
  Dim Crit As Criterion, CritTo As Criterion
  
  rFldTo.Alignment = Alignment
  rFldTo.BooleanFalse = BooleanFalse
  rFldTo.BooleanTrue = BooleanTrue
  rFldTo.Criteria.RemoveAll
  For i = 1 To Criteria.Count
    Set Crit = Criteria(i)
    Set CritTo = New Criterion
    Call rFldTo.Criteria.Add(CritTo)
    If Not Crit Is Nothing Then
      Call Crit.CopyTo(CritTo)
    End If
  Next i
  For i = 1 To rFldTo.Criteria.Count
    If rFldTo.Criteria(i).Comparison = 0 Then
      Call rFldTo.Criteria.Remove(i)
    End If
  Next i
  Call DataFont.CopyTo(rFldTo.DataFont) '
  'Set rFldTo.DataSet = DataSet '
  rFldTo.DataSetString = DataSetString
  'rFldTo.DataType = DataType
  rFldTo.FieldName = FieldName
  rFldTo.FixedWidth = FixedWidth
  rFldTo.Format = Format
  rFldTo.Group = Group
  rFldTo.GroupCount = GroupCount
  rFldTo.GroupHeader = GroupHeader
  rFldTo.GroupPageBreak = GroupPageBreak
  rFldTo.GroupTotal = GroupTotal
  Call HeadingFont.CopyTo(rFldTo.HeadingFont) '
  rFldTo.Hide = Hide
  If CopyUDMIndex Then rFldTo.IndexUDM = IndexUDM
  ' rFldTo.KeyString
  rFldTo.Name = Name
  rFldTo.NoSquash = NoSquash
  rFldTo.Wrap = Wrap
  'rFldTo.Order = Order
  rFldTo.Prefix = Prefix
  rFldTo.Selected = Selected '
  rFldTo.Sort = Sort
  rFldTo.Suffix = Suffix
  rFldTo.SumType = SumType
  rFldTo.TextWidth = TextWidth
  rFldTo.Width = Width
    
  
End Function


'cad p11d only a friend as for some reason this is public non crateable class
'cad p11d, maps to ReportField description, only used in display in tree etc
Friend Property Get Description() As String
  If (Len(m_Description) = 0) Then
    Description = Name
  Else
    Description = m_Description
  End If
End Property
Friend Property Let Description(ByVal NewValue As String)
  If NewValue = "Group 1" Then
    Dim i As Long
    i = 0
  End If
  
  m_Description = NewValue
End Property

Public Property Get FieldName() As String
  FieldName = m_FieldName
End Property
Public Property Let FieldName(ByVal NewValue As String)
  NewValue = Trim$(NewValue)
  If Len(NewValue) > 0 And StrComp(NewValue, Description) <> 0 Then
    m_FieldName = NewValue
  End If
End Property
  
