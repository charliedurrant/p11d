VERSION 5.00
Begin VB.UserControl QBEGrid 
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   ScaleHeight     =   1950
   ScaleWidth      =   6660
   Begin VB.HScrollBar scrQBE 
      Height          =   240
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1650
      Width           =   6615
   End
   Begin atc2QBE.QBEObj QBE 
      Height          =   1620
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   30
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2858
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Label lblCriteria 
      Caption         =   "Criteria"
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   645
      Width           =   800
   End
   Begin VB.Label lblSort 
      Caption         =   "Sort"
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   330
      Width           =   800
   End
   Begin VB.Label lblField 
      Caption         =   "Field"
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   15
      Width           =   800
   End
End
Attribute VB_Name = "QBEGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum FilterType
  FILTER_RECORD = 0
  FILTER_QUERY
End Enum

Private QBE_LEFT As Long
'Consts
Private Const QBEGRID_HEIGHT As Long = 1900
Private Const QBEGRID_WIDTH As Long = 6400
'Default Property Values:
Private Const m_def_sFilter = ""
Private Const m_def_sSort = ""
'Property Variables:
Private m_rsFilter As Recordset
Private m_qryFilter As QueryDef
Private m_sFilter As String
Private m_sSort As String
Private m_sSQL As String

Private Sub QBE_GotFocus(Index As Integer)
  If Index = QBE.Count - 1 Then
    AddQBE
  ElseIf Index = QBE.Count - 3 Then
    If (QBE(QBE.Count - 2).ListIndex = -1) And (QBE(QBE.Count - 1).ListIndex = -1) Then
      Unload QBE(QBE.Count - 1)
        If (QBE.Count * QBE(0).Width) > Width Then
          scrQBE.Visible = True
          scrQBE.Max = QBE.Count - 3
          scrQBE.Value = QBE.Count - 3
        Else
          scrQBE.Visible = False
        End If
    End If
  Else
    If QBE.Count < 3 Then
      scrQBE.Visible = False
      QBE_LEFT = 0
    End If
  End If
End Sub

Private Sub scrQBE_Change()
  Dim i As Long
  Static OldValue As Long
  
  QBE_LEFT = -(scrQBE.Value - OldValue) * QBE(0).Width
  'If QBE_LEFT > 0 Then QBE_LEFT = 0
  For i = 0 To QBE.Count - 1
    QBE(i).Visible = True
  Next i
  For i = 1 To scrQBE.Value
    QBE(i - 1).Visible = False
  Next i
  MoveQBE
  OldValue = scrQBE.Value
End Sub

Private Sub UserControl_Initialize()
  QBE_LEFT = 0
  QBE(0).Top = 0
  QBE(0).Left = lblField.Width
  scrQBE.Left = 0
  scrQBE.Top = Height - scrQBE.Height
  scrQBE.Width = Width
  scrQBE.Max = 1
  scrQBE.Visible = False
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = QBEGRID_HEIGHT
  UserControl.Width = QBEGRID_WIDTH
  scrQBE.Width = UserControl.Width
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
End Property

Public Property Set rsFilter(NewValue As Object)
  If Not Ambient.UserMode Then Err.Raise 383
  Set m_rsFilter = NewValue

  Call InitFields(FILTER_RECORD)
  Call PropertyChanged("rsFilter")
End Property

Public Property Get rsFilter() As Object
  Set rsFilter = m_rsFilter
End Property

Public Property Get qryFilter() As Object
  Set qryFilter = m_qryFilter
End Property

Public Property Get sFilter() As String
Attribute sFilter.VB_MemberFlags = "400"
  GetFilters
  If LTrim$(m_sFilter) <> "" Then
    sFilter = m_sFilter
  Else
    sFilter = ""
  End If
End Property

Public Property Let sFilter(ByVal New_sFilter As String)
  If Ambient.UserMode = False Then Err.Raise 382
  m_sFilter = New_sFilter
  SetFilters
  PropertyChanged "sFilter"
End Property

Public Property Get sSort() As String
Attribute sSort.VB_MemberFlags = "400"
  GetSorts
  If LTrim$(m_sSort) <> "" Then
    sSort = m_sSort
  Else
    sSort = ""
  End If
End Property

Public Property Let sSort(ByVal New_sSort As String)
  If Ambient.UserMode = False Then Err.Raise 382
  m_sSort = New_sSort
  SetSorts
  PropertyChanged "sSort"
End Property

Public Property Set qryFilter(ByVal New_qryFilter As Object)
  If Ambient.UserMode = False Then Err.Raise 383
  Set m_qryFilter = New_qryFilter
  PropertyChanged "qryFilter"
  Call InitFields(FILTER_QUERY)
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  Set Font = Ambient.Font
  m_sFilter = m_def_sFilter
  m_sSort = m_def_sSort
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  Set m_rsFilter = PropBag.ReadProperty("rsFilter", Nothing)
  m_sFilter = PropBag.ReadProperty("sFilter", m_def_sFilter)
  m_sSort = PropBag.ReadProperty("sSort", m_def_sSort)
  Set m_qryFilter = PropBag.ReadProperty("qryFilter", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("Font", Font, Ambient.Font)
  Call PropBag.WriteProperty("rsFilter", m_rsFilter, Nothing)
  Call PropBag.WriteProperty("sFilter", m_sFilter, m_def_sFilter)
  Call PropBag.WriteProperty("sSort", m_sSort, m_def_sSort)
  Call PropBag.WriteProperty("qryFilter", m_qryFilter, Nothing)
End Sub

Private Sub InitFields(FType As FilterType)
  Dim fld As Field
  Dim qbeSet As QBEObj
  
  Set qbeSet = QBE(0)
  If FType = FILTER_QUERY Then
    If Not m_qryFilter Is Nothing Then
      For Each fld In m_qryFilter.Fields
        Call qbeSet.AddItem(fld.Name)
      Next fld
    End If
  ElseIf FType = FILTER_RECORD Then
    If Not m_rsFilter Is Nothing Then
      For Each fld In m_rsFilter.Fields
        Call qbeSet.AddItem(fld.Name)
      Next fld
    End If
  End If
  Set qbeSet = Nothing
  Call AddQBE
End Sub

Private Sub AddQBE()
  Dim qbeSet As QBEObj
  Dim qbeSrc As QBEObj
  Dim i As Integer
  
  Load QBE(QBE.Count)
  Set qbeSet = QBE(QBE.Count - 1)
  Set qbeSrc = QBE(0)
  
  For i = 0 To qbeSrc.ListCount - 1
    qbeSet.AddItem (qbeSrc.List(i))
  Next i
  
  qbeSet.Filter1 = ""
  qbeSet.Filter2 = ""
  qbeSet.Filter3 = ""
  QBE(QBE.Count - 1).Visible = True
  QBE(QBE.Count - 1).Left = QBE(QBE.Count - 2).Left + QBE(QBE.Count - 2).Width
  
  'CheckSize
  If (QBE.Count * QBE(0).Width) > Width Then
    scrQBE.Visible = True
    scrQBE.Max = QBE.Count - 3
    scrQBE.Value = QBE.Count - 3
  Else
    scrQBE.Visible = False
  End If
  Set qbeSet = Nothing
  Set qbeSrc = Nothing
End Sub

Private Function CheckSize()
  If (lblField.Left + lblField.Width + (QBE.Count * QBE(0).Width)) > UserControl.Width Then
    scrQBE.Visible = True
    UserControl.Height = UserControl.Height + scrQBE.Height
  Else
    scrQBE.Visible = False
    UserControl.Height = UserControl.Height - scrQBE.Height
  End If
End Function

Private Sub GetFilters()
  Dim qbeitm As QBEObj
  Dim Filter1 As String
  Dim Filter2 As String
  Dim Filter3 As String
  
  Filter1 = ""
  For Each qbeitm In QBE
    If (qbeitm.ListIndex >= 0) And (qbeitm.Filter1 <> "") Then
      Filter1 = Filter1 & "[" & qbeitm.List(qbeitm.ListIndex) & "] " & qbeitm.Filter1 & " AND "
    End If
  Next qbeitm
  If Filter1 <> "" Then
    Filter1 = Left$(Filter1, Len(Filter1) - 5)
  End If
  
  Filter2 = ""
  For Each qbeitm In QBE
    If (qbeitm.ListIndex >= 0) And (qbeitm.Filter2 <> "") Then
      Filter2 = Filter2 & "[" & qbeitm.List(qbeitm.ListIndex) & "] " & qbeitm.Filter2 & " AND "
    End If
  Next qbeitm
  If Filter2 <> "" Then
    Filter2 = Left$(Filter2, Len(Filter2) - 5)
  End If
  
  Filter3 = ""
  For Each qbeitm In QBE
    If (qbeitm.ListIndex >= 0) And (qbeitm.Filter3 <> "") Then
      Filter3 = Filter3 & "[" & qbeitm.List(qbeitm.ListIndex) & "] " & qbeitm.Filter3 & " AND "
    End If
  Next qbeitm
  If Filter3 <> "" Then
    Filter3 = Left$(Filter3, Len(Filter3) - 5)
  End If
  
  If (Filter1 <> "" And Filter2 <> "") Or (Filter1 <> "" And Filter3 <> "") Then
    Filter1 = "(" & Filter1 & ") OR "
  End If
  
  If (Filter2 <> "" And Filter1 <> "" And Filter3 = "") Then
    Filter2 = "(" & Filter2 & ")"
  ElseIf (Filter2 <> "" And Filter3 <> "") Then
    Filter2 = "(" & Filter2 & ") OR"
  End If
  
  If (Filter3 <> "" And Filter1 <> "") Or (Filter3 <> "" And Filter2 <> "") Then
    Filter3 = "(" & Filter3 & ")"
  End If
  m_sFilter = Filter1 & " " & Filter2 & " " & Filter3
End Sub

Private Sub GetSorts()
  Dim qbeitm As QBEObj
  
  m_sSort = ""
  For Each qbeitm In QBE
    If qbeitm.ListIndex >= 0 Then
      If qbeitm.Sort > 0 Then
        m_sSort = m_sSort & "[" & qbeitm.List(qbeitm.ListIndex) & "] " & IIf(qbeitm.Sort = 1, "ASC", "DESC") & " , "
      End If
    End If
  Next qbeitm
  If m_sSort <> "" Then
    m_sSort = Left$(m_sSort, Len(m_sSort) - 3)
  End If
End Sub

Private Sub MoveQBE()
  Dim i As Long
  
  For i = 0 To QBE.Count - 1
    QBE(i).Left = QBE(i).Left + QBE_LEFT
  Next i
End Sub

Private Sub SetSorts()
  Dim i As Integer
  Dim j As Integer
  
  If m_sSort <> "" Then
    j = 0
    Do
      i = InStr(Right$(m_sSort, Len(m_sSort) - j), ",")
      If i > 0 Then
        QBE(QBE.Count - 2).ListIndex = GetFields(Mid$(m_sSort, j + 1, i - 1))
        QBE(QBE.Count - 2).Sort = GetSort(Mid$(m_sSort, j + 1, i - 1))
        AddQBE
      Else
        QBE(QBE.Count - 2).ListIndex = GetFields(Mid$(m_sSort, j + 1, Len(m_sSort)))
        QBE(QBE.Count - 2).Sort = GetSort(Mid$(m_sSort, j + 1, Len(m_sSort)))
      End If
      j = j + i
    Loop Until i = 0
  End If
End Sub

Private Function GetFields(sField As String) As Integer
  Dim i As Integer
  
  For i = 0 To QBE(QBE.Count - 2).ListCount - 1
    If InStr(sField, QBE(QBE.Count - 2).List(i)) Then
      GetFields = i
      Exit Function
    End If
  Next i
End Function

Private Function GetSort(sSort As String) As Integer
  If InStr(sSort, "ASC") Then
    GetSort = 1
  ElseIf InStr(sSort, "DESC") Then
    GetSort = 2
  Else
    GetSort = 0
  End If
End Function

Private Sub SetFilters()
  Dim sOR() As String
  Dim i As Integer
  Dim k As Integer
  Dim j As Integer
  Dim qbeitm As QBEObj
  
  If m_sFilter <> "" Then
    'Split into OR's - Tick!
    j = 1
    k = 0
    Do
      ReDim Preserve sOR(k)
      i = InStr(j, m_sFilter, "OR")
      If i > 0 Then
        sOR(UBound(sOR)) = Mid$(m_sFilter, j, i - 1)
      Else
        sOR(UBound(sOR)) = Mid$(m_sFilter, j, Len(m_sFilter) - i)
      End If
      j = i + 3
      k = k + 1
    Loop Until i = 0
    'Split into AND's - In progress
    For k = 0 To UBound(sOR)
      j = 1
      Do
        i = InStr(j, sOR(k), "AND")
        If i > 0 Then
          Set qbeitm = GetQBE(Mid$(sOR(k), j, i - j - 1), k)
        Else
          Set qbeitm = GetQBE(Mid$(sOR(k), j, Len(sOR(k)) - j + 1), k)
        End If
        j = i + 3
      Loop Until i = 0
    Next k
  End If
End Sub

Private Function GetQBE(txt As String, iFilterIndex As Integer) As QBEObj
  Dim qbeitm As QBEObj
  Dim tmp As String
  Dim i As Long
  Dim j As Long
  
  i = InStr(txt, "[")
  j = InStr(txt, "]")
  tmp = Mid$(txt, i + 1, j - i - 1)
  For Each qbeitm In QBE
    If InStr(qbeitm.List(qbeitm.ListIndex), "]") Then
      Set GetQBE = qbeitm
      Exit For
    End If
  Next qbeitm
  If GetQBE Is Nothing Then
    Set GetQBE = QBE(QBE.Count - 1)
    AddQBE
    GetQBE.ListIndex = GetFields(txt)
  End If
  Select Case iFilterIndex + 1
    
    Case 1
      GetQBE.Filter1 = GetCriteria(txt)
    Case 2
      GetQBE.Filter2 = GetCriteria(txt)
    Case 3
      GetQBE.Filter3 = GetCriteria(txt)
  End Select
End Function

Private Function GetCriteria(txt As String) As String
  GetCriteria = Trim$(Mid$(txt, InStr(1, txt, "]") + 1))
  If Mid$(GetCriteria, Len(GetCriteria)) = ")" Then
    GetCriteria = Mid$(GetCriteria, 1, Len(GetCriteria) - 1)
  End If
End Function

Public Sub Reset()
  Dim i As Integer
  
  For i = 2 To QBE.Count - 1
    Unload QBE(i)
  Next i
  For i = 0 To 1
    QBE(i).ListIndex = -1
    QBE(i).Sort = 0
    QBE(i).Filter1 = ""
    QBE(i).Filter2 = ""
    QBE(i).Filter3 = ""
  Next i
End Sub

Public Property Get SQL() As String
  'Not Implemented
End Property

Public Property Let SQL(ByVal sNewValue As String)
  'Not Implemented
End Property
