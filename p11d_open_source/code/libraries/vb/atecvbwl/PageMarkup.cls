VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageMarkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cadxx new
Option Explicit

Public Enum PAGE_MARKUP_TYPE
  PMT_HTML
  PMT_CSS
End Enum

Private m_iem As IPageMarkup
Private m_ElementsToSearchFor() As ELEMENT_TO_SEARCH_FOR
Private m_ElementsToSearchForMax As Long
Private m_ElementsToSearchForIndex As Long
'add the element name you wish to search for
Public Sub AddElementType(ByVal sname As String, Optional ByVal FindInnerHTML As Boolean = False)
  Dim i As Long
  Dim ssearch As String
  
On Error GoTo AddElementType_ERR
  
  sname = LCase$(sname)
  For i = 1 To m_ElementsToSearchForIndex
    If StrComp(m_ElementsToSearchFor(i).Name, sname) = 0 Then
      Exit Sub
    End If
  Next
  
  m_ElementsToSearchForIndex = m_ElementsToSearchForIndex + 1
  If m_ElementsToSearchForIndex > m_ElementsToSearchForMax Then
    m_ElementsToSearchForMax = m_ElementsToSearchForMax + L_RESIZE_INCREASE
    ReDim Preserve m_ElementsToSearchFor(1 To m_ElementsToSearchForMax)
  End If
  
  m_ElementsToSearchFor(m_ElementsToSearchForIndex).Name = sname
  
  ssearch = "<" & sname
  m_ElementsToSearchFor(m_ElementsToSearchForIndex).Search = ssearch
  m_ElementsToSearchFor(m_ElementsToSearchForIndex).LenSearch = Len(ssearch)
  m_ElementsToSearchFor(m_ElementsToSearchForIndex).FindInnerHTML = FindInnerHTML
  
AddElementType_END:
  Exit Sub
AddElementType_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "AddElementType"), Err.Description)
  Resume
End Sub
Private Function ElementstoSearchForDefaults()
  Dim i As Long
  For i = 1 To m_ElementsToSearchForIndex
    m_ElementsToSearchFor(i).SearchPos = 1
    m_ElementsToSearchFor(i).Position = -1
  Next
End Function
Private Function ElementPositions(eps() As ELEMENT_TO_SEARCH_FOR, ByRef sTextRaw As String, ByRef sTextLCase As String) As Long
  Dim iPos As Long, i As Long
  Dim iElementtoSearchForActOn As Long, iPosMin As Long
  Dim iElementPosition As Long, iElementPositionMax As Long
  Dim sNextChar As String
  Dim iLenText As Long
  Dim iChar As Long
  
On Error GoTo ElementPositions_ERR
  
  iElementPositionMax = 10
  ReDim eps(1 To iElementPositionMax)
  
  Call ElementstoSearchForDefaults
  iLenText = Len(sTextLCase)
  
  iElementtoSearchForActOn = -1
  Do While iElementtoSearchForActOn <> 0
    iElementtoSearchForActOn = 0
    iPosMin = 0
    For i = 1 To m_ElementsToSearchForIndex
      If m_ElementsToSearchFor(i).Position = -1 Then
try_again:
        m_ElementsToSearchFor(i).Position = InStr(m_ElementsToSearchFor(i).SearchPos, sTextLCase, m_ElementsToSearchFor(i).Search)
        If m_ElementsToSearchFor(i).Position > 0 Then
          Debug.Print Mid$(sTextLCase, m_ElementsToSearchFor(i).Position, 40)
          iPos = m_ElementsToSearchFor(i).Position + m_ElementsToSearchFor(i).LenSearch
          If iPos < iLenText Then
            sNextChar = Mid$(sTextLCase, iPos, 1)
            iChar = Asc(sNextChar)
            If Not ((iChar = 32) Or (iChar = 47) Or (iChar = 62)) Then
              m_ElementsToSearchFor(i).SearchPos = iPos
              GoTo try_again:
            End If
          End If
          If iPosMin = 0 Then
            iElementtoSearchForActOn = i
            iPosMin = m_ElementsToSearchFor(i).Position
          ElseIf m_ElementsToSearchFor(i).Position < iPosMin Then
            iElementtoSearchForActOn = i
            iPosMin = m_ElementsToSearchFor(i).Position
          End If
        End If

      ElseIf m_ElementsToSearchFor(i).Position <> 0 Then
        If iPosMin = 0 Then
          iElementtoSearchForActOn = i
          iPosMin = m_ElementsToSearchFor(i).Position
        ElseIf m_ElementsToSearchFor(i).Position < iPosMin Then
          iElementtoSearchForActOn = i
          iPosMin = m_ElementsToSearchFor(i).Position
        End If
      End If
    Next
    
    If iElementtoSearchForActOn <> 0 Then
      iElementPosition = iElementPosition + 1
      If iElementPosition = iElementPositionMax Then
        iElementPositionMax = iElementPositionMax + L_RESIZE_INCREASE
        ReDim Preserve eps(1 To iElementPositionMax)
      End If
      eps(iElementPosition) = m_ElementsToSearchFor(iElementtoSearchForActOn)
      m_ElementsToSearchFor(iElementtoSearchForActOn).SearchPos = m_ElementsToSearchFor(iElementtoSearchForActOn).Position + m_ElementsToSearchFor(iElementtoSearchForActOn).LenSearch
      m_ElementsToSearchFor(iElementtoSearchForActOn).Position = -1
    End If
  Loop
    
  If iElementPositionMax > iElementPosition And (iElementPosition > 0) Then
    ReDim Preserve eps(1 To iElementPosition)
  End If
  
ElementPositions_END:
  ElementPositions = iElementPosition
  Exit Function
ElementPositions_ERR:
  Call Err.Raise(ERR_ERROR, ErrorSource(Err, "ElementPositions"), Err.Description)
  Resume
End Function
Public Function Markup(ByRef sTextRaw As String, Optional ByVal PMT As PAGE_MARKUP_TYPE = PMT_HTML) As String
  On Error GoTo ERR_ERR
  
  If m_iem Is Nothing Then Call Err.Raise(380, "MarkupHTML", "Please set the IElementMarkup property to a valid object that implements IElementMarkup")
  
  Select Case PMT
    Case PMT_HTML
      Markup = MarkupHTML(sTextRaw)
    Case PMT_CSS
      
    Case Else
      Call Err.Raise(380, "Markup", "Invalid PAGE_MARKUP_TYPE:" & PMT)
  End Select
  
ERR_END:
  Exit Function
ERR_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "Markup"), Err.Description)
  Resume
End Function

Private Function MarkupHTML(ByRef sTextRaw As String) As String
  Dim qs As QString
  Dim i As Long, j As Long, k As Long, l As Long
  Dim eps() As ELEMENT_TO_SEARCH_FOR
  Dim iElementPositions As Long
  Dim sTextLCase As String
  
On Error GoTo ERR_ERR
  
  
  Set qs = New QString
  
  sTextLCase = LCase$(sTextRaw)
  iElementPositions = ElementPositions(eps, sTextRaw, sTextLCase)
  
  If iElementPositions > 0 Then
    k = 1
    For i = 1 To iElementPositions
      j = eps(i).Position
      Call qs.AppendFrom(sTextRaw, k, j - k)
      k = InStr(j, sTextRaw, ">", vbBinaryCompare)
      Call qs.Append(MarkupElement(Mid$(sTextRaw, j, (k - j) + 1), eps(i), sTextRaw, sTextLCase, j))
      k = k + 1
      If k > Len(sTextRaw) Then Exit For
    Next
    If k <= Len(sTextRaw) Then Call qs.Append(Mid$(sTextRaw, k))
    MarkupHTML = qs
  Else
    MarkupHTML = sTextRaw
  End If
  
ERR_END:
  Exit Function
ERR_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "MarkupHTML"), Err.Description)
  Resume
End Function
Private Function MarkupCSS(ByRef sTextRaw As String) As String
  Dim qs As QString
  Dim sTextLCase As String, surl As String
  Dim p0 As Long, p1 As Long, p2 As Long
  Dim iLen As String
  
On Error GoTo ERR_ERR
  
  Set qs = New QString
  
  iLen = Len(sTextRaw)
  sTextLCase = LCase$(sTextRaw)
  p0 = InStr(1, sTextLCase, "url(", vbBinaryCompare)
  p2 = 1
  If p0 = 0 Then
    MarkupCSS = sTextRaw
    GoTo ERR_END
  End If
  Do While p0 > 0
    p1 = InStr(p0, sTextLCase, ")", vbBinaryCompare)
    If p1 = 0 Then Call Err.Raise(ERR_PAGE_MARKUP, "MarkupCSS", "'url(' has not closing bracket, invaid page")
    Call qs.AppendFrom(sTextRaw, p2, p1 - p0)
    surl = Mid$(sTextRaw, p0, p1 - p0)
    'Call m_iem.CSSURL(url)
    p0 = InStr(p1 + 1, sTextLCase, "url(", vbBinaryCompare)
    p2 = p1 + 1
  Loop
  If (p2 < iLen) Then
    Call qs.AppendFrom(sTextRaw, p2)
  End If
  
  
ERR_END:
  Exit Function
ERR_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "MarkupCSS"), Err.Description)
  Resume
End Function



Private Function MarkupElement(selementtext As String, ByRef eps As ELEMENT_TO_SEARCH_FOR, ByRef sTextRaw As String, ByRef sTextLCase As String, ByVal CurrentPosition As Long)
  Dim Attributes As Collection
  Dim am As AttributeMarkup
  Dim bInQuote As Boolean
  Dim b() As Byte
  Dim sInnerHTML As String
  Dim iLen As Long, i As Long, iLenElementStart As Long
    
  Dim bstrIndex As Long, iQuoteChar As Long, iChar As Long, iLastChar As Long
  Dim pAttributeNameStart As Long, pAttributeNameEnd As Long, pAttributeValueStart As Long, pAttributeValueEnd As Long
  Dim bUnquotedAttrib As Boolean
  Dim TextAfter As String, TextBefore As String
  Dim sattributes As String, selementend As String, snewtext As String
  Dim ub As Long
On Error GoTo MarkupElement_ERR
  
  Set Attributes = New Collection
  iLastChar = -1
  iLen = Len(selementtext)
  iLenElementStart = eps.LenSearch
  
  If StrComp("/", Mid$(selementtext, iLen - 1, 1)) = 0 Then
    selementend = "/>"
    If iLen - 1 <> iLenElementStart + 1 Then
      sattributes = Mid$(selementtext, iLenElementStart + 1, iLen - (iLenElementStart + Len(selementend)))
    End If
    
    'if eps(i).asdadasd then find the inner HTML, pass byref to caller
  Else
    selementend = ">"
    If iLen <> eps.LenSearch + 1 Then
      sattributes = Mid$(selementtext, iLenElementStart + 1, iLen - ((iLenElementStart + 1)))
    End If
    If eps.FindInnerHTML Then
      i = InStr(CurrentPosition + iLen, sTextLCase, "</" & eps.Name & ">", vbBinaryCompare)
      If i <> 0 Then
        sInnerHTML = Mid$(sTextRaw, CurrentPosition + iLen, i - (CurrentPosition + iLen))
      End If
    End If
  End If
  
  b = sattributes
  ub = UBound(b)
  For i = LBound(b) To UBound(b) Step 2
    'If iChar <> 0 Then iLastChar = iChar
    iChar = b(i)
    bstrIndex = bstrIndex + 1
    If iChar = L_SINGLE_QUOTE Or iChar = L_DOUBLE_QUOTE And Not bUnquotedAttrib Then
      am.QuoteChar = iChar
      If bInQuote Then
        If iQuoteChar = iChar Then
          pAttributeValueEnd = bstrIndex - 1
          If pAttributeValueEnd <> pAttributeValueStart Then
ATTRIBUTE_VALUE_END:
            am.Value = Mid$(sattributes, pAttributeValueStart, (pAttributeValueEnd - pAttributeValueStart) + 1)
            Call Attributes.Add(am, LCase$(am.Name))
            pAttributeNameEnd = 0
            pAttributeNameStart = 0
            pAttributeValueEnd = 0
            pAttributeValueStart = 0
          End If
          bInQuote = False
          iQuoteChar = 0
        End If
      Else
        pAttributeValueStart = bstrIndex + 1
        bInQuote = True
        iQuoteChar = iChar
      End If
      
    ElseIf bUnquotedAttrib Then
     If iChar = L_SPACE Then
        pAttributeValueEnd = bstrIndex - 1
        am.Value = Mid$(sattributes, pAttributeValueStart, (pAttributeValueEnd - pAttributeValueStart) + 1)
        Call Attributes.Add(am, LCase$(am.Name))
        pAttributeNameEnd = 0
        pAttributeNameStart = 0
        pAttributeValueEnd = 0
        pAttributeValueStart = 0
        bUnquotedAttrib = False
     End If
    Else
      If Not bInQuote Then
        If iChar <> L_SPACE Then
          If pAttributeNameStart = 0 Then
            Set am = New AttributeMarkup
            pAttributeNameStart = bstrIndex
          Else
          
            If bstrIndex = Len(sattributes) Then
              If pAttributeValueStart <> 0 Then
                pAttributeValueEnd = bstrIndex
                GoTo ATTRIBUTE_VALUE_END
              Else
                GoTo ATTRIBUTE_END
              End If
            End If
            If iChar = L_EQUALS Then
              If i + 2 < ub Then
                If b(i + 2) <> L_SINGLE_QUOTE And b(i + 2) <> L_DOUBLE_QUOTE Then
                  bUnquotedAttrib = True
                End If
              End If
              If pAttributeNameStart <> 0 And (pAttributeNameEnd = 0) Then
                pAttributeValueStart = bstrIndex + 1
                pAttributeNameEnd = bstrIndex - 1
                am.HasEquals = True
                am.Name = Mid$(sattributes, pAttributeNameStart, (pAttributeNameEnd - pAttributeNameStart) + 1)
              End If
            End If
          End If
        Else
          If pAttributeNameStart <> 0 And (pAttributeNameEnd = 0) Then
            
ATTRIBUTE_END:
pAttributeNameEnd = bstrIndex
            am.HasEquals = False
            am.Name = Mid$(sattributes, pAttributeNameStart, (pAttributeNameEnd - pAttributeNameStart) + 1)
            Call Attributes.Add(am, LCase$(am.Name))
            pAttributeNameEnd = 0
            pAttributeNameStart = 0
            pAttributeValueEnd = 0
            pAttributeValueStart = 0
          End If
        End If
      End If
    End If
  Next
  
  Call m_iem.HTMLElement(Attributes, eps.Name, TextBefore, TextAfter, sInnerHTML)
  snewtext = eps.Search
  For i = 1 To Attributes.Count
    Set am = Attributes(i)
    snewtext = snewtext & " " & am.Name
    If am.HasEquals Then
      If am.QuoteChar = 0 Then
        snewtext = snewtext & "=" & am.Value
      Else
        snewtext = snewtext & "=" & Chr$(am.QuoteChar) & am.Value & Chr$(am.QuoteChar)
      End If
    End If
  Next
  MarkupElement = TextBefore & snewtext & selementend & TextAfter
  
MarkupElement_END:
  Exit Function
MarkupElement_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "MarkupElement"), Err.Description)
  Resume
End Function
Private Sub Class_Initialize()
  m_ElementsToSearchForMax = L_RESIZE_INCREASE
  ReDim m_ElementsToSearchFor(1 To m_ElementsToSearchForMax)
End Sub
Public Function AttributePresent(Attributes As Collection, ByVal Key As String, Optional ByRef am As AttributeMarkup) As Boolean
  
  On Error GoTo AttributePresent_ERR
  Set am = Nothing
  Set am = Attributes(LCase$(Key))
  AttributePresent = True
  
AttributePresent_ERR:
  Exit Function
End Function
Public Property Set IPageMarkup(ByRef NewValue As IPageMarkup)
  Set m_iem = NewValue
End Property


