VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Changed all non-public function to not include xSet/Return or error
'trapping as suggested by apf

Option Explicit
Private Const BSEP = "<"
Private Const ESEP = ">"
Private Const STRSEP = "#"
Private Const PBSEP = "{"
Private Const PESEP = "}"

Private mParseCols As New Collection
Public lTabs As Long
Public lCols As Long

Public Function Add(ByVal sField As String, ByVal sName As String) As Object
  Dim pcolNew As New ParseCol
  
  'Need to set defaults here
  mParseCols.Add pcolNew, sField
  pcolNew.lID = mParseCols.Count
  pcolNew.RealName = sField
  pcolNew.DisplayName = sReplaceChar(sName, "~", vbCrLf)
  Set Add = pcolNew

End Function

Public Function Remove(Index As Long)
  mParseCols.Remove Index
End Function

Public Function Count() As Long
  Count = mParseCols.Count
End Function

Public Function Item(ByVal Index As Variant) As Object
Attribute Item.VB_UserMemId = 0
  Set Item = mParseCols.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mParseCols.[_NewEnum]
End Function

Public Sub Parse(rs As Recordset, rsCodes As Recordset)
  Dim fld As Field
  Dim sParse As String
  Dim sCode As String
  Dim sName As String
  Dim pseCurrent As ParseCol
  Dim lStart As Long
  Dim lEnd As Long
  Dim i As Long
  
  On Error GoTo Parse_err
  xSet ("Parse")
  
  For Each fld In rs.Fields
    sName = ""
    sParse = ""
    Call GetCode(rsCodes, fld.Name, sName, sParse)
    Set pseCurrent = Add(fld.Name, sName)
    pseCurrent.DataType = fld.Type
    sParse = UCase(sParse)
    lStart = 0
    lEnd = -1
    For i = 1 To Len(sParse)
      If Mid(sParse, i, 1) = BSEP Then
        lStart = i
      End If
      If Mid(sParse, i, 1) = ESEP Then
        lEnd = i
      End If
      If lStart < lEnd Then
        sCode = Mid(sParse, lStart + 1, lEnd - (lStart + 1))
        Select Case Mid(sCode, 1, 1)
      
        Case "A"
          Call ParseAll(pseCurrent, sCode)
          
        Case "P"
          Call ParsePrint(pseCurrent, sCode)
          
        Case "S"
          Call ParseScreen(pseCurrent, sCode)
          
        Case "D"
          Call ParseDialog(pseCurrent, sCode)
          
        Case Else
          ECASE ("Parse")
          
        End Select
        lStart = 0
        lEnd = -1
      End If
    Next i
  Next fld
  Set fld = Nothing
  xReturn ("Parse")
  
  Exit Sub
  
Parse_err:
  Set fld = Nothing
  xReturn ("Parse")
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub ParseAll(ByRef pse As ParseCol, ByVal sParse As String)
  xSet ("ParseAll")
  
  Do While Len(sParse) > 1
    
    Select Case Mid(sParse, 2, 1)
          
    Case "P"
      pse.Precision = CLng(Mid(sParse, 3, NotInStr(sParse, "0123456789", 3) - 3))
      sParse = Mid(sParse, 1, 1) & Mid(sParse, NotInStr(sParse, "0123456789", 3))
      
    Case "S"
      pse.Significant = CLng(Mid(sParse, 3, NotInStr(sParse, "0123456789", 3) - 3))
      sParse = Mid(sParse, 1, 1) & Mid(sParse, NotInStr(sParse, "0123456789", 3))
      
    Case "T"
      
      Select Case Mid(sParse, 3, 3)
      
      Case "B"
        pse.DataType = vbBoolean
      
      Case Else
        ECASE ("ParseAll - DataType")
        'PC Need to take care of data type more effectively
        
      End Select
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 4)
      
    Case "D"
      pse.DateFormat = Mid(sParse, 4, InStr(4, sParse, "#") - 4)
      sParse = Mid(sParse, 1, 1) & Mid(sParse, InStr(4, sParse, "#") + 1)
    
    Case "A"
      pse.Alignment = GetAlign(Mid(sParse, 3, 1))
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 4)
      
    Case "H"
      pse.Hide = InStr(sParse, "H") <> 0
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 3)
      
    Case Else
      ECASE ("Parse")
        
    End Select
  
  Loop
  pse.NumberFormat = GetFormat(pse)
  xReturn ("ParseAll")

End Sub

Private Sub ParsePrint(ByRef pse As ParseCol, ByVal sParse As String)
  xSet ("ParsePrint")
  Do While Len(sParse) > 1
    
    Select Case Mid(sParse, 2, 1)
      
    Case "S"
      pse.Sum = InStr(1, sParse, "T") <> 0
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 3)
    
    Case "P"
      pse.PageBreak = InStr(1, sParse, "N") <> 0
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 3)
      
    Case "H"
      pse.Header = InStr(1, sParse, "H") <> 0
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 3)
      
    Case "T"
      pse.SubTotal = InStr(1, sParse, "S") <> 0
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 3)
      
    Case Else
      ECASE ("Parse")
              
    End Select
  
  Loop
  xReturn ("ParsePrint")
End Sub

Private Sub ParseScreen(ByRef pse As ParseCol, ByVal sParse As String)
  
  Dim lStart As Long
  Dim lEnd As Long
  Dim i As Long
  Dim sParam As String
  
xSet ("ParseScreen")
  Do While Len(sParse) > 1
    
    Select Case Mid(sParse, 2, 1)
      
    Case "V"
      lStart = InStr(4, sParse, PBSEP) + 1
      lEnd = InStr(4, sParse, PESEP) - lStart
      pse.QueryStyle = CLng(Mid(sParse, lStart, lEnd))
      lEnd = InStr(7, sParse, PBSEP) - 7
      If lEnd <= 0 Then
        lEnd = InStr(7, sParse, STRSEP) - 7
      End If
      pse.QueryName = Mid(sParse, 7, lEnd)
      
      lStart = 0
      lEnd = -1
      For i = 7 To Len(sParse)
        
        If Mid(sParse, i, 1) = PBSEP Then
          lStart = i
        End If
        If Mid(sParse, i, 1) = PESEP Then
          lEnd = i
        End If
        If lStart < lEnd Then
          sParam = Mid(sParse, lStart + 1, lEnd - (lStart + 1))
          Call pse.AddParam(sParam)
          lStart = 0
          lEnd = -1
        End If
      Next i
      
      sParse = Mid(sParse, 1, 1) & Mid(sParse, InStr(7, sParse, STRSEP) + 1)
    
    Case "D"
      pse.Drop = True
      lStart = 0
      lEnd = -1
      For i = 1 To Len(sParse)
        
        If Mid(sParse, i, 1) = PBSEP Then
          lStart = i
        End If
        If Mid(sParse, i, 1) = PESEP Then
          lEnd = i
        End If
        If lStart < lEnd Then
          sParam = Mid(sParse, lStart + 1, lEnd - (lStart + 1))
          Call pse.AddParam(sParam)
          lStart = 0
          lEnd = -1
        End If
      Next i
      sParse = Mid(sParse, 1, 1) & Mid(sParse, InStr(4, sParse, "#") + 1)
      
    Case "E"
      pse.NoEdit = InStr(1, sParse, "E") <> 0
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 3)
        
    Case "N"
      pse.NewRecord = Mid(sParse, 4, InStr(4, sParse, STRSEP) - 4)
      sParse = Mid(sParse, 1, 1) & Mid(sParse, InStr(4, sParse, STRSEP) + 1)
    
    Case "I"
      pse.Items = CInt(Mid(sParse, 3, NotInStr(sParse, "0123456789", 3) - 3))
      sParse = Mid(sParse, 1, 1) & Mid(sParse, NotInStr(sParse, "0123456789", 3))
      
    Case Else
      ECASE ("Parse")
      
    End Select
  
  Loop
  xReturn ("ParseScreen")
End Sub

Private Sub ParseDialog(ByRef pse As ParseCol, ByVal sParse As String)
  xSet ("ParseDialog")
  Do While Len(sParse) > 1
    
    Select Case Mid(sParse, 2, 1)
      
    Case "C"
      pse.NewColumn = InStr(1, sParse, "C") <> 0
      sParse = Mid(sParse, 1, 1) & Mid(sParse, 3)
      lCols = lCols + 1
      
    Case "T"
      pse.TabTitle = Mid(sParse, 4, InStr(4, sParse, "#") - 4)
      sParse = Mid(sParse, 1, 1) & Mid(sParse, InStr(4, sParse, "#") + 1)
      lTabs = lTabs + 1
      
    Case Else
      ECASE ("Parse")
            
    End Select
  
  Loop
    xReturn ("ParseDialog")

End Sub

Private Function GetFormat(ByRef pse As ParseCol) As String
  Dim sFormat As String
  Dim i As Long
  
  If pse.Precision Then
    sFormat = "."
    For i = 1 To pse.Precision
      sFormat = sFormat & "0"
    Next i
  End If
  If pse.Significant Then
    sFormat = "0" & sFormat
  End If
  GetFormat = sFormat
End Function

Private Function GetAlign(str As String) As String
  Select Case str
  
  Case "L"
    GetAlign = 0
    
  Case "R"
    GetAlign = 1
    
  Case "C"
    GetAlign = 2
    
  Case Else
    GetAlign = 3
    
  End Select
End Function

Private Sub GetCode(rsCodes As Recordset, sField As String, ByRef sName As String, ByRef sCode As String)
  xSet ("GetCode")
  
  If Not rsCodes Is Nothing Then
    If Not (rsCodes.BOF And rsCodes.EOF) Then
      rsCodes.MoveFirst
      Do
        If UCase(rsCodes.Fields("Field")) = UCase(sField) Then Exit Do
        rsCodes.MoveNext
      Loop Until rsCodes.EOF
      If rsCodes.EOF Then
        sName = sField
        sCode = ""
      Else
        ' gs change to avoid err 94 invalid null
        sName = "" & rsCodes.Fields("Caption")
        sCode = "" & rsCodes.Fields("Format")
      End If
    Else
      sName = sField
      sCode = ""
    End If
  Else
    sName = "" & sField
    sCode = ""
  End If
  xReturn ("GetCode")

End Sub
