Attribute VB_Name = "Parse"
Option Explicit

Public Function ParseEquals(sRet As String, sToParse As String, sToken As String) As Boolean
  Dim i As Long
  sToken = sToken & "="
  
  i = InStr(1, sToParse, sToken, vbBinaryCompare)
  If i > 0 Then
    If i < Len(sToParse) Then
      sRet = Trim(Right$(sToParse, Len(sToParse) - (i + Len(sToken) - 1)))
      sRet = ReplaceString(sRet, """", "")
      ParseEquals = (Len(sRet) > 0)
    End If
  End If
End Function

Public Sub ParseAutoDoc(sToParse As String, fi As FunctionItem)
  Dim sRet As String, s As String
  Dim param As Parameter
  Dim i As Long, j As Long
  
  On Error GoTo ParseAutoDoc_err
  If ParseAutoDocLine(sRet, sToParse, S_AUTODOC_DESCRIPTION) Then
    fi.Description = fi.Description & sRet & " "
  ElseIf ParseAutoDocLine(sRet, sToParse, S_AUTODOC_CATEGORY) Then
    fi.Category = sRet
  ElseIf ParseAutoDocLine(sRet, sToParse, S_AUTODOC_LONG_DESCRIPTION) Then
    fi.DescriptionLong = sRet
  ElseIf ParseAutoDocLine(sRet, sToParse, S_AUTODOC_RETURN_VALUE) Then
    fi.ReturnValueDescription = sRet
  ElseIf ParseAutoDocLine(sRet, sToParse, S_AUTODOC_VARNAME) Then
    i = InStr(1, sRet, " ")
    If i = 0 Then Err.Raise ERR_NOVARNAME, "ParseAutoDoc", "Parsing " & S_AUTODOC_VARNAME & " no variable name found"
    s = Left$(sRet, i - 1)
    For Each param In fi.Params
      If StrComp(s, param.Name, vbTextCompare) = 0 Then
        param.Description = Trim$(Mid$(sRet, i))
        Exit Sub
      End If
    Next
    Err.Raise ERR_NOVARNAME, "ParseAutoDoc", "Parsing " & S_AUTODOC_VARNAME & " variable name not found " & s
  End If
ParseAutoDoc_end:
  Exit Sub

ParseAutoDoc_err:
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ParseAutoDoc", "Parsing function", "Error parsing line:" & vbCrLf & sToParse)
  Resume ParseAutoDoc_end
  Resume
End Sub

Public Function ParseAutoDocLine(sRet As String, sToParse As String, sSearch As String) As Boolean
  Dim i As Long, j As Long, k As Long
    
  j = Len(sSearch)
  k = InStr(1, sToParse, sSearch, vbTextCompare)
  If k > 0 Then
    i = Len(sToParse)
    If i > j Then
      'ok we have a comment
       sRet = Trim(Right$(sToParse, i - (k + j) + 1))
       ParseAutoDocLine = True
    End If
  End If
  
End Function

Public Function ParseParam(sToParse As String, param As Parameter)
  Dim k As Long
  
  Set param = New Parameter
  
  param.IsOptional = InStr(1, sToParse, "Optional ") > 0
  sToParse = ReplaceString(sToParse, "Optional ", "", vbBinaryCompare)
  
  param.IsByval = InStr(1, sToParse, "ByVal ") > 0
  If param.IsByval Then sToParse = ReplaceString(sToParse, "ByVal ", "", vbBinaryCompare)
  param.IsArray = InStr(1, sToParse, "(") > 0
  If param.IsArray Then
    sToParse = ReplaceString(sToParse, "(", "", vbBinaryCompare)
    sToParse = ReplaceString(sToParse, ")", "", vbBinaryCompare)
  End If
  
  param.IsParamArray = InStr(1, sToParse, "ParamArray ") > 0
  If param.IsParamArray Then
    sToParse = ReplaceString(sToParse, "ParamArray ", "", vbBinaryCompare)
    param.IsArray = True
  End If
  k = InStr(1, sToParse, " As ")
  If k > 0 Then
    param.VarType = Right$(sToParse, Len(sToParse) - (k + Len(" As")))
    sToParse = ReplaceString(sToParse, " As " & param.VarType, "", vbBinaryCompare)
  Else
    param.VarType = "Variant"
  End If
  If param.IsOptional Then
    k = InStr(1, sToParse, "=")
    If k > 0 Then
      param.Name = Left$(sToParse, k - 1)
      param.DefaultValue = Right$(sToParse, Len(sToParse) - (k + 1))
      If Len(param.DefaultValue) = 0 Then
        param.DefaultValue = """"
      End If
    Else
      param.Name = sToParse
    End If
  Else
    param.Name = sToParse
  End If
  
  param.Name = Trim$(sToParse)
End Function


Public Function ParseString(s As String, ByVal ParseToken As String, Optional ByVal ChangeString As Boolean = True) As Boolean
  If InStr(1, s, ParseToken, vbTextCompare) = 1 Then
    If ChangeString Then s = Mid$(s, Len(ParseToken) + 1)
    ParseString = True
  End If
End Function
    
