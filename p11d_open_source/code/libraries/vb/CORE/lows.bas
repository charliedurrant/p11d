Attribute VB_Name = "lows"
Option Explicit


Public Function logfunction3(ByVal ErrorDateTime As Date, ByVal ErrorNumber As Long, ErrorName As String, ErrorText As String, FunctionName As String, FileExt As String) As Boolean
  Dim iFileNum As Integer
  Dim sFilePath As String, s As String
    
  On Error GoTo logfunction3_err
  iFileNum = -1
  logfunction3 = False
  ErrorText = Replace(ErrorText, vbCr, " ", , , vbBinaryCompare)
  ErrorText = LTrim$(Replace(ErrorText, vbLf, "", , , vbBinaryCompare))
  
  If Not m_IPreErrorFilter Is Nothing Then
    If Not m_IPreErrorFilter.FilterErrorMessage(GetNetUser_s(False), ErrorDateTime, ErrorNumber, ErrorName, ErrorText, FunctionName) Then
      GoTo logfunction3_err
    End If
  End If
  iFileNum = FreeFile
  sFilePath = mAppPath & "\" & mAppExeName & FileExt
  Open sFilePath For Append Lock Write As iFileNum
    Print #iFileNum, GetNetUser_s(False) & vbTab & Format$(ErrorDateTime, "hh:mm dd/mm/yyyy") & vbTab & ErrorName & vbTab & ErrorText & ": Function " & FunctionName
  logfunction3 = True
  
logfunction3_end:
  If iFileNum > 0 Then Close #iFileNum
  Exit Function
  
logfunction3_err:
  logfunction3 = False
  Resume logfunction3_end
End Function

'note: see VTEXT if changed
Public Function GetFullYear_CD(ByVal nyear As Integer, ConvStr As String) As Integer
  If nyear < 100 Then
    If nyear > YEAR1900CONV Then
      nyear = nyear + 1900
    ElseIf nyear < YEAR2000CONV Then
      nyear = nyear + 2000
    Else
      Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCrLf & "Two digit years greater than " & CStr(YEAR2000CONV) & " and less than " & CStr(YEAR1900CONV) & " must be in the format YYYY"
    End If
  End If
  GetFullYear_CD = nyear
End Function

Public Sub FillVersions(lblName As Label, lblVersion As Label)
  Dim lv As LibraryVersion
  Dim VerString As String
  
  On Error Resume Next
  
  'App
  lblName.Caption = mAppName & vbCrLf
  lblVersion.Caption = mAppVersion & vbCrLf
  
  'Core
  lblName.Caption = lblName.Caption & "Core library" & vbCrLf
  lblVersion.Caption = lblVersion.Caption & mTCSCoreVersion & vbCrLf
  
  For Each lv In LibraryVersions
    VerString = lv.Version
    #If DEBUGVER Then
      VerString = VerString & " Debug"
    #End If
    lblName.Caption = lblName.Caption & lv.Name & vbCrLf
    lblVersion.Caption = lblVersion.Caption & VerString & vbCrLf
  Next lv
End Sub
 

Public Function FileExistsEx(sFname As String, ByVal bDirectory As Boolean, ByVal bShowErrors As Boolean) As Boolean
  Dim Attrs As Long
  On Error GoTo fileexists_err
  FileExistsEx = False
  sFname = Replace$(sFname, Chr(34), "")
  Attrs = GetAttr(sFname)
  If bDirectory Then
    FileExistsEx = (Attrs And vbDirectory)
  Else
    If (Attrs And vbDirectory) = 0 Then
      FileExistsEx = True
    Else
      If bShowErrors Then Err.Raise 53
    End If
  End If
  
fileexists_end:
  Exit Function
  
fileexists_err:
  If bShowErrors Then
    Call ErrorMessageEx(ERR_ERROR, Err, "FileExists", "ERR_FILE_NOT_FOUND", "The file " & sFname & " cannot be located", False)
  End If
  FileExistsEx = False
  Resume fileexists_end
End Function

Public Function xStrPadEx(String1 As String, Pad As String, ByVal Length As Long, ByVal PadFront As Boolean) As String
  Dim padlen As Long
  Dim s As String
    
  padlen = Length - Len(String1)
  If padlen > 0 Then
    s = String$(padlen, Pad)
    If PadFront Then
      xStrPadEx = s & String1
    Else
      xStrPadEx = String1 & s
    End If
  Else
    xStrPadEx = left$(String1, Length)
  End If
End Function

Public Sub AddStaticEx(ByVal StaticName As String, Optional ByVal DefaultValue As String = "", Optional OverrideValue As Variant, Optional ByVal Persistent As Boolean = True)
  Dim sValue As String
  Dim sItem As New staticcls
  
  On Error Resume Next
  sValue = GetStaticEx(StaticName)
  If IsMissing(OverrideValue) Then
    If Len(sValue) = 0 Then
      If Persistent Then
        sValue = GetIniEntryEx(STATICS_SECTION, StaticName, DefaultValue, mStaticFileName)
      Else
        sValue = DefaultValue
      End If
    End If
  Else
    sValue = OverrideValue
  End If
  If StrComp(StaticName, "ApplicationName", vbTextCompare) = 0 Then mAppName = sValue
  If StrComp(StaticName, "Version", vbTextCompare) = 0 Then mAppVersion = sValue
  gstaticdata.Remove StaticName
  sItem.Name = StaticName
  sItem.Value = sValue
  sItem.bPersist = Persistent
  gstaticdata.Add sItem, StaticName
End Sub

Public Function GetStaticEx(ByVal StaticName As String, Optional ByVal ShowError As Boolean = False) As Variant
  Dim sItem As staticcls
  
  On Error GoTo GetStaticEx_err
  Set sItem = gstaticdata.Item(StaticName)
  If Not sItem Is Nothing Then
    GetStaticEx = sItem.Value
  Else
    GetStaticEx = ""
  End If

GetStaticEx_err:
  Exit Function
  
GetStaticEx_end:
  If ShowError Then Call ECASE_SYS("GetStatic: " & CStr(StaticName))
  Resume GetStaticEx_end
End Function

Public Function InCollectionEx(col As Object, vItem As Variant) As Boolean
  Dim v As Variant

  On Error Resume Next
  Set v = col.Item(vItem)
  InCollectionEx = Not (v Is Nothing)
End Function

Public Sub ReadAllStatics()
  Dim Keys() As String, MaxKey As Long, sItem As staticcls
  Dim sValue As String
  Dim i As Long
  
  On Error Resume Next
  MaxKey = GetIniKeyNamesInt(Keys, STATICS_SECTION, mStaticFileName)
  For i = 1 To MaxKey
    sValue = GetIniEntryEx(STATICS_SECTION, Keys(i), "", mStaticFileName)
    If Not InCollectionEx(gstaticdata, Keys(i)) Then
      Set sItem = New staticcls
      sItem.Name = Keys(i)
      sItem.Value = sValue
      sItem.bPersist = True
      gstaticdata.Add sItem, Keys(i)
    End If
  Next i
End Sub

Public Function GetTypedValueEx(v As Variant, ByVal dType As DATABASE_FIELD_TYPES) As Variant
  Select Case dType
    Case TYPE_STR
      GetTypedValueEx = CStr(v)
    Case TYPE_LONG
      GetTypedValueEx = CLng(v)
    Case TYPE_DOUBLE
      GetTypedValueEx = CDbl(v)
    Case TYPE_DATE
      GetTypedValueEx = Null
      If VarType(v) = vbDate Then
        GetTypedValueEx = v
      ElseIf VarType(v) = vbString Then
        If InStr(1, v, "/", vbBinaryCompare) > 0 Then
          If InStr(v, ":") > 0 Then
            GetTypedValueEx = ConvertDateEx(v, CONVERT_DELIMITED, "DMYHNS", "/", ":")
          Else
            GetTypedValueEx = ConvertDateEx(v, CONVERT_DELIMITED, "DMY", "/", ":")
          End If
        ElseIf InStr(1, v, "-", vbBinaryCompare) > 0 Then
          If InStr(v, ":") > 0 Then
            GetTypedValueEx = ConvertDateEx(v, CONVERT_DELIMITED, "DMYHNS", "-", ":")
          Else
            GetTypedValueEx = ConvertDateEx(v, CONVERT_DELIMITED, "YMD", "-", ":")
          End If
        End If
      End If
      If IsNull(GetTypedValueEx) Then GetTypedValueEx = CDate(v)
      If GetTypedValueEx = UNDATED Then Err.Raise ERR_INVALID_TYPE, "GetTypedValueEx", "Unable to convert " & CStr(v) & " to a date."
    Case TYPE_BOOL
      If Not IsBooleanEx(v) Then Err.Raise ERR_INVALID_TYPE, "GetTypedValueEx", "Unable to convert " & CStr(v) & " to a boolean value."
      GetTypedValueEx = CBooleanEx(v)
    Case Else
      Err.Raise ERR_INVALID_TYPE, "GetTypedValueEx", "Unrecognised type: " & CStr(dType)
  End Select
End Function

Public Function IsVarNumeric(v As Variant) As Boolean
  IsVarNumeric = (VarType(v) = vbInteger) Or (VarType(v) = vbLong) Or (VarType(v) = vbSingle) Or (VarType(v) = vbDouble) Or (VarType(v) = vbByte) Or (VarType(v) = vbDecimal) Or (VarType(v) = vbCurrency)
End Function

Public Function IsBooleanEx(ByVal v As Variant) As Boolean
  If IsVarNumeric(v) Or (VarType(v) = vbBoolean) Then
    IsBooleanEx = True
  ElseIf VarType(v) = vbString Then
    IsBooleanEx = (StrComp(v, "True", vbTextCompare) = 0) Or _
                (StrComp(v, "False", vbTextCompare) = 0) Or _
                (StrComp(v, "on", vbTextCompare) = 0) Or _
                (StrComp(v, "off", vbTextCompare) = 0) Or _
                (StrComp(v, "yes", vbTextCompare) = 0) Or _
                (StrComp(v, "no", vbTextCompare) = 0) Or _
                (StrComp(v, "-1", vbTextCompare) = 0) Or _
                (StrComp(v, "0", vbTextCompare) = 0)
  End If
End Function

Public Function CBooleanEx(ByVal v As Variant) As Boolean
  If VarType(v) = vbBoolean Then
    CBooleanEx = v
  ElseIf IsVarNumeric(v) Then
    CBooleanEx = Not (v = 0)
  ElseIf VarType(v) = vbString Then
    CBooleanEx = (StrComp(v, "True", vbTextCompare) = 0) Or (StrComp(v, "on", vbTextCompare) = 0) Or (StrComp(v, "yes", vbTextCompare) = 0) Or (StrComp(v, "-1", vbTextCompare) = 0)
  End If
End Function

Public Function IsDateTime(ByVal d0 As Date)
  IsDateTime = Not (Fix(d0) = d0)
End Function

Public Function IsArrayEx2(ByRef v As Variant) As Boolean
  On Error GoTo IsArrayEx2_err
  
  If IsArray(v) Then
    If UBound(v) >= LBound(v) Then IsArrayEx2 = True
  End If
  
IsArrayEx2_end:
  Exit Function
  
IsArrayEx2_err:
  Resume IsArrayEx2_end
End Function


Public Function isDirEx(ByVal FileName As String) As Boolean
  If Len(FileName) = 0 Then
    isDirEx = False
  Else
    isDirEx = (GetAttr(FileName) And vbDirectory) = vbDirectory
  End If
End Function

Public Function FindFilesEx(ByVal FileDirectory As String, ByVal FileMask As String, ByVal SubDirs As Boolean, ByVal IncludeDirectories As Boolean) As StringList
  Dim sFile As String, sFullPath As String
  Dim i As Long
  Dim sFiles As StringList, sDirs As StringList

#If DEBUGVER Then
  Call Tracer_XSet("FindFilesEx")
#End If
  If Not FileExistsEx(FileDirectory, True, False) Then Err.Raise ERR_FINDFILES, "FindFilesEx", "Directory " & FileDirectory & " does not exist"
  FileDirectory = FullPathEx(FileDirectory)
  Set sDirs = New StringList
  Set sFiles = New StringList
  Call sDirs.Add(FileDirectory)
  
process_dirs:
  For i = 1 To sDirs.Count
    FileDirectory = sDirs.Item(i)
    ' files first
    sFile = Dir$(FileDirectory & FileMask)
    Do While Len(sFile) > 0
      sFullPath = FileDirectory & sFile
      If Not isDirEx(sFullPath) Then Call sFiles.Add(sFullPath)
      sFile = Dir$
    Loop
    
    If SubDirs Then
      ' directories next
      sFile = Dir$(FileDirectory & "*.*", vbDirectory)
      Do While Len(sFile) > 0
        If Not ((StrComp(sFile, ".", vbBinaryCompare) = 0) Or (StrComp(sFile, "..", vbBinaryCompare) = 0)) Then
          sFullPath = FileDirectory & sFile
          If isDirEx(sFullPath) Then
            sFullPath = FullPathEx(sFullPath)
            Call sDirs.Add(sFullPath)
            If IncludeDirectories Then Call sFiles.Add(sFullPath)
          End If
        End If
        sFile = Dir$
      Loop
      Call sDirs.Remove(FileDirectory)
      GoTo process_dirs
    End If
  Next i
  Set FindFilesEx = sFiles
#If DEBUGVER Then
  Call Tracer_XReturn("FindFilesEx")
#End If
End Function
