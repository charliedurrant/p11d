Attribute VB_Name = "EstablishTemplate"
Option Explicit

Private Type FileDetails
  FileName As String
  FileType As String
  Year As String
End Type

Private Enum DB_TEMPLATE
  TEMP_UNINIT = -1
  AA = 0
  EXTERNAL = 1
End Enum

Private mFiles() As FileDetails
Private mMaxLine As Long
Private mDBTemplate As DB_TEMPLATE

Public Function GetDbTemplateFile(db As Database) As String
  Dim rs As Recordset
  
  On Error GoTo GetDbTemplateFile_Err
  Call xSet("GetDbTemplateFile")
  
  If db Is Nothing Then Err.Raise ERR_DB_NOTHING, "GetDbTemplateFile", "No database has been passed to the function."
  If TablePresent(db.TableDefs, "sys_Control") Then
    Set rs = db.OpenRecordset("sys_Control", dbOpenDynaset, dbFailOnError)
    rs.FindFirst ("Flag=" & StrSQL("TemplateFile"))
    If rs.NoMatch Then Err.Raise ERR_APPLY_FIXES, "GetDbTemplateFile", "The template filename could not be found."
    GetDbTemplateFile = GetStatic("NewFileTemplateDir") & "\" & rs!Text
    Set rs = Nothing
  Else
    Err.Raise ERR_APPLY_FIXES, "GetDbTemplateFile", "The control table cannot be found"
  End If

GetDbTemplateFile_End:
  Call xReturn("GetDbTemplateFile")
  Exit Function

GetDbTemplateFile_Err:
  Call xReturn("GetDbTemplateFile")
  Err.Raise ERR_APPLY_FIXES, ErrorSource(Err, "GetTemplateDbFile"), "Error getting the template database's file name."
End Function

Public Function EstablishTemplateFile(db As Database, Optional strOldTemplateFile As String, Optional bExternal As Boolean = False, Optional bInternal As Boolean = False) As String
  Dim rs As Recordset
  Dim strTPYear As String
  Dim strFileType As String
  Dim strTemplateFile As String
  Dim strTempFile As String
  Dim i As Long
  
  On Error GoTo EstablishTemplateFile_Err
  Call xSet("EstablishTemplateFile")
  
  If db Is Nothing Then Err.Raise ERR_DB_NOTHING, "EstablishTemplateFile", "No database has been passed to the function."
  If TablePresent(db.TableDefs, "sys_Control") Then
    Set rs = db.OpenRecordset("sys_Control", dbOpenDynaset, dbFailOnError)
    rs.FindFirst ("Flag=" & StrSQL("TPYear"))
    If rs.NoMatch Then
      Set rs = Nothing
      EstablishTemplateFile = ""
      Exit Function
    End If
    strTPYear = rs!Text
    rs.FindFirst ("Flag=" & StrSQL("FileType"))
    If rs.NoMatch Then
      Set rs = Nothing
      EstablishTemplateFile = ""
      Exit Function
    End If
    strFileType = rs!Text
    Call ReadFileList(strOldTemplateFile, bExternal, bInternal)
    For i = 0 To UBound(mFiles)
      If StrComp(mFiles(i).Year, strTPYear, vbTextCompare) = 0 Then
        If StrComp(mFiles(i).FileType, strFileType, vbTextCompare) = 0 Then
          strTempFile = mFiles(i).FileName
          strTemplateFile = GetStatic("NewFileTemplateDir") & "\" & strTempFile
          Exit For
        End If
      End If
    Next
    If FileExists(strTemplateFile) Then
      rs.FindFirst ("Flag=" & StrSQL("TemplateFile"))
      rs.Edit
        rs!Text = strTempFile
      rs.Update
      EstablishTemplateFile = strTemplateFile
      Set rs = Nothing
    Else
      EstablishTemplateFile = ""
      Set rs = Nothing
    End If
  Else
    Err.Raise ERR_APPLY_FIXES, "EstablishTemplateFile", "The control table cannot be found"
  End If

EstablishTemplateFile_End:
  Call xReturn("EstablishTemplateFile")
  Exit Function

EstablishTemplateFile_Err:
  Call xReturn("EstablishTemplateFile")
  Err.Raise ERR_APPLY_FIXES, ErrorSource(Err, "EstablishTemplateFile"), "Error establishing the template database's file name."
End Function

Private Function ReadFileList(Optional strOldTemplateFile As String, Optional bExternal As Boolean = False, Optional bInternal As Boolean = False) As Boolean
  Dim F As New FileSystemObject, ts As TextStream
  Dim Line As String
  Dim FileType As String
  Dim AvailableLines As Long
  Dim sDir As String
  
  On Error GoTo ReadFileList_Err
  Call xSet("ReadFileList")
    
  ReadFileList = True
  AvailableLines = 20: mMaxLine = 0: ReDim mFiles(0 To AvailableLines - 1)
  sDir = GetStatic("NewFileTemplateDir")
  If F.FolderExists(sDir) Then
    Set ts = F.OpenTextFile(AppPath & "\FileList.tpk", ForReading)
    Do While Not ts.AtEndOfStream
      Line = ts.ReadLine
      If mMaxLine = AvailableLines Then
        AvailableLines = AvailableLines + 20
        ReDim Preserve mFiles(0 To AvailableLines - 1)
      End If
      If ts.Line = 2 Then
        If StrComp(Right$(Line, 4), "True", vbTextCompare) = 0 Then mDBTemplate = EXTERNAL
        If StrComp(strOldTemplateFile, "template.mdb", vbTextCompare) = 0 Then mDBTemplate = AA
        If bExternal Then mDBTemplate = EXTERNAL
        If bInternal Then mDBTemplate = AA
        GoTo NextLine
      End If
      If Left$(Line, 1) = "[" Then
        FileType = Left$(Line, InStr(Line, "]") - 1)
        FileType = Right$(FileType, Len(FileType) - 1)
      Else
        mFiles(mMaxLine).FileType = FileType
        If mDBTemplate = EXTERNAL Then
          mFiles(mMaxLine).FileName = Right$(Line, Len(Line) - InStr(Line, ";"))
        Else
          mFiles(mMaxLine).FileName = Mid$(Line, InStr(Line, ",") + 1, InStr(Line, ";") - InStr(Line, ",") - 1)
        End If
        mFiles(mMaxLine).Year = Left$(Line, InStr(Line, ",") - 1)
        mMaxLine = mMaxLine + 1
      End If
NextLine:
    Loop
    
    If mMaxLine = 0 Then Call Err.Raise(ERR_RS_EMPTY, "ReadFileList", "Cannot find template files to allow creation of a new file")
  Else
    Call Err.Raise(ERR_NO_FILE, "ReadFileList", "Cannot find template files to allow creation of a new file")
  End If
        
ReadFileList_End:
  Call xReturn("ReadFileList")
  Exit Function

ReadFileList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadFileList", "Read file list", "Error reading file list.")
  ReadFileList = False
  Resume ReadFileList_End
  Resume
End Function

