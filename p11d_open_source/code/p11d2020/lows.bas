Attribute VB_Name = "lows"
Option Explicit
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Const L_USER_APP_DATA As Long = CSIDL_PERSONAL   ' needs to be hrere else causes circular ref

Public PreParser As Parser
Public PreRep As Reporter
Public PreAuto As AutoClass
Public PreADOAuto As AutoClass

Public Sub MathInit()
  Dim i As Long
  For i = LOW_POW To HIGH_POW
    Powers(i) = 10 ^ i
  Next i
End Sub



Function FileNameSafe(ByVal sFileName As String) As String
 Const sInvalidChars As String = "/\|<>:*?"""
 Dim lCt As Long
 
 sFileName = TrimEx(sFileName)
 
 If (Len(sFileName) = 0) Then
  FileNameSafe = sFileName
  Exit Function
 End If
 
 For lCt = 1 To Len(sInvalidChars)
  sFileName = Replace(sFileName, Mid(sInvalidChars, lCt, 1), "-")
 Next lCt
 FileNameSafe = sFileName
 
End Function

Public Function TrimEx(ByVal s As String)
  Dim p0 As Long, p1 As Long
  Dim i As Long, iLen As Long

On Error GoTo err_err

  p0 = -1
  p1 = -1

  iLen = Len(s)
  
  For i = 1 To iLen
    If Not abatecrt.IsSpaceStr(Mid$(s, i, 1)) Then
      p0 = i
      Exit For
    End If
  Next
  
  For i = iLen To 1 Step -1
    If Not abatecrt.IsSpaceStr(Mid$(s, i, 1)) Then
      p1 = i
      Exit For
    End If
  Next
  
  If (p0 = -1) Then
    GoTo err_end
  End If
  
  s = Mid$(s, p0, (p1 - p0) + 1)
  
err_end:
  TrimEx = s
  Exit Function
err_err:
  Resume err_end
End Function

Public Function ConvertUNDATEDDateSQL(sFieldName As String, ByVal bEndOfYear As Boolean) As String
  Dim sDefaultValue As String
  
  If bEndOfYear Then
    sDefaultValue = DateSQL(p11d32.Rates.value(TaxYearEnd))
  Else
    sDefaultValue = DateSQL(p11d32.Rates.value(TaxYearStart))
  End If
  ConvertUNDATEDDateSQL = "iif(" & sFieldName & " = " & DateSQL(UNDATED) & " or " & sFieldName & " is null, Null," & sFieldName & ")"
End Function
Public Function TrimMaxLength(ByVal s As String, MaxLength As Long) As String
  s = Trim$(s)
  
  If (Len(s) > MaxLength) Then
    s = Left$(s, MaxLength)
    s = Trim$(s)
  End If
  TrimMaxLength = s
End Function
Public Function TextFileLoad(ByVal sPathAndFile As String) As String
  Dim fr As TCSFileread
  Dim s As String
  
  On Error GoTo err_err
  Set fr = New TCSFileread
  
  If Not fr.OpenFile(sPathAndFile) Then Call Err.Raise(ERR_FILE_OPEN, "TextFileLoad", "Failed to open file " & sPathAndFile)
  Call fr.GetFile(s)
  TextFileLoad = s
  
  
err_end:
  Exit Function
err_err:
  Call Err.Raise(ERR_FILE_INVALID, ErrorSource(Err, "TextFileLoad"), Err.Description)
  
End Function

Public Sub TextFileSave(ByVal sPathAndFile As String, ByRef sText As String)
  Dim bClosing As Boolean
  'Dim fs As FileSystemObject
  'Dim ts As TextStream
  Dim ifile As Long
  
  On Error GoTo err_err
  
  ifile = 0
  ifile = FreeFile
  
  Open sPathAndFile For Output As #ifile
  Print #ifile, sText
  bClosing = True
  Close #ifile
  ifile = 0
  
  'Set fs = New FileSystemObject
  'Set ts = fs.CreateTextFile(sPathAndFile, True)
  
  'Call ts.Write(sText)
  'bClosing = True
  'Call ts.Close
  'Set ts = Nothing
  
  
err_end:
  Exit Sub
err_err:
  If (Not bClosing) And (ifile <> 0) Then Close #ifile
  Call Err.Raise(ERR_FILE_INVALID, ErrorSource(Err, "TextFileLoad"), Err.Description)
  Resume
End Sub

Public Function GuidNewEmployer() As String
  GuidNewEmployer = (Replace$(Replace$(Replace$(GenerateGUID, "{", ""), "}", ""), "-", ""))
End Function
Public Function ValidateFileFromTextBox(ByVal txt As textBox, ByVal bDir As Boolean, Optional ByVal sValidating As String = "'UNKNOWN ITEM'") As Boolean
  Dim sError As String
  
  On Error GoTo ValidateFileFromTextBox_ERR
  
  If Not FileExists(txt.Text, bDir) Then
    sError = sValidating & " "
    If bDir Then
      sError = "Directory"
    Else
      sError = "File"
    End If
    Call Err.Raise(ERR_FILE_NOT_EXIST, ErrorSource(Err, "ValidateFileFromTextBox"), sError & " does not exist, " & txt.Text)
  End If
  If bDir Then txt.Text = FullPath(txt.Text)
  
ValidateFileFromTextBox_END:
  Exit Function
ValidateFileFromTextBox_ERR:
  ValidateFileFromTextBox = True
  Call ErrorMessage(ERR_ERROR, Err, "ValidateFileFromTextBox", "Validate File From Text Box", "Error in ValidateFileFromTextBox.")
  Resume ValidateFileFromTextBox_END
  Resume
End Function
Public Function OpenDB(ws As Workspace, sPathAndFile As String, bExclusive As Boolean) As Database
  If bExclusive Then
    If IsFileOpen(sPathAndFile, True) Then Call Err.Raise(ERR_FILE_OPEN_EXCLUSIVE, "OpenDB", "The database " & sPathAndFile & " is opened exclusively please amend.")
  End If
  If Not FileExists(sPathAndFile) Then Call Err.Raise(ERR_FILE_NOT_EXIST, "OpenDB", "The database " & sPathAndFile & " is not present.")
  If ws Is Nothing Then Call Err.Raise(ERR_WORKSPACE_IS_NOTHING, "OpenPDDB", "Unable to open database, " & sPathAndFile & " as workspace is nothing.")
  If IsRunningInIDE Then Call RemoveReadOnlyFile(sPathAndFile) ' change file attribute from read-only to normal
  Set OpenDB = InitDB(ws, sPathAndFile, "DataBase", , , True)
  If OpenDB Is Nothing Then Call Err.Raise(ERR_DB_IS_NOTHING, "", "Unable to open database, " & sPathAndFile & ", for unknown reason, check if opened exclusively")
End Function
Public Sub SaveTextFile(ByVal path_and_file As String, Text As String)
  On Error GoTo err_err
  
  Dim fs As FileSystemObject
  Dim ts As TextStream
  
  Set fs = New FileSystemObject
  Set ts = fs.OpenTextFile(path_and_file, ForWriting, True)
  Call ts.Write(Text)
  Call ts.Close
  
err_end:
  Exit Sub
err_err:
  If (Not ts Is Nothing) Then
    Call ts.Close
  End If
  Call Err.Raise(Err.Number, "SaveTextFile", "Failed to save the text file:'" & path_and_file & "', " & Err.Description)
End Sub

Public Function IsLeapYear(ByVal vDate As Variant) As Boolean
  On Error GoTo IsLeapYear_ERR
  
  If VarType(vDate) = vbDate Then
     vDate = Year(vDate)
  End If
  
  IsLeapYear = (Day(DateSerial(vDate, 2, 28) + 1) = 29)
    
IsLeapYear_END:
  Exit Function
IsLeapYear_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IsLeapYear", "Is Leap Year", "Error determing if " & vDate & " is a leap year.")
  Resume IsLeapYear_END
End Function

Public Function EnumFiles(ByVal vUserData As Variant, ByVal sPath As String, sMaskToSearch As String, IEF As IEnumFiles) As Long
  Dim sFIle As String
  Dim sFullPath As String
  
  Dim l As Long
  
  On Error GoTo EnumFiles_ERR
  
  Call xSet("EnumFiles")
  
  If Not FileExists(sPath, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "EnumFiles", "The directory does not exist.")
  
  sPath = FullPath(sPath)
  sFullPath = sPath & sMaskToSearch
    
  sFIle = Dir$(sFullPath)
  Do While Len(sFIle) <> 0
    Call IEF.File(vUserData, sPath & sFIle, sFIle)
    l = l + 1
    sFIle = Dir$()
  Loop
  
EnumFiles_END:
  EnumFiles = l
  Call xReturn("EnumFiles")
  Exit Function
EnumFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EnumFiles", "Enum Files", "Error enumeration files, directory = " & sPath & ", mask = " & sMaskToSearch & ".")
  Resume EnumFiles_END
End Function
Public Function Is83FileName(ByVal sFileStr As String) As Boolean
   Dim sFIle As String
   Dim sExt As String
   
   Call SplitPath(sFileStr, , sFIle, sExt)
   Is83FileName = (Len(sFIle) < 9) And Len(sExt) < 5 'includes .
End Function

Public Function IRLoanRateAdjustmentDailyInterestRate(ByVal vSumOfRates As Variant, ByVal lDays As Long) As Double
  Dim dblDailyInterestRate
  
  Const dblIRFullYearRateMax As Double = 0.07166
  Const dblIRFullYearRateMin As Double = 0.07165
  
  If lDays > 0 Then
    dblDailyInterestRate = vSumOfRates / lDays
  Else
    dblDailyInterestRate = vSumOfRates
  End If
  
  'Hack for IR cad/aj
  'If (dblDailyInterestRate > dblIRFullYearRateMin) And (dblDailyInterestRate < dblIRFullYearRateMax) Then
  '  dblDailyInterestRate = 0.0716
  'Else
    dblDailyInterestRate = RoundN(dblDailyInterestRate, 4)
  'End If
    
  IRLoanRateAdjustmentDailyInterestRate = dblDailyInterestRate
    
End Function
Public Function IsFileOpen(FileAndPath As String, Optional Exclusive As Boolean = False) As Boolean
  Dim i As Integer
  
  On Error GoTo IsFileOpen_ERR
  
  i = FreeFile
  If FileExists(FileAndPath) Then
    If Exclusive Then
      Open FileAndPath For Input Shared As i
    Else
      Open FileAndPath For Input Lock Read Write As i
    End If
  End If
  
IsFileOpen_END:
  Close i
  Exit Function
IsFileOpen_ERR:
  IsFileOpen = True
  Resume IsFileOpen_END
End Function

Public Sub FileExistsAndNotOpenExclusive(ByVal sPathAndFile As String)
  If Not FileExists(sPathAndFile) Then Call Err.Raise(ERR_FILE_NOT_EXIST, "ViewFile", "The file " & sPathAndFile & " does not exist.")
  If IsFileOpen(sPathAndFile, True) Then Call Err.Raise(ERR_FILE_OPEN_EXCLUSIVE, "ViewFile", "The file " & sPathAndFile & " is open exclusively.")
End Sub
Public Function DateInRange(ByVal dDate As Date, ByVal dFrom As Date, ByVal dTo As Date) As Boolean
  DateInRange = (dDate >= dFrom) And (dDate <= dTo)
End Function

Public Function IsClientError(ErrNumber As Long) As Boolean
  If ErrNumber >= TCSCLIENT_ERROR And ErrNumber <= TCSCLIENT_ERROR_END Then IsClientError = True
End Function

Public Function GetFileText(ByVal sPathAndFile As String) As String
  Dim FSR As TCSFileread
  Dim s As String
  
  On Error GoTo GetFileText_ERR
  
  Call xSet("GetFileText")
  
  Set FSR = New TCSFileread
  
  
  If Not FSR.OpenFile(sPathAndFile) Then
    If Not FileExists(sPathAndFile) Then
      Call Err.Raise(ERR_FILE_NOT_EXIST, "GetFileText", "The file " & sPathAndFile & " does not exist.")
    Else
      Call Err.Raise(ERR_FILE_OPEN, "GetFileText", "Can not open the file " & sPathAndFile & " check rights to file.")
    End If
  Else
    Call FSR.GetFile(s)
    GetFileText = s
  End If
  
GetFileText_END:
  Call xReturn("GetFileText")
  Exit Function
GetFileText_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetFileText", "Get File Text", "Error reading the file " & sPathAndFile & ".")
  Resume GetFileText_END
  Resume
End Function


Public Function Records(rs As Recordset) As Long
  If Not rs Is Nothing Then
    With rs
      If Not (.EOF And .BOF) Then
        .MoveLast
        Records = .RecordCount
        .MoveFirst
      End If
    End With
  End If
End Function
Private Function xUBound(v As Variant) As Long
  On Error Resume Next
  xUBound = UBound(v)
End Function
Public Function GetFiles(Files() As String, ByVal Path As String, FileMask As String, Optional SubDirs As Boolean = False) As Long
  Dim s As String
  Dim FileCount As Long
  
  On Error GoTo GetFiles_ERR
  Call xSet("GetFiles")
  Path = FullPath(Path)
  FileCount = xUBound(Files)
  
  s = Dir$(Path & FileMask)
  Do While Len(s) > 0
    If (GetAttr(Path & s) And vbDirectory) <> vbDirectory Then
      GetFiles = GetFiles + 1
      ReDim Preserve Files(1 To GetFiles + FileCount) As String
      Files(GetFiles + FileCount) = Path & s
    End If
    s = Dir$()
  Loop
  
  If SubDirs Then
    s = Dir$(Path & "*.*")
    Do While Len(s) > 0
      If (GetAttr(Path & s) And vbDirectory) = vbDirectory Then
        If s <> "." And s <> ".." Then
          GetFiles = GetFiles + GetFiles(Files, Path & s, FileMask, SubDirs)
        End If
      End If
      s = Dir$
    Loop
  End If
    
GetFiles_END:
  Call xReturn("GetFiles")
  Exit Function
  
GetFiles_ERR:
  GetFiles = 0
  Call ErrorMessage(ERR_ERROR, Err, "GetFiles", "Error " & Err.Number, Err.Description)
  Resume GetFiles_END
End Function
Public Function FileAttributes(ByVal sPathAndFile As String)
  Dim fs As FileSystemObject
  Set fs = New FileSystemObject
  FileAttributes = fs.GetFile(sPathAndFile).Attributes
End Function

Public Function ReadOnly(ByVal sPathAndFile As String) As Boolean
  ReadOnly = ((FileAttributes(sPathAndFile) And vbReadOnly) = 1)
End Function

Public Function HiWord(ByVal l As Long) As Long
  HiWord = l \ &H10000 And &HFFFF&
End Function

Public Function LowWord(ByVal l As Long) As Long
  LowWord = l And &HFFFF&
End Function
Public Function LowWordToHiWord(l As Long) As Long
  'first take off the hiword portion
  l = LowWord(l)
  'now multiply up
  LowWordToHiWord = l * (2 ^ 16)
End Function
Public Function TwoLongsToHiAndLow(ByVal HiWordLong As Long, ByVal LowWordLong As Long) As Long
  TwoLongsToHiAndLow = LowWordToHiWord(HiWordLong) + LowWord(LowWordLong)
End Function
Public Function BenefitIsCollection(ben As IBenefitClass) As Boolean
  Dim o As ObjectList
  
  On Error GoTo BenefitIsCollection_END
   
  Set o = ben
  BenefitIsCollection = True
  
BenefitIsCollection_END:
  Exit Function
  
End Function
Public Sub SetSortOrder(lv As ListView, ColumnHeader As ColumnHeader, Optional OverrideSortOrder As ListSortOrderConstants = -1)
  Dim dt As DATABASE_FIELD_TYPES
  Dim bForceStringSort As Boolean
  On Error GoTo SetSortOrder_ERR
  
  If (Len(ColumnHeader.Tag) > 0) Then
    dt = ColumnHeader.Tag
  Else
    If InStr(1, ColumnHeader.Text, "Available", vbTextCompare) > 0 Then
      dt = TYPE_DATE
    ElseIf InStr(1, ColumnHeader.Text, "Benefit", vbTextCompare) > 0 Then
      dt = TYPE_LONG
    ElseIf InStr(1, ColumnHeader.Text, "Value", vbTextCompare) > 0 Then
      dt = TYPE_LONG
    ElseIf InStr(1, ColumnHeader.Text, "Expense", vbTextCompare) > 0 Then
      dt = TYPE_LONG
    ElseIf InStr(1, ColumnHeader.Text, "Date", vbTextCompare) > 0 Then
      dt = TYPE_DATE
    Else
      dt = TYPE_STR
    End If
  End If
  
  'only sort the employee reference if explicitly set so
  If lv.Parent Is F_Employees Then
    If (Not p11d32.SortEmployeeReferenceAsNumber) And ColumnHeader.Index = L_LV_COL_INDEX_EMPLOYEE_REFERENCE Then
      bForceStringSort = True
    End If
  End If
    
  With lv
    Call LockWindowUpdate(.hwnd)
    lv.SortKey = ColumnHeader.Index - 1
    If (dt = TYPE_STR) Or Not p11d32.DataTypeListViewSorting Or bForceStringSort Then
      If OverrideSortOrder = -1 Then
        If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
        Else
          .SortOrder = lvwAscending
        End If
      Else
        .SortOrder = OverrideSortOrder
      End If
      .Sorted = True
    Else
      Call ListViewSortByType(dt, lv, ColumnHeader, OverrideSortOrder)
    End If
  End With
  If lv.listitems.Count > 0 Then
    If Not lv.SelectedItem Is Nothing Then lv.SelectedItem.EnsureVisible
  End If
  
SetSortOrder_END:
  Call LockWindowUpdate(0&)
  DoEvents
  Exit Sub
SetSortOrder_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetSortOrder", "Set Sort Order", "Error setting the sort order for a list view.")
  Resume SetSortOrder_END
  Resume
End Sub
Private Function ListViewNumberFormat(sText As String) As String
  Dim d As Double
  Const S_NUMBER_FORMAT As String = ""
  Dim strFormat As String
  
  strFormat = String(30, "0") & "." & String(30, "0")
  
  sText = Replace$(sText, "£", "")
  If IsNumeric(sText) Then
    d = CDbl(sText)
    sText = Format(d, strFormat)
    If d < 0 Then
      sText = "&" & InvNumber(sText)
    End If
  Else
    sText = ""
  End If
  ListViewNumberFormat = sText
End Function
Private Function ListViewDateFormat(ByVal sText As String) As String
  Dim dt As Date
  Const S_FORMAT As String = "YYYYMMDD"
  
  dt = TryConvertDateDMY(sText, UNDATED)
  If dt <> UNDATED Then
    sText = Format(dt, S_FORMAT)
  Else
    sText = ""
  End If
  ListViewDateFormat = sText

End Function
Public Sub ListViewSortByType(ByVal dt As DATABASE_FIELD_TYPES, ByVal lv As ListView, ColumnHeader As ColumnHeader, ByVal lsoOverride As ListSortOrderConstants)
  Dim sText As String, s() As String
  
  Dim i As Long, iCount As Long
  
  Dim LVI As ListItem
  Dim lsu As ListSubItem
  Dim iColumnIndex As Long
  
On Error GoTo err_err

  iColumnIndex = ColumnHeader.Index - 1
  
  iCount = lv.listitems.Count
  For i = 1 To iCount
    Set LVI = lv.listitems(i)
    If iColumnIndex > 0 Then
      Set lsu = LVI.ListSubItems(iColumnIndex)
      sText = lsu.Text
    Else
      sText = LVI.Text
    End If
    
    
    LVI.Tag = sText & Chr$(0) & LVI.Tag
    
    If dt = TYPE_LONG Or dt = TYPE_DOUBLE Then
      sText = ListViewNumberFormat(sText)
    ElseIf dt = TYPE_DATE Then
      sText = ListViewDateFormat(sText)
    End If
    If (iColumnIndex > 0) Then
      lsu.Text = sText
    Else
      LVI.Text = sText
    End If
  Next

  If (lsoOverride = -1) Then
    lv.SortOrder = (lv.SortOrder + 1) Mod 2
  Else
    lv.SortOrder = lsoOverride
  End If
  
  lv.SortKey = iColumnIndex
  lv.Sorted = True

  For i = 1 To iCount
    Set LVI = lv.listitems(i)
    s = Split(LVI.Tag, Chr(0))
    If (iColumnIndex > 0) Then
      LVI.ListSubItems(iColumnIndex).Text = s(0)
    Else
      LVI.Text = s(0)
    End If
    
    LVI.Tag = s(1)
  Next
  
  
err_end:
  Exit Sub
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "ListViewSorter"), Err.Description)
  Resume
End Sub
Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function
Public Function DateTimeSQLLocal(ByVal vDateTime As Variant) As String
  If IsNull(vDateTime) Then
    DateTimeSQLLocal = "Null"
  Else
    DateTimeSQLLocal = "'" & Format$(vDateTime, "YYYY-MM-DD HH:NN:SS") & "'"
  End If
End Function

  
Public Sub ColumnWidths(lv As ListView, ParamArray P())
  Dim i As Long, j As Long, k As Long
  On Error GoTo ColumnWidths_Err
  Call xSet("ColumnWidths")
  
  i = LBound(P)
  For j = 1 To (lv.ColumnHeaders.Count)
    If k >= 100 Then P(i) = 0
    If P(i) <> 0 Then
      If k + P(i) >= 100 Then
        P(i) = (100 - k) - 1
        k = 100
      Else
       k = k + P(i)
      End If
    End If
    lv.ColumnHeaders(j).width = CLng(CDbl(P(i)) / 100 * CDbl(lv.width))
    i = i + 1
    If j + 1 = lv.ColumnHeaders.Count Then
      If k <= 100 Then lv.ColumnHeaders(j + 1).width = ((100 - k) / 100) * CDbl(lv.width) * 0.98
      Exit For
    End If
  Next j
  
  
  
ColumnWidths_End:
  Call xReturn("ColumnWidths")
  Exit Sub

ColumnWidths_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ColumnWidths", "Column Widths", "Error setting the column widths of a list view.")
  Resume ColumnWidths_End
  Resume
End Sub

Public Sub SelectItems(lv As ListView, ByVal Sm As SELECT_MODE, Optional sGroupCodeToSelect As String = "")
  Dim i As Long
  Dim lst As ListItem
  Dim ben As IBenefitClass
  Dim benEmployer As IBenefitClass
  Dim charMin As Long, charMax As Long
  Dim sSurname As String
  Dim c As Long
  
  On Error GoTo SelectItems_Err
  
  Call xSet("SelectItems")
  
  Set benEmployer = p11d32.CurrentEmployer
  
  If (Sm >= SELECT_ALPHABETICALLY_START And (Sm <= SELECT_ALPHABETICALLY_END)) Then
    If MsgBox("Do you wish to clear the current selected employees", vbYesNo, "Clear selections") = vbYes Then
      Call SelectItems(lv, SELECT_NONE)
    End If
  End If
  
  
  With p11d32.CurrentEmployer
    If Sm = SELECT_BY_REPORT Then
      Call p11d32.Help.ShowForm(F_SelectEmployeesByReport, vbModal)
      GoTo SelectItems_End:
    End If
    For i = 1 To lv.listitems.Count
      Set lst = lv.listitems(i)
      Set ben = .employees(lst.Tag)
      Select Case Sm
        Case Is <= SELECT_ALPHABETICALLY_END
           charMin = 65 + ((Sm - 1) * 3)
           If Sm = SELECT_ALPHABETICALLY_Y_Z Then
            charMax = charMin + 1
           Else
            charMax = charMin + 2
           End If
           sSurname = ben.value(ee_Surname_db)
           sSurname = Trim$(UCASE$(sSurname))
           If Len(sSurname) = 0 Then
             If Sm = SELECT_ALPHABETICALLY_Y_Z Then
               lst.Checked = True
               ben.value(ee_Selected) = True
             End If
           Else
             c = Asc(Left$(sSurname, 1))
             If (c >= charMin) And (c <= charMax) Then
               lst.Checked = True
               ben.value(ee_Selected) = True
             ElseIf ((c < 65) Or (c > (64 + 26))) And Sm = SELECT_ALPHABETICALLY_Y_Z Then
               lst.Checked = True
               ben.value(ee_Selected) = True
             End If
           End If
        Case SELECT_ALL
          If lst.Checked = False Then
            lst.Checked = True
            ben.value(ee_Selected) = True
            benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) + 1
          End If
        Case SELECT_NONE
          If lst.Checked = True Then
            benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) - 1
            lst.Checked = False
            ben.value(ee_Selected) = False
          End If
        Case SELECT_REVERSE
          If lst.Checked = True Then
            benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) - 1
            lst.Checked = False
            ben.value(ee_Selected) = False
          Else
            lst.Checked = True
            ben.value(ee_Selected) = True
            benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) + 1
          End If
        Case SELECT_GROUP_1
          Call SelectByGroup(lst, LV_EE_GROUP1, sGroupCodeToSelect, benEmployer, ben)
        Case SELECT_GROUP_2
          Call SelectByGroup(lst, LV_EE_GROUP2, sGroupCodeToSelect, benEmployer, ben)
        Case SELECT_GROUP_3
          Call SelectByGroup(lst, LV_EE_GROUP3, sGroupCodeToSelect, benEmployer, ben)
        Case SELECT_CURRENT_EMPLOYED
          Call SelectByProperty(lst, ben, benEmployer, ee_left_db, UNDATED, False)
        Case SELECT_LEFT
          Call SelectByProperty(lst, ben, benEmployer, ee_left_db, UNDATED, True)
        Case SELECT_NO_EMAIL
          Call SelectByProperty(lst, ben, benEmployer, ee_Email_db, "", False)
        Case SELECT_EMAIL
          Call SelectByProperty(lst, ben, benEmployer, ee_Email_db, "", True)
        Case Else
          ECASE "Unknown selection mode"
          GoTo SelectItems_End
      End Select
    Next i
  End With
  
  
  
  
SelectItems_End:
  If Not p11d32.CurrentEmployer Is Nothing Then Call p11d32.CurrentEmployer.SelectedPanel
  Call xReturn("SelectItems")
  Exit Sub

SelectItems_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SelectItems", "Select List Items", "Unable to complete selection.")
  Resume SelectItems_End
  Resume
End Sub
Private Sub SelectByGroup(lst As ListItem, ByVal lSubItemIndex As Long, sGroupCodeToSelect As String, benEmployer As IBenefitClass, benEmployee As IBenefitClass)
   If StrComp(lst.ListSubItems(lSubItemIndex), sGroupCodeToSelect) = 0 Then
    If Not lst.Checked Then
      benEmployee.value(ee_Selected) = True
      lst.Checked = True
      benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) + 1
    End If
   Else
    If lst.Checked Then
      benEmployee.value(ee_Selected) = False
      lst.Checked = False
      benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) - 1
    End If
   End If
End Sub
Private Sub SelectByProperty(lst As ListItem, benEmployee As IBenefitClass, benEmployer As IBenefitClass, ei As EmployeeItems, vPropertyMatch As Variant, bNotMatch As Boolean)
  Dim b As Boolean
  
  If (bNotMatch) Then
    b = benEmployee.value(ei) <> vPropertyMatch
  Else
    b = benEmployee.value(ei) = vPropertyMatch
  End If
   If b Then
    If Not lst.Checked Then
      benEmployee.value(ee_Selected) = True
      lst.Checked = True
      benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) + 1
    End If
   Else
    If lst.Checked Then
      benEmployee.value(ee_Selected) = False
      lst.Checked = False
      benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) - 1
    End If
   End If
End Sub

Public Sub ClearEdit(rs As Recordset)
  If Not rs Is Nothing Then
    If rs.EditMode <> dbEditNone Then rs.CancelUpdate
  End If
End Sub

Public Function GrisIsTooLong(ValidateMessage As String, RowBuf As RowBuffer, ByVal RowBufRowIndex, ByVal lCol As Long, Optional ByVal lMaxChars As Long = 50) As Boolean
  On Error GoTo GrisIsTooLong_ERR
  
  Call xSet("GrisIsTooLong")
  
  If Not IsNull(RowBuf.value(RowBufRowIndex, lCol)) Then
    If Len(RowBuf.value(RowBufRowIndex, lCol)) > lMaxChars + 1 Then
      ValidateMessage = "Data must be less than " & lMaxChars + 1 & " characters."
      GrisIsTooLong = True
    End If
  End If
    
GrisIsTooLong_END:
  Call xReturn("GrisIsTooLong")
  Exit Function
GrisIsTooLong_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GrisIsTooLong", "Gris Is Too Long", "Error determining if the buffer for col " & lCol & " is > " & lMaxChars + 1 & ".")
  Resume GrisIsTooLong_END
End Function
Public Function GridIsNotDate(ValidateMessage As String, vDate As Variant, ByVal ObjectListIndex As Long, Optional ByVal bTaxDate As Boolean) As Boolean
  Dim d0 As Date
  '// returns the column index that is invalid
  On Error GoTo GridIsNotTaxDate_Err
  Call xSet("GridIsNotTaxDate")

  If (Not IsNull(vDate)) Or (ObjectListIndex = -1) Then
    'ie it has not changed
    d0 = TryConvertDate(vDate)
    If d0 <> UNDATED Then
      If bTaxDate Then
        GridIsNotDate = Not ((d0 <= p11d32.Rates.value(TaxYearEnd)) And (d0 >= p11d32.Rates.value(TaxYearStart)))
        If GridIsNotDate Then
          ValidateMessage = "The date is not inside the tax year, (" & p11d32.Rates.value(TaxYearStart) & " - " & p11d32.Rates.value(TaxYearEnd) & ")."
        End If
      End If
    Else
      GridIsNotDate = True
      ValidateMessage = "The value is not a valid date."
    End If
  End If
  
GridIsNotTaxDate_End:
  Call xReturn("GridIsNotTaxDate")
  Exit Function
  
GridIsNotTaxDate_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GridIsNotTaxDate", "Grid Is Not Tax Date", "Error checking for a date.")
  Resume GridIsNotTaxDate_End
End Function

Public Function GridIsNotNumericOrLong(ValidateMessage As String, vNumber As Variant, ByVal ObejctListIndex As Long) As Boolean
  '// returns the column index that is invalid
  On Error GoTo GridIsNotNumericOrLong_Err
  
  Call xSet("GridIsNotNumericOrLong")

  If (Not IsNull(vNumber)) Or (ObejctListIndex = -1) Then
    'ie it has not changed
    If Not IsNumeric(vNumber) Then
      
      ValidateMessage = "The value is not a number."
      GridIsNotNumericOrLong = True
    Else
      If vNumber > L_MAX_LONG Then
        ValidateMessage = "The value is greater than " & CStr(L_MAX_LONG) & "."
        GridIsNotNumericOrLong = True
      End If
    End If
  End If
  
GridIsNotNumericOrLong_End:
  Call xReturn("GridIsNotNumericOrLong")
  Exit Function
GridIsNotNumericOrLong_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GridIsNotNumericOrLong", "GridIsNotNumericOrLong", "Error checking for a Numeric Value.")
  Resume GridIsNotNumericOrLong_End
End Function
Public Function GridIsZeroLength(ValidateMessage As String, v As Variant, ObejctListIndex As Long) As Boolean
  '// returns the column index that is invalid
  On Error GoTo GridIsZeroLength_Err
  
  Call xSet("GridIsZeroLength")

  If (Not IsNull(v)) Or ObejctListIndex = -1 Then
    'ie it has not changed
    If Len(CStr(v)) = 0 Then
      GridIsZeroLength = True
      ValidateMessage = "Zero length strings are not allowed."
    End If
  End If
  
GridIsZeroLength_End:
  Call xReturn("GridIsZeroLength")
  Exit Function
GridIsZeroLength_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GridIsZeroLength", "Grid Is Zero Length", "Undefined error.")
  Resume GridIsZeroLength_End
  Resume
End Function
Public Function ClearCollection(c As Collection)
  Dim i As Long
  For i = 1 To c.Count
    Call c.Remove(1)
  Next i
End Function

Public Sub ShowMaximized(frmShow As Form, frmHide As Form, Optional dt As DisplayType = [_INVALID_DISPPLAY_TYPE])
  
  On Error Resume Next
  Call LockWindowUpdate(MDIMain.hwnd)
  If Not (frmHide Is frmShow) And Not (frmHide Is Nothing) Then frmHide.Hide
  If frmShow Is Nothing Then Call ECASE("ShowMaximized - cannot show empty form")
  If dt > 0 Then Call DisplayEx(dt)
'  frmShow.Show
  Call p11d32.Help.ShowForm(frmShow)
  frmShow.WindowState = 2
  Call LockWindowUpdate(0)
  DoEvents
End Sub

Public Function dGetDateFactor(TotalDaysUnavailable As Long, ByVal dtFrom As Date, ByVal dtTo As Date, Optional ByVal ExcludeDays As Long = 0, Optional ByVal MinExcludeDays As Long = 0, Optional bForceLeapOpposite As Boolean = False, Optional outputDaysUsed As Long = 0) As Double
  Dim lDaysInYear As Long
  
  On Error GoTo dGetDateFactor_Err
  Call xSet("dGetDateFactor")
  lDaysInYear = IIf(bForceLeapOpposite, p11d32.Rates.value(DaysInYearLeap), p11d32.Rates.value(DaysInYear))
  outputDaysUsed = lDaysInYear - (DateDiff("d", p11d32.Rates.value(TaxYearStart), dtFrom) + DateDiff("d", dtTo, p11d32.Rates.value(TaxYearEnd)))
  If ExcludeDays >= MinExcludeDays Then outputDaysUsed = Max(outputDaysUsed - ExcludeDays, 0)
  TotalDaysUnavailable = lDaysInYear - outputDaysUsed
  dGetDateFactor = CDbl(outputDaysUsed) / lDaysInYear
    
  
  If dGetDateFactor > 1 Then Call Err.Raise(ERR_DAYS_INCONSISTENT, "dGetDateFactor", "The function has been passed inconsistent date data")
  If dGetDateFactor < 0 Then Call Err.Raise(ERR_DAYS_INCONSISTENT, "dGetDateFactor", "The function has been passed inconsistent date data")
  
dGetDateFactor_End:
  Call xReturn("dGetDateFactor")
  Exit Function
  
dGetDateFactor_Err:
  Resume dGetDateFactor_End
End Function


Public Function WriteDateItem(rs As Recordset, sField As String, vNewValue As Variant) As Boolean

  On Error GoTo WriteDateItem_Err
  Call xSet("WriteDateItem")

  If Len(vNewValue) Then
    rs.Fields(sField) = vNewValue
    WriteDateItem = True
  End If

WriteDateItem_End:
  Call xReturn("WriteDateItem")
  Exit Function

WriteDateItem_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteDateItem", "Write Date Item", "Error writing a date item to the database.")
  Resume WriteDateItem_End
  Resume
End Function
Public Function GetBenItem(ben As IBenefitClass, lItem As Long) As Variant
  If Not ben Is Nothing Then
    GetBenItem = ben.value(lItem)
    If p11d32.BenDataLinkDataType(ben.BenefitClass, lItem) = TYPE_DATE Then
      If GetBenItem = UNDATED Then GetBenItem = ""
    End If
  Else
    GetBenItem = ""
  End If
End Function
Public Function GetBenItemFWNRPT(ben As IBenefitClass, lItem As Long, Optional sCurrency As String = "£", Optional bNegative As Boolean = False) As Variant
  Dim dt As DATABASE_FIELD_TYPES
  
  If Not ben Is Nothing Then
    dt = p11d32.BenDataLinkDataType(ben.BenefitClass, lItem)
    If dt = TYPE_LONG Or dt = TYPE_DOUBLE Then
      GetBenItemFWNRPT = FormatWNRPT(ben.value(lItem), sCurrency, bNegative)
    Else
      GetBenItemFWNRPT = ben.value(lItem)
    End If
  Else
    GetBenItemFWNRPT = ""
  End If
End Function
Public Function CreateDefaultDirectory(sDefaultDirPath As String, sDefaultDirType As String, Optional bAskQuestion As Boolean = True) As Boolean

  On Error GoTo CreateDefaultDirectory_Err
  Call xSet("CreateDefaultDirectory")
  Dim sDirType_Name As String
  
  If (bAskQuestion) Then
  
    If (MsgBox("A sub-directory does not exist in your working directory" & _
                " for " & sDefaultDirType & "." & vbCrLf & _
                "Do you want the default directory: " & sDefaultDirPath & " created?", vbOKCancel, "Missing directory") = vbOK) Then
        xMkdir (sDefaultDirPath)
      
    End If
  Else
    xMkdir (sDefaultDirPath)
  End If
  CreateDefaultDirectory = True
CreateDefaultDirectory_End:
  Call xReturn("CreateDefaultDirectory")
  Exit Function

CreateDefaultDirectory_Err:
  CreateDefaultDirectory = False
  Call Err.Raise(ERR_DIRECTORY_CREATE, "CreateDefaultDirectory", "Could not create the directory " & sDefaultDirPath & " check rights to directory.")
  'Call ErrorMessage(ERR_ERROR, Err, "CreateDefaultDirectory", "Error in CreateDefaultDirectory", "Undefined error.")
  Resume CreateDefaultDirectory_End
  Resume
End Function

Public Sub RemoveReadOnlyFolder(sFolder As String, Optional ByVal bNoErrors = False)
  Dim fso As FileSystemObject
  Dim f As folder
  
On Error GoTo err_err

  Set fso = New FileSystemObject
  Set f = fso.GetFolder(sFolder)
  If ((1 And f.Attributes) = 1) Then 'remove the read only if present
    f.Attributes = f.Attributes - 1
  End If
  
err_end:
  Exit Sub
err_err:
  If bNoErrors Then Resume err_end
  Call Err.Raise(ERR_DIRECTORY_CREATE, "RemoveReadOnlyFolder", "Failed to remove the read only attribute for the folder: " & sFolder)

End Sub

Public Sub RemoveReadOnlyFile(s As String)
  Dim fso As FileSystemObject
  Dim f As File
  
On Error GoTo err_err

  Set fso = New FileSystemObject
  Set f = fso.GetFile(s)
  If ((1 And f.Attributes) = 1) Then 'remove the read only if present
    f.Attributes = f.Attributes - 1
  End If
  
err_end:
  Exit Sub
err_err:
  Call Err.Raise(ERR_DIRECTORY_CREATE, "RemoveReadOnlyFile", "Failed to remove the read only attribute for the file: " & s)
  Resume
End Sub

Public Sub MkDirEx(ByVal sPath As String)

On Error GoTo err_err

  If (Not xMkdir(sPath)) Then
    Call Err.Raise(ERR_DIRECTORY_CREATE, "MKDirEx", "Failed to create the folder " & sPath)
  End If
    
'
'
'  Dim v As Variant
'  Dim sPathNew As String, s As String
'  Dim i As Long
'
'
'
'  sPathNew = ""
'  sPath = Trim$(sPath)
'  'deal with UNC paths
'  If (Len(sPath) > 2) Then
'    If (Left$(sPath, 2) = "\\") Then
'      'first is always the server so ignore
'      i = InStr(3, sPath, "\")
'      sPathNew = Left$(sPath, i)
'      sPath = Mid$(sPath, i + 1)
'    End If
'
'  End If
'
'
'  For i = 1 To GetDelimitedValues(v, sPath, True, True, "\")
'    s = Trim$(v(i))
'    If (Len(s) > 0) Then
'      sPathNew = FullPath(sPathNew & s)
'      If (Not FileExists(sPathNew, True)) Then
'
'        Call MkDir(sPathNew)
'      End If
'    End If
'  Next
'
err_end:
  Exit Sub
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "MkDirEx"), Err.Description)
  Resume
End Sub

Public Sub ChDriveUNC(sCDir As String)
  On Error Resume Next
  Call ChDrive(sCDir)

End Sub
'deals with negative numbers
Public Function RoundDownEx(ByVal d As Double, Optional ByVal DecimalPlaces As Long = 2)
  If (d < 0) Then
    d = d * -1
    d = RoundUp(d, DecimalPlaces) 'make value lower for negative numbers
    d = d * -1
  Else
    d = RoundDown(d, DecimalPlaces)
  End If
  RoundDownEx = d
End Function



Public Function GetSpecialFolderEx(spt As CSIDLConstants) As String
  Const MAX_PATH = 260
  Const S_OK = 0

   Dim sPath As String
   Dim pidl As Long
   
  'fill the idl structure with the specified folder item
   If SHGetSpecialFolderLocation(0, spt, pidl) = S_OK Then
     
     'if the pidl is returned, initialize
     'and get the path from the id list
      sPath = Space$(MAX_PATH)
      
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
        'return the path
         GetSpecialFolderEx = FullPath(Left(sPath, InStr(sPath, Chr$(0)) - 1))
      End If
    
     'free the pidl
      Call CoTaskMemFree(pidl)

    End If
   
End Function

Public Function PayeOnlineNameSpaceSearchString(ByVal Namespace As PayeOnlneNameSpace)
 Select Case Namespace
    Case PayeOnlneNameSpace.Errors
      PayeOnlineNameSpaceSearchString = S_PAYE_ONLINE_NAMESPACE_EFILER_ERROR
    Case PayeOnlneNameSpace.GoTalk
      PayeOnlineNameSpaceSearchString = S_PAYE_ONLINE_NAMESPACE_EFILER_GOVTALK
    Case Else
      Call Err.Raise(ERR_ERROR, "NameSpaceSearchStringEFiler", "Invalid namsspace search enum")
  End Select
End Function
Public Function PayeOnlineSearchString(ByVal search As String, Optional ByVal Namespace As PayeOnlneNameSpace = PayeOnlneNameSpace.GoTalk) As String
  Dim sReplace As String
  Dim searchNew As String
  Dim iStart As Long
  
  search = Trim(search)
  If (Len(search) > 0) Then
    If (Mid(search, 1, 1) <> "/") Then
      search = "/" + search
    End If
  End If
  
  sReplace = PayeOnlineNameSpaceSearchString(Namespace)
  sReplace = "/" & sReplace & ":"
  If (InStr(1, search, "//", vbBinaryCompare) = 1) Then
    iStart = 2
  Else
    iStart = 1
  End If
  
  
  
  searchNew = Replace(search, "/", sReplace, iStart)
  If (iStart > 1) Then
    searchNew = Mid(search, 1, 1) & searchNew
  End If
  
  PayeOnlineSearchString = searchNew
End Function



Public Function DOMDocumentNewEFiler() As DOMDocument60
  Dim xmlDoc As DOMDocument60
  Set xmlDoc = New DOMDocument60
  xmlDoc.setProperty "SelectionNamespaces", "xmlns:" & S_PAYE_ONLINE_NAMESPACE_EFILER_GOVTALK & "='http://www.govtalk.gov.uk/CM/envelope' xmlns:" & S_PAYE_ONLINE_NAMESPACE_EFILER_ERROR & "='http://www.govtalk.gov.uk/CM/errorresponse'"
  Set DOMDocumentNewEFiler = xmlDoc
End Function

Public Sub XMLError(ByRef xml As DOMDocument60, ByVal sDOCName As String, Optional sSrcXML As String = "")
  Dim vLines As Variant
  Dim sDescription As String
  Dim sLine As String
  Dim sXML As String
  Dim p0 As Long
  Dim iLen As Long
  Call xml.Validate
  If (Len(xml.parseError.reason) > 0) Then
    sXML = xml.xml
    If Len(sXML) = 0 Then sXML = sSrcXML
    sDescription = "Error in XML document: '" & sDOCName & "'" & vbCrLf
    sDescription = sDescription & xml.parseError.reason & " at line " & xml.parseError.line & ", position " & xml.parseError.linepos & " (see *)" & vbCrLf & vbCrLf
    p0 = xml.parseError.linepos - 30
    iLen = 30
    If p0 < 1 Then
      iLen = 30 + p0
      p0 = 1
    End If
    If (Len(sXML) > 0) Then
      vLines = Split(sXML, vbCrLf)
      sLine = vLines(xml.parseError.line - 1)
      sDescription = sDescription & Mid$(sLine, p0, iLen) & "*" & Mid$(sLine, xml.parseError.linepos, 30)
    End If
    
    
    
    Call Err.Raise(ERR_XMLNUMBERTOOBIG, "Submit_EfilerCom", sDescription)
  End If
End Sub


Public Function GetPropertyFromString(ByVal sSearchString As String, ByVal sProperty) As String
  Dim p0 As Integer
  Dim p1 As Integer
  
  On Error GoTo GetPropertyFromString_End
  sProperty = LCase$(sProperty)
  
  sProperty = sProperty & S_STRING_PROPERTY_OPEN
  p0 = InStr(sSearchString, sProperty)
  If p0 > 0 Then
    p1 = InStr(p0, sSearchString, S_STRING_PROPERTY_CLOSE)
    GetPropertyFromString = Mid(sSearchString, p0 + Len(sProperty), p1 - p0 - Len(sProperty))
  Else
    GetPropertyFromString = ""
  End If
  
  
GetPropertyFromString_End:
  Exit Function
GetPropertyFromString_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "MkDirEx"), Err.Description)
  Resume GetPropertyFromString_End

End Function

Public Function SetPropertiesFromString(ByVal sSearchString As String, ParamArray P()) As String
  Dim iLB As Long, iUB As Long, i As Long
  Dim s As String
  On Error GoTo err_err
  
  s = sSearchString
    iLB = LBound(P)
    iUB = UBound(P)
    For i = iLB To iUB Step 2
      s = SetPropertyFromString(s, P(i), IsNullEx(P(i + 1), ""))
    Next
  SetPropertiesFromString = s
  
err_end:
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "SetPropertiesFromString"), Err.Description)
  Resume
End Function

Public Function SetPropertyFromString(ByVal sSearchString As String, ByVal sProperty, ByVal sNewValue As String) As String
  Dim p0 As Integer
  Dim p1 As Integer
  
  On Error GoTo err_err
  sProperty = LCase$(sProperty)
  sProperty = sProperty & S_STRING_PROPERTY_OPEN
  p0 = InStr(sSearchString, sProperty)
  If p0 > 0 Then
    p1 = InStr(p0, sSearchString, S_STRING_PROPERTY_CLOSE)
    SetPropertyFromString = Left$(sSearchString, p0 + Len(sProperty) - 1)
    SetPropertyFromString = SetPropertyFromString & sNewValue & Mid$(sSearchString, p1)
  Else
    SetPropertyFromString = sSearchString & sProperty & sNewValue & S_STRING_PROPERTY_CLOSE
  End If
  
err_end:
  Exit Function
err_err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "MkDirEx"), Err.Description)
  Resume err_end

End Function
Public Function ReporterNew() As Reporter
  Dim rep As Reporter
  
  Set rep = New Reporter
  rep.A4Force = p11d32.ReportPrint.A4ForcePrint
  Set ReporterNew = rep
End Function
Public Function ReportWizardNew() As ReportWizard
  Dim repw As ReportWizard
  
  
  Set repw = New ReportWizard
  repw.A4Force = p11d32.ReportPrint.A4ForcePrint
  Set ReportWizardNew = repw
End Function

Public Function GetEnvironmentInfo(MyMail As Mail) As String
Dim sOutput As String

'ATC Mail environment information
  sOutput = sOutput & "MailApplication: " & MyMail.MailApplication & vbCrLf
  sOutput = sOutput & "MailSystem: " & MyMail.MailSystem & vbCrLf

'CORE environment information, copied from Public Function UpdateSys(frmsys As frmMailTest) As Boolean
  Dim l0 As Long, l1 As Long, l2 As Long, l3 As Long
  Dim pid As OS_TYPE ', locInfo As LocaleInfo
  Dim s0 As String, s1 As String, s2 As String, ret As Boolean
  Dim d0 As Double, d1 As Double, d2 As Double
  
  On Error GoTo GetEnvironmentInfo_err
  Call GetSysInfo(s0)
  sOutput = sOutput & "Processor type: " & s0 & vbCrLf
  'frmsys.lblSysInfo(0).Caption = s0
  If GetWindowsVersion(l0, l1, l2, pid, s0) Then
    s1 = l0 & "." & l1 & "." & l2
    If Len(s0) > 0 Then s1 = s1 & " (" & s0 & ")"
    'sOutput = sOutput & s1 & vbCrLf
    'frmsys.lblSysInfo(1).Caption = s1
    Select Case pid
      Case OS_NT4
        s2 = "Microsoft Windows NT"
      Case OS_WIN95
        s2 = "Microsoft Windows 95"
      Case OS_WIN98
        s2 = "Microsoft Windows 98"
      Case OS_W2000
        s2 = "Microsoft Windows 2000"
      Case Else
        s2 = "Unknown OS"
    End Select
    sOutput = sOutput & s2 & " version: " & s1 & vbCrLf
    'frmsys.lblInformation(1) = s0
  End If
  l0 = GetPhysicalMemory(d0, d1, MEGABYTES)
  sOutput = sOutput & "Total physical memory available: " & Format$(d0, "#,###0.00 Mb ") & vbCrLf
  sOutput = sOutput & "Free physical memory available: " & Format$(d1, "#,###0.00 Mb ") & vbCrLf
  sOutput = sOutput & "Overall memory usage: " & CStr(l0) & "% "
  'frmsys.lblSysInfo(2).Caption = Format$(d0, "#,###0.00 Mb ")
  'frmsys.lblSysInfo(3).Caption = Format$(d1, "#,###0.00 Mb ")
  'frmsys.lblSysInfo(8).Caption = CStr(l0) & "% "
'  s0 = UCase$(Left$(mHomeDirectory, 3))
'  ret = GetDiskSpaceEx(s0, d0, d1, d2, MEGABYTES)
''  frmsys.lblInformation(5).Visible = True
''  frmsys.lblSysInfo(4).Visible = True
''  frmsys.lblSysInfo(6).Visible = True
'  If ret Then
'    sOutput = sOutput & "Application drive " & s0 & vbCrLf
'    sOutput = sOutput & Format$(d0, "#,###0.00 Mb ") & vbCrLf
'    sOutput = sOutput & Format$(d1, "#,###0.00 Mb ") & vbCrLf
''    frmsys.lblInformation(5).Caption = "Application drive " & s0
''    frmsys.lblSysInfo(4).Caption = Format$(d0, "#,###0.00 Mb ")
''    frmsys.lblSysInfo(6).Caption = Format$(d1, "#,###0.00 Mb ")
'  Else
''    sOutput = sOutput & "Application drive " & s0 & vbCrLf
''    sOutput = sOutput & "Unavailable" & vbCrLf
''    sOutput = sOutput & "Unavailable" & vbCrLf
''    frmsys.lblInformation(5).Caption = "Application drive " & s0
''    frmsys.lblSysInfo(4).Caption = "Unavailable"
''    frmsys.lblSysInfo(6).Caption = "Unavailable"
'  End If
'  s1 = UCase$(Left$(CurDir$, 3))
'  ret = ret And GetDiskSpaceEx(s1, d0, d1, d2, MEGABYTES)
'  If StrComp(s1, s0, vbTextCompare) <> 0 And ret Then
''    frmsys.lblInformation(6).Visible = True
''    frmsys.lblStatic(1).Visible = True
''    frmsys.lblStatic(3).Visible = True
''    frmsys.lblSysInfo(5).Visible = True
''    frmsys.lblSysInfo(7).Visible = True
'    sOutput = sOutput & "Current drive " & s0 & vbCrLf
'    sOutput = sOutput & Format$(d0, "#,###0.00 Mb ") & vbCrLf
'    sOutput = sOutput & Format$(d1, "#,###0.00 Mb ") & vbCrLf
''    frmsys.lblInformation(6).Caption = "Current drive " & s0
''    frmsys.lblSysInfo(5).Caption = Format$(d0, "#,###0.00 Mb ")
''    frmsys.lblSysInfo(7).Caption = Format$(d1, "#,###0.00 Mb ")
'  Else
''    frmsys.lblInformation(6).Visible = False
''    frmsys.lblStatic(1).Visible = False
''    frmsys.lblStatic(3).Visible = False
''    frmsys.lblSysInfo(5).Visible = False
''    frmsys.lblSysInfo(7).Visible = False
'  End If
'  Set locInfo = New LocaleInfo
'  s0 = "System Locale ID " & locInfo.GetSystemDefaultLcid & vbCrLf
'  s0 = s0 & "Country " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SENGCOUNTRY) & vbCrLf
'  s0 = s0 & "Language " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SENGLANGUAGE) & vbCrLf
'  s0 = s0 & "Currency " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SCURRENCY) & " (" & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL) & ")" & vbCrLf
'  s0 = s0 & "Short Date " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SSHORTDATE) & vbCrLf
'  s0 = s0 & "Long Date " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SLONGDATE) & vbCrLf
'  frmsys.lblLocaleSys.Caption = s0
'
'  s0 = "User Locale ID " & locInfo.GetUserDefaultLcid & vbCrLf
'  s0 = s0 & "Country " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY) & vbCrLf
'  s0 = s0 & "Language " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE) & vbCrLf
'  s0 = s0 & "Currency " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY) & " (" & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL) & ")" & vbCrLf
'  s0 = s0 & "Short Date " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE) & vbCrLf
'  s0 = s0 & "Long Date " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE) & vbCrLf
'  frmsys.lblLocalUser.Caption = s0
  GetEnvironmentInfo = sOutput
  
GetEnvironmentInfo_exit:
  Exit Function
GetEnvironmentInfo_err:
  GetEnvironmentInfo = "Error retrieving environment information"
  Resume GetEnvironmentInfo_exit
End Function

Public Sub GetSysInfo(sProcessor As String)
  Dim lpSysInfo As SYSTEM_INFO
  
  Call GetSystemInfo(lpSysInfo)
  Select Case lpSysInfo.dwProcessorType
    Case PROCESSOR_INTEL_386
      sProcessor = "Intel 386"
    Case PROCESSOR_INTEL_486
      sProcessor = "Intel 486"
    Case PROCESSOR_INTEL_PENTIUM
      sProcessor = "Intel Pentium"
    Case Else
      sProcessor = "Information unavailable"
  End Select
End Sub

Public Function GetWindowsVersion(lMajorVer As Long, lMinorVer As Long, lBuild As Long, PlatformID As OS_TYPE, sCSDVersion As String) As Boolean
  Dim lpVI As OSVERSIONINFO
  
  lpVI.dwOSVersionInfoSize = Len(lpVI)  ' 148
  GetWindowsVersion = GetVersionEx(lpVI) <> 0
  If GetWindowsVersion Then
    lMajorVer = lpVI.dwMajorVersion
    lMinorVer = lpVI.dwMinorVersion
    lBuild = lpVI.dwBuildNumber
    PlatformID = OS_UNKNOWN
    If lpVI.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      If lMajorVer = 4 Then
        PlatformID = OS_NT4
      Else
        PlatformID = OS_W2000
      End If
    ElseIf lpVI.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
      If lMinorVer = 0 Then
        PlatformID = OS_WIN95
      Else
        PlatformID = OS_WIN98
      End If
    End If
    'AM RTrimChar not available
    sCSDVersion = Replace(lpVI.szCSDVersion, vbNullChar, "")
    'sCSDVersion = RTrimChar(lpVI.szCSDVersion, vbNullChar)
  End If
End Function

Public Function GetPhysicalMemory(dTotalPhysical As Double, dFreePhysical As Double, ByVal nMemUnit As MemoryUnit) As Long
  Dim lpBuffer As MEMORYSTATUS
  Dim retval As Long
  
  lpBuffer.dwLength = LenB(lpBuffer)
  Call GlobalMemoryStatus(lpBuffer)
  dTotalPhysical = lpBuffer.dwTotalPhys
  dFreePhysical = lpBuffer.dwAvailPhys
  Select Case nMemUnit
    Case GIGABYTES
      dTotalPhysical = dTotalPhysical / TWO_POW_30
      dFreePhysical = dFreePhysical / TWO_POW_30
    Case MEGABYTES
      dTotalPhysical = dTotalPhysical / TWO_POW_20
      dFreePhysical = dFreePhysical / TWO_POW_20
  End Select
  dTotalPhysical = RoundDouble(dTotalPhysical, 2, R_NORMAL)
  dFreePhysical = RoundDouble(dFreePhysical, 2, R_NORMAL)
  GetPhysicalMemory = lpBuffer.dwMemoryLoad
End Function

Public Function RoundDouble(ByVal Number As Double, ByVal DecimalPlaces As Long, ByVal rType As ROUND_TYPE) As Double
  Dim d As Double
  Dim TenPow As Double
  
  If (DecimalPlaces >= LOW_POW) And (DecimalPlaces <= HIGH_POW) Then
    TenPow = Powers(DecimalPlaces)
  Else
    TenPow = 10 ^ DecimalPlaces
  End If
  
  Select Case rType
    Case R_NORMAL
      RoundDouble = Int((Number * TenPow) + 0.5) / TenPow
    Case R_UP, R_DOWN
      d = Number * TenPow
      If Int(d) <> d Then
        If rType = R_UP Then d = d + 1
        RoundDouble = Int(d) / TenPow
      Else
        RoundDouble = d
      End If
    Case R_BANKERS
      RoundDouble = CLng(Number * TenPow) / TenPow
  End Select
End Function

Public Function SetupOpraInput(lbl As Label, textBox As ValText, Optional ByVal sLabelAddtionalText As String = "")
  
  lbl.Caption = S_UDM_OPRA_AMOUNT_FOREGONE & sLabelAddtionalText
  
  lbl.ToolTipText = S_UDM_OPRA_AMOUNT_FOREGONE_HELP
  textBox.TXTAlign = TXT_RIGHT
  textBox.TypeOfData = VT_LONG
  textBox.ToolTipText = S_UDM_OPRA_AMOUNT_FOREGONE_HELP
End Function


Public Sub SendKeysEx(Text As Variant, Optional Wait As Boolean = False)
  If (IsRunningInIDE()) Then
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.SendKeys CStr(Text), Wait
    Set WshShell = Nothing
  Else
    Call SendKeys(Text, Wait)
  End If
  
  
End Sub
