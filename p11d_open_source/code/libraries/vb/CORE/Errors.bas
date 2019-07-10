Attribute VB_Name = "Errors"
Option Explicit

Private Stk_ErrorNumber() As Long
Private Stk_ErrorDescription() As String
Private Stk_ErrorSource() As String
Private StackMaxTop As Long
Private StackTop As Long

Private IgnoreErrorList() As Long
Private CurIgnoreErrorList As Long
Private MaxIgnoreErrorList As Long

Private Const userdeferror As String = "Application-defined or object-defined error"

Public Sub PushErrorMessage(ErrObj As ErrObject)
  StackTop = StackTop + 1
  If StackTop > StackMaxTop Then
    StackMaxTop = StackMaxTop + 10
    ReDim Preserve Stk_ErrorNumber(1 To StackMaxTop) As Long
    ReDim Preserve Stk_ErrorDescription(1 To StackMaxTop) As String
    ReDim Preserve Stk_ErrorSource(1 To StackMaxTop) As String
  End If
  Stk_ErrorNumber(StackTop) = ErrObj.Number
  Stk_ErrorDescription(StackTop) = ErrObj.Description
  Stk_ErrorSource(StackTop) = ErrObj.Source
End Sub

Public Sub PopErrorMessageErr(ErrObj As ErrObject)
  If StackTop > 0 Then
    Call ErrObj.Clear
    ErrObj.Number = Stk_ErrorNumber(StackTop)
    ErrObj.Description = Stk_ErrorDescription(StackTop)
    ErrObj.Source = Stk_ErrorSource(StackTop)
    StackTop = StackTop - 1
  End If
End Sub

Public Function ErrorSourceEx(ErrObj As ErrObject, FunctionName As String) As String
  Dim Source As String

  If Not ErrObj Is Nothing Then Source = ErrObj.Source
  If Len(Source) = 0 Then
    ErrorSourceEx = FunctionName
  Else
    If InStr(1, Source, FunctionName, vbTextCompare) <> 1 Then
      ErrorSourceEx = FunctionName & ";" & Source
    Else
      ErrorSourceEx = Source
    End If
  End If
End Function

Public Function ErrorMessageEx(ByVal ErrType As errnumbersenum, ErrObj As ErrObject, ByVal sFunctionName As String, ByVal ErrTitle As String, ByVal ErrMessage As String, ByVal UseStack As Boolean) As Boolean
  Dim ErrNo As Long, ErrSrc As String, ErrDesc As String, p As Long
   
  Call Tracer_suspend
  If Not ErrObj Is Nothing Then
    If UseStack Then Call PopErrorMessageErr(ErrObj)
    ErrNo = ErrObj.Number
    ErrDesc = ErrObj.Description
    ErrSrc = ErrObj.Source
    If Len(sFunctionName) = 0 Then
      p = InStr(1, ErrSrc, ";", vbBinaryCompare)
      If p = 0 Then
        sFunctionName = ErrSrc
        ErrSrc = ""
      Else
        sFunctionName = left$(ErrSrc, p - 1)
        ErrSrc = Mid$(ErrSrc, p + 1)
      End If
    End If
    If Len(ErrDesc) <> 0 Then
      If StrComp(ErrDesc, ErrMessage, vbTextCompare) <> 0 Then
        ErrMessage = ErrMessage & vbCrLf & vbCrLf & ErrDesc
      End If
    End If
    Call ErrObj.Clear
  Else
    ErrNo = 0
  End If
  ErrorMessageEx = ErrMainMsg(ErrNo, ErrType, sFunctionName, ErrSrc, ErrTitle, ErrMessage)
  Call Tracer_restart
End Function

Public Function ErrMainMsg(ByVal ErrNo As Long, ByVal ErrType As errnumbersenum, FunctionName As String, ErrSource As String, ErrTitle As String, ErrMsg As String) As Boolean
  Dim ErrCaption As String
  Dim s As String, CurDate As Date
  Dim bSetInFlag As Boolean, logret As Boolean
  Dim AllowIgnore As Boolean, AllowOther As Boolean, AllowRetry As Boolean, AllowCancel As Boolean
  Dim i As Long, OriginalErrNo As Long
  
  On Error GoTo ErrMainMsg_err
  bSetInFlag = False
  If inErrMainMsg Then GoTo ErrMainMsg_err
  inErrMainMsg = True: bSetInFlag = True
  
  If IgnoreError(ErrNo) Then GoTo ErrMainMsg_end
  
  OriginalErrNo = ErrNo
  Call GetErrorNumber(ErrNo, ErrCaption)
  CurDate = Now
  
  AllowIgnore = (ErrType And ERR_ALLOWIGNORE) = ERR_ALLOWIGNORE
  AllowRetry = (ErrType And ERR_ALLOWRETRY) = ERR_ALLOWRETRY
  AllowCancel = (ErrType And ERR_ALLOWCANCEL) = ERR_ALLOWCANCEL
  AllowOther = (ErrType And ERR_ALLOWOTHER) = ERR_ALLOWOTHER
  
  ErrType = ErrType And (Not ERR_ALLOWIGNORE)
  ErrType = ErrType And (Not ERR_ALLOWRETRY)
  ErrType = ErrType And (Not ERR_ALLOWCANCEL)
  ErrType = ErrType And (Not ERR_ALLOWOTHER)
   
  ' Do not pass contact info to ErrorFilter
  If Not m_IPreErrorFilter Is Nothing Then
    logret = logfunction3(CurDate, ErrNo, ErrTitle, ErrMsg, FunctionName, ErrFileExt)
  Else
    If ErrType < ERR_INFO Then
      logret = logfunction3(CurDate, ErrNo, ErrCaption & " " & ErrTitle, ErrMsg & vbTab & GetStaticEx("CONTACT") & vbTab & mAppExeName & " " & mAppVersion, FunctionName, ErrFileExt)
    Else
      logret = logfunction3(CurDate, ErrNo, ErrCaption & " " & ErrTitle, ErrMsg & vbTab & GetStaticEx("CONTACT") & vbTab & mAppExeName & " " & mAppVersion, FunctionName, LogFileExt)
    End If
  End If
  If Not logret Then GoTo ErrMainMsg_end
  
  'setup form and show
  If (Not mSilentError) And ((ErrType = ERR_ERROR) Or (ErrType = ERR_INFO)) Then
    If (AllowCancel + AllowRetry + AllowOther) < True Then Call ECASE_SYS("Cancel, Retry and Other buttons on errormessages are mutually exclusive")
    
    With frmErr
      .chkIgnoreError.Value = vbUnchecked
      .chkIgnoreError.Visible = AllowIgnore
      
      .cmdExtra.Visible = AllowRetry Or AllowCancel Or AllowOther
      If AllowOther Then
        .cmdExtra.Caption = "&Other"
      ElseIf AllowRetry Then
        .cmdExtra.Caption = "&Retry"
      ElseIf AllowCancel Then
        .cmdExtra.Caption = "&Cancel"
      End If
      If Len(m_OtherCaption) > 0 Then .cmdExtra.Caption = m_OtherCaption
      
      .OtherButton = False
      
      .Caption = ErrTitle
      
      If ErrType = ERR_INFO Then
        .fraErr.Caption = ""
      Else
        .lblErrType.Caption = ErrCaption
      End If
      
      .lblHelp.Caption = GetStaticEx("Contact")
      .lblApplication.Caption = UCase$(mAppName) & " " & mAppVersion
      .txtPath.Text = "Path : " & UCase$(mHomeDirectory)
      .lblExeName.Caption = "Running : " & UCase$(mAppExeName)
      .lblFunction.Caption = "Function : " & FunctionName
      
      Call FillErrorSource(.ErrorSource, ErrSource)
      i = .lblErrMsg.Height
      .lblErrMsg.Caption = ErrMsg
      i = .lblErrMsg.Height - i
      If mFormattedErrorStrings Then
        .picErrMsg.left = .lblErrMsg.left
        .picErrMsg.top = .lblErrMsg.top
        .picErrMsg.Height = .lblErrMsg.Height
        .picErrMsg.Width = .lblErrMsg.Width
        .Message = ErrMsg
      End If
      .ClipMessage = ErrMsg & vbTab & FunctionName & vbTab & mAppExeName & " " & mAppVersion
      .picErrMsg.Visible = mFormattedErrorStrings
      .lblErrMsg.Visible = Not mFormattedErrorStrings
      
      .Height = .Height + i
      .fraErr.Height = .fraErr.Height + i
      .fraContact.top = .fraContact.top + i
      .fraApp.top = .fraApp.top + i
      .chkIgnoreError.top = .chkIgnoreError.top + i
      .cmdExtra.top = .cmdExtra.top + i
      
      .fraDetails.top = .fraDetails.top + i
      .cmdOK.top = .cmdOK.top + i
      .cmdDetails.top = .cmdDetails.top + i
      .lstStack.Clear
      Call Tracer_FillList(.lstStack)
    End With
    Call SetCursorEx(vbArrow, "")
    Call CentreInFormEx(frmErr, Nothing)
    
    If Not isMDI_Minimized Then
      frmErr.Show vbModal
      If frmErr.chkIgnoreError.Value = vbChecked Then Call AddIgnoreError(OriginalErrNo)
      ErrMainMsg = frmErr.OtherButton
    End If
    
    Call ClearCursorEx(False)
    Unload frmErr
    Set frmErr = Nothing

  End If
  
ErrMainMsg_end:
  If bSetInFlag Then
    inErrMainMsg = False
    If Not m_IPostErrorProcess Is Nothing Then Call m_IPostErrorProcess.PostErrorProcess(ErrNo)
  End If
  Exit Function
  
ErrMainMsg_err:
  Call Err.Clear
  Call SetCursorEx(vbArrow, "")
  Call MsgBox("An Error has occurred while processing an Error" & vbCrLf & "Error: " & ErrMsg & vbCrLf & "Function: " & FunctionName, vbCritical + vbOKOnly + vbSystemModal, "Recursive Error handling")
  Call ClearCursorEx(False)
  GoTo ErrMainMsg_end
End Function

Public Sub ECASE_SYS(ByVal s As String, Optional ByVal IDEOnly As Boolean = False)
  On Error Resume Next
  If IsRunningInIDEEx Or (Not IDEOnly) Then
    Call MsgBox(s & vbCrLf & GetStaticEx("Contact") & vbCrLf & UCase$(mAppName) & " " & mAppVersion, vbApplicationModal + vbOKOnly + vbExclamation, "System Error")
  End If
End Sub
   
Private Function InErrorRange(ByVal Value As Long, ByVal base As Long) As Boolean
  InErrorRange = (Value >= base) And (Value < (base + ERROR_INCREMENT))
End Function

' get error string from errnumber
Private Sub GetErrorNumber(eno As Long, sCaption As String)
  If eno >= 0 Then
    sCaption = "Visual Basic error: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSCLIENT_ERROR) Then
    eno = eno - TCSCLIENT_ERROR
    sCaption = "Application error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSTWIST_ERROR) Then
    eno = eno - TCSTWIST_ERROR
    sCaption = "abatec Twist Control error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSUBGRD_ERROR) Then
    eno = eno - TCSUBGRD_ERROR
    sCaption = "abatec UB Object Grid error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSSIZE_ERROR) Then
    eno = eno - TCSSIZE_ERROR
    sCaption = "abatec Size error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSAUTOREPORTER_ERROR) Then
    eno = eno - TCSAUTOREPORTER_ERROR
    sCaption = "abatec AutoData error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSDB_ERROR) Then
    eno = eno - TCSDB_ERROR
    sCaption = "abatec database error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSPARSER_ERROR) Then
    eno = eno - TCSPARSER_ERROR
    sCaption = "abatec Parser error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSREPORTER_ERROR) Then
    eno = eno - TCSREPORTER_ERROR
    sCaption = "abatec Reporter error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSIMPORT_ERROR) Then
    eno = eno - TCSIMPORT_ERROR
    sCaption = "abatec Importer error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSCORE_ERROR) Then
    eno = eno - TCSCORE_ERROR
    sCaption = "abatec core error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSCCORE_ERROR) Then
    eno = eno - TCSCCORE_ERROR
    sCaption = "abatec C core library error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSREPWIZ_ERROR) Then
    eno = eno - TCSREPWIZ_ERROR
    sCaption = "abatec Report Wizard error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSSTAT_ERROR) Then
    eno = eno - TCSSTAT_ERROR
    sCaption = "abatec Status bar error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSMAIL_ERROR) Then
    eno = eno - TCSMAIL_ERROR
    sCaption = "abatec Mail control error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSRECENTFILE_ERROR) Then
    eno = eno - TCSRECENTFILE_ERROR
    sCaption = "abatec recent file list error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSRECENTFILE_ERROR) Then
    eno = eno - TCSRECENTFILE_ERROR
    sCaption = "abatec recent file list error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSDA_ERROR) Then
    eno = eno - TCSDA_ERROR
    sCaption = "abatec Abacus+ data access error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSOMGR_ERROR) Then
    eno = eno - TCSOMGR_ERROR
    sCaption = "abatec object manager error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSWHERE_ERROR) Then
    eno = eno - TCSWHERE_ERROR
    sCaption = "abatec Where control error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSDMENU_ERROR) Then
    eno = eno - TCSDMENU_ERROR
    sCaption = "abatec dynamic menu control error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSADOAUTO_ERROR) Then
    eno = eno - TCSADOAUTO_ERROR
    sCaption = "abatec AutoData ADO error number: " & CStr(eno)
  ElseIf InErrorRange(eno, TCSADOIMPORT_ERROR) Then
    eno = eno - TCSADOIMPORT_ERROR
    sCaption = "abatec ADO importer error number: " & CStr(eno)
  Else
    sCaption = "Unrecognised Error: " & CStr(eno) & " (0x" & Hex$(eno) & ")"
  End If
End Sub

Private Function IsTCSError(ByVal ErrNo As Long) As Boolean
  IsTCSError = ((ErrNo > [_MIN_ERROR_NUMBER]) And (ErrNo < [_MAX_ERROR_NUMBER]))
End Function

Private Function IgnoreError(ByVal ErrNo As Long) As Boolean
  Dim i As Long
  
  For i = 1 To CurIgnoreErrorList
    If ErrNo = IgnoreErrorList(i) Then
      IgnoreError = True
      Exit Function
    End If
  Next i
  IgnoreError = False
End Function

Private Sub AddIgnoreError(ByVal ErrNo As Long)
  Dim i As Long
    
  CurIgnoreErrorList = CurIgnoreErrorList + 1
  If CurIgnoreErrorList > MaxIgnoreErrorList Then
    MaxIgnoreErrorList = MaxIgnoreErrorList + 1024
    If IsArrayEx2(IgnoreErrorList) Then
      ReDim Preserve IgnoreErrorList(1 To MaxIgnoreErrorList) As Long
    Else
      ReDim IgnoreErrorList(1 To MaxIgnoreErrorList) As Long
    End If
  End If
  IgnoreErrorList(CurIgnoreErrorList) = ErrNo
End Sub

Public Sub ClearIgnoreErrors(ByVal ErrorType As Long)
  Dim i As Long, j As Long, minError As Long, maxError As Long
    
  If ErrorType = -1 Then
    CurIgnoreErrorList = 0
  Else
    minError = ErrorType
    maxError = ErrorType + ERROR_INCREMENT - 1
    For i = 1 To CurIgnoreErrorList
      If (IgnoreErrorList(i) >= minError) And (IgnoreErrorList(i) <= maxError) Then
        IgnoreErrorList(i) = 0
      End If
    Next i
  End If
  ' compact error list
  For j = CurIgnoreErrorList To 1 Step -1
    If IgnoreErrorList(j) <> 0 Then Exit For
  Next j
  For i = 1 To j
    If IgnoreErrorList(i) = 0 Then
      IgnoreErrorList(i) = IgnoreErrorList(j)
      j = j - 1
      For j = j To 1 Step -1
        If IgnoreErrorList(j) <> 0 Then Exit For
      Next j
    End If
  Next i
  CurIgnoreErrorList = j
End Sub

Private Sub FillErrorSource(lstErrorSource As ListBox, ErrSource As String)
  Dim nextstring As String
  Dim p1 As Long, p0 As Long
  
  lstErrorSource.Clear
  p0 = 1
  Do
    p1 = InStr(p0, ErrSource, ";", vbBinaryCompare)
    If p1 = 0 Then
      nextstring = Mid$(ErrSource, p0)
    Else
      nextstring = Mid$(ErrSource, p0, (p1 - p0))
    End If
    lstErrorSource.AddItem nextstring
    p0 = p1 + 1
  Loop Until p1 = 0
End Sub
