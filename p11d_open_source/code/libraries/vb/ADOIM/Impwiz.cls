VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImportWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Enum SET_ROWCOLTYPE
  SET_ROW = 1
  SET_COLUMN
End Enum

Public Enum IMPORT_GOTOFORM
  TCSIMP_CANCEL = -1
  TCSIMP_UNDEFINED = 0
  TCSIMP_DEST
  TCSIMP_SOURCE
  TCSIMP_DLIM
  TCSIMP_FW
  TCSIMP_MISC
  TCSIMP_LINK
  TCSIMP_END
  TCSIMP_VLOG
  TCSIMP_GOTOFORM_N
End Enum

Public Enum DOIMPORT_FLAGS
   FIELD_NAME = 1
   LINKED_NAME = 2
   FIELD_TYPE = 4
   ROW_LINENUMBERS = 8
End Enum

Private m_DefaultStatics As Collection
Public ImpParent As ImportClass
Public CurrentDest As DestRecordSet
Private m_ImportForms As Collection
Private m_DestRecSets As Collection
Private m_FieldSpecs As FieldSpecs
Private m_SavedFieldSpecs As FieldSpecs
Private m_PreviewLines As Long
Private m_TitlePrefix As String

Private m_Imported As Boolean

Public GotoForm As IMPORT_GOTOFORM
Private m_AllowChangeUpdateType As Boolean
Private m_AllowSourceChange As Boolean
Private m_AutoOpenSourceFile As Boolean
Private m_AllowMoveColumns As Boolean
Private m_AllowMergeColumns As Boolean

Private m_SourcePath As String
Private m_SpecPath As String
Private m_SpecFile As String
Private m_SpecLoadedOK As Boolean

Private m_PostProcessingLinesInError As Long
Private m_NoReviewSpec As Boolean

' Variables for tracking source file
Private m_SourceFile As String
Private m_NumLines As Long
Private m_MaxLineLength As Long
Private m_FirstLine As Long
Private m_LineArray() As String
Private m_FGDrop As FlexGridDropInfo
Private m_NumDestRecs As Long

Private m_DestPreviewRecLimit As Long
Private m_AllowMultipleImports As Boolean
Private m_ImportAnother As Boolean
Private m_bFileExtensionsDefaults As Boolean
Private m_JustImportedSpec As Boolean
Private m_FileOpenExtensions As String
Private m_FileOpenFilter As Long

Private m_ShowSimpleLinkView As Boolean

Private m_UseSpec As Boolean  'MPSMarch
Private m_GotoSourceScreen As Boolean
Private m_SavedSpec As Boolean

Public Property Get SavedSpec() As Boolean
  SavedSpec = m_SavedSpec
End Property

Public Property Let FileOpenFilter(ByVal NewValue As Long)
  m_FileOpenFilter = NewValue
End Property

Public Sub AddFileExtension(ByVal DisplayName As String, ByVal FileExtension As String, Optional ByVal DefaultChoice As Boolean = False, Optional ByVal ClearExisting As Boolean = False)
  If ClearExisting Then
    m_FileOpenExtensions = ""
    m_FileOpenFilter = 1
  End If
  
  If Len(FileExtension) Then
    If DefaultChoice Then
      m_FileOpenFilter = 1
      m_FileOpenExtensions = DisplayName & "|" & FileExtension & IIf(Len(m_FileOpenExtensions), "|", "") & m_FileOpenExtensions
    Else
      m_FileOpenExtensions = m_FileOpenExtensions & IIf(Len(m_FileOpenExtensions), "|", "") & DisplayName & "|" & FileExtension
    End If
  End If
End Sub

Private Sub Class_Initialize()
  Dim iFrm As IImportForm
  
  Set m_ImportForms = New Collection
  Set iFrm = Frm_Dest
  Call m_ImportForms.Add(iFrm, CStr(iFrm.FormType))
  Set iFrm = Frm_DLim
  Call m_ImportForms.Add(iFrm, CStr(iFrm.FormType))
  Set iFrm = Frm_End
  Call m_ImportForms.Add(iFrm, CStr(iFrm.FormType))
  Set iFrm = Frm_FW
  Call m_ImportForms.Add(iFrm, CStr(iFrm.FormType))
  Set iFrm = Frm_Link
  Call m_ImportForms.Add(iFrm, CStr(iFrm.FormType))
  Set iFrm = Frm_Misc
  Call m_ImportForms.Add(iFrm, CStr(iFrm.FormType))
  Set iFrm = Frm_Source
  Call m_ImportForms.Add(iFrm, CStr(iFrm.FormType))
  
  Set m_DestRecSets = New Collection
  m_FileOpenExtensions = "Text Files (*.txt)|*.txt|Comma Separated Variable Files (*.csv)|*.csv|All Files (*.*)|*.*"
  m_FileOpenFilter = 1
End Sub

Private Sub Class_Terminate()
  Call Me.KillWizard
End Sub

Public Sub KillWizard()
  If Not ImpParent Is Nothing Then
    Set ImpParent.ImportWizard = Nothing
    Set ImpParent = Nothing
  End If
  Set CurrentDest = Nothing
  Set m_ImportForms = Nothing
  Set m_DestRecSets = Nothing
  Set m_FieldSpecs = Nothing
  Set m_SavedFieldSpecs = Nothing
End Sub

Public Property Get TitlePrefix() As String
  TitlePrefix = m_TitlePrefix
End Property

Public Property Let TitlePrefix(ByVal NewValue As String)
  m_TitlePrefix = NewValue
End Property

Public Property Let AllowChangeUpdateType(ByVal Value As Boolean)
  m_AllowChangeUpdateType = Value
End Property

Public Property Get AllowChangeUpdateType() As Boolean
  AllowChangeUpdateType = m_AllowChangeUpdateType
End Property

Public Property Get AllowMoveColumns()
  AllowMoveColumns = m_AllowMoveColumns
End Property

Public Property Get PostProcessingLinesInError()
  PostProcessingLinesInError = m_PostProcessingLinesInError
End Property

Public Property Get SourcePath() As String
  SourcePath = m_SourcePath
End Property

Public Property Let SourcePath(ByVal NewValue As String)
  m_SourcePath = NewValue
End Property

Public Property Get SpecPath() As String
  SpecPath = m_SpecPath
End Property

Public Property Let SpecPath(ByVal NewValue As String)
  m_SpecPath = NewValue
End Property

Public Property Get PreviewLines() As Long
  PreviewLines = m_PreviewLines
End Property

Public Property Let PreviewLines(ByVal NewValue As Long)
  m_PreviewLines = NewValue
End Property

Public Property Get NumLines() As Long
  NumLines = m_NumLines
End Property

Public Property Let NumLines(ByVal NewValue As Long)
  m_NumLines = NewValue
End Property

Public Property Get NoReviewSpec() As Boolean
  NoReviewSpec = m_NoReviewSpec
End Property

Public Property Let NoReviewSpec(ByVal NewValue As Boolean)
  m_NoReviewSpec = NewValue
End Property

Public Property Get AutoOpenSourceFile() As Boolean
  AutoOpenSourceFile = m_AutoOpenSourceFile
End Property

Public Property Let AutoOpenSourceFile(ByVal NewValue As Boolean)
  m_AutoOpenSourceFile = NewValue
End Property

Public Property Get ShowSimpleLinkView() As Boolean
  ShowSimpleLinkView = m_ShowSimpleLinkView
End Property

Public Property Let ShowSimpleLinkView(ByVal NewValue As Boolean)
  m_ShowSimpleLinkView = NewValue
End Property

Public Property Get AllowMultipleImports() As Boolean
  AllowMultipleImports = m_AllowMultipleImports
End Property

Public Property Let AllowMultipleImports(ByVal NewValue As Boolean)
  m_AllowMultipleImports = NewValue
End Property

Public Property Get Imported() As Boolean
  Imported = m_Imported
End Property

Public Property Get ImportAnother() As Boolean
  ImportAnother = m_ImportAnother
End Property

Public Property Let ImportAnother(ByVal NewValue As Boolean)
  m_ImportAnother = NewValue
End Property

Public Property Get DefaultStatics() As Collection
  Set DefaultStatics = m_DefaultStatics
End Property

Public Property Set DefaultStatics(ByVal NewValue As Collection)
  Set m_DefaultStatics = NewValue
End Property

Public Property Get DestPreviewRecLimit() As Long
  DestPreviewRecLimit = m_DestPreviewRecLimit
End Property

Public Property Let DestPreviewRecLimit(ByVal NewValue As Long)
  m_DestPreviewRecLimit = NewValue
End Property

Public Property Get SourceFile() As String
  SourceFile = m_SourceFile
End Property

Public Property Let SourceFile(ByVal NewValue As String)
  If Not FileExists(NewValue) Then Err.Raise 53  ' file does not exist
  m_SourceFile = NewValue
End Property

Public Property Get SpecFile() As String
  SpecFile = m_SpecFile
End Property

Public Property Let SpecFile(ByVal NewValue As String)
  m_SpecFile = NewValue
End Property

Public Function RemoveSQL(ByVal Name As String) As Boolean
  RemoveSQL = RemoveRS(Name)
End Function

Public Function RemoveRS(ByVal Name As String) As Boolean
  Dim drs As DestRecordSet
  Dim iFrm As IImportForm, frm As Form, i As Long
  
  On Error GoTo RemoveRS_err
  m_DestRecSets.Remove Name
  RemoveRS = True
  
RemoveRS_end:
  If RemoveRS Then
    Set iFrm = m_ImportForms.Item(CStr(TCSIMP_DEST))
    Set frm = iFrm
    For i = 1 To frm.Cbo_Table.ListCount
      frm.Cbo_Table.RemoveItem 0
    Next i
  End If
  Exit Function
  
RemoveRS_err:
  RemoveRS = False
  Resume RemoveRS_end
End Function

Public Function AddRS(ByVal rsAdd As Recordset, ByVal Name As String, Optional ByVal Tag As Variant, Optional ByVal HelpText As String, Optional ByVal UpdateType As IMPORT_UPDATE_TYPE = NO_UPDATES, Optional ByVal ProcessImport As IProcessImport, Optional ByVal DisplayName As String) As Boolean
  If Len(Name) = 0 Then Name = "Recordset"
  If Len(DisplayName) = 0 Then DisplayName = Name
  AddRS = AddDest(rsAdd, "", Nothing, Name, DisplayName, Tag, HelpText, UpdateType, ProcessImport)
End Function

Public Function AddSQL(ByVal SQLString As String, ByVal CN As Connection, ByVal Name As String, Optional ByVal Tag As Variant, Optional ByVal HelpText As String, Optional ByVal UpdateType As IMPORT_UPDATE_TYPE = NO_UPDATES, Optional ByVal ProcessImport As IProcessImport, Optional ByVal DisplayName As String) As Boolean
  If Len(Name) = 0 Then Name = SQLString
  If Len(DisplayName) = 0 Then DisplayName = Name
  AddSQL = AddDest(Nothing, SQLString, CN, Name, DisplayName, Tag, HelpText, UpdateType, ProcessImport)
End Function

' UPDATE types
'* note If a Field is a key field it is required !!
Private Function AddDest(ByVal rsAdd As Recordset, ByVal SQLString As String, ByVal CN As Connection, ByVal Name As String, ByVal DisplayName As String, ByVal Tag As Variant, ByVal HelpText As String, ByVal UpdateType As IMPORT_UPDATE_TYPE, ByVal ProcessImport As IProcessImport) As Boolean
  Dim drs As DestRecordSet, i As Long
  Dim FieldName As String
   
  On Error GoTo AddDest_err
  Call xSet("AddDest")
  AddDest = True
  Set drs = New DestRecordSet
  drs.Name = Name
  drs.DisplayName = DisplayName
  drs.HelpText = HelpText
  
  If Not rsAdd Is Nothing Then
    If Not rsAdd.Supports(adUpdate) Then Err.Raise ERR_ADDFAILED, "AddRS", "Appended Recordset is not updateable"
    Set drs.rs = rsAdd
    drs.SQLString = ""
    Set drs.CN = Nothing
  End If
  
  If Len(SQLString) > 0 Then
    Set drs.rs = Nothing
    drs.SQLString = SQLString
    Set drs.CN = CN
  End If
  drs.UpdateType = UpdateType
  If Not IsMissing(Tag) Then drs.Tag = Tag
  Set drs.ProcessImport = ProcessImport
  If m_DestRecSets Is Nothing Then Set m_DestRecSets = New Collection
  Call m_DestRecSets.Add(drs, drs.Name)
  AddDest = True
  
AddDest_end:
  Call xReturn("AddDest")
  Exit Function
  
AddDest_err:
  AddDest = False
  Call ErrorMessage(ERR_ERROR, Err, "AddDest", "Adding new Destination to Import Wizard", "Unable to add destination '" & Name & "'")
  Resume AddDest_end
End Function


Public Sub StartWizard()
  Dim FreeFil As Long
  Dim iFrm As IImportForm, frm As Form, i As Long
  Dim frmEnd As Frm_End
  Dim Started As Boolean

  On Error GoTo StartWizard_Err
  Call xSet("StartWizard")
  m_ImportAnother = False
  m_SavedSpec = False
  Set m_FieldSpecs = New FieldSpecs
  Set m_SavedFieldSpecs = New FieldSpecs
  Set CurrentDest = Nothing
  
  Set frmEnd = m_ImportForms.Item(CStr(TCSIMP_END))

  m_AllowMoveColumns = False
  m_AllowMergeColumns = False
  
  For Each iFrm In m_ImportForms
    Set iFrm.ImpWiz = Me
  Next iFrm
    
  If m_DestRecSets.Count = 0 Then Call Err.Raise(ERR_NOIMPDEST, "StartWizard", "No updateable recordsets were specified.")
  
  If m_PreviewLines <= 0 Then m_PreviewLines = 5
  If m_DestPreviewRecLimit <= 0 Then m_DestPreviewRecLimit = 100
  m_Imported = False
  m_AllowSourceChange = (Len(m_SourceFile) = 0)
  If Len(m_SourcePath) = 0 Then m_SourcePath = AppPath
  If Len(m_SpecPath) = 0 Then m_SpecPath = AppPath
  
  ' Check that these paths are valid !
  ImpParent.Delimiter = Chr$(0)
  ImpParent.HeaderCount = -1
  ImpParent.ImportType = IMPORT_DELIMITED
  Call m_FieldSpecs.RemoveAll
  Call m_SavedFieldSpecs.RemoveAll
   
  GotoForm = TCSIMP_DEST
  If ReCalc_Dest Then
    If Len(m_SpecFile) > 0 Then Call LoadSpec(m_SpecFile)
    Do
      If GotoForm = TCSIMP_CANCEL Then
        Exit Do
      End If
      
      If (m_DestRecSets.Count = 1) And GotoForm = TCSIMP_DEST And (Not Started) Then
        'Frm_Source.Cmd_Back.Visible = False  'User should be able to go back to see existing recordset contents
        Frm_Dest.Cmd_Next.Value = True
        GotoForm = TCSIMP_SOURCE
      Else
        Set frm = m_ImportForms.Item(CStr(GotoForm))
        frm.Show vbModal
      End If
      Started = True
    Loop
  End If
  
  For Each iFrm In m_ImportForms
    Set iFrm.ImpWiz = Nothing
    Set frm = iFrm
    Unload frm
  Next iFrm
  If Not m_Imported Then Set CurrentDest = Nothing
  
StartWizard_End:
  Call xReturn("StartWizard")
  Exit Sub

StartWizard_Err:
  Call ErrorMessage(ERR_ERROR, Err, "StartWizard", "Start Import Wizard", "Import wizard failed to start.")
  Resume StartWizard_End
  Resume
End Sub

Friend Function DoImport(ByVal flx As MSFlexGrid, ByVal Flags As DOIMPORT_FLAGS, ByVal DisplayErrors As Boolean, Optional ByVal Actual As Boolean = False) As Boolean
  Dim i As Long, j As Long, lLinesOK As Long, lAddedRecordCount As Long, lUpdatedRecordCount As Long
  Dim isp As ImportSpec
  Dim ImpVariant() As Variant, ImpFieldNamesVariant As Variant, vtmp As Variant
  Dim ErrorCount As Long
  Dim InRetry As Boolean
  Dim NumHeaderRows As Long
  Dim s As String, FieldType  As String, DoRollback As Boolean
  Dim OldCol As Long
  Dim CurrentRow As Long
  Const COLUMN_WIDTH As Long = 1500
  Const COLUMN0_WIDTH As Long = 600
  Dim frmEnd As Frm_End
  Dim frmdlim As Frm_DLim
  Dim flxRedraw As Boolean
  Dim ActualUpdateType As IMPORT_UPDATE_TYPE
  Dim ActualUseTransactionFlag As Boolean
  Dim OldErrorFilter As IErrorFilter
  
  On Error GoTo DoImport_Err
  Call xSet("DoImport")
  Call SetCursor
  ActualUpdateType = ImpParent.UpdateType
  ActualUseTransactionFlag = ImpParent.UseTransaction
  Set frmEnd = m_ImportForms.Item(CStr(TCSIMP_END))
  Set frmdlim = m_ImportForms.Item(CStr(TCSIMP_DLIM))
  InRetry = False
retry:
  ErrorCount = 1
  
  If Not Actual Then
    ImpParent.UpdateType = NO_UPDATES
    ImpParent.UseTransaction = False
  End If
  
  If ImpParent.InitImport(m_SourceFile, ImpParent.ImportType, , ImpParent.UpdateType) Then
    For i = 1 To m_FieldSpecs.Count
      Set isp = m_FieldSpecs(i)
      isp.Tempkey = isp.key
      If Actual Then
        If Len(isp.DestField) = 0 Then isp.key = KEY_SKIPFIELD
      End If
      Call ImpParent.AddColumn(isp)
    Next i
    ImpParent.TestImport = Not Actual
    If Actual Then
      If Not CurrentDest.ProcessImport Is Nothing Then
        m_PostProcessingLinesInError = CurrentDest.ProcessImport.PreProcess(ImpParent)
      End If
      m_NumDestRecs = SetDestDataGrid(CurrentDest.rs, TCSIMP_END)
      
      If ImpParent.UseTransaction Then Call TransactionADO(ImpParent.Connection, TRANS_BEGIN)
      ErrorCount = ImpParent.ImportFile(CurrentDest.rs, frmEnd.PBar_Import)
      For i = 1 To m_FieldSpecs.Count
        Set isp = m_FieldSpecs(i)
        isp.key = isp.Tempkey
      Next i
    Else
      ErrorCount = ImpParent.ImportLines(ImpVariant(), UBound(m_LineArray) + ImpParent.FooterCount)
    End If
    
    If Not Actual And ErrorCount <> 0 Then
      Call m_FieldSpecs.RemoveAll
      If InRetry Or (m_SavedFieldSpecs.Count = 0) Then
        InRetry = True
        Err.Raise ERR_RESTOREIMPORT, "TestImport", "Unable to roll back to previous import spec"
      End If
      For i = 1 To m_SavedFieldSpecs.Count
        Set isp = m_FieldSpecs.Add
        Call m_SavedFieldSpecs(i).Copy(isp)
      Next i
      Call FixUpCopyFields(m_FieldSpecs)
      InRetry = True
      GoTo retry
    End If
  End If
  
  If Actual Then
    'frmEnd.Cmd_Back.Enabled = False
    m_Imported = True
    
    If Not CurrentDest.ProcessImport Is Nothing Then
      If Not ImpParent.ErrorFilter Is Nothing Then
        Set OldErrorFilter = ErrorFilter
        Set ErrorFilter = ImpParent.ErrorFilter
      End If
      
      lAddedRecordCount = ImpParent.AddedRecordCount 'will become less post processing error, done by users PostProcessingInterface
      lLinesOK = ImpParent.LinesOK
      lUpdatedRecordCount = ImpParent.UpdatedRecordCount
      m_PostProcessingLinesInError = CurrentDest.ProcessImport.PostProcess(ImpParent, DoRollback, lLinesOK, lAddedRecordCount, lUpdatedRecordCount)
      
      If Not ImpParent.ErrorFilter Is Nothing Then
        Set ErrorFilter = OldErrorFilter
      End If
    End If
    If ImpParent.UseTransaction Then
      If DoRollback Then
        Call TransactionADO(ImpParent.Connection, TRANS_ROLLBACK)
        'Call Transaction(ImpParent.WorkSpace, TRANS_ROLLBACK)
        ImpParent.AddedRecordCount = 0
        ImpParent.LinesOK = lLinesOK
        ImpParent.UpdatedRecordCount = 0
      Else
        ImpParent.AddedRecordCount = lAddedRecordCount
        ImpParent.LinesOK = lLinesOK
        ImpParent.UpdatedRecordCount = lUpdatedRecordCount
        'apf cd error if commit fails?
        Call TransactionADO(ImpParent.Connection, TRANS_COMMIT)
        'Call Transaction(ImpParent.WorkSpace, TRANS_COMMIT)
      End If
      
    End If
    Call ReCalc_End(True)
  Else
     ' apf save spec here
     Call UpdateImportDescription(True)
     m_SavedFieldSpecs.RemoveAll
     For i = 1 To m_FieldSpecs.Count
       Set isp = m_SavedFieldSpecs.Add
       Call m_FieldSpecs(i).Copy(isp)
     Next i
     Call FixUpCopyFields(m_SavedFieldSpecs)
     
     flxRedraw = flx.Redraw
     flx.Redraw = False
     OldCol = flx.Col
  
     NumHeaderRows = 0
     If (Flags And FIELD_NAME) = FIELD_NAME Then NumHeaderRows = NumHeaderRows + 1
     If (Flags And FIELD_TYPE) = FIELD_TYPE Then NumHeaderRows = NumHeaderRows + 1
      
     Call SetFlexRowCol(flx, SET_ROW, UBound(m_LineArray) + NumHeaderRows, NumHeaderRows)
     For i = 0 To (FIXED_COLCOUNT - 1)
       If (Flags And ROW_LINENUMBERS) = ROW_LINENUMBERS Then
         flx.ColWidth(0) = COLUMN0_WIDTH
       Else
         flx.ColWidth(i) = 0
       End If
     Next i
     
     'Fixed rows headers
     For i = 1 To m_FieldSpecs.Count
       Set isp = m_FieldSpecs(i)
       CurrentRow = 0
'         If Not m_FieldSpecs(i).MergeField Is Nothing Then
'           .ColWidth(i) = 0
'         Else
       flx.ColWidth(i) = COLUMN_WIDTH
       If isp.Hide Then flx.ColWidth(i) = 0
       flx.ColAlignment(i) = flexAlignLeftCenter
       If (Flags And FIELD_NAME) = FIELD_NAME Then
         s = m_FieldSpecs(i).FldName
         If (Flags And LINKED_NAME) = LINKED_NAME And Len(m_FieldSpecs(i).DestField) > 0 Then
           s = s & " (" & m_FieldSpecs(i).DestField & ")"
         End If
         flx.TextMatrix(CurrentRow, GetFlexColumn(flx, isp.FieldKey)) = s
         CurrentRow = CurrentRow + 1
       End If
       If (Flags And FIELD_TYPE) = FIELD_TYPE Then
         If Not m_FieldSpecs(i).CopyField Is Nothing Then
           FieldType = DataTypeName(m_FieldSpecs(i).CopyField.vartype)
         Else
           FieldType = DataTypeName(m_FieldSpecs(i).vartype)
         End If
         'MPSMarch2
         If (FieldType = DataTypeName(TYPE_LONG) Or FieldType = DataTypeName(TYPE_DOUBLE)) And IsNumeric(m_FieldSpecs(i).actionparam) Then
           If m_FieldSpecs(i).actionparam <> 1 Then
             FieldType = FieldType & " (" & Chr$(215) & CStr(m_FieldSpecs(i).actionparam) & ")"
           End If
         End If
         If IsSpecialFieldType(isp.key) Then FieldType = FieldType & " (" & SpecialFieldTypeName(isp.key) & ")"  'MPSMarch2
         flx.TextMatrix(1, GetFlexColumn(flx, isp.FieldKey)) = FieldType
         CurrentRow = CurrentRow + 1
       End If
     Next i
     ' data
     For i = CurrentRow To UBound(m_LineArray) + NumHeaderRows - 1
       flx.TextMatrix(i, 0) = i - CurrentRow + m_FirstLine
       For j = 1 To m_FieldSpecs.Count
         Set isp = m_FieldSpecs(j)
         vtmp = ImpVariant(i - CurrentRow + 1, j)
         If IsNull(vtmp) Then
           s = ""  ' (Null) here ? apf
         Else
           If isp.vartype = TYPE_DATE Then
             s = Format$(vtmp, "D MMM YYYY")
           Else
             s = CStr(vtmp)
           End If
         End If
         flx.TextMatrix(i, GetFlexColumn(flx, isp.FieldKey)) = s
       Next j
     Next i
     If ImpParent.TakeFieldNamesFromLine > 0 Then
       Call ImportFieldNames(ImpFieldNamesVariant)
       i = 0 'CurrentRow - 1
       For j = 1 To m_FieldSpecs.Count
         Set isp = m_FieldSpecs(j)
         If isp.key <> KEY_CONSTANT And isp.key <> KEY_COPIEDFIELD Then
           isp.FldName = "" & ImpFieldNamesVariant(j)
           flx.TextMatrix(i, GetFlexColumn(flx, isp.FieldKey)) = "" & CStr("" & isp.FldName)
         End If
       Next j
       Call MakeFieldNamesUnique(m_FieldSpecs)
     End If
     If OldCol >= flx.cols Then OldCol = flx.cols - 1
     flx.Col = OldCol
     flx.Row = CurrentRow
     flx.RowSel = flx.Rows - 1
     If flx.ColWidth(flx.Col) = 0 Then
       flx.Col = FIXED_COLCOUNT
       flx.LeftCol = FIXED_COLCOUNT
       flx.Row = FIXED_ROWCOUNT
       flx.RowSel = flx.Rows - 1
     End If
     flx.Redraw = flxRedraw
     If Not flx.ColIsVisible(flx.Col) Then
       flx.LeftCol = flx.Col
     End If
  End If

DoImport_End:
  'If Not ImpParent.WorkSpace Is Nothing Then Call Transaction(ImpParent.WorkSpace, TRANS_ROLLBACK)
  If ImpParent.UseTransaction Then Call TransactionADO(ImpParent.Connection, TRANS_ROLLBACK)
  ImpParent.UpdateType = ActualUpdateType
  ImpParent.UseTransaction = ActualUseTransactionFlag
  DoImport = Not InRetry
  Call ClearCursor
  Call xReturn("DoImport")
  Exit Function

DoImport_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DoImport", "Test Import specification", "Import specification does not define a valid import.")
  Resume DoImport_End
  Resume
End Function

Private Function GetdestbyDisplayname(ByVal DisplayName As String) As DestRecordSet
  Dim drs As DestRecordSet
  
  For Each drs In m_DestRecSets
    If StrComp(DisplayName, drs.DisplayName, vbTextCompare) = 0 Then
      Set GetdestbyDisplayname = drs
      Exit Function
    End If
  Next drs
End Function

Friend Function ReCalc_Dest(Optional ByVal NewDest As DestRecordSet) As Boolean 'MPSMarch
  Dim frmDest As Frm_Dest
  Dim frmlink As Frm_Link
  Dim frmEnd As Frm_End
  Dim OldDest As DestRecordSet
  Dim NumDestFields As Long
  Dim fld As Field, fSpec As ImportSpec
  Dim i As Long ', updateimp As Boolean
  Dim NumDestRecs As Long
  Dim drs As DestRecordSet

  On Error GoTo ReCalc_Dest_Err
  Call xSet("ReCalc_Dest")
  If mDisableRecalc Then GoTo ReCalc_Dest_End
  ReCalc_Dest = False
  Set frmDest = m_ImportForms.Item(CStr(TCSIMP_DEST))
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  Set frmEnd = m_ImportForms.Item(CStr(TCSIMP_END))
  
  frmDest.Caption = m_TitlePrefix & "Destination Recordset"
  
ReCalc_redo:
  Set OldDest = CurrentDest
  If frmDest.Cbo_Table.ListCount > 0 Then
    If frmDest.Cbo_Table.ListIndex < 0 Then frmDest.Cbo_Table.ListIndex = 0
  End If
  If frmDest.Cbo_Table.ListCount = 0 Then
    If Not GetDestTableInfo(frmDest) Then GoTo ReCalc_Dest_End
  End If
      
  If Not NewDest Is Nothing Then
    Set OldDest = Nothing
    Set CurrentDest = NewDest
    frmDest.Cbo_Table.Text = CurrentDest.DisplayName
  Else
    Set CurrentDest = GetdestbyDisplayname(frmDest.Cbo_Table.Text)
  End If
  If CurrentDest.rs Is Nothing Then
    If Not OpenSQLspecifiedRS(CurrentDest.SQLString) Then
      Call RemoveSQL(CurrentDest.Name)
      Set CurrentDest = Nothing
      GoTo ReCalc_redo
    End If
  End If
  ' have valid recordset check CurrentRS values
  If Not ValidateDestination(CurrentDest) Then
    Call RemoveSQL(CurrentDest.Name)
    Set CurrentDest = Nothing
    GoTo ReCalc_redo
  End If
    
  ImpParent.UpdateType = CurrentDest.UpdateType
  If Not OldDest Is CurrentDest Then
    Set m_FieldSpecs = New FieldSpecs
    Set m_SavedFieldSpecs = New FieldSpecs
    frmDest.FlG_Dest.Visible = True
    NumDestRecs = SetDestDataGrid(CurrentDest.rs, TCSIMP_DEST)
    Call SetDestDataGrid(CurrentDest.rs, TCSIMP_LINK)
    
    m_NumDestRecs = SetDestDataGrid(CurrentDest.rs, TCSIMP_END)
    
    Call SetFlexRowCol(frmlink.FlG_DestInfo, SET_ROW, 2, 1)
    Call SetFlexRowCol(frmlink.FlG_DestInfo, SET_COLUMN, 4, 0)
    With frmlink.FlG_DestInfo
      .TextMatrix(0, 0) = "Field"
      .ColWidth(0) = .Width / 100 * 50
      .TextMatrix(0, 1) = "Type"
      .ColWidth(1) = .Width / 100 * 18
      .TextMatrix(0, 2) = "Linked to"
      .ColWidth(2) = .Width / 100 * 16
      .TextMatrix(0, 3) = "Key"
      .ColWidth(3) = .Width / 100 * 8
      NumDestFields = 0
      For Each fld In CurrentDest.rs.Fields
        NumDestFields = NumDestFields + 1
        .Rows = 1 + NumDestFields
        .TextMatrix(NumDestFields, 0) = fld.Name
        .TextMatrix(NumDestFields, 1) = TypeName(fld.Type)
        If TypeName(fld.Type) = "String" Then
          .TextMatrix(NumDestFields, 1) = .TextMatrix(NumDestFields, 1) & ": " & fld.DefinedSize
        End If
        .TextMatrix(NumDestFields, 2) = ""
        .TextMatrix(NumDestFields, 3) = ""
        If CurrentDest.HiddenFields.IsPresent(fld.Name) Then
          .RowHeight(NumDestFields) = 0
        End If
      Next fld
    End With
  
    For Each drs In m_DestRecSets
      If (Not drs Is CurrentDest) And Len(drs.SQLString) > 0 And (Not drs.rs Is Nothing) Then
        drs.rs.Close
        Set drs.rs = Nothing
      End If
    Next drs
  End If
  ImpParent.FooterCount = -1
  Call SetUpdateType(frmDest, CurrentDest.UpdateType, m_AllowChangeUpdateType, CurrentDest.LockUpdate)
  ReCalc_Dest = True
  'frmEnd.Lbl_RecsImported.Caption = ""
  
ReCalc_Dest_End:
  Call xReturn("ReCalc_Dest")
  Exit Function

ReCalc_Dest_Err:
  ReCalc_Dest = False
  If Not frmDest Is Nothing Then frmDest.FlG_Dest.Visible = False
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_Dest", "Recalc Destination Form", "Unable to recalc the Destination form")
  Resume ReCalc_Dest_End
End Function
 

Private Function OpenSQLspecifiedRS(ByVal SQL As String) As Boolean

  On Error GoTo OpenSQLspecifiedRS_Err
  Call xSet("OpenSQLspecifiedRS")
  OpenSQLspecifiedRS = False
  Set CurrentDest.rs = New Recordset
  Call CurrentDest.rs.Open(SQL, CurrentDest.CN, adOpenStatic, adLockOptimistic)
  If Not CurrentDest.rs.Supports(adUpdate) Then
    Set CurrentDest.rs = Nothing
    Call Err.Raise(ERR_RSNOTUPDATEABLE, "OpenSQLspecifiedRS", "SQL specified recordset is not updateable:" & vbCrLf & SQL)
  End If
  OpenSQLspecifiedRS = True
  
OpenSQLspecifiedRS_End:
  Call xReturn("OpenSQLspecifiedRS")
  Exit Function

OpenSQLspecifiedRS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenSQLspecifiedRS", "OpenSQL import destination", "Error opening SQL destination specified by: " & vbCrLf & SQL)
  Resume OpenSQLspecifiedRS_End
  Resume
End Function

Friend Function ReCalc_Src(ByVal frmSrcObj As Object) As Boolean
  Dim FrmSrc As Frm_Source
  Dim Path As String, File As String, Ext As String
  Dim i As Long, j As Long

  On Error GoTo ReCalc_Src_Err
  Call xSet("ReCalc_Src")
  If mDisableRecalc Then GoTo ReCalc_Src_End
  Set FrmSrc = frmSrcObj
  If FrmSrc Is Nothing Then Set FrmSrc = m_ImportForms.Item(CStr(TCSIMP_SOURCE))
  
  FrmSrc.Caption = m_TitlePrefix & "Source Data File"
  
  If ImpParent.ImportType = IMPORT_DELIMITED Then
    FrmSrc.Opt_Format(0).Value = True
  Else
    FrmSrc.Opt_Format(1).Value = True
  End If
  If Len(m_SourceFile) > 0 Then
    If m_NumLines = 0 Then Call FileInfo
    FrmSrc.Lbl_SourceInst.Caption = "Data will be imported from the file shown below."
    Call SplitPath(m_SourceFile, Path, File, Ext)
    File = File & Ext
    'FrmSrc.Lbl_SourcePath.Caption = Path & vbCrLf & File
    FrmSrc.Lbl_SourcePath.Caption = m_SourceFile
    FrmSrc.Lbl_SourcePath.Visible = True
    If m_AllowSourceChange Then
      FrmSrc.Lbl_SourceInst = FrmSrc.Lbl_SourceInst & vbCrLf & "Press the Open button to select a different file."
      FrmSrc.Cmd_OpenSource.Visible = True
    Else
      FrmSrc.Cmd_OpenSource.Visible = False
     End If
    m_FirstLine = CopyLines(m_SourceFile, m_LineArray(), 1, m_PreviewLines)
    If Not ((m_FirstLine = 0) Or (m_NumLines = 0)) Then
      FrmSrc.FW_Source.RestoreOriginalSize
      FrmSrc.FW_Source.LinesCopied = UBound(m_LineArray)
      FrmSrc.FW_Source.CharsWide = m_MaxLineLength
      FrmSrc.FW_Source.LinesIn = CVar(m_LineArray)
      FrmSrc.FW_Source.FWFontName = "Courier"
      FrmSrc.FW_Source.FWFontSize = 10
      FrmSrc.FW_Source.Refresh
      FrmSrc.FW_Source.Visible = True
      FrmSrc.Lbl_SrcContents.Caption = "Selected file is " & m_NumLines & " lines long and a maximum of " & m_MaxLineLength & " characters wide:"
      FrmSrc.Lbl_SrcContents.Visible = True
      FrmSrc.Cmd_Next.Enabled = True
      FrmSrc.Fra_Format.Visible = True
    End If
  Else
    FrmSrc.Lbl_SourcePath.Visible = False
    FrmSrc.Cmd_OpenSource.Enabled = True
    FrmSrc.Lbl_SrcContents.Visible = False
    FrmSrc.Cmd_Next.Enabled = False
    FrmSrc.Fra_Format.Visible = False
    FrmSrc.FW_Source.Visible = False
    FrmSrc.Lbl_SourceInst.Caption = "Press the Open button to select a file to import from."
    If m_AutoOpenSourceFile Then Call OpenSourceFile(FrmSrc)
  End If
  
ReCalc_Src_End:
  Call xReturn("ReCalc_Src")
  Exit Function

ReCalc_Src_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_Src", "Recalc of Source Form", "Error in Source Form recalc.")
  Resume ReCalc_Src_End
End Function

Public Sub OpenSourceFile(ByVal frmSrcObj As Object)
  Dim s As String, sFilesOfTypes As String
  Dim FrmSrc As Frm_Source
  Dim frmDest As Frm_Dest
  Dim l As Long
  
  On Error GoTo OpenSourceFile_Err
  Call xSet("OpenSourceFile")
  Set FrmSrc = frmSrcObj
  If FrmSrc.Visible = False Then
    Set frmDest = m_ImportForms.Item(CStr(TCSIMP_DEST))
    frmDest.Visible = False
    FrmSrc.Show
  End If
    
RetryFileOpen:
  s = FileOpenDlgFilter(m_FileOpenFilter, "Choose a Source File to Open", m_FileOpenExtensions, m_SourcePath)
  If Len(s) > 0 Then
    'MPSMarch
    If FileLen(s) = 0 Then
      Call MsgBox("The file '" & s & "' is empty.", vbOKOnly + vbExclamation, "Open Source File")
      GoTo RetryFileOpen
    End If
    m_SourceFile = s
    Call SplitPath(s, m_SourcePath)
    m_NumLines = 0
    ImpParent.HeaderCount = -1
    Call ReCalc_Src(FrmSrc)
  End If
  FrmSrc.Visible = False

OpenSourceFile_End:
  Call xReturn("OpenSourceFile")
  Exit Sub

OpenSourceFile_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenSourceFile", "Open source file", "Unable to open source file.")
  Resume OpenSourceFile_End
End Sub

Friend Function ReCalc_DLim(ByVal ForceInit As Boolean) As Boolean
  Dim frmdlim As Frm_DLim
  Dim i As Long
  Dim s As String
  Dim Offset As Long, nFields As Long, nMaxFields As Long, iLen As Long
  Dim nSourceFieldSpecs As Long
  Dim isp As ImportSpec
  Dim ReInit As Boolean
  Dim TestImportFlags As DOIMPORT_FLAGS
  
  On Error GoTo ReCalc_DLim_Err
  Call xSet("ReCalc_DLim")
  If mDisableRecalc Then GoTo ReCalc_DLim_End
  ReInit = ForceInit
  Set frmdlim = m_ImportForms.Item(CStr(TCSIMP_DLIM))
  
  frmdlim.Caption = m_TitlePrefix & "Delimit Source File"

  frmdlim.UpD_Omit(0).Max = m_NumLines - 1
  frmdlim.UpD_Omit(1).Max = m_NumLines - 1
  frmdlim.UpD_FieldNames.Max = m_NumLines - 1
    
  If (ImpParent.TakeFieldNamesFromLine <> CLng(Val(frmdlim.Txt_FieldNames))) Or (ImpParent.HeaderCount <> CLng(frmdlim.Txt_Omit(0))) Or (ImpParent.FooterCount <> CLng(frmdlim.Txt_Omit(1))) Or (UBound(m_LineArray) <= 0) Then
    ReInit = True
    ImpParent.HeaderCount = frmdlim.Txt_Omit(0)
    ImpParent.FooterCount = frmdlim.Txt_Omit(1)
    m_FirstLine = CopyLines(m_SourceFile, m_LineArray(), ImpParent.HeaderCount + 1, Min(m_PreviewLines, m_NumLines - ImpParent.FooterCount - ImpParent.HeaderCount))
    ImpParent.TakeFieldNamesFromLine = CLng(Val(frmdlim.Txt_FieldNames))
  End If

  If StrComp(ImpParent.Delimiter, frmdlim.GetDelimiters, vbTextCompare) <> 0 Then
    ReInit = True
    ImpParent.Delimiter = frmdlim.GetDelimiters
  End If
  If StrComp(ImpParent.EscapeChar, frmdlim.GetTextQualifier, vbTextCompare) <> 0 Then
    ReInit = True
    ImpParent.EscapeChar = frmdlim.GetTextQualifier
  End If

  If ReInit Then
    If Not m_JustImportedSpec Then m_UseSpec = False 'MPSMarch
    nMaxFields = 0
    For i = 1 To UBound(m_LineArray)
      Offset = 1
      nFields = 0
      iLen = Len(m_LineArray(i))
      Do
        Offset = GetDelimitedValue(s, m_LineArray(i), Offset, ImpParent.Delimiter, ImpParent.EscapeChar)
        nFields = nFields + 1
      Loop Until Offset > iLen
      nMaxFields = Max(nMaxFields, nFields)
    Next i
    nSourceFieldSpecs = m_FieldSpecs.Count
    For i = 1 To m_FieldSpecs.Count
      'If m_FieldSpecs(i).key = KEY_CONSTANT Then nSourceFieldSpecs = nSourceFieldSpecs - 1
      'If m_FieldSpecs(i).key = KEY_COPIEDFIELD Then nSourceFieldSpecs = nSourceFieldSpecs - 1
      'MPSMarch2
      Select Case m_FieldSpecs(i).key
        Case KEY_CONSTANT, KEY_COPIEDFIELD, KEY_DATENOW, KEY_FILENAME, KEY_FILEPATH, KEY_FILEDATE, KEY_LINENUMBER, KEY_CFGFILENAME, KEY_USERNAME, KEY_IMPDATE
          nSourceFieldSpecs = nSourceFieldSpecs - 1
        Case Else
      End Select
    Next i
    If (Not m_JustImportedSpec) Or (nMaxFields <> nSourceFieldSpecs) Then
      If m_JustImportedSpec Then
        Call ErrorMessage(ERR_INFO, Nothing, "LoadSpec", "Load specification", "Specification has " & CStr(nSourceFieldSpecs) & " field" & String$(Sgn(nSourceFieldSpecs - 1), "s") & "," & vbCrLf & "Source file has " & CStr(nMaxFields) & " field" & String$(Sgn(nMaxFields - 1), "s") & "." & vbCrLf & vbCrLf & "IMPORT ABORTED" & vbCrLf)
        m_SpecLoadedOK = False
        m_JustImportedSpec = False
      End If
      Call SetFlexRowCol(frmdlim.FlG_Source, SET_COLUMN, nMaxFields + FIXED_COLCOUNT, FIXED_COLCOUNT)
      Call m_SavedFieldSpecs.RemoveAll
      Call m_FieldSpecs.RemoveAll
      For i = 1 To nMaxFields
        Set isp = m_FieldSpecs.Add
        isp.key = KEY_FIELD
        isp.vartype = TYPE_STR
        isp.FldName = FieldName(i, nMaxFields)
        isp.FieldKey = i ' apf used to reference this column < 0, for constants/copies
        frmdlim.FlG_Source.Coldata(i + FIXED_COLCOUNT - 1) = isp.FieldKey
      Next i
      'Call AddInDefaultStatics(frmdlim.Name)
    End If
    m_JustImportedSpec = False
    
    If Not m_UseSpec Then Call AddInDefaultStatics(frmdlim.Name)  'MPSMarch
    Call ResetFlexGrids
    
    TestImportFlags = FIELD_NAME + ROW_LINENUMBERS
    ReCalc_DLim = DoImport(frmdlim.FlG_Source, TestImportFlags, True)
  End If
  
  Call FrameEnable(frmdlim, frmdlim.Fra_Omit, Not m_Imported)
  frmdlim.Fra_Omit.Enabled = Not m_Imported
  Call FrameEnable(frmdlim, frmdlim.Fra_Delimiters, Not m_Imported)
  frmdlim.Fra_Delimiters.Enabled = Not m_Imported
  frmdlim.Cbo_TextQual.Enabled = Not m_Imported
  frmdlim.Lbl_TextQual.Enabled = Not m_Imported
  frmdlim.Fra_FieldNames.Enabled = Not m_Imported
  
ReCalc_DLim_End:
  Call xReturn("ReCalc_DLim")
  Exit Function

ReCalc_DLim_Err:
  ReCalc_DLim = False
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_DLim", "Recalculate Delimiter", "Unable to Recalculate Delimiter")
  Resume ReCalc_DLim_End
End Function

Friend Function ReCalc_FW() As Boolean
  Dim FrmFW As Frm_FW
  Dim ForceInit As Boolean

  On Error GoTo ReCalc_FW_Err
  Call xSet("ReCalc_FW")
  If mDisableRecalc Then GoTo ReCalc_FW_End
  Set FrmFW = m_ImportForms.Item(CStr(TCSIMP_FW))
  
  FrmFW.Caption = m_TitlePrefix & "Fixed Width Source File"
  
  ForceInit = False
  FrmFW.UpD_Omit(0).Max = m_NumLines - 1
  FrmFW.UpD_Omit(1).Max = m_NumLines - 1

  If ImpParent.HeaderCount <> FrmFW.Txt_Omit(0) Or ImpParent.FooterCount <> FrmFW.Txt_Omit(1) Or UBound(m_LineArray) <= 0 Then
    ForceInit = True
    ImpParent.HeaderCount = FrmFW.Txt_Omit(0)
    ImpParent.FooterCount = FrmFW.Txt_Omit(1)
    m_FirstLine = CopyLines(m_SourceFile, m_LineArray(), ImpParent.HeaderCount + 1, m_PreviewLines)
  End If
  If m_FieldSpecs.Count = 0 Then ForceInit = True
  Call ReCalc_FW2(FrmFW, ForceInit)

ReCalc_FW_End:
  Call xReturn("ReCalc_FW")
  Exit Function

ReCalc_FW_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_FW", "Recalc Fixed width columns", "Unable to recalculate fixed width columns on reimport.")
  Resume ReCalc_FW_End
End Function

Friend Function ReCalc_FW2(ByVal FrmFW As Frm_FW, ByVal ForceInit As Boolean) As Boolean
  Dim v As Variant
  Dim isp As ImportSpec
  Dim i As Long, nMaxFields As Long
  On Error GoTo ReCalc_FW2_Err
  Call xSet("ReCalc_FW2")
  If mDisableRecalc Then GoTo ReCalc_FW2_End
  If FrmFW Is Nothing Then Set FrmFW = m_ImportForms.Item(CStr(TCSIMP_FW))
  If ForceInit Then
    FrmFW.FW_FWGrid.UsePointers = True
    FrmFW.FW_FWGrid.LinesCopied = UBound(m_LineArray)
    FrmFW.FW_FWGrid.LinesIn = CVar(m_LineArray)
    FrmFW.FW_FWGrid.CharsWide = m_MaxLineLength
    FrmFW.FW_FWGrid.FWFontName = "Courier"
    FrmFW.FW_FWGrid.FWFontSize = 10
    FrmFW.FW_FWGrid.Refresh
  End If
  v = Frm_FW.FW_FWGrid.Lines
  Call m_FieldSpecs.RemoveAll
  Call m_SavedFieldSpecs.RemoveAll
  nMaxFields = UBound(v, 2)
  For i = 1 To nMaxFields
    Set isp = m_FieldSpecs.Add
    isp.key = KEY_FIELD
    isp.vartype = TYPE_STR
    isp.FldName = FieldName(i, nMaxFields)
    isp.FieldKey = i
    isp.Offset = v(1, i)
    isp.Length = v(2, i)
  Next i
  If Not m_UseSpec Then Call AddInDefaultStatics(FrmFW.Name)  'MPSMarch
  Call ResetFlexGrids
  
ReCalc_FW2_End:
  Call xReturn("ReCalc_FW2")
  Exit Function

ReCalc_FW2_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_FW2", "Recalc Fixed width columns", "Unable to recalculate fixed width columns.")
  Resume ReCalc_FW2_End
End Function

' If FieldKey < 0 then create a static column else take copy of column
Friend Function CopyColumn(ByVal FieldKey As Long) As Boolean
  Dim isp As ImportSpec, SrcIsp As ImportSpec
  Dim NextKey As Long, i As Long, lastcolumn As Long
  Dim frmMisc As Frm_Misc

  On Error GoTo CopyColumn_Err
  Call xSet("CopyColumn")
  NextKey = m_FieldSpecs.Count + 1
  For i = 1 To m_FieldSpecs.Count
    If FieldKey = m_FieldSpecs(i).FieldKey Then
      Set SrcIsp = m_FieldSpecs(i)
      Exit For
    End If
  Next i
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  Call SetFlexRowCol(frmMisc.FlG_Source, SET_COLUMN, m_FieldSpecs.Count + 1 + FIXED_COLCOUNT, FIXED_COLCOUNT)
  For i = 1 To m_FieldSpecs.Count
    Set isp = m_FieldSpecs(i)
    frmMisc.FlG_Source.Coldata(i + FIXED_COLCOUNT - 1) = isp.FieldKey
  Next i
  lastcolumn = frmMisc.FlG_Source.cols - 1
  frmMisc.FlG_Source.Col = lastcolumn
  frmMisc.FlG_Source.Coldata(lastcolumn) = NextKey
  frmMisc.FlG_Source.Row = FIXED_ROWCOUNT
  'frmMisc.FlG_Source.RowSel = frmMisc.FlG_Source.Rows - 1
  Set isp = m_FieldSpecs.Add
  isp.FieldKey = NextKey
  isp.FldName = FieldName(NextKey, NumSourceFields)
   
  If SrcIsp Is Nothing Then ' Static column
    isp.key = KEY_CONSTANT
    isp.vartype = TYPE_STR
        
    frmMisc.FlG_Source.Enabled = False
    frmMisc.Cmd_Cancel.Enabled = False
    frmMisc.Cmd_Back.Enabled = False
    frmMisc.Cmd_Next.Enabled = False
    frmMisc.Lbl_Misc1.Visible = False
    frmMisc.Fra_Static.Visible = True
    frmMisc.Fra_Static.ZOrder
    frmMisc.Txt_Static.Text = ""
  Else ' Copy column
    isp.key = KEY_COPIEDFIELD
    Set isp.CopyField = SrcIsp
    isp.CopyFieldKey = FieldKey
    isp.vartype = SrcIsp.vartype
  End If
  frmMisc.FlG_Source.RowSel = frmMisc.FlG_Source.Rows - 1
  Call ReCalc_Misc(False)
  frmMisc.Cbo_Static.Text = frmMisc.Cbo_Static.List(0) 'MPSMarch2
  frmMisc.Opt_Static(0).Value = True  'MPSMarch2
  frmMisc.Txt_Static.SetFocus 'MPSMarch2 - Moved from above
  
CopyColumn_End:
  Call xReturn("CopyColumn")
  Exit Function

CopyColumn_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CopyColumn", "Copy Columns", "Unable to copy columns.")
  Resume CopyColumn_End
End Function

Friend Function DeleteColumn(ByVal FieldKey As Long) As Boolean
  Dim i As Long
  Dim frmMisc As Frm_Misc
  
  On Error GoTo DeleteColumn_Err
  Call xSet("DeleteColumn")
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  
  For i = 1 To m_FieldSpecs.Count
    If m_FieldSpecs(i).FieldKey = FieldKey Then
      If m_FieldSpecs(i).key = KEY_FIELD Then
        MsgBox "Selected field comes from the the source file:" & vbCrLf & "It cannot be deleted."
      Else
        Call m_FieldSpecs.Remove(i)
      End If
      Exit For
    End If
  Next i
  Call ReCalc_Misc(True)

DeleteColumn_End:
  Call xReturn("DeleteColumn")
  Exit Function

DeleteColumn_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DeleteColumn", "Delete Columns", "Unable to delete column: " & FieldKey & ".")
  Resume DeleteColumn_End
End Function


Friend Sub ReCalc_Misc(ByVal ForceRecalc As Boolean)
  Dim i As Long, isp As ImportSpec
  Dim frmMisc As Frm_Misc
  Dim TestImportFlags As DOIMPORT_FLAGS
  Dim s As String, b As String
  
  On Error GoTo ReCalc_Misc_Err
  Call xSet("ReCalc_Misc")
  If mDisableRecalc Then GoTo ReCalc_Misc_End
    
  b = "*" 'Chr$(149)
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  frmMisc.Caption = m_TitlePrefix & "Field Manipulation"
  s = "Select a field or fields and click the right mouse button to see a menu of options. You can:"
  s = s & vbCrLf & b & "  Add or delete a 'static' field (i.e. a field containing a fixed value such as some text or a number)"
  s = s & vbCrLf & b & "  Make a copy of an existing field"
  If m_AllowMoveColumns Then s = s & vbCrLf & b & "  Rearrange the order of fields"
  If m_AllowMergeColumns Then s = s & vbCrLf & b & "  Merge adjacent fields together, or separate already merged fields"
  frmMisc.Lbl_Misc1.Caption = s
  If frmMisc.Fra_Static.Visible = True Then
    frmMisc.Lbl_Misc1.Visible = False
  Else
    frmMisc.Lbl_Misc1.Visible = True
  End If
  
  frmMisc.HSc_FGLeft.Visible = m_AllowMoveColumns
  frmMisc.HSc_Right.Visible = m_AllowMoveColumns
  If (frmMisc.FlG_Source.cols = 0) Or ForceRecalc Then
    Call SetFlexRowCol(frmMisc.FlG_Source, SET_COLUMN, m_FieldSpecs.Count + FIXED_COLCOUNT, FIXED_COLCOUNT)
    For i = 1 To m_FieldSpecs.Count
      Set isp = m_FieldSpecs(i)
      frmMisc.FlG_Source.Coldata(i + FIXED_COLCOUNT - 1) = isp.FieldKey
    Next i
  End If
  TestImportFlags = FIELD_NAME + FIELD_TYPE + ROW_LINENUMBERS
  'CAD
  Call DoImport(frmMisc.FlG_Source, TestImportFlags, True)
  Call FormatField
  frmMisc.Lbl_Misc1.Enabled = Not m_Imported

ReCalc_Misc_End:
  Call xReturn("ReCalc_Misc")
  Exit Sub

ReCalc_Misc_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_Misc", "Recalculate Miscellaneous", "Unable to recalculate form.")
  Resume ReCalc_Misc_End
End Sub

Private Sub FileInfo()
  Dim rf As New TCSFileread
    
  Call xSet("FileInfo")
  If Not rf.OpenFile(m_SourceFile) Then Err.Raise ERR_IMPORT, "ImportFile", "Unable to open file " & m_SourceFile
  m_NumLines = rf.LineCount
  m_MaxLineLength = rf.MaxLineLength
  Set rf = Nothing
End Sub

Private Function GetDestTableInfo(ByVal frmDest As Frm_Dest) As Boolean
  Dim drs As DestRecordSet
  
  On Error GoTo GetDestTableInfo_Err
  Call xSet("GetDestTableInfo")
  GetDestTableInfo = False
  If m_DestRecSets.Count = 0 Then
    Call Err.Raise(ERR_NO_DESTRS, "GetDestTableInfo", "No usable destinations have been specified.")
    GoTo GetDestTableInfo_End
  End If
  
  frmDest.Cbo_Table.Clear
  For Each drs In m_DestRecSets
    frmDest.Cbo_Table.AddItem drs.DisplayName
  Next drs
  frmDest.Cbo_Table.Text = frmDest.Cbo_Table.List(0)
  If m_DestRecSets.Count > 1 Then
    frmDest.Lbl_TableInst2 = "Use the selector if you wish to choose a different recordset."
  Else
    frmDest.Lbl_TableInst2 = ""
  End If
  GetDestTableInfo = True
  
GetDestTableInfo_End:
  Call xReturn("GetDestTableInfo")
  Exit Function

GetDestTableInfo_Err:
  GetDestTableInfo = False
  Call ErrorMessage(ERR_ERROR, Err, "GetDestTableInfo", "Get all import destinations", "Unable to retrieve all the import destinations.")
  Resume GetDestTableInfo_End
End Function

Friend Sub SaveSpec(ByVal FileName As String)
  Dim i As Long, hFile As Integer
  Dim isp As ImportSpec
  Dim s As String, ss As String
  
  On Error GoTo SaveSpec_Err
  Call xSet("SaveSpec")
  hFile = -1
  hFile = FreeFile
  Open FileName For Output As #hFile
  
  If Not CurrentDest Is Nothing Then Print #hFile, "(DATASOURCE=" & CurrentDest.Name & ")"
  If ImpParent.ImportType = IMPORT_DELIMITED Then
    Print #hFile, "(DELIMITEDIMPORT)"
  Else
    Print #hFile, "(FIXEDWIDTHIMPORT)"
  End If
  
  Print #hFile, "(DELIMITER=" & AddEscapeChars(ImpParent.Delimiter) & ")"
  If ImpParent.EscapeChar = Chr$(34) Then
    Print #hFile, "(ESCAPECHAR=" & String$(4, Chr$(34)) & ")"
  Else
    Print #hFile, "(ESCAPECHAR=" & AddEscapeChars(ImpParent.EscapeChar) & ")"
  End If
  Print #hFile, "(SKIPHEADER=" & ImpParent.HeaderCount & ")"
  Print #hFile, "(SKIPFOOTER=" & ImpParent.FooterCount & ")"
  If ImpParent.TakeFieldNamesFromLine > 0 Then
    Print #hFile, "(FIELDNAMESLINE=" & ImpParent.TakeFieldNamesFromLine & ")"
  End If
  If ImpParent.RightTrimFields Then
    Print #hFile, "(RIGHTTRIMFIELDS)"
  End If
  For i = 1 To m_FieldSpecs.Count
    Set isp = m_FieldSpecs(i)
    s = isp.GetColumn(ImpParent.ImportType)
    If Len(s) > 0 Then Print #hFile, s
  Next i
  m_SavedSpec = True
  m_SpecFile = FileName
SaveSpec_End:
  If hFile > 0 Then Close #hFile
  Call xReturn("SaveSpec")
  Exit Sub

SaveSpec_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SaveSpec", "Save Import Specification", "Error saving import specification to " & vbCrLf & FileName & vbCrLf)
  Resume SaveSpec_End
End Sub

Friend Sub LoadSpec(Optional ByVal SpecFile As String)
  Dim i As Long, j As Long
  Dim FrmSrc As Frm_Source
  Dim frmdlim As Frm_DLim
  Dim FrmFW As Frm_FW
  Dim frmlink As Frm_Link
  Dim isp As ImportSpec
  
  On Error GoTo LoadSpec_Err
  Call xSet("LoadSpec")
  Set frmdlim = m_ImportForms.Item(CStr(TCSIMP_DLIM))
  Set FrmFW = m_ImportForms.Item(CStr(TCSIMP_FW))
  Set FrmSrc = m_ImportForms.Item(CStr(TCSIMP_SOURCE))
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
 
  If Len(SpecFile) = 0 Then SpecFile = FileOpenDlgFilter(0, "Choose a Specification File to Open", "Specification Files (*.imp)|*.imp|All Files (*.*)|*.*", m_SpecPath)
  If Len(SpecFile) > 0 Then
    Call SplitPath(SpecFile, m_SpecPath)
    m_SpecFile = SpecFile 'MPSMarch
    If ImpParent.InitImport(m_SourceFile, [_IMPORT_UNDEFINED], SpecFile, ImpParent.UpdateType) Then
      mDisableRecalc = True
      If ImpParent.TakeFieldNamesFromLine >= 1 And ImpParent.TakeFieldNamesFromLine <= m_NumLines Then
        frmdlim.Txt_FieldNames.Text = Trim$(CStr(ImpParent.TakeFieldNamesFromLine))
      Else
        frmdlim.Txt_FieldNames.Text = "0"
      End If
      If ImpParent.TakeFieldNamesFromLine > 0 Then
        frmdlim.Chk_FieldNames.Value = vbChecked
      Else
        frmdlim.Chk_FieldNames.Value = vbUnchecked
      End If
      If ImpParent.HeaderCount + ImpParent.FooterCount < m_NumLines Then
        frmdlim.Txt_Omit(0).Text = Trim$(CStr(ImpParent.HeaderCount))
        frmdlim.Txt_Omit(1).Text = Trim$(CStr(ImpParent.FooterCount))
        FrmFW.Txt_Omit(0).Text = Trim$(CStr(ImpParent.HeaderCount))
        FrmFW.Txt_Omit(1).Text = Trim$(CStr(ImpParent.FooterCount))
      Else
        MsgBox "Header and footer count in the specification are not compatible with the source file: They will be set to zero"
        frmdlim.Txt_Omit(0).Text = "0"
        frmdlim.Txt_Omit(1).Text = "0"
        FrmFW.Txt_Omit(0).Text = "0" 'Trim$(CStr(ImpParent.HeaderCount))
        FrmFW.Txt_Omit(1).Text = "0" 'Trim$(CStr(ImpParent.FooterCount))
      End If
      mDisableRecalc = False
      
      Call ReCalc_Dest(Me.CurrentDest) 'MPSMarch
      
      Call ReCalc_Src(Nothing)
      m_JustImportedSpec = True
      m_SpecLoadedOK = True
      m_FieldSpecs.RemoveAll
      For i = 1 To ImpParent.Icols.Count
        Set isp = m_FieldSpecs.Add
        Call ImpParent.Icols(i).Copy(isp)
        'If isp.key <> KEY_CONSTANT And isp.key <> KEY_COPIEDFIELD Then
        '  isp.key = KEY_FIELD
        'End If
        'MPSMarch2
        Select Case isp.key
          Case KEY_CONSTANT, KEY_COPIEDFIELD, KEY_DATENOW, KEY_FILENAME, KEY_FILEPATH, KEY_FILEDATE, KEY_LINENUMBER, KEY_CFGFILENAME, KEY_USERNAME, KEY_IMPDATE
          Case Else
            isp.key = KEY_FIELD
        End Select
      Next i
      Call FixUpCopyFields(m_FieldSpecs)
      If ImpParent.ImportType = IMPORT_DELIMITED Then
        Call SetFlexRowCol(frmdlim.FlG_Source, SET_COLUMN, ImpParent.Icols.Count + FIXED_COLCOUNT, FIXED_COLCOUNT)
        For i = 1 To m_FieldSpecs.Count
          Set isp = m_FieldSpecs(i)
          frmdlim.FlG_Source.Coldata(i + FIXED_COLCOUNT - 1) = isp.FieldKey
          If Len(isp.DestField) > 0 Then
            For j = 1 To CurrentDest.rs.Fields.Count
              If StrComp(isp.DestField, CurrentDest.rs.Fields(j - 1).Name, vbTextCompare) = 0 Then
                frmlink.FlG_DestInfo.TextMatrix(j, 2) = isp.FldName
                Exit For
              End If
            Next j
            'MPSMarch
            If j = (CurrentDest.rs.Fields.Count + 1) Then
              Call ErrorMessage(ERR_INFO, Nothing, "LoadSpec", "Load specification", "The destination field '" & isp.DestField & "' in the specification file does not exist in the dataset '" & CurrentDest.DisplayName & "'." & vbCrLf & vbCrLf & "IMPORT ABORTED" & vbCrLf)
              m_SpecLoadedOK = False
              m_JustImportedSpec = False
              Exit For
            End If
          End If
        Next i
        Call frmdlim.SetDelimsEscChar(ImpParent.Delimiter, ImpParent.EscapeChar)
        ReDim m_LineArray(0)
        m_UseSpec = m_JustImportedSpec 'MPSMarch
                
        If Not ReCalc_DLim(True) Then
          m_JustImportedSpec = False
          m_SpecLoadedOK = False
          Call ReCalc_DLim(True)
        End If
        Call ReCalc_Misc(True)
        Call ReCalc_Link
  
        'MPSMarch
        'Call AddInDefaultCopiedFields(frmlink.Name)
        'Call AddInDefaultConstraints(frmlink.Name)
        'Call AddInDefaultValues(frmlink.Name)
        'Call ReCalc_Link 'MPSMarch
      Else
        Call ECASE("Fixed Width specification file not yet implemented!")
        m_SpecLoadedOK = False
      End If
      If m_SpecLoadedOK Then Call SwitchForm(FrmSrc, TCSIMP_LINK)
    End If
  End If

LoadSpec_End:
  m_UseSpec = m_SpecLoadedOK  'MPSMarch
  mDisableRecalc = False
  Call xReturn("LoadSpec")
  Exit Sub

LoadSpec_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadSpec", "Load specification", "Error loading specification file " & SpecFile)
  Resume LoadSpec_End
End Sub

Friend Sub ReCalc_Link(Optional ByVal ReMakeLinkedFieldList As Boolean = False)
  Dim frmlink As Frm_Link, frmMisc As Frm_Misc
  Dim isp As ImportSpec
  Dim i As Long, j As Long
  Dim colnum As Long, RowNum As Long
  Dim TestImportFlags As DOIMPORT_FLAGS
  Const Shift As Long = 3360 - 840
  
  On Error GoTo ReCalc_Link_Err
  Call xSet("ReCalc_Link")
  If mDisableRecalc Then GoTo ReCalc_Link_End
  Call SetCursor
  
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  
  frmlink.Caption = m_TitlePrefix & "Link Fields"
  
  frmlink.FlG_Source.Redraw = False
  frmlink.FlG_DestInfo.Redraw = False
  frmlink.FlG_Source.HighLight = flexHighlightNever
  frmlink.FlG_DestInfo.HighLight = flexHighlightNever
    
  colnum = frmlink.FlG_Source.Col
  If colnum < 1 Then colnum = 1
  RowNum = frmlink.FlG_DestInfo.Row
  If RowNum < 1 Then RowNum = 1
   
  If frmlink.FlG_Source.cols <> m_FieldSpecs.Count + FIXED_COLCOUNT Then
    Call SetFlexRowCol(frmlink.FlG_Source, SET_COLUMN, m_FieldSpecs.Count + FIXED_COLCOUNT, FIXED_COLCOUNT)
  End If
  For i = 1 To m_FieldSpecs.Count
    frmlink.FlG_Source.Coldata(i + FIXED_COLCOUNT - 1) = frmMisc.FlG_Source.Coldata(i + FIXED_COLCOUNT - 1)
  Next i
        
  TestImportFlags = FIELD_NAME + FIELD_TYPE + LINKED_NAME + ROW_LINENUMBERS
  Call DoImport(frmlink.FlG_Source, TestImportFlags, True)
  
  For i = 1 To frmlink.FlG_DestInfo.Rows - 1
    frmlink.FlG_DestInfo.Row = i
    frmlink.FlG_DestInfo.Col = 0
    frmlink.FlG_DestInfo.ColSel = frmlink.FlG_DestInfo.cols - 1
    frmlink.FlG_DestInfo.TextMatrix(i, 2) = ""
    For j = 1 To m_FieldSpecs.Count
      Set isp = m_FieldSpecs(j)
      If StrComp(isp.DestField, frmlink.FlG_DestInfo.TextMatrix(i, 0), vbTextCompare) = 0 Then
        frmlink.FlG_DestInfo.TextMatrix(i, 2) = isp.FldName
        Exit For
      End If
    Next j
    If Len(frmlink.FlG_DestInfo.TextMatrix(i, 2)) = 0 Then
      frmlink.FlG_DestInfo.CellBackColor = 0
    Else
      frmlink.FlG_DestInfo.CellBackColor = QBColor(8)
    End If
  Next i
  
  frmlink.FlG_DestInfo.Row = RowNum
  frmlink.FlG_DestInfo.Col = 0
  frmlink.FlG_DestInfo.ColSel = frmlink.FlG_DestInfo.cols - 1
  For i = 1 To m_FieldSpecs.Count
    frmlink.FlG_Source.Col = GetFlexColumn(frmlink.FlG_Source, m_FieldSpecs(i).FieldKey)
    frmlink.FlG_Source.Row = 0
    frmlink.FlG_Source.RowSel = frmlink.FlG_Source.Rows - 1
    If Len(m_FieldSpecs(i).DestField) = 0 Then
      frmlink.FlG_Source.CellBackColor = 0
    Else
      frmlink.FlG_Source.CellBackColor = QBColor(8)
    End If
  Next i
  
  frmlink.FlG_Source.Col = colnum
  frmlink.FlG_Source.Row = LINK_FIXED_ROWCOUNT
  frmlink.FlG_Source.ColSel = colnum
  frmlink.FlG_Source.RowSel = frmlink.FlG_Source.Rows - 1

  If m_ShowSimpleLinkView And frmlink.Fra_Link1.Top < frmlink.FlG_Source.Top Then
    frmlink.Fra_Link1.Visible = False
    frmlink.FlG_Source.Top = frmlink.FlG_Source.Top - Shift
    frmlink.FlG_Source.Height = frmlink.FlG_Source.Height + Shift
    frmlink.Lbl_DelimitedFile.Top = frmlink.Lbl_DelimitedFile.Top - Shift
    frmlink.Cbo_Link.Top = frmlink.Cbo_Link.Top - Shift
  End If
  If Not m_ShowSimpleLinkView And frmlink.Fra_Link1.Top > frmlink.FlG_Source.Top Then
    frmlink.Fra_Link1.Visible = True
    frmlink.FlG_Source.Top = frmlink.FlG_Source.Top + Shift
    frmlink.FlG_Source.Height = frmlink.FlG_Source.Height - Shift
    frmlink.Lbl_DelimitedFile.Top = frmlink.Lbl_DelimitedFile.Top + Shift
    frmlink.Cbo_Link.Top = frmlink.Cbo_Link.Top + Shift
  End If
  Call ReCalc_Link2(ReMakeLinkedFieldList)

ReCalc_Link_End:
  Call ClearCursor
  Call xReturn("ReCalc_Link")
  Exit Sub

ReCalc_Link_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_Link", "Recalculate links", "Unable to recalculate field links.")
  Resume ReCalc_Link_End
End Sub

Friend Sub ReCalc_Link2(ByVal ReMakeLinkedFieldList As Boolean)
  Dim i As Long, j As Long, k As Long
  Dim LinkedTo As Long
  Dim ColNo As Long
  Dim CWFactor As Double
  Dim ItemStr() As String
  Dim DestFieldLinkedTo As String
  Dim frmlink As Frm_Link
  Dim fld As Field
  Dim Linked As Boolean, ReqFieldsLinked As Boolean
  Dim ReqField As Boolean, PKeyField As Boolean
  Dim s As String
  
  On Error GoTo ReCalc_Link2_Err
  Call xSet("ReCalc_Link2")
  If mDisableRecalc Then GoTo ReCalc_Link2_End
  Call SetCursor
  
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  frmlink.FlG_Source.Redraw = False
  frmlink.FlG_DestInfo.Redraw = False
  frmlink.FlG_Source.HighLight = flexHighlightNever
  frmlink.FlG_DestInfo.HighLight = flexHighlightNever
   
  ReqFieldsLinked = True
  ReqField = False
  PKeyField = False
  For i = 1 To frmlink.FlG_DestInfo.Rows - 1
    If CurrentDest.RequiredFields.IsPresent(frmlink.FlG_DestInfo.TextMatrix(i, 0)) Then
      frmlink.FlG_DestInfo.TextMatrix(i, 3) = "*"
      If frmlink.FlG_DestInfo.TextMatrix(i, 2) = "" Then ReqFieldsLinked = False
      ReqField = True
    End If
    If CurrentDest.PrimaryKeys.IsPresent(frmlink.FlG_DestInfo.TextMatrix(i, 0)) Then
      frmlink.FlG_DestInfo.TextMatrix(i, 3) = "**"
      frmlink.FlG_DestInfo.ToolTipText = "Fields marked with * in the key field column MUST have a source field linked to them."
      PKeyField = True
    End If
  Next i
  
  s = ""
  
  If ReqField = True Then s = s & "Fields marked with * in the key field column MUST have a source field linked to them."
  If PKeyField = True And s <> "" Then s = s & " "
  If PKeyField = True Then s = s & "Fields marked with ** in the key field column are primary key fields."
  
  frmlink.FlG_DestInfo.ToolTipText = s
  frmlink.Cmd_Finish.Enabled = (ReqFieldsLinked And (FieldsLinkCount(m_FieldSpecs, True) > 0))
  frmlink.Cmd_ShowLinked.Enabled = (FieldsLinkCount(m_FieldSpecs, False) > 0)
  
  If frmlink.Fra_Link1.Top < frmlink.FlG_Source.Top Then
    frmlink.Lbl_Inst.Caption = "Link a source field to a destination field by using the drop-down selector in the source field display, or by dragging a source field to a destination field." & vbCrLf & "Break a link by using the drop-down selector or by double-clicking on the source field."
  Else
    frmlink.Lbl_Inst.Caption = "Link a source field to a destination field by using the drop-down selector in the source field display." & vbCrLf & "Break a link by using the drop-down selector or by double-clicking on the source field."
  End If
   
  DestFieldLinkedTo = m_FieldSpecs(frmlink.FlG_Source.Coldata(Frm_Link.FlG_Source.Col)).DestField
  
  With frmlink
    .Cbo_Link.Visible = False
    .Cbo_Link.Top = .FlG_Source.Top '- .FlG_Source.RowHeight(0)
    .Cbo_Link.Left = .FlG_Source.Left + .FlG_Source.ColPos(.FlG_Source.Col) + 2 * Screen.TwipsPerPixelX
    .Cbo_Link.Width = .FlG_Source.ColWidth(.FlG_Source.Col) + 2 * Screen.TwipsPerPixelX
    If .Cbo_Link.Left >= .FlG_Source.Left + .FlG_Source.ColWidth(0) And .Cbo_Link.Left + .Cbo_Link.Width <= .FlG_Source.Left + .FlG_Source.Width And .Cmd_ShowLinked.Caption = "Show Only Linked Fields" Then
      .Cbo_Link.Visible = True
    End If
  End With
  
  ReDim ItemStr(frmlink.FlG_DestInfo.Rows - 1)
  
  If Len(DestFieldLinkedTo) > 0 Then
    ItemStr(0) = "LINKED to " & DestFieldLinkedTo
    ItemStr(1) = "UNLINK to " & DestFieldLinkedTo
    j = 2
  Else
    ItemStr(0) = "UNLINKED"
    j = 1
    For Each fld In CurrentDest.rs.Fields
      Linked = False
      For i = 1 To m_FieldSpecs.Count
        If StrComp(m_FieldSpecs(i).DestField, fld.Name, vbTextCompare) = 0 Then
          Linked = True
          Exit For
        End If
      Next i
      If Linked = False Then
        If Not CurrentDest.HiddenFields.IsPresent(fld.Name) Then
          ItemStr(j) = "LINK to " & fld.Name
          j = j + 1
        End If
      End If
    Next fld
  End If
            
  With frmlink.Cbo_Link
                        
  If .ListCount <> j Then
    ReMakeLinkedFieldList = True
  Else
    For i = 0 To j - 1
      If ItemStr(i) <> .List(i) Then
        ReMakeLinkedFieldList = True
        Exit For
      End If
    Next i
  End If
  
  If ReMakeLinkedFieldList Then
    .Clear
    For i = 0 To j - 1
      .AddItem ItemStr(i)
    Next i
    .Text = .List(0)
  End If
  
  End With
  Call DisplayDField
  
  frmlink.FlG_Source.Refresh
  frmlink.FlG_DestInfo.Refresh
  frmlink.FlG_Source.HighLight = flexHighlightAlways
  frmlink.FlG_DestInfo.HighLight = flexHighlightAlways
  frmlink.FlG_Source.Redraw = True
  frmlink.FlG_DestInfo.Redraw = True
  
  frmlink.Cbo_Link.Enabled = Not m_Imported
  frmlink.Lbl_Inst.Enabled = Not m_Imported
  
  Call ClearCursor

ReCalc_Link2_End:
  Call xReturn("ReCalc_Link2")
  Exit Sub

ReCalc_Link2_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_Link2", "Re Calc Link 2", "Error racalulating the links.")
  Resume ReCalc_Link2_End
End Sub

Friend Sub SetStatic(ByVal colnum As Long, ByVal Value As String)
  m_FieldSpecs(colnum).actionparam = Value
End Sub

'MPSMarch2
Friend Sub SetSpecialFieldKey(colnum As Long, SpecialFieldKey As IMPORTFIELD_KEY)
  m_FieldSpecs(colnum).key = SpecialFieldKey
End Sub

Friend Sub ToggleShowOnlyLinked()
  Dim i As Long
  Dim FirstLinkedCol As Long
  Dim NumNotLinked As Long
  Dim frmlink As Frm_Link
    
  On Error GoTo ToggleShowOnlyLinked_Err
  Call xSet("ToggleShowOnlyLinked")
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  With frmlink
    If .Cmd_ShowLinked.Caption = "Show Only Linked Fields" Then
      If FieldsLinkCount(m_FieldSpecs, False) > 0 Then
        For i = 1 To m_FieldSpecs.Count
          If (Len(m_FieldSpecs(i).DestField) > 0) And Not m_FieldSpecs(i).Hide Then
            FirstLinkedCol = i
            Exit For
          End If
        Next i
        .FlG_Source.Redraw = False
        For i = 1 To m_FieldSpecs.Count
          If Len(m_FieldSpecs(i).DestField) = 0 Then .FlG_Source.ColWidth(i) = 0
        Next i
        .FlG_Source.Col = FirstLinkedCol
        .FlG_Source.ColSel = FirstLinkedCol
        .FlG_Source.Row = LINK_FIXED_ROWCOUNT
        .FlG_Source.RowSel = .FlG_Source.Rows - 1
        .FlG_Source.LeftCol = .FlG_Source.Col
        .FlG_Source.Redraw = True
        .Cmd_ShowLinked.Caption = "Show All Fields"
        .Cbo_Link.Visible = False
        frmlink.Lbl_Inst.Caption = "Press the Show All Fields button to modify links between source and destination fields."
      End If
    Else
      .Cmd_ShowLinked.Caption = "Show Only Linked Fields"
      Call ReCalc_Link(True)
    End If
  End With

ToggleShowOnlyLinked_End:
  Call xReturn("ToggleShowOnlyLinked")
  Exit Sub

ToggleShowOnlyLinked_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ToggleShowOnlyLinked", "Toggle Show Only Linked", "Error displaying the linked/unlinked fields.")
  Resume ToggleShowOnlyLinked_End
End Sub

Friend Sub SetupFlexGrid(ByVal FGrid As MSFlexGrid, ByVal NumCols As Long, ByVal NumRows As Long, ByVal NumHeaderCols As Long, ByVal NumHeaderRows As Long, CWidth() As Long, HideCol() As Boolean)

  Dim i As Long
  
  On Error GoTo SetupFlexGrid_Err
  Call xSet("SetupFlexGrid")
    
  Call SetFlexRowCol(FGrid, SET_ROW, NumRows + NumHeaderRows, NumHeaderRows)
  Call SetFlexRowCol(FGrid, SET_COLUMN, NumCols + NumHeaderCols, NumHeaderCols)
  With FGrid
    For i = 0 To .cols - 1
      .ColWidth(i) = CWidth(i) * ((Not HideCol(i)) * -1)
    Next
  End With

SetupFlexGrid_End:
  Call xReturn("SetupFlexGrid")
  Exit Sub

SetupFlexGrid_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetupFlexGrid", "Setup Flex Grid", "Error setting up the flex grid (columns and widths)")
  Resume SetupFlexGrid_End
End Sub

Friend Sub MoveSourceCol(ByVal FromCol As Long, ByVal ToCol As Long)

  Dim i As Long, j As Long, k As Long
  Dim frmMisc As Frm_Misc
  
  On Error GoTo MoveSourceCol_Err
  Call xSet("MoveSourceCol")

  If ToCol = FromCol Then GoTo MoveSourceCol_End
  
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  
  j = frmMisc.FlG_Source.Coldata(FromCol)
  k = Sgn(ToCol - FromCol)
  For i = FromCol + k To ToCol Step k
    frmMisc.FlG_Source.Coldata(i - k) = frmMisc.FlG_Source.Coldata(i)
  Next i
  frmMisc.FlG_Source.Coldata(ToCol) = j
  
  
  Call ReCalc_Misc(False)
  
MoveSourceCol_End:
  Call xReturn("MoveSourceCol")
  Exit Sub

MoveSourceCol_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MoveSourceCol", "Move Source Col", "Error moving the source cell.")
  Resume MoveSourceCol_End
End Sub

Friend Sub ReCalc_End(Optional ByVal FullRecalc As Boolean = False)
  Dim i As Long
  Dim s As String
  Dim frmEnd As Frm_End
  Dim NewNumDestRecs As Long
  
  On Error GoTo ReCalc_End_Err
  Call xSet("ReCalc_End")
  If mDisableRecalc Then GoTo ReCalc_End_End
  Set frmEnd = m_ImportForms.Item(CStr(TCSIMP_END))
  frmEnd.Caption = m_TitlePrefix & "Import Data"
  
  If FullRecalc Then
    Call SetDestDataGrid(CurrentDest.rs, TCSIMP_DEST)
    Call SetDestDataGrid(CurrentDest.rs, TCSIMP_LINK)
    NewNumDestRecs = SetDestDataGrid(CurrentDest.rs, TCSIMP_END)
    
    Call UpdateImportDescription(False)
    
    m_NumDestRecs = NewNumDestRecs
    frmEnd.Cmd_Another.Visible = m_AllowMultipleImports
  End If
  
  frmEnd.Cmd_Import.Enabled = Not m_Imported
  
  If m_Imported Then
    If frmEnd.FlG_Dest.Rows > 0 Then
      s = "The data in the source file has been imported into the database table shown below." & vbCrLf 'MPSMarch
    Else
      s = "The data in the source file has been imported."
    End If
    frmEnd.Cmd_Cancel.Caption = "Exit"
  Else
    s = "Press the Import button to import the data in the source file into the database shown below."
    frmEnd.Cmd_Cancel.Caption = "Cancel"
  End If
  s = s & vbCrLf '& vbCrLf  'MPSMarch
  s = s & "Destination recordset:  " & CurrentDest.DisplayName & vbCrLf & vbCrLf
  s = s & "Source file:  " & m_SourceFile & vbCrLf
  
  frmEnd.Lbl_Info1.Caption = s
  
  If (ImpParent.UpdateType And NO_UPDATES) = NO_UPDATES Then
    frmEnd.Cmd_Import.Caption = "Import"
  Else
    frmEnd.Cmd_Import.Caption = "Update Import"
  End If
  If Not ImpParent.ErrorFilter Is Nothing And ((ImpParent.ProcessingLinesInError > 0) Or (m_PostProcessingLinesInError > 0)) Then
    frmEnd.Cmd_ErrPrint.Visible = ImpParent.ErrorFilter.PrintAvailable
    frmEnd.Cmd_ErrView.Visible = ImpParent.ErrorFilter.ViewAvailable
  End If
  
  If Not m_NoReviewSpec And m_Imported Then
    frmEnd.Cmd_Back.Enabled = False
  Else
    frmEnd.Cmd_Back.Enabled = True
  End If
  
ReCalc_End_End:
  Call xReturn("ReCalc_End")
  Exit Sub

ReCalc_End_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReCalc_End", "Re Calc End", "Error at end of recalculation.")
  Resume ReCalc_End_End
End Sub

Friend Sub DisplayDField()
  Dim i As Long
  Dim frmlink As Frm_Link
  Const CWFactor As Single = 0.8
  
  On Error GoTo DisplayDField_Err
  Call xSet("DisplayDField")
  
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
    
  If frmlink.FlG_DestInfo.Row = 0 Or frmlink.FlG_DestInfo.Row <> frmlink.FlG_DestInfo.RowSel Then
    frmlink.FlG_DestInfo.Row = 1
    frmlink.FlG_DestInfo.Col = 0
    frmlink.FlG_DestInfo.ColSel = frmlink.FlG_DestInfo.cols - 1
  End If
  
  If frmlink.FlG_DField.Enabled Then
    frmlink.FlG_DField.Redraw = False
    For i = 0 To (CurrentDest.rs.Fields.Count - 1)
      frmlink.FlG_DField.ColWidth(i) = 0
    Next i
    frmlink.FlG_DField.ColWidth(frmlink.FlG_DestInfo.Row - 1) = frmlink.FlG_DField.Width * CWFactor
    frmlink.FlG_DField.Visible = True
    frmlink.Lbl_Link.Caption = "Existing contents of:" & vbCrLf & "  " & frmlink.FlG_DestInfo.TextMatrix(frmlink.FlG_DestInfo.Row, 0)
    frmlink.Lbl_Link.Visible = True
    frmlink.FlG_DField.Redraw = True
  Else
    frmlink.FlG_DField.Visible = False
  End If
  
DisplayDField_End:
  Call xReturn("DisplayDField")
  Exit Sub
  
DisplayDField_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DisplayDField", "Display D Field", "Error tyring to display contents of selected destination field.")
  Resume DisplayDField_End
End Sub


Friend Sub SetDragParams(ByVal FForm As String, ByVal FFG As String, ByVal FCol As Long, ByVal FRow As Long, ByVal TForm As String, ByVal TFG As String, ByVal TCol As Long, ByVal TRow As Long)

  On Error GoTo SetDragParams_Err
  Call xSet("SetDragParams")
  
  'apf not cleared out if prev set
  If (StrComp(FForm, "-") <> 0) And (StrComp(FFG, "-") <> 0) Then
    m_FGDrop.FromForm = FForm
    m_FGDrop.FromFG = FFG
    m_FGDrop.FromCol = FCol
    m_FGDrop.FromRow = FRow
  End If
  m_FGDrop.ToForm = TForm
  m_FGDrop.ToFG = TFG
  m_FGDrop.ToCol = TCol
  m_FGDrop.ToRow = TRow
  
SetDragParams_End:
  Call xReturn("SetDragParams")
  Exit Sub

SetDragParams_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetDragParams", "Set Drag Params", "Error setting drag / drop parameters.")
  Resume SetDragParams_End
End Sub

Public Function ProcessDrag() As Boolean
  Dim ActionTaken As Boolean
  
  On Error GoTo ProcessDrag_Err
  Call xSet("ProcessDrag")

  ActionTaken = False

  ' Move a source column
  If StrComp(m_FGDrop.FromForm, Frm_Misc.Name) = 0 Then
    If (StrComp(m_FGDrop.FromFG, "FlG_Source") = 0) And (StrComp(m_FGDrop.ToFG, "FlG_Source") = 0) Then
      If m_FGDrop.FromCol <> 0 And m_FGDrop.ToCol <> 0 And m_FGDrop.FromCol <> m_FGDrop.ToCol Then
        If m_AllowMoveColumns Then
          ActionTaken = True
          Call MoveSourceCol(m_FGDrop.FromCol, m_FGDrop.ToCol)
        End If
      End If
    End If
  End If
  
  ' Link / delete link between source and destination field
  If StrComp(m_FGDrop.FromForm, Frm_Link.Name) = 0 Then
    If m_FGDrop.FromFG <> m_FGDrop.ToFG And m_FGDrop.FromFG <> "*" And m_FGDrop.ToFG <> "*" Then
      If MakeLink Then ActionTaken = True
    End If
  End If
 
  ProcessDrag = ActionTaken
 
ProcessDrag_End:
  Call xReturn("ProcessDrag")
  Exit Function

ProcessDrag_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ProcessDrag", "Process Drag", "Error procesing a drag/drop.")
  Resume ProcessDrag_End
End Function


Private Function MakeLink() As Boolean
  Dim i As Long
  Dim OldFieldType As DATABASE_FIELD_TYPES
  Dim frmlink As Frm_Link
  Dim FieldNum As Long
  Dim TestImportFlags As DOIMPORT_FLAGS
  Dim CWFactor As Single
  Dim ColNo As Long
  Dim LinkedTo As Long
  Dim fld As Field
  Dim Link1 As String, Link2 As String
  Dim ispec As ImportSpec
  Dim dStatic As DefaultStatic
  Static InMakeLink
  
  On Error GoTo MakeLink_Err
  Call xSet("MakeLink")
  
  MakeLink = False
  
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  
  If m_Imported Then GoTo MakeLink_End
  
  frmlink.FlG_Source.BackColorSel = QBColor(1)
  'MPSPHD check that the column dest is a valid column
  Select Case m_FGDrop.FromFG
    Case "FlG_DestInfo"  ' Make link
      FieldNum = GetColIndex(frmlink.FlG_Source.Coldata(m_FGDrop.ToCol))
      If m_FGDrop.ToCol < 1 Or Len(frmlink.FlG_DestInfo.TextMatrix(m_FGDrop.FromRow, 2)) > 0 Or Len(m_FieldSpecs(FieldNum).DestField) > 0 Then GoTo MakeLink_End
      If m_FieldSpecs(FieldNum).key = KEY_COPIEDFIELD Then
        If Len(m_FieldSpecs(m_FieldSpecs(FieldNum).CopyFieldKey).DestField) = 0 Then
          'apf mps not ECASE
          Call ECASE("A copied source field cannot be linked before the original source field has been linked." & vbCrLf)
          GoTo MakeLink_End
        End If
        Link1 = "": Link2 = ""
        Set ispec = m_FieldSpecs(FieldNum)
        For i = 1 To frmlink.FlG_DestInfo.Rows - 1
          If StrComp(frmlink.FlG_DestInfo.TextMatrix(i, 0), ispec.CopyField.DestField, vbTextCompare) = 0 Then
            Link1 = frmlink.FlG_DestInfo.TextMatrix(i, 1)
          End If
'          If frmlink.FlG_DestInfo.TextMatrix(i, 2) = ispec.DestField Then
'            Link2 = frmlink.FlG_DestInfo.TextMatrix(i, 1)
'          End If
        Next i
        Link2 = frmlink.FlG_DestInfo.TextMatrix(m_FGDrop.FromRow, 1)
        If (StrComp(Link1, Link2, vbTextCompare) <> 0) Or (Len(Link1) = 0) Then Err.Raise ERR_NOLINK, "MakeLink", "A copied source field cannot be linked to a destination field of a different type to the original source field."
      End If
      OldFieldType = m_FieldSpecs(FieldNum).vartype
      
      'm_FieldSpecs(FieldNum).key = KEY_FIELD
      
      m_FieldSpecs(FieldNum).DestField = CurrentDest.rs.Fields(m_FGDrop.FromRow - 1).Name
      m_FieldSpecs(FieldNum).vartype = ADOtoDatatype(CurrentDest.rs.Fields(m_FGDrop.FromRow - 1).Type)
      m_FieldSpecs(FieldNum).KeyField = CurrentDest.PrimaryKeys.IsPresent(m_FieldSpecs(FieldNum).DestField)
      m_FieldSpecs(FieldNum).Required = CurrentDest.RequiredFields.IsPresent(m_FieldSpecs(FieldNum).DestField)
      
      If (Not InMakeLink) And (Not m_UseSpec) Then  'MPSMarch
        InMakeLink = True
        Call AddInDefaultCopiedFields(frmlink.Name)
        Call AddInDefaultConstraints(frmlink.Name)
        Call AddInDefaultValues(frmlink.Name)
        InMakeLink = False
      End If
     
      TestImportFlags = FIELD_NAME + FIELD_TYPE + LINKED_NAME + ROW_LINENUMBERS
      If Not DoImport(frmlink.FlG_Source, TestImportFlags, False) Then
        Call SetText(frmlink.Lbl_Status, "The selected fields have incompatible formats: They cannot be linked.")
        frmlink.Cbo_Link.Text = frmlink.Cbo_Link.List(0)
      Else
        MakeLink = True
        Call SetText(frmlink.Lbl_Status, "")
        frmlink.FlG_DestInfo.TextMatrix(m_FGDrop.FromRow, 2) = m_FieldSpecs(FieldNum).FldName
      End If

    Case "FlG_Source"  ' Remove Link
      FieldNum = GetColIndex(frmlink.FlG_Source.Coldata(m_FGDrop.FromCol))
      If Len(m_FieldSpecs(FieldNum).DestField) = 0 Then GoTo MakeLink_End
      frmlink.FlG_DestInfo.TextMatrix(DestFieldItem(m_FieldSpecs(FieldNum).DestField), 2) = ""
      m_FieldSpecs(FieldNum).DestField = ""
      m_FieldSpecs(FieldNum).vartype = TYPE_STR
      m_FieldSpecs(FieldNum).KeyField = False
      m_FieldSpecs(FieldNum).Required = False
      If Not CurrentDest.DefaultStatics Is Nothing Then
        For i = CurrentDest.DefaultStatics.Count To 1 Step -1
          Set dStatic = CurrentDest.DefaultStatics(i)
          If Len(dStatic.CopyFieldName) = 0 And dStatic.FieldName = m_FieldSpecs(FieldNum).FldName Then
            Call CurrentDest.DefaultStatics.Remove(i)
          End If
        Next i
      End If
      MakeLink = True
  End Select

'  If Not InMakeLink Then
'    InMakeLink = True
'    Call AddInDefaultCopiedFields(frmlink.Name)
'    Call AddInDefaultConstraints(frmlink.Name)
'    Call AddInDefaultValues(frmlink.Name)
'    InMakeLink = False
'  End If

MakeLink_End:
  Call xReturn("MakeLink")
  Exit Function

MakeLink_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MakeLink", "Make Link", "Unable to make link.")
  Resume MakeLink_End
  Resume
End Function

Public Function DestFieldItem(ByVal FieldName As String) As Long
  Dim fld As Field
  Dim i As Long
  DestFieldItem = 0
  For i = 1 To CurrentDest.rs.Fields.Count
    If StrComp(CurrentDest.rs.Fields(i - 1).Name, FieldName, vbTextCompare) = 0 Then
      DestFieldItem = i
      Exit For
    End If
  Next i
End Function

Friend Function SetDestDataGrid(ByVal rs As Recordset, ByVal FormType As IMPORT_GOTOFORM) As Long
  Dim frmDest As Frm_Dest
  Dim frmlink As Frm_Link
  Dim frmEnd As Frm_End
  Dim i As Long, Limit As Long, NumRecs As Long
        
  On Error GoTo SetDestDataGrid_err
  Call SetCursor
  Limit = m_DestPreviewRecLimit
  If Limit <= 0 Then Limit = 0
  SetDestDataGrid = -1
  
  If Not (rs Is Nothing) Then
    'MPSFEB
    rs.Requery
    If Not rs.BOF Then rs.MoveFirst
    If Not (rs.BOF And rs.EOF) Then
      rs.MoveLast
      rs.MoveFirst
      SetDestDataGrid = rs.RecordCount
    Else
      SetDestDataGrid = 0
    End If
    Select Case FormType
      Case TCSIMP_DEST
        Set frmDest = m_ImportForms.Item(CStr(TCSIMP_DEST))
        NumRecs = FillFlGFromRS(frmDest.FlG_Dest, rs, m_DestPreviewRecLimit)
        If NumRecs > m_DestPreviewRecLimit Then
          frmDest.Lbl_NumRecs.Caption = "Number of Records: " & NumRecs & "  (only first " & m_DestPreviewRecLimit & " displayed)"
        Else
          frmDest.Lbl_NumRecs.Caption = "Number of Records: " & NumRecs
        End If
        frmDest.Lbl_TabContents.Caption = "Contents of selected recordset:"
        
      Case TCSIMP_LINK
        Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
        NumRecs = FillFlGFromRS(frmlink.FlG_DField, rs, m_DestPreviewRecLimit)

      Case TCSIMP_END
        Set frmEnd = m_ImportForms.Item(CStr(TCSIMP_END))
        NumRecs = FillFlGFromRS(frmEnd.FlG_Dest, rs, m_DestPreviewRecLimit)
        'If NumRecs > m_DestPreviewRecLimit Then
        '  frmEnd.Lbl_NumRecs.Caption = "Number of Records: " & NumRecs & "  (only first " & m_DestPreviewRecLimit & " displayed)"
        'Else
        '  frmEnd.Lbl_NumRecs.Caption = "Number of Records: " & NumRecs
        'End If
        frmEnd.Lbl_Dest.Caption = "Destination database:"
        
      Case Else
        Call ECASE("SetDestDataGrid unknown form type")
    End Select
  End If
SetDestDataGrid_end:
  Call ClearCursor
  Exit Function
  
SetDestDataGrid_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetDestDataGrid", "Set Form data control", "Unable to set the data control")
  Resume SetDestDataGrid_end
End Function

Friend Sub FormatField()
  Dim i As Long
  Dim ColNo As Long
  Dim s As String, ss As String
  Dim InList As Boolean
  Dim frmMisc As Frm_Misc
  Dim FieldNum As Long

  On Error GoTo FormatField_Err
  Call xSet("FormatField")
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  
  frmMisc.FlG_Source.Redraw = False
  
  frmMisc.FlG_Source.ColSel = frmMisc.FlG_Source.Col
  frmMisc.FlG_Source.Row = FIXED_ROWCOUNT
  'frmMisc.FlG_Source.RowSel = UBound(m_LineArray)
  frmMisc.FlG_Source.RowSel = frmMisc.FlG_Source.Rows - 1
  If frmMisc.FlG_Source.ColIsVisible(frmMisc.FlG_Source.Col) = False Then
    frmMisc.FlG_Source.LeftCol = frmMisc.FlG_Source.Col
  End If
  frmMisc.Fra_Format.Caption = "Format of field  " & frmMisc.FlG_Source.TextMatrix(0, frmMisc.FlG_Source.Col)
  FieldNum = GetColIndex(frmMisc.FlG_Source.Coldata(frmMisc.FlG_Source.Col))
  If Len(m_FieldSpecs(FieldNum).DestField) > 0 Then
    frmMisc.Fra_Format.Caption = frmMisc.Fra_Format.Caption & " - Linked to destination field '" & m_FieldSpecs(FieldNum).DestField & "'"
  End If
  If Len(m_FieldSpecs(FieldNum).DestField) > 0 Or m_Imported Or m_FieldSpecs(FieldNum).key = KEY_COPIEDFIELD Or m_FieldSpecs(FieldNum).key = KEY_CONSTANT Or IsSpecialFieldType(m_FieldSpecs(FieldNum).key) Then 'MPSMarch2
    Call FrameEnable(frmMisc, frmMisc.Fra_Format, False)
    frmMisc.Fra_Format.Enabled = False
  Else
    Call FrameEnable(frmMisc, frmMisc.Fra_Format, True)
    frmMisc.Fra_Format.Enabled = True
  End If
  frmMisc.Cbo_Type.Clear
  For i = 1 To DATABASE_FIELD_TYPES_COUNT
    frmMisc.Cbo_Type.AddItem DataTypeName(i)
  Next i
  frmMisc.Cbo_Date.Clear
  frmMisc.Cbo_Date.AddItem "DMY"
  frmMisc.Cbo_Date.AddItem "MDY"
  frmMisc.Cbo_Date.AddItem "DMYHNS"
  frmMisc.Cbo_Date.AddItem "MDYHNS"
  
  frmMisc.Cbo_DateDLim.Clear
  frmMisc.Cbo_DateDLim.AddItem "/"
  frmMisc.Cbo_DateDLim.AddItem "{None}"
  frmMisc.Cbo_DateDLim.AddItem "{Space}"
  frmMisc.Cbo_DateDLim.AddItem "- {Dash}"
  frmMisc.Cbo_DateDLim.AddItem ", {Comma}"
  frmMisc.Cbo_DateDLim.AddItem ". {Full Stop}"
  frmMisc.Cbo_DateDLim.AddItem ":"
  frmMisc.Cbo_DateDLim.AddItem ";"
  
  frmMisc.Cbo_TimeDLim.Clear
  frmMisc.Cbo_TimeDLim.AddItem ":"
  frmMisc.Cbo_TimeDLim.AddItem "{None}"
  frmMisc.Cbo_TimeDLim.AddItem "{Space}"
  frmMisc.Cbo_TimeDLim.AddItem "- {Dash}"
  frmMisc.Cbo_TimeDLim.AddItem ", {Comma}"
  frmMisc.Cbo_TimeDLim.AddItem ". {Full Stop}"
  frmMisc.Cbo_TimeDLim.AddItem "/"
  frmMisc.Cbo_TimeDLim.AddItem ";"
  
  frmMisc.Cbo_Boolean(0).Clear
  frmMisc.Cbo_Boolean(0).AddItem "True"
  frmMisc.Cbo_Boolean(0).AddItem "Yes"
  
  frmMisc.Cbo_Boolean(1).Clear
  frmMisc.Cbo_Boolean(1).AddItem "False"
  frmMisc.Cbo_Boolean(1).AddItem "No"
  
  'MPSMarch2
  frmMisc.Cbo_Static.Clear
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_DATENOW)
  frmMisc.Cbo_Static.ItemData(0) = KEY_DATENOW
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_FILENAME)
  frmMisc.Cbo_Static.ItemData(1) = KEY_FILENAME
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_FILEPATH)
  frmMisc.Cbo_Static.ItemData(2) = KEY_FILEPATH
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_FILEDATE)
  frmMisc.Cbo_Static.ItemData(3) = KEY_FILEDATE
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_LINENUMBER)
  frmMisc.Cbo_Static.ItemData(4) = KEY_LINENUMBER
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_CFGFILENAME)
  frmMisc.Cbo_Static.ItemData(5) = KEY_CFGFILENAME
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_USERNAME)
  frmMisc.Cbo_Static.ItemData(6) = KEY_USERNAME
  frmMisc.Cbo_Static.AddItem SpecialFieldTypeName(KEY_IMPDATE)
  frmMisc.Cbo_Static.ItemData(7) = KEY_IMPDATE
  
  ColNo = GetColIndex(frmMisc.FlG_Source.Coldata(frmMisc.FlG_Source.Col))
  frmMisc.Cbo_Type.Text = DataTypeName(m_FieldSpecs(ColNo).vartype)

  If frmMisc.Cbo_Type.Text = DataTypeName(TYPE_DATE) And (m_FieldSpecs(ColNo).key <> KEY_CONSTANT) Then
    s = m_FieldSpecs(ColNo).actionparam
    If m_FieldSpecs(ColNo).varaction = CONVERT_FIXEDDATE Then
      frmMisc.Cbo_DateDLim = "{None}"
      frmMisc.Chk_4FigYear.Value = -(InStr(s, "YYYY") > 0)
      s = ReplaceChar(s, "YYYY", "Y")
      s = ReplaceChar(s, "YY", "Y")
      s = ReplaceChar(s, "MM", "M")
      s = ReplaceChar(s, "DD", "D")
      s = ReplaceChar(s, "HH", "H")
      s = ReplaceChar(s, "NN", "N")
      s = ReplaceChar(s, "SS", "S")
    Else
      ss = m_FieldSpecs(ColNo).DateDelimiter
      If ss = " " Then ss = "{Space}"
      If ss = "-" Then ss = "- {Dash}"
      If ss = "," Then ss = ", {Comma}"
      If ss = "." Then ss = ". {Full Stop}"
      frmMisc.Cbo_DateDLim = ss
      ss = m_FieldSpecs(ColNo).TimeDelimiter
      If ss = " " Then ss = "{Space}"
      If ss = "-" Then ss = "- {Dash}"
      If ss = "," Then ss = ", {Comma}"
      If ss = "." Then ss = ". {Full Stop}"
      frmMisc.Cbo_TimeDLim = ss
    End If
    frmMisc.Cbo_Date = s
  Else
    frmMisc.Cbo_DateDLim.Text = frmMisc.Cbo_DateDLim.List(0)
    frmMisc.Cbo_TimeDLim.Text = frmMisc.Cbo_TimeDLim.List(0)
    frmMisc.Cbo_Date.Text = frmMisc.Cbo_Date.List(0)
  End If
  If frmMisc.Cbo_Type.Text = DataTypeName(TYPE_BOOL) Then
    InList = False
    For i = 0 To frmMisc.Cbo_Boolean(0).ListCount - 1
      If frmMisc.Cbo_Boolean(0).List(i) = m_FieldSpecs(ColNo).BooleanTrue Then InList = True
    Next i
    If InList = False Then
      frmMisc.Cbo_Boolean(0).AddItem m_FieldSpecs(ColNo).BooleanTrue
    End If
    frmMisc.Cbo_Boolean(0).Text = m_FieldSpecs(ColNo).BooleanTrue
    InList = False
    For i = 0 To frmMisc.Cbo_Boolean(1).ListCount - 1
      If frmMisc.Cbo_Boolean(1).List(i) = m_FieldSpecs(ColNo).BooleanFalse Then InList = True
    Next i
    If InList = False Then
      frmMisc.Cbo_Boolean(1).AddItem m_FieldSpecs(ColNo).BooleanFalse
    End If
    frmMisc.Cbo_Boolean(1).Text = m_FieldSpecs(ColNo).BooleanFalse
  Else
    frmMisc.Cbo_Boolean(0).Text = frmMisc.Cbo_Boolean(0).List(0)
    frmMisc.Cbo_Boolean(1).Text = frmMisc.Cbo_Boolean(1).List(0)
  End If

  'End With
  frmMisc.FlG_Source.Redraw = True

FormatField_End:
    Call xReturn("FormatField")
    Exit Sub

FormatField_Err:
    Call ErrorMessage(ERR_ERROR, Err, "FormatField", "Format Field", "Unable to format field.")
    Resume FormatField_End
End Sub

Friend Sub FormatFieldOK()
  Dim ColNo As Long
  Dim s As String
  Dim i As Long
  Dim j1 As String, j2 As String
  Dim frmMisc As Frm_Misc

  On Error GoTo FormatFieldOK_Err
  Call xSet("FormatFieldOK")
  
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))

  With frmMisc
    ColNo = GetColIndex(.FlG_Source.Coldata(.FlG_Source.Col))

    m_FieldSpecs(ColNo).vartype = .Cbo_Type.ListIndex + 1

    If .Cbo_Type = DataTypeName(TYPE_BOOL) Then
      m_FieldSpecs(ColNo).BooleanTrue = .Cbo_Boolean(0).Text
      m_FieldSpecs(ColNo).BooleanFalse = .Cbo_Boolean(1).Text
    End If
    
    If .Cbo_Type = DataTypeName(TYPE_DATE) Then
      If .Cbo_DateDLim.Text = "{None}" Then
        m_FieldSpecs(ColNo).varaction = CONVERT_FIXEDDATE
        s = .Cbo_Date.Text
        s = ReplaceChar(s, "D", "DD")
        s = ReplaceChar(s, "M", "MM")
        If .Chk_4FigYear.Value = vbChecked Then
          s = ReplaceChar(s, "Y", "YYYY")
        Else
          s = ReplaceChar(s, "Y", "YY")
        End If
        s = ReplaceChar(s, "H", "HH")
        s = ReplaceChar(s, "N", "NN")
        s = ReplaceChar(s, "S", "SS")
        m_FieldSpecs(ColNo).actionparam = s
      Else
        m_FieldSpecs(ColNo).varaction = CONVERT_DELIMITED
        Select Case .Cbo_DateDLim.Text
          Case "{Space}"
            m_FieldSpecs(ColNo).DateDelimiter = " "
          Case "- {Dash}"
            m_FieldSpecs(ColNo).DateDelimiter = "-"
          Case ", {Comma}"
            m_FieldSpecs(ColNo).DateDelimiter = ","
          Case ". {Full Stop}"
            m_FieldSpecs(ColNo).DateDelimiter = "."
          Case Else
            m_FieldSpecs(ColNo).DateDelimiter = .Cbo_DateDLim.Text
        End Select
        Select Case .Cbo_TimeDLim.Text
          Case "{Space}"
            m_FieldSpecs(ColNo).TimeDelimiter = " "
          Case "- {Dash}"
            m_FieldSpecs(ColNo).TimeDelimiter = "-"
          Case ", {Comma}"
            m_FieldSpecs(ColNo).TimeDelimiter = ","
          Case ". {Full Stop}"
            m_FieldSpecs(ColNo).TimeDelimiter = "."
          Case Else
            m_FieldSpecs(ColNo).TimeDelimiter = .Cbo_TimeDLim.Text
        End Select
        s = .Cbo_Date.Text
        m_FieldSpecs(ColNo).actionparam = s
      End If
    End If
    
    'MPSMarch2
    If .Cbo_Type = DataTypeName(TYPE_LONG) Or .Cbo_Type = DataTypeName(TYPE_DOUBLE) Then
      If .Chk_Factor.Value = vbChecked Then
        m_FieldSpecs(ColNo).varaction = ACTION_FACTOR
        If .Cbo_Type = DataTypeName(TYPE_LONG) Then
          m_FieldSpecs(ColNo).actionparam = CLng(.Txt_Factor)
        Else
          m_FieldSpecs(ColNo).actionparam = CDbl(.Txt_Factor)
        End If
      Else
        m_FieldSpecs(ColNo).varaction = 0
      End If
    End If
    
    Call ReCalc_Misc(False)
    
  End With

FormatFieldOK_End:
    Call xReturn("FormatFieldOK")
    Exit Sub

FormatFieldOK_Err:
    Call ErrorMessage(ERR_ERROR, Err, "FormatFieldOK", "Format Field OK", "Error in format field OK")
    Resume FormatFieldOK_End
End Sub

Public Function NumSourceFields() As Long

Dim i As Long
Dim Num As Long

Num = 0
For i = 1 To m_FieldSpecs.Count
  If m_FieldSpecs(i).key = 1 Then Num = Num + 1
Next i
NumSourceFields = Num

End Function

Friend Sub CreateMiscMenu()

  Dim frmMisc As Frm_Misc
  Dim ColL As Long, ColR As Long
  Dim i As Long
  
  On Error GoTo CreateMiscMenu_Err
  Call xSet("CreateMiscMenu")
  
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))

  If m_Imported Then GoTo CreateMiscMenu_End

  ColL = frmMisc.FlG_Source.Col
  ColR = frmMisc.FlG_Source.ColSel
  If ColR < ColL Then
    i = ColL
    ColL = ColR
    ColR = i
    'Switch round selection here
    frmMisc.FlG_Source.Col = ColL
    frmMisc.FlG_Source.Row = FIXED_ROWCOUNT
    frmMisc.FlG_Source.ColSel = ColR
    frmMisc.FlG_Source.RowSel = frmMisc.FlG_Source.Rows - 1
  End If
  
  ' Static field can always be added
  frmMisc.mnuAddStatic.Enabled = True
  
  ' Hide Delete Static Field if first selected field
  ' is not a static field.
  frmMisc.mnuDelStatic.Enabled = False
  If ColL = ColR Then
    For i = 1 To m_FieldSpecs.Count
      If m_FieldSpecs(i).FieldKey = ColR Then
        If m_FieldSpecs(i).key <> KEY_FIELD Then
          frmMisc.mnuDelStatic.Enabled = True
        End If
        Exit For
      End If
    Next i
  End If

  frmMisc.mnuFormat.Visible = False
  'frmMisc.mnuFormat.Enabled = False
  'If ColL = ColR Then
  '  frmMisc.mnuFormat.Enabled = True
  'End If
  
  frmMisc.mnuCopyField.Enabled = False
  If ColL = ColR Then
    i = GetColIndex(frmMisc.FlG_Source.Coldata(frmMisc.FlG_Source.Col))
    If m_FieldSpecs(i).key = KEY_FIELD Then
      frmMisc.mnuCopyField.Enabled = True
    End If
  End If
  
  If m_AllowMergeColumns = True Then
    frmMisc.mnuMerge.Enabled = False
    frmMisc.mnuMergeSpaces.Enabled = False
    If ColL <> ColR Then
    End If
    frmMisc.mnuSeparate.Enabled = False
    If ColL = ColR Then
    End If
  Else
    frmMisc.mnuMerge.Visible = False
    frmMisc.mnuMergeSpaces.Visible = False
    frmMisc.mnuSeparate.Visible = False
  End If
  
  frmMisc.PopupMenu frmMisc.mnuMisc

CreateMiscMenu_End:
  Call xReturn("CreateMiscMenu")
  Exit Sub

CreateMiscMenu_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CreateMiscMenu", "Create Misc Menu", "Error creating the misc menu.")
  Resume CreateMiscMenu_End
End Sub

'* IF MoveCols is allowed this function must remember ColData after update
' APF MPSPHD
Private Sub SetFlexRowCol(ByVal flx As MSFlexGrid, ByVal rctype As SET_ROWCOLTYPE, ByVal Value As Long, ByVal FixedValue As Long)
  If (Value <= 0) Or (FixedValue <= 0) Then
    If rctype = SET_ROW Then
      flx.Rows = 0
    ElseIf rctype = SET_COLUMN Then
      flx.cols = 0
    End If
  End If
  If Value > 0 Then
    If FixedValue > Value Then
      FixedValue = 1
      If Value = 1 Then FixedValue = 0
    End If
    If FixedValue > 0 Then
      If rctype = SET_ROW Then
        flx.Rows = Max(FixedValue + 1, flx.FixedRows + 1)
        flx.FixedRows = FixedValue
      ElseIf rctype = SET_COLUMN Then
        flx.cols = Max(FixedValue + 1, flx.FixedCols + 1)
        flx.FixedCols = FixedValue
      End If
    End If
    If rctype = SET_ROW Then
      flx.Rows = Value
    ElseIf rctype = SET_COLUMN Then
      flx.cols = Value
    End If
  End If
End Sub

  ' Find which column in the flex grid has the field key
Private Function GetFlexColumn(ByVal flx As MSFlexGrid, ByVal key As Long) As Long
  Dim i As Long

  GetFlexColumn = -1
  For i = flx.FixedCols To (flx.cols - 1)
    If flx.Coldata(i) = key Then
      GetFlexColumn = i: Exit For
    End If
  Next i
End Function

  ' Find which column in the flex grid has the field key
Private Function GetColIndex(ByVal Coldata As Long) As Long
  Dim i As Long

  GetColIndex = -1
  For i = 1 To m_FieldSpecs.Count
    If m_FieldSpecs(i).FieldKey = Coldata Then
      GetColIndex = i: Exit For
    End If
  Next i
End Function

Private Sub ResetFlexGrids()
  Dim frmMisc As Frm_Misc
  Dim frmlink As Frm_Link
  
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  Call SetFlexRowCol(frmMisc.FlG_Source, SET_COLUMN, 0, 0)
  Call SetFlexRowCol(frmlink.FlG_Source, SET_COLUMN, 0, 0)
End Sub

'apf fix errors
Public Function AddDefaultStaticLink(ByVal Name As String, ByVal FieldName As String, ByVal DataType As DATABASE_FIELD_TYPES, ByVal Value As Variant, Optional ByVal GridFieldName As String = "") As Boolean
  On Error GoTo AddDefaultStaticLink_Err
  Call xSet("AddDefaultStaticLink")
  AddDefaultStaticLink = False
  If AddStaticLink(Name, FieldName, DataType, Value, GridFieldName) Then
    AddDefaultStaticLink = AddDefaultValue(Name, FieldName, DataType, Value)
    Call AddRequiredField(Name, FieldName)
  End If
  
AddDefaultStaticLink_End:
  Call xReturn("AddDefaultStaticLink")
  Exit Function

AddDefaultStaticLink_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddDefaultStaticLink", "Add default static link", "Error trying to add default static '" & Value & "' to recordset '" & Name & "'.")
  Resume AddDefaultStaticLink_End
End Function


Public Function AddStaticLink(ByVal Name As String, ByVal FieldName As String, ByVal DataType As DATABASE_FIELD_TYPES, ByVal Value As Variant, Optional ByVal GridFieldName As String = "", Optional ByVal Hide As Boolean = False) As Boolean
  Dim DefStatic As DefaultStatic
  Dim drs As DestRecordSet
  
  On Error GoTo AddStaticLink_Err
  Call xSet("AddStaticLink")
  AddStaticLink = False
  Set drs = m_DestRecSets.Item(Name)
  Set DefStatic = New DefaultStatic
  DefStatic.DestField = FieldName
  DefStatic.FieldName = GridFieldName
  DefStatic.CopyFieldName = ""
  DefStatic.Value = Value
  DefStatic.DataType = DataType
  If Hide And (Len(FieldName) = 0) Then Call Err.Raise(ERR_DEFSTATIC, "AddDefaultStatic", "Cannot hide an unlinked field")
 
  If Hide Then Call HideField(Name, FieldName)
  Call drs.DefaultStatics.Add(DefStatic)
  AddStaticLink = True
  
AddStaticLink_End:
  Call xReturn("AddStaticLink")
  Exit Function

AddStaticLink_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddStaticLink", "Add default static link", "Error trying to add default static '" & Value & "' to recordset '" & Name & "'.")
  Resume AddStaticLink_End
  Resume
End Function

Public Function AddDefaultStatic(ByVal Name As String, ByVal DataType As DATABASE_FIELD_TYPES, ByVal Value As Variant, Optional ByVal FieldName As String) As Boolean
  Dim DefStatic As DefaultStatic
  Dim drs As DestRecordSet
  
  On Error GoTo AddDefaultStatic_Err
  Call xSet("AddDefaultStatic")
  AddDefaultStatic = False
  Set drs = m_DestRecSets.Item(Name)
  Set DefStatic = New DefaultStatic
  DefStatic.CopyFieldName = ""
  DefStatic.Value = Value
  DefStatic.DataType = DataType
  DefStatic.FieldName = FieldName
  Call drs.DefaultStatics.Add(DefStatic)
  AddDefaultStatic = True
  
AddDefaultStatic_End:
  Call xReturn("AddDefaultStatic")
  Exit Function

AddDefaultStatic_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddDefaultStatic", "Add Static value", "Error trying to add default static '" & Value & "' to recordset '" & Name & "'.")
  Resume AddDefaultStatic_End
End Function

Private Sub AddInDefaultStatics(ByVal FormName As String)
  Dim DefStatic As DefaultStatic, isp As ImportSpec
  Dim frmdlim As Frm_DLim, frmlink As Frm_Link, i As Long

  On Error GoTo AddInDefaultStatics_Err
  Call xSet("AddInDefaultStatics")
 
  Set frmdlim = m_ImportForms.Item(CStr(TCSIMP_DLIM))
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  If Not CurrentDest.DefaultStatics Is Nothing Then
    For Each DefStatic In CurrentDest.DefaultStatics
      If Len(DefStatic.CopyFieldName) > 0 Then
        Set isp = m_FieldSpecs.Add
        isp.key = KEY_CONSTANT
        isp.actionparam = "[COPY!]"
        isp.vartype = TYPE_STR
        isp.FldName = DefStatic.FieldName
        isp.DestField = ""
        DefStatic.CopyFieldKey = m_FieldSpecs.Count
      Else
        Set isp = m_FieldSpecs.Add
        isp.key = KEY_CONSTANT
        isp.vartype = DefStatic.DataType
        isp.FldName = DefStatic.FieldName
        isp.Hide = False
        isp.actionparam = DefStatic.Value
        isp.DestField = ""
        For i = 1 To frmlink.FlG_DestInfo.Rows - 1
          If StrComp(frmlink.FlG_DestInfo.TextMatrix(i, 0), DefStatic.DestField, vbTextCompare) = 0 Then
            isp.DestField = DefStatic.DestField
            isp.Hide = CurrentDest.HiddenFields.IsPresent(DefStatic.DestField)
            Exit For
          End If
        Next i
      End If
  
      isp.FieldKey = m_FieldSpecs.Count
      If FormName = frmdlim.Name Then
        frmdlim.FlG_Source.cols = frmdlim.FlG_Source.cols + 1
        frmdlim.FlG_Source.Coldata(m_FieldSpecs.Count + FIXED_COLCOUNT - 1) = isp.FieldKey
      End If
        
    Next DefStatic
  End If
  Call MakeFieldNamesUnique(m_FieldSpecs)
  
AddInDefaultStatics_End:
  Call xReturn("AddInDefaultStatics")
  Exit Sub

AddInDefaultStatics_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddInDefaultStatics", "Add In Default Statics", "Error adding the default statics")
  Resume AddInDefaultStatics_End
End Sub


Private Function FillFlGFromRS(ByVal FlG As MSFlexGrid, ByVal rs As Recordset, ByVal MaxLines As Long) As Long
  Dim j As Long, NumRecs As Long
  
  On Error GoTo FillFlGFromRS_Err
  Call xSet("FillFlGFromRS")

  FlG.Redraw = False
  If Not (rs Is Nothing) Then
    'rs.Requery
    DoEvents
    If Not (rs.BOF And rs.EOF) Then
      rs.MoveLast
      rs.MoveFirst
      rs.MoveLast
      NumRecs = rs.RecordCount
    Else
      NumRecs = 0
    End If
    FillFlGFromRS = NumRecs
    Call SetFlexRowCol(FlG, SET_ROW, Min(NumRecs, MaxLines) + 1, 1)
    Call SetFlexRowCol(FlG, SET_COLUMN, rs.Fields.Count, 0)
    For j = 0 To rs.Fields.Count - 1
      FlG.TextMatrix(0, j) = rs.Fields(j).Name
      If CurrentDest.HiddenFields.IsPresent(rs.Fields(j).Name) Then
        FlG.ColWidth(j) = 0
      End If
    Next j
    If NumRecs > 0 Then
      Call RecordsToGrid(FlG, rs, MaxLines)
    End If
  Else
    FillFlGFromRS = -1
  End If
  FlG.Redraw = True
  
FillFlGFromRS_End:
  Call xReturn("FillFlGFromRS")
  Exit Function

FillFlGFromRS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "FillFlGFromRS", "Fill FlG From RS", "Error filling the flex grid from a recordset.")
  Resume FillFlGFromRS_End
End Function

Private Function RecordsToGrid(ByVal msg As MSFlexGrid, ByVal rs As Recordset, ByVal MaxLines As Long) As Boolean
  Dim l As Long, m As Long
  
  On Error GoTo RecordsToGrid_Err
  Call xSet("RecordsToGrid")
  
  rs.MoveFirst
  
  With msg
    Do
      l = l + 1
      For m = 0 To rs.Fields.Count - 1
       .TextMatrix(l, m) = "" & rs.Fields(m).Value
      Next m
NEXT_RECORD:
      If Not rs.EOF Then rs.MoveNext
    Loop Until rs.EOF Or l = MaxLines Or l = .Rows - 1
  End With

RecordsToGrid_End:
  Call xReturn("RecordsToGrid")
  Exit Function

RecordsToGrid_Err:
  If Err.Number = 3167 Then
    msg.Rows = msg.Rows - 1
    l = l - 1
    Resume NEXT_RECORD
  Else
    Call ErrorMessage(ERR_ERROR, Err, "RecordsToGrid", "RecordsToGrid", "Error placing the records to the grid.")
  End If
  Resume RecordsToGrid_End
End Function

Public Function HideField(ByVal Name As String, ByVal FieldName As String) As Boolean
  Dim drs As DestRecordSet
  
  On Error GoTo HideField_Err
  Call xSet("HideField")
  Set drs = m_DestRecSets(Name)
  Call drs.HiddenFields.Add(FieldName)

HideField_End:
  Call xReturn("HideField")
  Exit Function

HideField_Err:
  Call ErrorMessage(ERR_ERROR, Err, "HideField", "Hide Field", "Error trying to hide field '" & FieldName & "' in recordset '" & Name & "'")
  Resume HideField_End
  Resume
End Function

Friend Function UpdateImportDescription(ByVal bHide As Boolean) As Boolean
  On Error GoTo UpdateImportDescription_Err
  Call xSet("UpdateImportDescription")
  
  With Frm_End
    If bHide Then
      .fraImportAnalysis.Visible = False
    Else
      .lblFileLines = "Lines processed: " & ImpParent.AttemptedLines
      .lblLinesInError = "Processing lines in error: " & ImpParent.ProcessingLinesInError
      .lblPostProcessingLinesInError = "Post processing lines in error: " & PostProcessingLinesInError
      .lblLinesOK = "Lines OK: " & ImpParent.LinesOK
      .lblRecordsAdded = "Records added: " & ImpParent.AddedRecordCount
      .lblRecordsUpdated = "Records updated: " & ImpParent.UpdatedRecordCount
      .fraImportAnalysis.Visible = True
    End If
  End With
        
UpdateImportDescription_End:
  Call xReturn("UpdateImportDescription")
  Exit Function

UpdateImportDescription_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateImportDescription", "Update Import Description", "Error updating the import description.")
  Resume UpdateImportDescription_End
End Function

Friend Function ImportFieldNames(ImpFNamesVariant As Variant) As Boolean
  Dim ImpVariant() As Variant
  Dim ErrorCount As Long
  Dim hCount As Long
  Dim i As Long, isp As ImportSpec, ispTmp As ImportSpec

  On Error GoTo ImportFieldNames_Err
  Call xSet("ImportFieldNames")
  ImpParent.UpdateType = NO_UPDATES
  
  If ImpParent.InitImport(m_SourceFile) Then
    For i = 1 To m_FieldSpecs.Count
      Set isp = m_FieldSpecs(i)
      Set ispTmp = New ImportSpec
      Call isp.Copy(ispTmp)
      If ispTmp.key = KEY_COPIEDFIELD Then ispTmp.key = KEY_CONSTANT  'MPSMarch2
      ispTmp.vartype = TYPE_STR
      ispTmp.Constraint = [_CONSTRAINT_INVALID]
      ispTmp.ConstraintList.RemoveAll
      ispTmp.Required = False 'MPSMarch
      Call ImpParent.AddColumn(ispTmp)
    Next i
    ImpParent.TestImport = True
    hCount = ImpParent.HeaderCount
    ImpParent.HeaderCount = 0
    ErrorCount = ImpParent.ImportLines(ImpVariant(), ImpParent.TakeFieldNamesFromLine)
    ImpParent.HeaderCount = hCount
    ReDim ImpFNamesVariant(1 To m_FieldSpecs.Count)
    For i = 1 To m_FieldSpecs.Count
      ImpFNamesVariant(i) = ImpVariant(ImpParent.TakeFieldNamesFromLine, i)
    Next i
  End If
  
ImportFieldNames_End:
  Call xReturn("ImportFieldNames")
  Exit Function

ImportFieldNames_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ImportFieldNames", "Import Field Names", "Error importing the field names.")
  Resume ImportFieldNames_End
End Function

Public Function AddDefaultCopiedField(ByVal Name As String, ByVal CopyFieldName As String, Optional ByVal FieldName As String = "", Optional ByVal DestField As String = "", Optional ByVal Hide As Boolean = False) As Boolean
  Dim DefStatic As DefaultStatic
  Dim drs As DestRecordSet
  
  On Error GoTo AddDefaultCopiedField_Err
  Call xSet("AddDefaultCopiedField")

  AddDefaultCopiedField = False
  Set drs = m_DestRecSets.Item(Name)
  Set DefStatic = New DefaultStatic
  DefStatic.CopyFieldName = CopyFieldName
  DefStatic.FieldName = FieldName
  
  If Len(DestField) = 0 Then Call Err.Raise(ERR_DEFCOPIEDFIELD, "AddDefaultCopiedField", "Unable to add default copied field. A destination field is required")
  If Hide And (Len(DestField) = 0) Then Call Err.Raise(ERR_DEFCOPIEDFIELD, "AddDefaultCopiedField", "Cannot hide an unlinked field")
  DefStatic.DestField = DestField
  If Hide Then Call HideField(Name, DestField)
  
  Call drs.DefaultStatics.Add(DefStatic)
  AddDefaultCopiedField = True

AddDefaultCopiedField_End:
  Call xReturn("AddDefaultCopiedField")
  Exit Function

AddDefaultCopiedField_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddDefaultCopiedField", "Add Default Copied Field", "Error trying to add a default copied field named '" & CopyFieldName & "' to recordset '" & Name & "'.")
  Resume AddDefaultCopiedField_End
End Function

Private Sub AddInDefaultCopiedFields(ByVal FormName As String)
  Dim DefStatic As DefaultStatic, isp As ImportSpec
  Dim ispCopy As ImportSpec
  Dim frmlink As Frm_Link, i As Long, j As Long
  Static InDefCopyFields As Boolean

  On Error GoTo AddInDefaultCopiedFields_Err
  Call xSet("AddInDefaultCopiedFields")
 
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  
  If Not InDefCopyFields Then
  
  InDefCopyFields = True
  
  If Not CurrentDest.DefaultStatics Is Nothing Then
    For Each DefStatic In CurrentDest.DefaultStatics
      For i = 1 To m_FieldSpecs.Count
        Set ispCopy = m_FieldSpecs(i)
        If ispCopy.FieldKey = DefStatic.CopyFieldKey Then
          Exit For
        End If
      Next i
      m_FGDrop.FromForm = ""
      If ispCopy.FieldKey = DefStatic.CopyFieldKey Then
        If Len(DefStatic.CopyFieldName) > 0 Then
          For i = 1 To m_FieldSpecs.Count
            Set isp = m_FieldSpecs(i)
            If StrComp(isp.DestField, DefStatic.CopyFieldName, vbTextCompare) = 0 Then
              ' Make link
              Set ispCopy.CopyField = isp
              ispCopy.CopyFieldKey = isp.FieldKey
              ispCopy.key = KEY_COPIEDFIELD
              If StrComp(ispCopy.DestField, DefStatic.DestField, vbTextCompare) <> 0 Then
                ispCopy.DestField = DefStatic.DestField
                m_FGDrop.FromForm = frmlink.Name
                m_FGDrop.FromFG = frmlink.FlG_DestInfo.Name
                m_FGDrop.FromRow = -1
                For j = 1 To frmlink.FlG_DestInfo.Rows - 1
                  If StrComp(frmlink.FlG_DestInfo.TextMatrix(j, 0), ispCopy.DestField, vbTextCompare) = 0 Then
                    m_FGDrop.FromRow = j
                  End If
                Next j
                ispCopy.DestField = ""
                
                m_FGDrop.FromCol = 0
                m_FGDrop.ToForm = frmlink.Name
                m_FGDrop.ToFG = frmlink.FlG_Source.Name
                m_FGDrop.ToRow = 0
                m_FGDrop.ToCol = -1
                For j = 1 To frmlink.FlG_Source.cols - 1
                  If frmlink.FlG_Source.Coldata(j) = ispCopy.FieldKey Then
                    m_FGDrop.ToCol = j
                  End If
                Next j
                'Exit For
              End If
              Exit For
            End If
          Next i
          If StrComp(isp.DestField, DefStatic.CopyFieldName, vbTextCompare) <> 0 Then
            ' Remove link
            ispCopy.key = KEY_CONSTANT
            m_FGDrop.FromForm = frmlink.Name
            m_FGDrop.FromFG = frmlink.FlG_Source.Name
            m_FGDrop.FromRow = 0
            m_FGDrop.FromCol = -1
            For j = 1 To frmlink.FlG_Source.cols - 1
              If frmlink.FlG_Source.Coldata(j) = ispCopy.FieldKey Then
                m_FGDrop.FromCol = j
              End If
            Next j
            m_FGDrop.ToForm = frmlink.Name
            m_FGDrop.ToFG = frmlink.FlG_DestInfo.Name
            m_FGDrop.ToRow = -1
            For j = 1 To frmlink.FlG_DestInfo.Rows - 1
              If StrComp(frmlink.FlG_DestInfo.TextMatrix(j, 0), ispCopy.DestField, vbTextCompare) = 0 Then
                m_FGDrop.ToRow = j
              End If
            Next j
            m_FGDrop.ToCol = 0
            'Exit For
          End If
        End If
        If Len(m_FGDrop.FromForm) > 0 Then Call MakeLink
      End If
    Next DefStatic
  End If
  
  'Call MakeLink
  InDefCopyFields = False
  Call MakeFieldNamesUnique(m_FieldSpecs)
  
  End If
     
AddInDefaultCopiedFields_End:
  Call xReturn("AddInDefaultCopiedFields")
  Exit Sub

AddInDefaultCopiedFields_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddInDefaultCopiedFields", "Add In Default Copied Fields", "Error adding in the default copied fields.")
  Resume AddInDefaultCopiedFields_End
End Sub

Public Function AddConstraint(ByVal Name As String, ByVal FieldName As String, ByVal DataType As DATABASE_FIELD_TYPES, ByVal ImportConstraint As IMPORT_CONSTRAINT, ByVal Value As Variant) As Boolean
  Dim drs As DestRecordSet, drstemp As DestRecordSet
  Dim Con As DefaultConstraint, AddIn As Boolean
  
  On Error GoTo AddConstraint_Err
  Call xSet("AddConstraint")
  AddConstraint = False
  Set drs = m_DestRecSets.Item(Name)
  If IsMissing(Value) Then Call Err.Raise(ERR_ADDCONSTRAINT, "AddConstraint", "Attempt to add constraint wihout Value.")
  If InCollection(drs.Constraints, FieldName) Then
    Set Con = drs.Constraints.Item(FieldName)
    AddIn = False
  Else
    Set Con = New DefaultConstraint
    AddIn = True
  End If
  Select Case ImportConstraint
    Case CONSTRAINT_MIN
      If (Con.Constraint And CONSTRAINT_MIN) = CONSTRAINT_MIN Then Call Err.Raise(ERR_ADDCONSTRAINT, "AddConstraint", "Min constraint is already set.")
      Con.minvalue = Value
    Case CONSTRAINT_MAX
      If (Con.Constraint And CONSTRAINT_MAX) = CONSTRAINT_MAX Then Call Err.Raise(ERR_ADDCONSTRAINT, "AddConstraint", "Max constraint is already set.")
      Con.maxvalue = Value
    Case CONSTRAINT_MASK
      If (Con.Constraint And CONSTRAINT_MASK) = CONSTRAINT_MASK Then Call Err.Raise(ERR_ADDCONSTRAINT, "AddConstraint", "Mask constraint is already set.")
      Con.Mask = Value
    Case CONSTRAINT_INLIST
      If (Con.Constraint And CONSTRAINT_EXCLUDELIST) = CONSTRAINT_EXCLUDELIST Then Call Err.Raise(ERR_ADDCONSTRAINT, "AddConstraint", "Exclude List constraint is already set. Cannot set include list constraint.")
      Call Con.ConstraintList.Add(CStr(Value)) 'apf check stringlist change
    Case CONSTRAINT_EXCLUDELIST
      If (Con.Constraint And CONSTRAINT_INLIST) = CONSTRAINT_INLIST Then Call Err.Raise(ERR_ADDCONSTRAINT, "AddConstraint", "Include List constraint is already set. Cannot set exclude list constraint.")
      Call Con.ConstraintList.Add(CStr(Value))
    Case CONSTRAINT_IGNORE_CONVERSIONERRORS
    Case Else
      Call ECASE("AddContraint - unrecognised constraint: " & CStr(ImportConstraint))
  End Select
  Con.Constraint = Con.Constraint Or ImportConstraint
  Con.DestField = FieldName
  Con.DataType = DataType
  If AddIn Then Call drs.Constraints.Add(Con, Con.DestField)
  AddConstraint = True

AddConstraint_End:
  Call xReturn("AddConstraint")
  Exit Function

AddConstraint_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddConstraint", "Add Constraint to Field", "Error trying to add a constraint to field '" & FieldName & "' in recordset '" & Name & "'.")
  Resume AddConstraint_End
End Function

Private Sub AddInDefaultConstraints(ByVal FormName As String)
  Dim DefConstr As DefaultConstraint, isp As ImportSpec, i As Long, j As Long
  
  On Error GoTo AddInDefaultConstraints_Err
  Call xSet("AddInDefaultConstraints")
  
  For i = 1 To m_FieldSpecs.Count
    Set isp = m_FieldSpecs(i)
    If Len(isp.DestField) = 0 Then
      isp.Constraint = 0
    Else
      If Not CurrentDest.Constraints Is Nothing Then
        For Each DefConstr In CurrentDest.Constraints
          If StrComp(isp.DestField, DefConstr.DestField, vbTextCompare) = 0 Then
            isp.Constraint = DefConstr.Constraint
            isp.Mask = DefConstr.Mask
            isp.minvalue = DefConstr.minvalue
            isp.maxvalue = DefConstr.maxvalue
            isp.ConstraintList.RemoveAll
            For j = 1 To DefConstr.ConstraintList.Count
              isp.ConstraintList.Add (DefConstr.ConstraintList.Item(j))
            Next j
          End If
        Next DefConstr
      End If
    End If
  Next i
  
AddInDefaultConstraints_End:
  Call xReturn("AddInDefaultConstraints")
  Exit Sub

AddInDefaultConstraints_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddInDefaultConstraints", "Add In Default Constraints", "Error trying to add in default cnstraints to the import.")
  Resume AddInDefaultConstraints_End
End Sub

Public Function AddLockUpdate(ByVal Name As String, ByVal LockUpdate As IMPORT_UPDATE_TYPE) As Boolean
  Dim drs As DestRecordSet
  On Error GoTo AddLockUpdate_err
  
  Set drs = m_DestRecSets.Item(Name)
  drs.LockUpdate = LockUpdate
  
AddLockUpdate_end:
  Exit Function
  
AddLockUpdate_err:
  Call ErrorMessage(ERR_ERROR, Err, "AddLockUpdate", "Add Lock Update", "Error trying to add a lockupdate value to '" & Name & "'.")
  Resume AddLockUpdate_end
End Function

Public Function AddRequiredField(ByVal Name As String, ByVal FieldName As String, Optional ByVal PrimaryKey As Boolean = False) As Boolean
  Dim drs As DestRecordSet
  
  On Error GoTo AddRequiredField_Err
  Call xSet("AddRequiredField")
  AddRequiredField = False
  Set drs = m_DestRecSets.Item(Name)
  If PrimaryKey Then drs.PrimaryKeys.Add (FieldName)
  drs.RequiredFields.Add (FieldName)
  AddRequiredField = True

AddRequiredField_End:
  Call xReturn("AddRequiredField")
  Exit Function

AddRequiredField_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddRequiredField", "Add required field", "Error trying to add a required field '" & FieldName & "' in recordset '" & Name & "'.")
  Resume AddRequiredField_End
End Function

Public Function AddDefaultValue(ByVal Name As String, ByVal FieldName As String, ByVal DataType As DATABASE_FIELD_TYPES, ByVal Value As Variant, Optional ByVal NonNullValue As Variant) As Boolean
  Dim dType As DATABASE_FIELD_TYPES
  Dim drs As DestRecordSet, drstemp As DestRecordSet
  Dim DVal As DefaultValue
  
  On Error GoTo AddDefaultValue_Err
  Call xSet("AddDefaultValue")

  AddDefaultValue = False
  Set drs = m_DestRecSets.Item(Name)
  Set DVal = New DefaultValue
  DVal.DestField = FieldName
  DVal.Value = Value
  If Not IsMissing(NonNullValue) Then DVal.NonNullValue = NonNullValue
  Call drs.DefaultValues.Add(DVal)
  AddDefaultValue = True

AddDefaultValue_End:
  Call xReturn("AddDefaultValue")
  Exit Function

AddDefaultValue_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddDefaultValue", "Add default value", "Error trying to add a default value to field '" & FieldName & "' in recordset '" & Name & "'.")
  Resume AddDefaultValue_End
End Function


Private Function AddInDefaultValues(ByVal FormName As String) As Boolean
  Dim frmlink As Frm_Link, DVal As DefaultValue, isp As ImportSpec
  Dim i As Long

  On Error GoTo AddInDefaultValues_Err
  Call xSet("AddInDefaultValues")
  
  Set frmlink = m_ImportForms.Item(CStr(TCSIMP_LINK))
  For i = 1 To m_FieldSpecs.Count
    Set isp = m_FieldSpecs(i)
    If Len(isp.DestField) > 0 Then
      For Each DVal In CurrentDest.DefaultValues
        If StrComp(DVal.DestField, isp.DestField, vbTextCompare) = 0 Then
          isp.DefaultValue = DVal.Value
          isp.DefaultNonNullValue = DVal.NonNullValue
          Exit For
        End If
      Next DVal
    End If
  Next i

AddInDefaultValues_End:
  Call xReturn("AddInDefaultValues")
  Exit Function

AddInDefaultValues_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddInDefaultValues", "Add In Default Values", "Error adding in the default values.")
  Resume AddInDefaultValues_End
End Function

Public Property Get ProgressBar() As Object
  Dim frm As Frm_End
  Set frm = m_ImportForms.Item(CStr(TCSIMP_END))
  Set ProgressBar = frm.PBar_Import
End Property

'MPSMarch
Public Sub SetDataSource(ByVal DataSourceName As String)
  Dim drs As DestRecordSet
  
  For Each drs In m_DestRecSets
    If StrComp(DataSourceName, drs.Name, vbTextCompare) = 0 Then
      Set CurrentDest = drs
      Exit Sub
    End If
  Next drs
End Sub


'MPSMarch2
Friend Sub DisplayFactorFormatting()
  Dim ColNo As Long, frmMisc As Frm_Misc
  
  Set frmMisc = m_ImportForms.Item(CStr(TCSIMP_MISC))
  ColNo = GetColIndex(frmMisc.FlG_Source.Coldata(frmMisc.FlG_Source.Col))
  
  frmMisc.Chk_Factor.Value = -(m_FieldSpecs(ColNo).varaction = ACTION_FACTOR)
  frmMisc.Lbl_Factor.Enabled = (m_FieldSpecs(ColNo).varaction = ACTION_FACTOR) And (frmMisc.Fra_Format.Enabled)
  frmMisc.Txt_Factor.Enabled = (m_FieldSpecs(ColNo).varaction = ACTION_FACTOR) And (frmMisc.Fra_Format.Enabled)
  If (m_FieldSpecs(ColNo).varaction = ACTION_FACTOR) Then
    frmMisc.Txt_Factor.Text = m_FieldSpecs(ColNo).actionparam
  Else
    frmMisc.Txt_Factor.Text = "1"
  End If
End Sub


