VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private m_DoEventsCount As Long
'cad xx
Private m_Linecount As Long
Private m_ReportTarget As REPORT_TARGET
Private m_frmCancel As frmCancel
Private m_Cancel As Boolean
Private m_UseCancel As Boolean

Private m_DoReportOnCancel As Boolean
Private m_AllowSeparateBranchFieldSelection As Boolean

Public Enum REPORTW_GOTOFORM
  TCSREPWIZ = 1
End Enum

Public Enum REPORTW_TARGET
  REPORTW_LOAD_FILE = 0
  REPORTW_PRINT = PRINT_REPORT                        ' print to printer
  REPORTW_PREVIEW = PREPARE_REPORT                    ' preview to screen
  [_REPORTW_EXPORT] = 4096
  REPORTW_PREPARE_EXPORT = PREPARE_REPORT + [_REPORTW_EXPORT] ' used to prepare for export
End Enum

Public Enum SELECTITEM_ACTION
  SELECTITEM_ERROR = 0
  SELECTITEM_OK
  SELECTITEM_NO_REMOVE
  SELECTITEM_NO_ADD
End Enum

'  PREPARE_REPORT = 255
'  PRINT_REPORT

Private m_ReportDataSets As DataSetCollection
Private m_SessionFileName As String
Private m_IRepPP As IReportPreProcess
Private m_ReportFileName As String

' Details and Fields
Private mReportDetails As ReportDetails
Private mReportFields As Collection

Private m_ReportFilesPath As String
Private m_AutoPreviewTarget As REPORTW_TARGET

Private mAuto As AutoClass
Private mRep As Reporter
Private m_IgnoreUserError As Boolean

Private rFrm As IReportForm
Private FrmRepWiz As Frm_RepWiz
Private m_ReportForms As Collection
Private m_CurrentFrame As String

Private m_ReportInterface As ITCSUDM
Private m_SessionOpened As Boolean
Private m_ReportNotifyInterface As IBaseNotify

Private m_TitlePrefix As String
Private m_DisallowFileSelection As Boolean
Private m_AllowDataTypeChange As Boolean

' Export file
Private m_ExportFileName As String
Private m_ExportPath As String
Private m_ExportType As REPORT_EXPORTS

Private m_SPanel As TCSPANEL
Private m_CtrlCodes As Variant
Private m_CtrlCodesText As Variant
Private m_CtrlCodesDesc As Variant

Private m_SelectedNodeCount As Long
Private m_NotifyLineMeetsCriteria As Boolean
Private m_ReporterInternal As Reporter
Private m_ReportInError As Boolean
' Abacus Reporter properties
#If AbacusReporter Then
  Private m_FileGroupContainer As ARFileGroupContainer
#End If
Friend Property Get ReporterInternal() As Reporter
  Set ReporterInternal = m_ReporterInternal
End Property
Friend Property Let Cancel(NewValue As Boolean)
  m_Cancel = NewValue
End Property

Friend Function RaiseCancel(Optional bRaiseError As Boolean = True) As Boolean
 Static bInFunction As Boolean
 
 If bInFunction Then Exit Function
 
 bInFunction = True
 m_DoEventsCount = m_DoEventsCount + 1
 If m_DoEventsCount Mod L_DOEVENTS_MOD_COUNT = 0 Then
   DoEvents
 End If
 If m_Cancel Then
   m_Cancel = False
   RaiseCancel = True
   If bRaiseError Then
     bInFunction = False
     Call Err.Raise(ERR_CANCELREPORT, "Cancel", "Cancel report preparation")
   End If
 End If
 bInFunction = False
End Function

Public Property Get IncludeBlankLines()  'apf change next time
  If mReportDetails Is Nothing Then Err.Raise 380, "IncludeBlankLines", "Not in report Wizard, no report details object available"
  IncludeBlankLines = mReportDetails.LeftJoinLeaves
End Property

Public Function ExitWizard() As Boolean
  Dim frm As Form
  On Error Resume Next
  
  #If AbacusReporter Then
    If m_FileGroupContainer.Dirty Then
      Call m_FileGroupContainer.Save(FrmRepWiz, mReportDetails)
    End If
  #End If
  If Not m_ReportInterface Is Nothing Then Call m_ReportInterface.CloseSession
  If Not m_ReportForms Is Nothing Then
    For Each rFrm In m_ReportForms
      Set rFrm.ReportWizard = Nothing
      Set rFrm.ReportDetails = Nothing
      Set rFrm.ReportFields = Nothing
      Set frm = rFrm
      Unload frm
    Next rFrm
  End If
  Set m_ReportForms = Nothing
  Set mReportDetails = Nothing
  If Not mAuto Is Nothing Then
    mAuto.Kill
    Set mAuto = Nothing
  End If
  Set mRep = Nothing
  If (Not m_ReportInterface Is Nothing) And m_SessionOpened Then
    Call m_ReportInterface.CloseSession
    m_SessionOpened = False
  End If
  Set m_ReportInterface = Nothing
  Set m_ReportNotifyInterface = Nothing
  'apf ms all other kill/cleanup here
End Function

Public Property Get CtrlCodes() As Variant
  CtrlCodes = m_CtrlCodes
End Property

Public Property Get CtrlCodesText() As Variant
  CtrlCodesText = m_CtrlCodesText
End Property

Public Property Get CtrlCodesDesc() As Variant
  CtrlCodesDesc = m_CtrlCodesDesc
End Property

Private Sub Class_Initialize()
  On Error Resume Next
  Set m_ReportForms = New Collection
  
  Set rFrm = Frm_RepWiz
  Call m_ReportForms.Add(rFrm, CStr(rFrm.FormType))
  m_ExportPath = AppPath
  m_ReportFileName = ""
  m_UseCancel = True
  Set mAuto = New AutoClass
  Set RepParser = New Parser
  Set RepParser.ParseSettings = New ReportWizardParserSettings
    
  Call RepParser.AddParseItem(New ParseAlignHeaders)
  Call RepParser.AddParseItem(New ParseAlignment)
  Call RepParser.AddParseItem(New ParseAutoWidth)
  Call RepParser.AddParseItem(New ParseBooleanFalse)
  Call RepParser.AddParseItem(New ParseBooleanTrue)
  Call RepParser.AddParseItem(New ParseCollapseAll)
  Call RepParser.AddParseItem(New ParseCriterion)
  Call RepParser.AddParseItem(New ParseDataType)
  Call RepParser.AddParseItem(New ParseDisplay)
  Call RepParser.AddParseItem(New ParseFieldName)
  Call RepParser.AddParseItem(New ParseFitToPage)
  Call RepParser.AddParseItem(New ParseFixedWidth)
  Call RepParser.AddParseItem(New ParseFont)
  Call RepParser.AddParseItem(New ParseFormat)
  Call RepParser.AddParseItem(New ParseGHDelimiter)
  Call RepParser.AddParseItem(New ParseGHSeparator)
  Call RepParser.AddParseItem(New ParseGroup)
  Call RepParser.AddParseItem(New ParseHideGHTypes)
  Call RepParser.AddParseItem(New ParseLeftJoinLeaves)
  Call RepParser.AddParseItem(New ParseIncludeDate)
  Call RepParser.AddParseItem(New ParseIncludeHeaders)
  Call RepParser.AddParseItem(New ParseIncludePageFooter)
  Call RepParser.AddParseItem(New ParseIncludePageHeader)
  Call RepParser.AddParseItem(New ParseIncludePageNums)
  Call RepParser.AddParseItem(New ParseIncludeRepHeader)
  Call RepParser.AddParseItem(New ParseIncludeTime)
  Call RepParser.AddParseItem(New ParseKey)
  Call RepParser.AddParseItem(New ParseLocateDate)
  Call RepParser.AddParseItem(New ParseLocatePageNums)
  Call RepParser.AddParseItem(New ParseLocateTime)
  Call RepParser.AddParseItem(New ParseName)
  Call RepParser.AddParseItem(New ParseNoSquash)
  Call RepParser.AddParseItem(New ParseOrientation)
  Call RepParser.AddParseItem(New ParsePageFooterC)
  Call RepParser.AddParseItem(New ParsePageFooterL)
  Call RepParser.AddParseItem(New ParsePageFooterR)
  Call RepParser.AddParseItem(New ParsePageHeaderC)
  Call RepParser.AddParseItem(New ParsePageHeaderL)
  Call RepParser.AddParseItem(New ParsePageHeaderR)
  Call RepParser.AddParseItem(New ParsePrefix)
  Call RepParser.AddParseItem(New ParsePrintRecordCount)
  Call RepParser.AddParseItem(New ParseRepHeaderC)
  Call RepParser.AddParseItem(New ParseRepHeaderL)
  Call RepParser.AddParseItem(New ParseRepHeaderR)
  Call RepParser.AddParseItem(New ParseRHonAllPages)
  Call RepParser.AddParseItem(New ParseSort)
  Call RepParser.AddParseItem(New ParseSuffix)
  Call RepParser.AddParseItem(New ParseSumType)
  Call RepParser.AddParseItem(New ParseTextWidth)
  Call RepParser.AddParseItem(New ParseTrimHeadings)
  Call RepParser.AddParseItem(New ParseWidth)
  Call RepParser.AddParseItem(New ParseWrap)
  
  'RK New AR parseitems
  #If AbacusReporter Then
    Call RepParser.AddParseItem(New ARParseFileGroups)
    Call RepParser.AddParseItem(New ARParseFileGroupSelection)
    Call RepParser.AddParseItem(New ARParseAbacusProductType)
    If AR Then Set m_FileGroupContainer = New ARFileGroupContainer
  #End If
End Sub

Public Property Set ReportInterface(ByVal NewValue As TCSBASE.ITCSUDM)
  Set m_ReportInterface = NewValue
  If Not m_ReportInterface.Notify Is Nothing Then
    Set m_ReportNotifyInterface = m_ReportInterface.Notify
  End If
End Property

Public Property Get ReportInterface() As TCSBASE.ITCSUDM
  Set ReportInterface = m_ReportInterface
End Property

Public Property Let ExportFileName(ByVal NewValue As String)
  Dim ePath As String, eFile As String, eExt As String
  
  ' full path or valid filename
  Call SplitPath(NewValue, ePath, eFile, eExt)
  ' attempt to set the path
  If Len(ePath) > 0 Then Me.ExportPath = ePath
  m_ExportFileName = eFile & eExt
End Property

Public Property Let ExportPath(ByVal NewValue As String)
  If Not FileExists(NewValue, True) Then Call Err.Raise(380, "ExportPath", "Export path '" & NewValue & "' does not exist")
  m_ExportPath = NewValue
End Property

Public Property Let ExportType(ByVal ExportType As REPORT_EXPORTS)
  m_ExportType = ExportType
End Property

Public Property Let SessionFileName(ByVal NewValue As String)
  If Not FileExists(NewValue) Then Call Err.Raise(380, "SessionFileName", "Session file '" & NewValue & "' does not exist")
  m_SessionFileName = NewValue
End Property

Public Property Get ReportParser() As Parser
  Set ReportParser = RepParser
End Property

Public Property Set Repr(ByVal NewValue As Reporter)
  Set mRep = NewValue
End Property
Public Property Get Repr() As Reporter
  Set Repr = mRep
End Property
Public Property Set SPanel(ByVal NewValue As TCSPANEL)
  Set m_SPanel = NewValue
End Property

Public Property Get SPanel() As TCSPANEL 'RK REMOVE!
  Set SPanel = m_SPanel
End Property

'RK Addition for m_frmCancel
Friend Property Get ProgressBar() As Object
  'Following addition of m_frmCancel
  '(Frm_RepWiz.SBar replace with m_frmCancel.PBar & m_frmCancel.lblStatus)
  If m_UseCancel Then
   Set ProgressBar = m_frmCancel.PBar
  Else
   Set ProgressBar = FrmRepWiz.SBar.prg
  End If
End Property

Friend Sub ProgressBar_StopPrg()
  'RK Addition sub to replicate StatusBar functionality
  'Replaces Me.ProgressBar_StopPrg
  
  With Me.ProgressBar
    .Indicator = None
    .Caption = ""
    .Value = .Min
    .Max = .Min + 1
  End With
End Sub

Public Property Let ReportFileName(ByVal NewValue As String)
  If Not FileExists(NewValue) Then Call Err.Raise(380, "ReportFileName", "Report file '" & NewValue & "' does not exist")
  m_ReportFileName = NewValue
End Property

Public Property Get ReportFileName() As String
  ReportFileName = m_ReportFileName
End Property

Public Property Get AutoPreviewTarget() As REPORTW_TARGET
  AutoPreviewTarget = m_AutoPreviewTarget
End Property

Public Property Let AutoPreviewTarget(ByVal NewValue As REPORTW_TARGET)
  m_AutoPreviewTarget = NewValue
End Property

Public Property Get IgnoreUserError() As Boolean
  IgnoreUserError = m_IgnoreUserError
End Property

Public Property Let IgnoreUserError(ByVal NewValue As Boolean)
  m_IgnoreUserError = NewValue
End Property

Public Property Let ReportFilesPath(ByVal NewValue As String)
  If Not FileExists(NewValue, True) Then Call Err.Raise(380, "ReportFilesPath", "Report path '" & NewValue & "' does not exist")
  m_ReportFilesPath = NewValue
End Property

Public Property Get ReportFilesPath() As String
  ReportFilesPath = m_ReportFilesPath
End Property

Public Function StartReportWizard() As Boolean
  'Dim frm As Form
  
  Dim RepVar As Variant, i As Long
  Dim hwndTV As Long
  Static Started As Boolean
  
  If Started Then
    Call ECASE("Cannot call StartReportWizard more than once in same ReportWizard instance!")
    GoTo StartReportWizard_End
  Else
    Started = True
  End If
  
  On Error GoTo StartReportWizard_Err
  Call xSet("StartReportWizard")
  
  Set m_ReportDataSets = New DataSetCollection
  Set mReportDetails = New ReportDetails
  Set mReportFields = New Collection

  Set FrmRepWiz = m_ReportForms.Item(CStr(TCSREPWIZ))
  
  Set m_SPanel = FrmRepWiz.SBar.AddPanel(50)

  For Each rFrm In m_ReportForms
    Set rFrm.ReportWizard = Me
    Set rFrm.ReportDetails = mReportDetails
    Set rFrm.ReportFields = mReportFields
    Set rFrm.ReportDataSets = m_ReportDataSets
  Next rFrm
  
  'cad review, AbacusReporter only?
  'RK addition to allow override of Report Wizard name
  If Len(m_TitlePrefix) = 0 Then
    m_TitlePrefix = m_ReportInterface.Name & " Report Wizard - "
  Else
    m_TitlePrefix = m_TitlePrefix & " - "
  End If
    
  Set mRep = New Reporter
  hwndTV = FrmRepWiz.BeginWiz
  
  m_SessionOpened = m_ReportInterface.OpenSession(m_SessionFileName, False)
  Call ReadReport
  
  'Populate Treeview
  FrmRepWiz.TrV_Fields.Nodes.Clear
  FrmRepWiz.lstFields.Clear
  Call FillTreeView(FrmRepWiz.TrV_Fields, m_ReportDataSets)
  If FileExists(m_ReportFileName) Then
    If Not LoadReportDetails(RepParser, Me, FrmRepWiz, m_ReportDataSets, mReportDetails, mReportFields, m_ReportFileName) Then GoTo StartReportWizard_End
    If m_AutoPreviewTarget <> REPORTW_LOAD_FILE Then
      Call PrepareReport(-1, m_AutoPreviewTarget)
      GoTo StartReportWizard_End
    End If
  End If
  
  'Apply File Group Selection from .rep file
  #If AbacusReporter Then
    Call ApplyFileGroupSelection(Me, FrmRepWiz, mReportDetails)
  #End If
  
  FrmRepWiz.InFillFormat = True
  FrmRepWiz.Show vbModal
 
StartReportWizard_End:
  Call ExitWizard
  Call xReturn("StartReportWizard")
  Exit Function
  
StartReportWizard_Err:
  If Err.Number = 401 Then FrmRepWiz.Show vbModal
  Call ErrorMessage(ERR_ERROR, Err, "StartReportWizard", "Start report wizard", "Error in start report wizard")
  Resume StartReportWizard_End
  Resume
End Function

Private Sub ReadReport()
  Dim DataSets As Variant, Fields As Variant
  Dim CurDataSet() As Variant, fldBase As FieldDetails
  Dim rFld As ReportField
  Dim i As Long, j As Long
  Dim dSet As ReportDataSet
    
  On Error GoTo ReadReport_Err
  
  Call m_ReportDataSets.RemoveAll
  DataSets = m_ReportInterface.QueryAllTopLevelDataSets
  
  If Not IsArray(DataSets) Then Call Err.Raise(ERR_READREPORT, "ReadReport", "No top level datasets found in interface.")
  ReDim CurDataSet(LBound(DataSets) To LBound(DataSets))
  For i = LBound(DataSets) To UBound(DataSets)
    Set dSet = New ReportDataSet
    dSet.Name = DataSets(i)
    CurDataSet(LBound(DataSets)) = dSet.Name
    dSet.CurrentDataSet = CurDataSet
    Call dSet.GetFieldDetails(m_ReportInterface)
    Call dSet.ReadChildren(m_ReportInterface)
    m_ReportDataSets.Add dSet, dSet.CurrentDataSetString
  Next i
ReadReport_End:
  Exit Sub
  
ReadReport_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadReport", "Read report data sets", "Unable to retrieve dataset information." & STR_CR & "Report Interface: " & m_ReportInterface.Name)
  Resume ReadReport_End
  Resume
End Sub

Private Sub SplitKey(DataSet As String, Name As String, ByVal Key As String)
  Dim p As Long
  
  p = InStrRev(Key, DATASET_SEP)
  If p = 0 Then Call Err.Raise(ERR_SPLITKEY, "SplitKey", "Invalid key " & Key)
  DataSet = left$(Key, p - 1)
  Name = Mid$(Key, p + 1)
End Sub

Private Function IsValidSelection(ByVal Key As String) As Boolean
  Dim Name As String, DataSet As String
  Dim rFld As ReportField
  Dim SelDataSet As String, CompareLen As Long
  
  On Error GoTo IsValidSelection_Err
  
    'RK: This logic prevents you selecting nodes on different root branches of a UDM Hierarchy.
    'This logic maybe required for Tracker where the UDM hierarchy is generated dynamically.
    'Have added a property to remove this check for Reporter.
  If Not m_AllowSeparateBranchFieldSelection Then
    Call SplitKey(DataSet, Name, Key)
    For Each rFld In mReportFields
      SelDataSet = rFld.DataSet.CurrentDataSetString
      If rFld.DataSet.IsLeafNode Then
        If Not rFld.DataSet.Parent Is Nothing Then SelDataSet = rFld.DataSet.Parent.CurrentDataSetString
      End If
      CompareLen = Min(Len(SelDataSet), Len(DataSet))
      IsValidSelection = (StrComp(left$(SelDataSet, CompareLen), left$(DataSet, CompareLen), vbTextCompare) = 0)
      If Not IsValidSelection Then Exit Function
    Next rFld
  End If
  IsValidSelection = True
  
IsValidSelection_End:
  Exit Function
  
IsValidSelection_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsValidSelection", "IsValidSelection", "Unable to determine validity of selected field.")
  IsValidSelection = False
  Resume IsValidSelection_End
End Function

Private Function GetDataSet(ByVal DataSet As String, Optional Depth As Long) As ReportDataSet
  Dim dSet As ReportDataSet, dSetParent As ReportDataSet
  Dim p As Long, NextDataSet As String
  
  On Error GoTo GetDataSet_Err
  Depth = 1
  p = InStr(1, DataSet, DATASET_SEP, vbTextCompare)
  If p = 0 Then
    Set dSet = m_ReportDataSets(DataSet)
  Else
    NextDataSet = left$(DataSet, p - 1)
    Set dSetParent = m_ReportDataSets(NextDataSet)
    Do
      p = InStr(p + 1, DataSet, DATASET_SEP, vbTextCompare)
      If p = 0 Then
        NextDataSet = DataSet
      Else
        NextDataSet = left$(DataSet, p - 1)
      End If
      Set dSet = dSetParent.Children(NextDataSet)
      Depth = Depth + 1
      Set dSetParent = dSet
    Loop Until p = 0
  End If
  Set GetDataSet = dSet
    
GetDataSet_End:
  Exit Function
  
GetDataSet_Err:
  Set GetDataSet = Nothing
  Resume GetDataSet_End
End Function

Public Function SelectItem(ByVal Key As String, ByVal bSelectItem As Boolean) As SELECTITEM_ACTION
  Dim nod As node
  'Dim DataSet As String, FieldName As String
  Dim dSet As ReportDataSet
  Dim rFld As ReportField
   
  On Error GoTo SelectItem_Err
  SelectItem = SELECTITEM_ERROR
  If IsValidSelection(Key) Then
    'Call SplitKey(DataSet, FieldName, Key)
    'Set DSet = GetDataSet(DataSet)
    'Set rFld = DSet.Fields(Key)
    Set rFld = GetReportFieldFromKey(dSet, Key)
    rFld.Selected = Not bSelectItem
    If rFld.Selected Then
      SelectItem = SELECTITEM_NO_ADD
      If Not InCollection(mReportFields, rFld.KeyString) Then
        rFld.Order = mReportFields.Count + 1
        mReportFields.Add rFld, rFld.KeyString
        Call rFld.Reinitialise
        SelectItem = SELECTITEM_OK
        Do
          Set nod = Frm_RepWiz.TrV_Fields.Nodes(dSet.CurrentDataSetString)
          nod.Image = SELECTED_PARENT
          Set dSet = dSet.Parent
        Loop Until dSet Is Nothing
      End If
    Else
      SelectItem = SELECTITEM_NO_REMOVE
      If InCollection(mReportFields, rFld.KeyString) Then
        Call mReportFields.Remove(rFld.KeyString)
        SelectItem = SELECTITEM_OK
        Do
          If dSet.Selected Then Exit Do
          Set nod = Frm_RepWiz.TrV_Fields.Nodes(dSet.CurrentDataSetString)
          nod.Image = NOTSELECTED_PARENT
          Set dSet = dSet.Parent
        Loop Until dSet Is Nothing
      End If
    End If
    FrmRepWiz.Cmd_Preview.Enabled = (mReportFields.Count > 0)
    FrmRepWiz.Cmd_Print.Enabled = (mReportFields.Count > 0)
    FrmRepWiz.Cmd_Export.Enabled = (mReportFields.Count > 0)
    FrmRepWiz.Cmd_SaveReport.Enabled = (mReportFields.Count > 0)
  End If
SelectItem_End:
  Exit Function
  
SelectItem_Err:
  SelectItem = SELECTITEM_ERROR
  Call ErrorMessage(ERR_ERROR, Err, "SelectItem", "Select data field", "Unable to select field")
  Resume SelectItem_End
  Resume
End Function

Private Function GetTreeStructureAndLeafNode(NodeCount As Long, TreeDepth As Long) As ReportDataSet
  Dim rFld As ReportField, dSet As ReportDataSet, i As Long
    
  TreeDepth = 0: NodeCount = 0
  For Each rFld In mReportFields
    If rFld.DataSet.Depth > TreeDepth Then
      TreeDepth = rFld.DataSet.Depth
      Set dSet = rFld.DataSet
    End If
    rFld.DataSet.NodeUsed = False
  Next rFld
  For Each rFld In mReportFields
    If Not rFld.DataSet.NodeUsed Then
      NodeCount = NodeCount + 1
      rFld.DataSet.NodeUsed = True
    End If
  Next rFld
  For Each rFld In mReportFields
    rFld.DataSet.NodeUsed = False
  Next rFld
  
  m_SelectedNodeCount = NodeCount
  
  Set GetTreeStructureAndLeafNode = dSet
  For i = TreeDepth To 1 Step -1  ' do path from leaf to root
    If Not dSet.Selected(True) Then
      NodeCount = NodeCount + 1
    End If
    Set dSet = dSet.Parent
  Next i
  
End Function

Private Sub OpenExportDialog(SaveFile As String, ExportType As REPORT_EXPORTS)
  Dim sFilter As String
  Dim i As Long, j As Long
  Dim FilterIndex As Long
  
  On Error GoTo OpenExportDialog_Err
  ExportType = 0
  For i = [_REPORT_EXPORTS_FIRST] To [_REPORT_EXPORTS_LAST]
    If mRep.ExportAvailable(i) Then sFilter = sFilter & mRep.ExportTypeStr(i) & " (*" & mRep.ExportTypeExt(i) & ")|*" & mRep.ExportTypeExt(i) & "|"
  Next i
  sFilter = left$(sFilter, Len(sFilter) - 1)
  SaveFile = FileSaveAsDlgFilter(FilterIndex, "Export Report", sFilter, m_ExportPath)
  If Len(SaveFile) > 0 Then
    j = 1
    For i = [_REPORT_EXPORTS_FIRST] To [_REPORT_EXPORTS_LAST]
      If mRep.ExportAvailable(i) Then
        If j = FilterIndex Then
          ExportType = i
          Exit For
        End If
        j = j + 1
      End If
    Next i
  End If

OpenExportDialog_End:
  Exit Sub
  
OpenExportDialog_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenExportDialog", "Choose Export File", "Error choosing export file" & vbCrLf & SaveFile)
  Resume OpenExportDialog_End
End Sub

Friend Sub PrepareReport(ByVal LineCount As Long, Optional ByVal rt As REPORTW_TARGET = REPORTW_PRINT)
  Dim sExportFile As String, re As REPORT_EXPORTS
  Dim bSaveAutoWidth As Boolean
  Dim bExcepted As Boolean, bEndedReport As Boolean, bAborted As Boolean
  On Error GoTo Err_Err
  
  m_Cancel = False
  If mRep Is Nothing Then
    Set m_ReporterInternal = New Reporter
  Else
    Set m_ReporterInternal = mRep
  End If
  
  m_Linecount = LineCount
  m_ReportTarget = rt
  
  bSaveAutoWidth = mReportDetails.AutoWidth
  Call SetCursor
  Call ResetIgnoreError
  
  If (m_ReportTarget = REPORTW_PREPARE_EXPORT) Then
    If m_ExportType = 0 Then  ' MPS APF invalid export type
      Call OpenExportDialog(sExportFile, re)
    Else
      sExportFile = FullPath(m_ExportPath) & m_ExportFileName
      re = m_ExportType
    End If
    If Len(sExportFile) = 0 Then GoTo Err_End
    If re = 0 Then
      Call ECASE("No export to that type of file")
    End If
    Call SplitPath(sExportFile, m_ExportPath)
  End If
  
    
  If m_UseCancel Then
    Set m_frmCancel = New frmCancel 'For status bar updating
    If m_frmCancel Is Nothing Then Call Err.Raise(ERR_ERROR, "ProgressBar", "frmCancel is Nothing.")
    Set m_frmCancel.rw = Me
    m_frmCancel.Show 1
    bExcepted = m_frmCancel.Excepted
    Unload m_frmCancel
    Set m_frmCancel = Nothing
    If bExcepted Then
      Call ErrorMessagePopErr(Err)
      Call Err.Raise(Err.Number, Err.Source, Err.Description)
    End If
  Else
    Call PrepareReportSub
  End If
  
  If Not m_ReporterInternal Is Nothing Then
    If m_ReportInError Then
      Call m_ReporterInternal.AbortReport
      bAborted = True
    End If
    Call m_ReporterInternal.EndReport(m_ReportInError)
    bEndedReport = True
    If Not m_ReportInError Then
      If (rt = REPORTW_PREPARE_EXPORT) Then
        Call m_ReporterInternal.ExportReport(sExportFile, re, True)
      End If
      If (m_ReportTarget = REPORTW_PREVIEW) Then
        Call m_ReporterInternal.PreviewReport
      End If
    End If
    m_ReportInError = False
  End If
  
Err_End:
  Call PrepareReportEnd(bSaveAutoWidth, bExcepted, bEndedReport, bAborted)
  Exit Sub
Err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PrepareReport", "Prepare output report", "Error preparing output report " & mReportDetails.ReportName)
  Resume Err_End
  Resume
End Sub
Private Sub PrepareReportEnd(ByVal bSaveAutoWidth As Boolean, ByVal bExcepted As Boolean, ByVal bEndedReport As Boolean, ByVal bAborted As Boolean)
  Call ClearCursor
      
  If Not mReportDetails Is Nothing Then mReportDetails.AutoWidth = bSaveAutoWidth
  
  If bExcepted And m_ReportInError Then
    If Not bAborted Then
      Call m_ReporterInternal.AbortReport
    End If
    If Not bEndedReport Then
      Call m_ReporterInternal.EndReport(True)
    End If
    m_ReportInError = False
    Set m_ReporterInternal = Nothing
  End If
  If Not m_frmCancel Is Nothing Then
    Unload m_frmCancel
    Set m_frmCancel = Nothing
  End If
End Sub

Friend Sub PrepareReportSub()
  
  
  Dim rt As REPORT_TARGET
  Dim t0 As Long
  t0 = GetTicks()
  m_ReportInError = True
  
  'take off export
  rt = (m_ReportTarget And (Not [_REPORTW_EXPORT]))
  If m_ReporterInternal.ReportActive Then rt = m_ReporterInternal.ReportTarget
     
  Status = "Connecting to printer"
  Call ProgressBar_StopPrg
  If m_ReporterInternal.InitReport("TestReport", rt, mReportDetails.Orientation) Then
    If (m_ReportTarget = REPORTW_PREPARE_EXPORT) Then mReportDetails.AutoWidth = False ' apf changed so that export does not get widthed
    m_ReportInError = Not PrepareReportEx(t0)
  End If
  
PrepareReport_End:
  Exit Sub
PrepareReport_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "StartUser"), Err.Description)
  
End Sub
Private Sub SavePreviousAutoWidth(ByVal bAutoWidth As Boolean)

End Sub

Private Function PrepareReportEx(ByVal t0 As Long) As Boolean
  Dim PrintOrder As Variant, PrintLine As Variant, FieldOffset As Long
  Dim DataSetPath() As ReportDataSet, MaxTreeDepth As Long, MaxTreeNodes As Long
  Dim ReportFields() As ReportField
  Dim dSet As ReportDataSet, i As Long, j As Long, AllEOF As Boolean, CurDepth As Long
  Dim MaxCol As Long
  Dim LastKey As Variant, SavedLastKey As Variant, SavedKeyPath As Variant
  Dim rFld As ReportField
  Dim MaxCriteria As Long
  Dim CurLine As Long
  Dim LineNum As Long, t1 As Long
  Dim bSetStatusForLines As Boolean
  #If AbacusReporter Then
  Dim tStart As Double
  #End If
  
  On Error GoTo PrepareReportEx_Err
  Call xSet("PrepareReport")

  m_DoEventsCount = 0
  If mReportFields.Count = 0 Then Call Err.Raise(ERR_PREPAREREPORT, "PrepareReportEx", "No fields selected to report")
  MaxCol = mReportFields.Count
  ReDim PrintLine(1 To MaxCol)
  ReDim PrintOrder(1 To MaxCol)   ' mapping from UDM order to PrintLine(i)
  ReDim ReportFields(1 To MaxCol) ' in Print order
  
  #If AbacusReporter Then
   Call CheckFileGroupSettings(Me, FrmRepWiz, mReportDetails)
   Call g_AbacusReporter.Session.Reset
   Call ARFileGroupCode.UpdateFileCatalogue(Me, FrmRepWiz, mReportDetails)
   Call g_AbacusReporter.Session.PopulateTempTable
  #End If
  
  Set dSet = GetTreeStructureAndLeafNode(MaxTreeNodes, MaxTreeDepth)
  If dSet Is Nothing Then Call Err.Raise(ERR_PREPAREREPORT, "PrepareReportEx", "Unable to acquire reference to leaf node")
  
  ReDim DataSetPath(1 To MaxTreeNodes) As ReportDataSet
  For i = MaxTreeDepth To 1 Step -1  ' do path from leaf to root
    Set DataSetPath(i) = dSet
    #If DEBUGVER Then
      Call OutputDebug("PrepareReportEx", dSet.Name & " Criteria: " & dSet.CriteriaSQL)
    #End If
    dSet.NodeUsed = True
    Set dSet = dSet.Parent
  Next i
  Set dSet = DataSetPath(MaxTreeDepth).Parent
  If (dSet Is Nothing) And (MaxTreeDepth <> MaxTreeNodes) Then Call Err.Raise(ERR_PREPAREREPORT, "PrepareReportEx", "Unable to traverse dataset tree, if multiple leaf nodes are selected then a parent node must exist")
  i = MaxTreeDepth   ' do all leaf nodes
  For Each rFld In mReportFields
    If Not rFld.DataSet.NodeUsed Then
      i = i + 1
      Set DataSetPath(i) = rFld.DataSet
      rFld.DataSet.NodeUsed = True
    End If
  Next rFld
  If (i <> MaxTreeNodes) Then Call Err.Raise(ERR_PREPAREREPORT, "PrepareReportEx", "Unable to traverse dataset tree, could not find all leaf nodes")

  ' Initialise data sets on DataSetPath in order
  ' Goes down the tree consisting of all the parent datasets of the deepest
  ' selected field, its parent data set, its parent data set etc.
  ' For each dataset, the FieldsSelected collection is created consisting
  ' of all the selected fields immediately within that dataset.
  ' Private property mFieldOffset is set to FieldOffset and FieldOffset is set to
  ' 1 plus the number of ALL selected fields already added in for this report.
  If Not m_IRepPP Is Nothing Then
    If m_IRepPP.BeforeReportPrepare(m_SessionFileName) Then Err.Raise ERR_CANCELREPORT, "PrepareReportEx", "Cancel report preparation - from BeforeReportPrepare"
  End If
  MaxCriteria = GetMaxCriteria(mReportFields)

  FieldOffset = 1
  For i = 1 To MaxTreeNodes
    Set dSet = DataSetPath(i)
    Call dSet.PrepareOutput(mReportFields, m_ReportInterface, FieldOffset, PrintOrder, ReportFields, MaxCriteria)
  Next i
  If (FieldOffset - 1) <> mReportFields.Count Then If dSet Is Nothing Then Err.Raise ERR_PREPAREREPORT, "PrepareReportEx", "Error in Prepare output"

  Call mReportDetails.AddIntoAuto(mAuto, Me)
  If Not InitReportColumns(mAuto, ReportFields, MaxCol, mReportDetails) Then Err.Raise ERR_PREPAREREPORT, "PrepareReportEx", "Unable to initialise report columns"
   
  Call mAuto.InitAutoReport(PrintLine, m_ReporterInternal)

  Status = "Initialising source application's data"
  #If AbacusReporter Then
    Call g_AbacusReporter.Session.Start
    'Call Err.Raise(ERR_ERROR, "monkey", "nuts")
    Call AutomateDataProviders(Me, FrmRepWiz, mReportDetails)
    If g_AbacusReporter.Session.ExceptionCount > 0 And g_AbacusReporter.IncludeExceptionsReport Then
      Call LoadExceptionsReport(Me, True)
    End If
    'RK QUERY: Force exceptions display when in no results are to be displayed?
    If g_AbacusReporter.IncludeResultsReport = False Then
      GoTo PREPARE_REPORT:
    End If
  #End If

  Me.ProgressBar.Value = 0
  LineNum = 0: CurLine = 0
  CurDepth = 1
  
  If Not m_IRepPP Is Nothing Then
    If m_IRepPP.BeforeReportOutput(m_SessionFileName, DataSetPath) Then Err.Raise ERR_CANCELREPORT, "PrepareReportEx", "Cancel report preparation - from BeforeReportOutput"
  End If
  
  Call Me.RaiseCancel
  Me.ProgressBar.Indicator = None
  Do
    #If AbacusReporter Then
      tStart = GetTickCount
    #End If
    If Not IsEmpty(SavedLastKey) Then
      LastKey = SavedLastKey
    Else
      LastKey = Empty
    End If
    CurDepth = GetNextLine(PrintLine, CurDepth, DataSetPath, MaxTreeDepth, MaxTreeNodes, LastKey, PrintOrder, SavedLastKey, SavedKeyPath)
    
    If Not bSetStatusForLines Then
      Me.Status = "Extracting data from source application"
      bSetStatusForLines = True
    End If
    
    If CurDepth = MaxTreeDepth Then
      LineNum = LineNum + 1
      If PrepareOutputLine(mAuto, m_ReporterInternal, ReportFields, PrintLine, MaxCol, MaxCriteria, Me) Then
      'If PrepareOutputLine(mAuto, rep, ReportFields, PrintLine, MaxCol, MaxCriteria) Then
        CurLine = CurLine + 1
      End If
      If m_ReportNotifyInterface Is Nothing Then
        If Me.ProgressBar.Max <> m_ReportInterface.RecordCount Then
          Me.ProgressBar.Max = m_ReportInterface.RecordCount
        End If
        If Me.ProgressBar.Value >= m_ReportInterface.RecordCount Then
          Me.ProgressBar.Value = 0
        End If
        Call Me.ProgressBar.StepCaption("Processing Line " & Trim$(LineNum))
      Else
        Call m_ReportNotifyInterface.Notify(CurLine, m_ReportInterface.RecordCount, "Processing Line " & Trim$(LineNum))
      End If
    End If
    If RaiseCancel(Not m_DoReportOnCancel) Then
      Exit Do
    End If
    #If AbacusReporter Then
    Call SetReportLineTimings(GetTickCount - tStart)
    #End If
  Loop Until (CurDepth <= 0) Or (CurLine = m_Linecount)

    
#If AbacusReporter Then
PREPARE_REPORT:   'RK Line Label to avoid process results report
#End If
  Me.ProgressBar_StopPrg
  Status = "Preparing Report"
  If mReportDetails.AutoWidth Then
    Call mAuto.PreviewDoColumnSpacing
    'For i = 1 To MaxCol
    '  ReportFields(i).Width = PercentDP(mAuto.Item(i).MaxWidth * 100 / rep.ReportPageWidth, 1)
    'Next i
  End If
  t1 = GetTicks()
  Status = "Data retrieved in " & Format$((t1 - t0) / 1000, "#,###.00") & " seconds."
  Call mAuto.PreviewOutput
  Call mAuto.EndAutoLine
  Me.ProgressBar.Value = 0
  If Not m_IRepPP Is Nothing Then Call m_IRepPP.AfterReportOutput(m_SessionFileName, DataSetPath)
  PrepareReportEx = True

PrepareReportEx_End:
  Call xReturn("PrepareReport")
  Exit Function

PrepareReportEx_Err:
  'RK QUERY: TEST WITH ERRORS RAISED In PROCESSING
  Call Err.Raise(Err.Number, ErrorSource(Err, "PrepareReportEx"), Err.Description)
  Resume PrepareReportEx_End
  Resume
End Function

Private Function GetNextLine(CurrentLine As Variant, ByVal CurDepth As Long, DataSetPath() As ReportDataSet, ByVal MaxTreeDepth As Long, ByVal MaxTreeNodes As Long, ByVal LastKey As Variant, PrintOrder As Variant, SavedLastKey As Variant, SavedKeyPath As Variant) As Long
  Static BlankLineOutput As Boolean
  Dim i As Long, dSet As ReportDataSet
  Dim CurKey As Variant
  Dim NotAllEOF As Boolean, OnLeafNode As Boolean
    
  On Error GoTo GetNextLine_Err
  OnLeafNode = False
  If CurDepth = MaxTreeDepth Then
    NotAllEOF = False
    For i = MaxTreeDepth To MaxTreeNodes
      Set dSet = DataSetPath(i)
      NotAllEOF = NotAllEOF Or dSet.GetLineEx(CurrentLine, CurKey, m_ReportInterface, LastKey, PrintOrder, m_IgnoreUserError)
    Next i
    OnLeafNode = NotAllEOF Or (mReportDetails.LeftJoinLeaves And _
                     (m_SelectedNodeCount > 1) And _
                     (Not BlankLineOutput))
    If OnLeafNode Then
      GetNextLine = CurDepth
      Call SetSavedLastKey(SavedLastKey, SavedKeyPath, (GetNextLine - 1))
    Else
      GetNextLine = CurDepth - 1
    End If
    BlankLineOutput = True
    GoTo GetNextLine_End
  Else
    BlankLineOutput = False
    Set dSet = DataSetPath(CurDepth)
    If Not dSet.GetLineEx(CurrentLine, CurKey, m_ReportInterface, LastKey, PrintOrder, m_IgnoreUserError) Then
      GetNextLine = CurDepth - 1
      GoTo GetNextLine_End
    End If
  End If
  Call SetMaxKeyPath(SavedKeyPath, CurKey)
  GetNextLine = GetNextLine(CurrentLine, CurDepth + 1, DataSetPath, MaxTreeDepth, MaxTreeNodes, CurKey, PrintOrder, SavedLastKey, SavedKeyPath)

GetNextLine_End:
  Call SetMaxKeyPath(SavedKeyPath, CurKey)
  Debug.Print "Depth: " & GetNextLine & "  (" & GetKeyAsString(SavedKeyPath) & ")"
  If GetNextLine < CurDepth Then
    If GetNextLine > 1 Then
      Call SetSavedLastKey(SavedLastKey, SavedKeyPath, (GetNextLine - 1))
      Debug.Print "Using CurrentID " & (GetNextLine - 1) & " *= " & GetNextLine & "  (" & GetKeyAsString(SavedLastKey) & ")"
    Else
      SavedLastKey = Empty
    End If
    'If Not OnLeafNode Then SavedKeyPath = Empty
  End If
  Exit Function
  
GetNextLine_Err:
  Resume GetNextLine_End
End Function

Private Sub SetSavedLastKey(SavedLastKey As Variant, SavedKeyPath As Variant, ByVal MaxIndex As Long)
  If (Not IsEmpty(SavedKeyPath)) And (MaxIndex > 0) Then
    SavedLastKey = SavedKeyPath
    ReDim Preserve SavedLastKey(1 To MaxIndex)
  Else
    SavedLastKey = Empty
  End If
End Sub

Private Sub SetMaxKeyPath(KeyPath As Variant, CurrentKey As Variant)
  Dim i As Long
  
  If Not IsEmpty(CurrentKey) Then
    If IsEmpty(KeyPath) Then
      KeyPath = CurrentKey
    ElseIf UBound(CurrentKey) > UBound(KeyPath) Then
      KeyPath = CurrentKey
    Else
      'Debug.Print "CurrentKey: " & GetKeyAsString(CurrentKey) & "  KeyPath: " & GetKeyAsString(KeyPath)
      For i = LBound(CurrentKey) To UBound(CurrentKey)
        KeyPath(i) = CurrentKey(i)
      Next i
    End If
  End If
End Sub

Private Function GetKeyAsString(vKey As Variant) As String
  Dim i As Long, s As String
  
  If IsEmpty(vKey) Then
    s = "empty"
  Else
    For i = LBound(vKey) To UBound(vKey)
      s = s & vKey(i)
      If i <> UBound(vKey) Then s = s & ", "
    Next i
  End If
  GetKeyAsString = s
End Function

'.....apf
Public Property Let AllowDataTypeChange(ByVal NewValue As Boolean)
  m_AllowDataTypeChange = NewValue
End Property

Public Property Get AllowDataTypeChange() As Boolean
  AllowDataTypeChange = m_AllowDataTypeChange
End Property

Public Property Let DisallowFileSelection(ByVal NewValue As Boolean)
  Call ECASE("'DisallowFileSelection' not implemented")
  m_DisallowFileSelection = NewValue
End Property

Public Property Get DisallowFileSelection() As Boolean
  Call ECASE("'DisallowFileSelection' not implemented")
  DisallowFileSelection = m_DisallowFileSelection
End Property

Public Property Let Title(ByVal NewValue As String)
  m_TitlePrefix = NewValue
End Property

Public Property Get Title() As String
  Title = m_TitlePrefix
End Property

Public Property Get FRW() As Object
  Set FRW = FrmRepWiz
End Property

Public Property Set FRW(ByVal vNewValue)
  Set FrmRepWiz = vNewValue
End Property

Public Property Get CurrentFrame() As String
  CurrentFrame = m_CurrentFrame
End Property

Public Property Let CurrentFrame(ByVal NewValue As String)
  m_CurrentFrame = NewValue
End Property
     
Public Sub SaveReport()
  Dim FileName As String
  Dim FileNum As Integer
  Dim rFld As ReportField
  Dim i As Long

  On Error GoTo SaveReport_Err
  Call xSet("SaveReport")
  'RK TODO InitialDirectory is not being set in FileSaveAsDlg
  FileName = FileSaveAsDlg("Save Report As", "Report Wizard Files (*.rep)|*.rep|All Files (*.*)|*.*", m_ReportFilesPath, m_ReportFileName)
  If Len(FileName) > 0 Then
    m_ReportFileName = FileName
    FileNum = FreeFile
    Open m_ReportFileName For Output As #FileNum
    Print #FileNum, mReportDetails.SaveDetails
    For i = 1 To mReportFields.Count
      For Each rFld In mReportFields
        If rFld.Order = i Then
          Print #FileNum, rFld.SaveDetails
        End If
      Next rFld
    Next i
    Status = "Specification saved successfully"
  End If
  
SaveReport_End:
  Close #FileNum
  'If Frm_RepWiz.chkSaveReadOnly And Len(FileName) <> 0 Then 'RK Removed functionality as instructed by CAD
  '  Call SetAttr(FileName, vbReadOnly)
  'End If
  Call xReturn("SaveReport")
  Exit Sub

SaveReport_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SaveReport", "ERR_SAVEREPORT", "Error saving report specification.")
  Resume SaveReport_End
  Resume
End Sub

Public Function GetReportFieldFromKey(DataSet As Object, ByVal KeyString As String) As Object
  Dim DataSetString As String, Name As String
  
  Call SplitKey(DataSetString, Name, KeyString)
  Set DataSet = GetDataSet(DataSetString)
  Set GetReportFieldFromKey = DataSet.cFields(KeyString)
End Function

Public Function AddControlCode(CtrlCode As String, CtrlCodeText As String, CtrlCodeDesc As String) As Boolean
  Dim i As Long
  
  If IsEmpty(m_CtrlCodes) Then
    i = 1
    ReDim m_CtrlCodes(1 To i)
    ReDim m_CtrlCodesText(1 To i)
    ReDim m_CtrlCodesDesc(1 To i)
  Else
    i = UBound(m_CtrlCodes) + 1
    ReDim Preserve m_CtrlCodes(1 To i)
    ReDim Preserve m_CtrlCodesText(1 To i)
    ReDim Preserve m_CtrlCodesDesc(1 To i)
  End If
  m_CtrlCodes(i) = CtrlCode
  m_CtrlCodesText(i) = CtrlCodeText
  m_CtrlCodesDesc(i) = CtrlCodeDesc
End Function

Public Function ReplaceCtrlChars(ByVal s As String) As String
  Dim i As Long
  
  s = ReplaceChar(s, "~", vbCrLf)
  s = ReplaceChar(s, "{USER}", GetNetUser)
  s = ReplaceChar(s, "{APPLICATION}", GetStatic("ApplicationName"))
  s = ReplaceChar(s, "{VERSION}", GetStatic("Version"))
  s = ReplaceChar(s, "{SPECFILENAME}", m_ReportFileName)
  If Not IsEmpty(m_CtrlCodes) Then
    For i = 1 To UBound(m_CtrlCodes)
      s = ReplaceChar(s, m_CtrlCodes(i), m_CtrlCodesText(i))
    Next i
  End If
  ReplaceCtrlChars = s
End Function

Public Property Set IRepProcess(ByVal NewValue As IReportPreProcess)
  Set m_IRepPP = NewValue
End Property

Public Property Get IRepProcess() As IReportPreProcess
  Set IRepProcess = m_IRepPP
End Property

Public Property Let DoReportOnCancel(ByVal NewValue As Boolean)
  m_DoReportOnCancel = NewValue
End Property

Public Property Get DoReportOnCancel() As Boolean
  DoReportOnCancel = m_DoReportOnCancel
End Property

Friend Property Get AR() As Boolean
  #If AbacusReporter Then
    AR = True
  #End If
End Property
'
'Property exposed for AbacusReporter to restrict to selected fields
#If AbacusReporter Then
  Friend Property Get ReportDataSets() As DataSetCollection
    Set ReportDataSets = m_ReportDataSets
  End Property

  Friend Property Get ReportDetails() As ReportDetails
    Set ReportDetails = mReportDetails
  End Property
  
  Friend Property Get FileGroupContainer() As ARFileGroupContainer
    If FileGroupContainer Is Nothing Then
      Set FileGroupContainer = m_FileGroupContainer
    End If
  End Property
  
  Friend Sub SetReportLineTimings(ByVal tIndividual As Double, Optional ByVal bReset As Boolean)
    If m_UseCancel Then
      If m_frmCancel Is Nothing Then Call Err.Raise(ERR_ERROR, "ProgressBar", "frmCancel is Nothing.")
      Call m_frmCancel.SetReportLineTimings(tIndividual, bReset)
    End If
  End Sub
  
  Friend Sub SetFileGroupTimings(ByVal tIndividual As Double, Optional ByVal bReset As Boolean)
    If m_UseCancel Then
      If m_frmCancel Is Nothing Then Call Err.Raise(ERR_ERROR, "ProgressBar", "frmCancel is Nothing.")
      Call m_frmCancel.SetFileGroupTimings(tIndividual, bReset)
    End If
  End Sub
  
  Friend Sub StartTimings()
    If m_UseCancel Then
      If m_frmCancel Is Nothing Then Call Err.Raise(ERR_ERROR, "ProgressBar", "frmCancel is Nothing.")
      Call m_frmCancel.StartTimings
    End If
  End Sub
  
  Friend Sub SetFileTimings(ByVal tFileOpen As Double, ByVal tFileExtract As Double, ByVal tFileClose As Double, ByVal tFileSave As Double, Optional ByVal bReset As Boolean)
    If m_UseCancel Then
      If m_frmCancel Is Nothing Then Call Err.Raise(ERR_ERROR, "ProgressBar", "frmCancel is Nothing.")
      Call m_frmCancel.SetFileTimings(tFileOpen, tFileExtract, tFileClose, tFileSave, bReset)
    End If
  End Sub

#End If

Private Sub Class_Terminate()
  #If AbacusReporter Then
    Set m_FileGroupContainer = Nothing
  #End If
End Sub

Public Property Get NotifyLineMeetsCriteria() As Boolean
  NotifyLineMeetsCriteria = m_NotifyLineMeetsCriteria
End Property

Public Property Let NotifyLineMeetsCriteria(ByVal NewValue As Boolean)
  m_NotifyLineMeetsCriteria = NewValue
End Property

Public Property Get AllowSeparateBranchFieldSelection() As Boolean
  AllowSeparateBranchFieldSelection = m_AllowSeparateBranchFieldSelection
End Property

Public Property Let AllowSeparateBranchFieldSelection(ByVal NewValue As Boolean)
  m_AllowSeparateBranchFieldSelection = NewValue
End Property

Friend Property Let Status(NewVal As String)
  'Replaces SPanel.Caption
  m_SPanel.Caption = NewVal
  If m_UseCancel Then
    If Not m_frmCancel Is Nothing Then
      m_frmCancel.lblStatus = NewVal
      DoEvents
    End If
  Else
    m_SPanel.Caption = NewVal
  End If
End Property
Public Property Let UseCancel(NewVal As Boolean)
  m_UseCancel = NewVal
End Property
Public Property Get UseCancel() As Boolean
  UseCancel = m_UseCancel
End Property
#If AbacusReporter Then
  Friend Sub SetAbacusTimings(ByVal tEvaluate As Double, ByVal tExecute As Double, ByVal tSaveCursor As Double)
    Call m_frmCancel.SetAbacusTimings(tEvaluate, tExecute, tSaveCursor)
  End Sub
#End If


