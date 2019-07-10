Attribute VB_Name = "LoadSave"
Option Explicit

Public Function LoadReportDetails(p As Parser, RepWiz As ReportWizard, FrmRepWiz As Frm_RepWiz, RepDataSets As DataSetCollection, RepDetails As ReportDetails, RepFields As Collection, ByVal FileName As String) As Boolean
  Dim rFld As ReportField, rFldTmp As ReportField, nod As node
  Dim DataSet As ReportDataSet
  Dim RepFieldsTmp As Collection
  
  On Error GoTo LoadReportDetails_Err
  
  Call SetCursor
  Set RepFieldsTmp = New Collection
  Call p.ParseReset
  Set p.ParseSettings.ParseParameter(1) = RepFieldsTmp
  Set p.ParseSettings.ParseParameter(2) = RepDetails
  
  Call p.ParseFile(FileName)
  If CheckLoadValid(RepDataSets, RepDetails, RepFieldsTmp) Then
    Call FrmRepWiz.ClearAllFields
    For Each rFldTmp In RepFieldsTmp
      Set rFld = RepWiz.GetReportFieldFromKey(DataSet, rFldTmp.KeyString)
      Set nod = FrmRepWiz.TrV_Fields.nodes(rFld.KeyString)
      Call FrmRepWiz.TrV_Fields_NodeClick(nod)
      Call rFldTmp.Copy(rFld, False)
      Set rFld.DataSet = DataSet
    Next rFldTmp
    #If AbacusReporter Then
      Call RepWiz.FileGroupContainer.Load(RepDetails.ARFileGroups)
      Call LoadAvailablePacks(RepWiz, FrmRepWiz)
      Call LoadAvailableFileGroups(RepWiz, FrmRepWiz, RepDetails)
      Call ApplyFileGroupSelection(RepWiz, FrmRepWiz, RepDetails)
      Call FrmRepWiz.SetButtons
    #End If
    Call ClearCollection(RepFieldsTmp)
    Call FrmRepWiz.FillReportDetails
    FrmRepWiz.lbl_SpecFilename = "Last specification file loaded: " & FileName
    LoadReportDetails = True
  Else
    Call DisplayMessage(FrmRepWiz, "The report specification in file '" & FileName & "' is not valid for the data available." & vbCrLf & "Load aborted.", "Load Report Details", "Ok", "")
  End If
  
LoadReportDetails_End:
  Set RepFieldsTmp = Nothing
  Call ClearCursor
  Exit Function
  
LoadReportDetails_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadReportDetails", "Load report specification", "Error loading report specification from file: " & FileName)
  Resume LoadReportDetails_End
  Resume
End Function

Private Function CheckLoadValid(RepDataSets As DataSetCollection, RepDetails As ReportDetails, RepFields As Collection) As Boolean
  Dim Fld As ReportField
  Dim i As Long, dSet As ReportDataSet
  
  On Error GoTo CheckLoadValidity_Err
  Call xSet("CheckLoadValid")
  CheckLoadValid = True
  For Each Fld In RepFields
    Fld.Selected = False
    If Not IsDataSetWithin(Fld.DataSetString, RepDataSets, dSet) Then Call Err.Raise(ERR_CHECKLOAD, "CheckLoadValid", "The data set " & Fld.DataSetString & " is not in the data collection.")
    If Not InCollection(dSet.cFields, Fld.KeyString) Then Call Err.Raise(ERR_CHECKLOAD, "CheckLoadValid", "The field " & Fld.Name & " is not in the data set " & Fld.DataSetString & ".")
    Fld.Selected = True
next_dataset:
  Next Fld
  For i = RepFields.Count To 1 Step -1
    Set Fld = RepFields(i)
    If Not Fld.Selected Then Call RepFields.Remove(i)
  Next i
  
  #If AbacusReporter Then
    If RepDetails.ARAbacusProductType <> g_AbacusReporter.AbacusDataProvider.ProductType Then
      Call Err.Raise(ERR_CHECKLOAD, "CheckLoadValid", "This report was prepared using the Abacus Product Type: " & GetAbacusProductDescription(RepDetails.ARAbacusProductType) & " which differs from the current application.")
    End If
  #End If
  

CheckLoadValidity_End:
  Call xReturn("CheckLoadValid") 'RK Amended 10/02/05
  Exit Function

CheckLoadValidity_Err:
  CheckLoadValid = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckLoadValid", "Check report", "Error checking validity of report to be loaded")
  If Err.Number = ERR_CHECKLOAD Then Resume next_dataset
  Resume CheckLoadValidity_End
End Function

Private Function IsDataSetWithin(ByVal FindDataSetString As String, WithinDataSets As Object, DataSet As ReportDataSet) As Boolean
  Dim dSet As ReportDataSet
  
  For Each dSet In WithinDataSets
    If StrComp(FindDataSetString, dSet.CurrentDataSetString, vbTextCompare) = 0 Then
      IsDataSetWithin = True
      Set DataSet = dSet
      Exit Function
    End If
    IsDataSetWithin = IsDataSetWithin(FindDataSetString, dSet.Children, DataSet)
    If IsDataSetWithin Then Exit Function
  Next dSet
End Function

