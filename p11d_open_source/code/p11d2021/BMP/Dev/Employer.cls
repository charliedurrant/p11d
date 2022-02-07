VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Employer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0

Implements IBenefitClass

Public Enum LoadEmployeeMode
  LEMDefault = 0
  LEMForceSave
  LEMForceReload
  LEMBeforeDelete
End Enum

Private Enum DBVERSION
  DBVER_OK
  DBVER_DIFFERENT_APP
  DBVER_TEMPLATE_DB
  DBVER_LOWER_FIX
  DBVER_HIGHER_FIX
End Enum

Public Enum BENEFIT_TABLES
  TBL_FIRST_ITEM = 1
  TBL_VANS = TBL_FIRST_ITEM
  TBL_COMPANY_CARS
  TBL_SHARED_VANS
  TBL_OTHER
  TBL_CDB_LINKS
'MP DB  TBL_HOMELOAN
  TBL_BENLOAN
  TBL_PHONES
  TBL_ASSETSATDISPOSAL
  TBL_SERVICESPROVIDED
  TBL_ACCOM
  TBL_RELOC
  TBL_CDRNEW
  TBL_EMPLOYEE_CARS
  TBL_TRANS
  TBL_ALLBENEFITS
  TBL_LASTITEM = TBL_ALLBENEFITS
End Enum

Public Enum Employer_ITEMS
  employer_Name_db = 1 'ITEM_DESC
  employer_Value '= ITEM_VALUE
  employer_MadeGood_Net '= ITEM_MADEGOOD_NET
  employer_Benefit '= ITEM_BENEFIT
  employer_MadeGood '= ITEM_MADEGOOD
  employer_ActualAmountMadeGood '= ITEM_ACTUALAMOUNTMADEGOOD
  employer_Benefit_Reportable '= ITEM_BENEFIT_REPORTABLE
  employer_UDM_BENEFIT_TITLE '= ITEM_UDM_BENEFIT_TITLE
  employer_BoxNumber '= ITEM_BOX_NUMBER
  employer_MadeGood_Is_TaxDeducted '=ITEM_MADEGOOD_IS_TAXDEDUCTED
  employer_Class1A_Adjustment '=ITEM_CLASS1A_ADJUSTMENT
  employer_NIC_Class1A_Able '=ITEM_NIC_CLASS1A_ABLE
  employer_NIC_Class1A_Value '= ITEM_NIC_CLASS1A_VALUE
  employer_NIC_Class1A_Benefit '= ITEM_NIC_CLASS1A_BENEFIT
  
  employer_PathAndFile

  employer_Payeref_db
  employer_Contact_db
  employer_Signatory_db
  employer_contactnumber_db
  employer_CT_db
  employer_PassWord_db
  
  employer_PathAndFileDEL
  employer_BroughtForward_db
  employer_NoOfSelectedEmployees
  employer_FileName
  employer_FixLevel_db
  employer_ApplicationName_db
  employer_LastApplicationName_db
  employer_ApplicationVersion_db
  employer_EmployeesWithSharedVan
  employer_SomeoneEditingCDBS_db
  employer_MeEditingCDBS_db
  employer_FileIsOpenBySomeOneElse
  employer_SubmitterRef_db
  employer_SubmitterName_db
  employer_PayeOnlineID_db
  employer_PayeOnlinePwd_db
  employer_PayeOnlineUser_db
  employer_PayeOnlineEmail_db
  employer_PayeOnlineSubID_db  'Submission ID guid from efiler
  employer_PayeOnlineSubmissionDate_db
  employer_PayeOnlineSubmissionType_db
  employer_AddressLine1_db
  employer_AddressLine2_db
  employer_AddressLine3_db
  employer_AddressLine4_db
  employer_AddressPostCode_db
  employer_EerDeclSect2Chk1_db  'Section 2 of P11b form - employer declaration
  employer_EerDeclSect2Chk2_db
  employer_EerDeclSect2Chk3_db
  
  employer_MagneticMediaSelected
  employer_IntranetSelected
  employer_EmployeeResponseDate_db
  
  employer_PAYEOnlineSelected
  employer_hasP46Cars
  
  employer_IRTaxOffice_db
  employer_IRAddressLine1_db
  employer_IRAddressLine2_db
  employer_IRAddressLine3_db
  employer_IRAddressLine4_db
  employer_IRPostcode_db
  employer_IRTelephoneNumber_db
  employer_IRYourReference_db
  employer_IRAccountsOfficeReference_db
  employer_AddClass1AAmounts_db
  employer_AddClass1ADescription_db
  employer_deductClass1AAmounts_db
  employer_deductClass1ADescription_db
  employer_GUID_db
  employer_NIC_AdjustmentAdd
  employer_NIC_AdjustmentDeduct

  employer_Year
  employer_TreatEmployeeCarsAsUnderApprovedFPCS_db
  employer_EmployeesCount
  employer_CheapTaxableLoansAllNon_db
  employer_GroupCode1Alias_db
  employer_GroupCode2Alias_db
  employer_GroupCode3Alias_db
  
  employer_Class1A_due 'EK added 1/04 TTP#51
  employer_LAST_ITEM = employer_Class1A_due

End Enum

Private m_BenItems(1 To employer_LAST_ITEM)
Private m_CDCs() As String
Private m_Parent As Object
Private m_rsEmpployeesRecordset As Recordset

Public employees As ObjectList
Public CDBEmployee As Employee

Private m_rsBenTables(1 To TBL_LASTITEM) As Recordset
Private m_dbfile As Database
Private m_NeedToCalculate As Boolean
Private m_dirty As Boolean

Public FPCSchemes As ObjectList
Public CurrentEmployee As Employee

Public SharedVans As SharedVans
Public CopyEmployeeIndex As Long
Public CopyBenIndex As Long
Public CopyBenClass As BEN_CLASS
Public CopyBenCut As Boolean
Public CDCs As ObjectList
Public Sub RepairAndCompact()
  Dim ben As IBenefitClass
  
  On Error GoTo RepairAndCompact_ERR
  
  Call xSet("RepairAndCompact")
  
  'If Not P11d32.CurrentEmployer Is Nothing Then Call Err.Raise(ERR_NOT_NOTHING, "RepairAndCompact", "The current emloyer is not nothing.")
  Call OnlyFromForm(F_Employers)
  Set ben = Me
  If Not RepairCompactDB(ben.value(employer_PathAndFile), MODE_REPAIR + MODE_COMPACT) Then Call Err.Raise(ERR_REPAIR_COMPACT, "RepairAndCompact", "Failed to repair and compact " & ben.value(employer_PathAndFile))
  
RepairAndCompact_END:
  Call xReturn("RepairAndCompact")
  Exit Sub
RepairAndCompact_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RepairAndCompact", "Repair And Compact", "Error repairing and compacting an employer.")
  Resume RepairAndCompact_END
End Sub

Private Sub Class_Terminate()
 'Debug.Print"Employer - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant
End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Employers"
End Property

Private Property Let IBenefitClass_LinkBen(RHS As Boolean)

End Property

Private Property Get IBenefitClass_LinkBen() As Boolean

End Property

Public Function EmployeeCheck(ByVal bChecked, ByVal lListItemTag As Long)
  Dim ben As IBenefitClass
  Dim benEmployer As IBenefitClass
  
On Error GoTo EmployeeCheck_ERR

  Call xSet("EmployeeCheck")
  
  Set ben = employees(lListItemTag)
  Set benEmployer = Me
  
  If Not ben Is Nothing Then
    ben.value(ee_Selected) = bChecked
    If bChecked Then
      benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) + 1
    Else
      benEmployer.value(employer_NoOfSelectedEmployees) = benEmployer.value(employer_NoOfSelectedEmployees) - 1
    End If
    Call SelectedPanel
  End If
    
EmployeeCheck_END:
  Call xReturn("EmployeeCheck")
  Exit Function
EmployeeCheck_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeCheck", "Employee Check", "Unable to select the employee.")
  Resume EmployeeCheck_END
  Resume
End Function

Public Sub SelectedPanel()
  Call SetPanel1(m_BenItems(employer_NoOfSelectedEmployees) & " employee(s) selected")
End Sub
  
Private Function IBenefitClass_CanBringForward() As Boolean

End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass
  Dim ey As IBenefitClass
  
  On Error GoTo Copy_END
  Call xSet("Copy")
  
  Set ey = New Employer
  If CopyBenData(ey, Me) Then
    Set ey.Parent = Parent
    Call CopyBenStandardEnd(ey)
    Set IBenefitClass_Copy = ey
  End If
  
Copy_END:
  Call xReturn("Copy")
  Exit Function
Copy_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Copy", "Copy", "Error copying an employer.")
  Resume Copy_ERR
End Function
Public Sub CopyReset()
  CopyBenClass = 0
  CopyBenIndex = 0
  CopyEmployeeIndex = 0
  CopyBenCut = False
End Sub
Friend Sub KillEmployeesBenefits()
  Dim ee As Employee
  Dim i As Long
   
  On Error GoTo KillEmployeesBenefits_ERR
  
  Call xSet("KillEmployeesBenefits")
   
  Call SetCursor(vbArrowHourglass)
  
  For i = 1 To employees.Count
    Set ee = employees(i)
    ee.KillBenefits
  Next
    
KillEmployeesBenefits_END:
  Call ClearCursor
  Call xReturn("KillEmployeesBenefits")
  Exit Sub
KillEmployeesBenefits_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "KillEmployeesBenefits", "Kill Employees Benefits", "Error killing all the employees benefits.")
  Resume KillEmployeesBenefits_END
End Sub
Public Property Get PayeRefValid() As Boolean
  
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  PayeRefValid = Validations.ValidatePAYE(ben.value(employer_Payeref_db))
  
End Property
Public Property Get TaxOfficeNumeral() As String
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  TaxOfficeNumeral = Validations.TaxOfficeNumeral(ben.value(employer_Payeref_db))
  
  
End Property
Public Property Get EmployerRefNo() As String
  Dim ben As IBenefitClass
  
  Set ben = Me
    
  EmployerRefNo = Validations.EmployerRefNo(ben.value(employer_Payeref_db))
  
End Property

Public Function ReloadCurrentEmployee() As Boolean
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass
  Dim frm As Form
  
  On Error GoTo ReloadCurrentEmployee_Err
  Call xSet("ReloadCurrentEmployee")
  
  If CurrentEmployee Is Nothing Then GoTo ReloadCurrentEmployee_End
  Set ben = CurrentEmployee
  
  Call SetPanel2("Loadinge benefits for " & CurrentEmployee.FullName)
  
  Set ibf = p11d32.GetBenefitForm
  
  If Not ibf Is Nothing Then
    Call ben.ReadDB
    CurrentEmployee.BenefitsLoaded = False
    ReloadCurrentEmployee = CurrentEmployee.LoadBenefits(TBL_ALLBENEFITS)
    ben.InvalidFields = 0
    ben.Dirty = False
    
    If ibf.benclass = BC_EMPLOYEE Then
      Call ibf.BenefitOn
    Else
      Call ibf.BenefitsToListView
      Call SelectBenefitByBenefitIndex(ibf, -1)
    End If
    MDIMain.ClearConfirmUndo
  End If
  
ReloadCurrentEmployee_End:
  Call SetPanel2("")
  Set ben = Nothing
  Set ibf = Nothing
  Call xReturn("ReloadCurrentEmployee")
  Exit Function
ReloadCurrentEmployee_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReloadCurrentEmployee", "Reload Current Employee", "Error .")
  Resume ReloadCurrentEmployee_End
  Resume
End Function

Public Sub CreatePathAndFile(ByVal NewValue As String)
  Dim sExt As String
  Dim sFIle As String
  Dim sDir As String
  Dim ben As IBenefitClass
  
  Call SplitPath(NewValue, sDir, sFIle, sExt)
  
  m_BenItems(employer_PathAndFile) = NewValue
  m_BenItems(employer_PathAndFileDEL) = FullPath(sDir) & sFIle & ".DEL"
  m_BenItems(employer_FileName) = sFIle & sExt
End Sub

Public Function WriteFPCS() As Boolean
  Dim rs As Recordset
  Dim l As Long, m As Long
  Dim Band As FPCSBand
  Dim FPCS As FPCS
  Dim ibf As IBenefitForm2
  
  On Error GoTo WriteFPCS_Err
  Call xSet("WriteFPCS")
  
  If db Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "WriteFPCS", "The database is nothing.")
  
  Call db.Execute(sql.Queries(DELETE_ER_FPCS))
  
  Set rs = Me.db.OpenRecordset(sql.Queries(SELECT_FPCS), dbOpenDynaset, dbFailOnError)
  
  For l = 2 To FPCSchemes.Count
    Set FPCS = FPCSchemes(l)
    If Not FPCS Is Nothing Then
      FPCS.Bands.Compact
      For m = 1 To FPCS.Bands.Count
        Set Band = FPCS.Bands(m)
        rs.AddNew
        rs.Fields("FPCS").value = FPCS.Name
        rs.Fields("CCBand").value = Band.Name
        rs.Fields("EngineAbove").value = Band.GreaterThanCC
        rs.Fields("MilesAbove").value = Band.GreaterThanMiles
        rs.Fields("RateMiles").value = Band.Rate
        rs.Fields("Approved").value = FPCS.Approved
        rs.Update
        'get the key for backward compatibility
        rs.MoveLast
        Band.Key = rs.Fields("CCBandKey").value
      Next
    End If
  Next
    
  'notify user of changes and report on cars effected?
  
  If CurrentForm Is F_EmployeeCar Then
    Set ibf = CurrentForm
    'we must redo the FPCS combo
    If Not ibf.lv.SelectedItem Is Nothing Then
      Call F_EmployeeCar.FPCSToCombo(True)
      ibf.benefit.Calculate
      Call ibf.BenefitToScreen(ibf.lv.SelectedItem.Tag)
      Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ibf.benefit)
    Else
      Call F_EmployeeCar.FPCSToCombo(False)
    End If
  End If
    
WriteFPCS_End:
  Set rs = Nothing
  Call xReturn("WriteFPCS")
  Exit Function

WriteFPCS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteFPCS", "Write FPCS", "Error writing the car scheme to the database.")
  Resume WriteFPCS_End
  Resume
End Function

Public Function ReadCDCs() As Long
  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo ReadCDCs_Err
  
  Call xSet("ReadCDCs")
  
  CDCs.RemoveAll
  Set rs = db.OpenRecordset(sql.Queries(SELECT_CDCS))
  If Not rs Is Nothing Then
    Do While Not rs.EOF
      Set ben = New CDC
      ben.ReadFromDB = True
      Set ben.Parent = Me
      ben.value(cdc_name_db) = rs!CDCItem
      ben.value(cdc_Key_db) = rs!CDCKey
      ben.value(cdc_IsUsed) = False
      Call CDCs.Add(ben)
      ReadCDCs = ReadCDCs + 1
      rs.MoveNext
    Loop
  End If
  
  
ReadCDCs_End:
  Call xReturn("ReadCDCs")
  Exit Function
ReadCDCs_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadCDCs", "Read CDCs", "Error reading the company defined categories from the database.")
  Resume ReadCDCs_End
  Resume
End Function
Public Function WriteCDCs() As Long
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo WriteCDCs_Err

  Call xSet("WriteCDCs")

  If p11d32.BringForward.Yes Then db.Execute (sql.Queries(DELETE_CDCS))
  
  Set rs = db.OpenRecordset(sql.Queries(SELECT_CDCS), dbOpenDynaset)
  If Not rs Is Nothing Then
    For i = 1 To CDCs.Count
      Set ben = CDCs(i)
      If Not ben Is Nothing Then
        If Not ben.ReadFromDB Or p11d32.BringForward.Yes Then
          rs.AddNew
          rs!CDCItem = ben.value(cdc_name_db)
          ben.value(cdc_Key_db) = rs!CDCKey
          ben.ReadFromDB = True
          rs.Update
        ElseIf ben.ReadFromDB And ben.Dirty Then
          rs.MoveFirst
          Do While Not rs.EOF
            If rs!CDCKey = ben.value(cdc_Key_db) Then
              rs.Edit
              rs!CDCItem = ben.value(cdc_name_db)
              Call rs.Update
              ben.Dirty = False
              Exit Do
            End If
            rs.MoveNext
          Loop
          
        End If
      End If
    Next
  End If

WriteCDCs_End:
  Call xReturn("WriteCDCs")
  Exit Function
WriteCDCs_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteCDCs", "Write CDCs", "Error writing the company defined categories to the database.")
  Resume WriteCDCs_End
  Resume
End Function

Public Function ReadFPCS() As Long
  Dim bFirstOne As Boolean
  Dim rs As Recordset
  Dim FPCS As FPCS
  Dim FPCSBand As FPCSBand
  Dim sSchemeName As String
  
  On Error GoTo ReadERFPCS_Err
  Call xSet("ReadERFPCS_Err")
  FPCSchemes.RemoveAll
  
  Call FPCSchemes.Add(p11d32.Rates.IRFPCS)
  
  Set rs = Me.db.OpenRecordset(sql.Queries(SELECT_FPCS), dbOpenForwardOnly, dbFailOnError)
  Do While Not rs.EOF
    If StrComp(sSchemeName, "" & rs.Fields("FPCS").value, vbTextCompare) <> 0 Then
      sSchemeName = "" & rs.Fields("FPCS").value
      Set FPCS = New FPCS
      FPCS.Name = sSchemeName
      FPCS.Approved = rs.Fields("Approved").value
      Call FPCSchemes.Add(FPCS)
    End If
    Set FPCSBand = New FPCSBand
    FPCSBand.Name = "" & rs.Fields("CCBand").value
    FPCSBand.GreaterThanCC = IIf(IsNull(rs.Fields("EngineAbove").value), 0, rs.Fields("EngineAbove").value)
    If FPCSBand.GreaterThanCC < 0 Then FPCSBand.GreaterThanCC = 0
    FPCSBand.GreaterThanMiles = IIf(IsNull(rs.Fields("MilesAbove").value), 0, rs.Fields("MilesAbove").value)
    If FPCSBand.GreaterThanMiles < 0 Then FPCSBand.GreaterThanMiles = 0
    FPCSBand.Rate = IIf(IsNull(rs.Fields("RateMiles").value), 0, rs.Fields("RateMiles").value)
    FPCSBand.Key = IIf(IsNull(rs.Fields("CCBandKey").value), -1, rs.Fields("CCBandKey").value)
    Call FPCS.Bands.Add(FPCSBand)
    rs.MoveNext
    ReadFPCS = ReadFPCS + 1
  Loop
  
ReadERFPCS_End:
  Call xReturn("ReadERFPCS_Err")
  Exit Function
ReadERFPCS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReadFPCS", "Read FPCS", "Error reading the car schemes from the database.")
  Resume ReadERFPCS_End
End Function

Public Function EditFPCS() As Boolean
  Dim benfrm2 As IBenefitForm2
  
  On Error GoTo EditFPCS_ERR
  
  Call xSet("EditFPCS")
  
  If Not MoveMenuUpdateEmployee Then GoTo EditFPCS_END
  
  Set benfrm2 = F_FPCS
  benfrm2.BenefitsToListView
  Call SelectBenefitByBenefitIndex(benfrm2, 1)
'  F_FPCS.Show vbModal
  Call p11d32.Help.ShowForm(F_FPCS, vbModal)
  EditFPCS = True
  
EditFPCS_END:
  Call xReturn("EditFPCS")
  Exit Function
EditFPCS_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EditFPCS", "Edit FPCS", "Error initialising FPCS schemes.")
  Resume EditFPCS_END
End Function

Public Function MoveMenuUpdateEmployee() As Boolean
  Call UpdateBenefitFromTags
  MoveMenuUpdateEmployee = SaveCurrentEmployee
End Function

Public Function EditCDC() As Boolean
  Dim ibf As IBenefitForm2
  Dim i As Long
  Dim ben As IBenefitClass
  Dim rs As Recordset
  Dim bFound As Boolean
  
  On Error GoTo EditCDC_ERR
  
  Call xSet("EditCDC")
   
  
  If Not MoveMenuUpdateEmployee Then GoTo EditCDC_END
  
  Set ibf = F_CDC
  Set rs = db.OpenRecordset(sql.Queries(SELECT_CDCS_EMPLOYEES))
  If rs Is Nothing Then Call Err.Raise(ERR_RS_IS_NOTHING, "EditCDC", "The rs is nothing when trying to determine if CDCs being used by employees.")
  
  For i = 1 To CDCs.Count
    Set ben = CDCs(i)
    If Not ben Is Nothing Then
      If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
      Do While Not rs.EOF
'MP DB
'          If rs!CDCKey = ben.value(cdc_Key) Then
        If rs!CDCKey = ben.value(cdc_Key_db) Then
          bFound = rs!Count > 0
        End If
        rs.MoveNext
      Loop
      ben.value(cdc_IsUsed) = bFound
      bFound = False
    End If
    
  Next
  
  Call BenScreenSwitchEnd(ibf)
'  F_CDC.Show vbModal
  Call p11d32.Help.ShowForm(F_CDC, vbModal)
  If CurrentForm Is F_Other Then
    Call BenScreenSwitchEnd(CurrentForm)
  End If
  
EditCDC_END:
  Call xReturn("EditCDC")
  Exit Function
EditCDC_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EditCDC", "Edit CDC", "Error editing CDC's.")
  Resume EditCDC_END
  Resume
End Function

Public Function EditSharedVans() As Boolean
  Dim PreviousCurrentForm As Form
  Dim ibf As IBenefitForm2
  Dim frm As Form
  Dim ben As IBenefitClass
    
  On Error GoTo EditSharedVans_ERR
  
  Call xSet("EditSharedVans")
  
  If Not MoveMenuUpdateEmployee Then GoTo EditSharedVans_END
  
  Set PreviousCurrentForm = CurrentForm
  Set CurrentForm = F_SharedVans
  Set ibf = CurrentForm
  ibf.benclass = BC_SHAREDVAN_G
  ibf.BenefitsToListView
  Call SelectBenefitByBenefitIndex(ibf, -1)
'  F_SharedVans.Show vbModal
  Call p11d32.Help.ShowForm(F_SharedVans, vbModal)
  Unload F_SharedVans
  Set F_SharedVans = Nothing
  Set ben = p11d32.CurrentEmployer.SharedVans
  Call ben.WriteDB
  Set CurrentForm = PreviousCurrentForm
  Set ibf = p11d32.GetBenefitForm
  If Not ibf Is Nothing Then
    If ibf.benclass = BC_NONSHAREDVANS_G Then
      Set frm = CurrentForm
      Call frm.UpdateTotalVansBenefit
    End If
  End If
  
  EditSharedVans = True
  
EditSharedVans_END:
  Set ben = Nothing
  Set PreviousCurrentForm = Nothing
  Set frm = Nothing
  Set ibf = Nothing
  Call xReturn("EditSharedVans")
  Exit Function
EditSharedVans_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EditSharedVans", "Edit Shared Vans", "Error editing shared vans.")
  Resume EditSharedVans_END
End Function
Public Property Get EmployeesRecordSet() As Recordset
  Set EmployeesRecordSet = m_rsEmpployeesRecordset
End Property
Public Property Let EmployeesRecordSet(rs As Recordset)
  Set m_rsEmpployeesRecordset = rs
End Property

Public Function FindEmployee(sPNum As String) As Employee
  Dim ee As Employee
  Dim i As Long
  On Error GoTo FindEmployee_ERR
  
  Set FindEmployee = Nothing
  
  For i = 1 To employees.Count
    Set ee = employees(i)
    If ee.PersonnelNumber = sPNum Then
      Set FindEmployee = ee
      Exit For
    End If
  Next i
    
FindEmployee_END:
  Exit Function
FindEmployee_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FindEmployee", "Find Employee", "Error - could not finding employee" & sPNum)
  Resume
End Function


Public Function LoadEmployees(Optional bToScreen As Boolean = True) As Long
  Dim emp As Employee
  Dim ben As IBenefitClass, benEmployer As IBenefitClass
  Dim i As Long, lMax As Long
  Dim P46Cars As ObjectList
  
  
  On Error GoTo LoadEmployees_Err
  Call xSet("Lo adEmployees")
  
  Call SetCursor(vbHourglass)
  Call employees.RemoveAll
  Set benEmployer = Me
  Call benEmployer.SetCalcDefaults
  'benEmployer.value(employer_hasP46Cars) = False
  
  If m_dbfile Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "LoadEmployees", "The database is nothing.")
  EmployeesRecordSet = m_dbfile.OpenRecordset(sql.Queries(SELECT_EMPLOYEES), dbOpenSnapshot)
  
  With m_rsEmpployeesRecordset
    If Not (.EOF And .BOF) Then
      If bToScreen Then
        Call PrgStartCaption(benEmployer.value(employer_EmployeesCount), "Loading Employees...", , ValueOfMax)
        Call PrgAlignment(TCSPROG.ALIGN_RIGHT)
      End If
      Do While Not .EOF
        DoEvents
        Set ben = New Employee
        Set ben.Parent = Me
        Call ben.ReadDB
        Set emp = ben
        If emp.InitOK Then
          If IsCBDEmployee(emp.PersonnelNumber) Then
            Set p11d32.CurrentEmployer.CurrentEmployee = emp
            Set p11d32.CurrentEmployer.CDBEmployee = emp
            Set p11d32.CurrentEmployer.CurrentEmployee = Nothing
            'setting flag if employee has p46 cars - used in PAYEONLINE
            
          Else
            Call employees.Add(ben)
            i = i + 1
            If bToScreen Then Call PrgStepCaption(emp.FullName)
          End If
        End If
        Set emp = Nothing
        .MoveNext
      Loop
    End If
  End With
      
  
LoadEmployees_End:
  Set m_rsEmpployeesRecordset = Nothing
  Call PrgStopCaption
  Call PrgAlignment(ALIGN_LEFT)
  LoadEmployees = i
  Call DBEngine.Idle(dbFreeLocks)
  Call ClearCursor
  Call xReturn("LoadEmployees")
  Exit Function
  
LoadEmployees_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadEmployees", "Load Employees", "Error loading the employees from the database. Employee database index = " & i)
  i = 0
  Resume LoadEmployees_End
  Resume
End Function


Public Sub DeleteSelectedEmployee()
  Dim ibf As IBenefitForm2
  
  On Error GoTo DeleteSelectedEmployee_Err
  
  Call xSet("DeleteSelectedEmployee")
  
  Set ibf = F_Employees
  With ibf.lv
    If Not .SelectedItem Is Nothing Then
      Call ibf.RemoveBenefit(.SelectedItem.Tag)
    End If
  End With
  
  
DeleteSelectedEmployee_End:
  Set ibf = Nothing
  Call DBEngine.Idle(dbFreeLocks + dbRefreshCache)
  Call xReturn("DeleteSelectedEmployee")
  Exit Sub
DeleteSelectedEmployee_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DeleteSelectedEmployee", "Delete Selected Employee", "Error deleting the selected employee.")
  Resume DeleteSelectedEmployee_End
End Sub
Public Function RunFixesYearCheck(ByVal lFixNumber As Long) As Boolean
  Dim sYear As String
  Dim ben As IBenefitClass
  
  On Error GoTo RunFixesYearCheck_ERR
  
  Call xSet("RunFixesYearCheck")
  
  Set ben = Me
  
  sYear = ben.value(employer_Year)
  
  Select Case lFixNumber
    Case Is <= LastFixLevel(0)
      RunFixesYearCheck = (sYear = "00")
    Case Is <= LastFixLevel(1)
      RunFixesYearCheck = (sYear = "01")
    Case Is <= LastFixLevel(2)
      RunFixesYearCheck = (sYear = "02")
    Case Is <= LastFixLevel(3)
      RunFixesYearCheck = (sYear = "03")
    Case Else
      RunFixesYearCheck = (sYear = "04")
  End Select
  
RunFixesYearCheck_END:
  Call xReturn("RunFixesYearCheck")
  Exit Function
RunFixesYearCheck_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RunFixesYearCheck", "Run Fixes Year Check", "Error deciding what fixes to run for each year.")
  Resume RunFixesYearCheck_END
End Function

Private Sub RunAllFixes(ByVal RunUpdateFixes As Boolean)
  Dim rs As Recordset, ibf As IBenefitForm2
  Dim i As Long
  Dim prg As Object
  Dim ben As IBenefitClass
  
  Set prg = MDIMain.sts.prg
  Set ben = Me
  
  If RunUpdateFixes Then
    For i = 0 To p11d32.TemplateFixlevel 'm_fixlevel
      prg.Min = 0
      prg.Max = p11d32.TemplateFixlevel
      prg.value = i
      prg.Indicator = ValueOfMax
      prg.Caption = "Running update fix"
      Call RunFixes(i, True)
    Next i
  End If
  
  If ben.value(employer_FixLevel_db) < p11d32.TemplateFixlevel Then
    For i = (ben.value(employer_FixLevel_db) + 1) To p11d32.TemplateFixlevel
      prg.Min = ben.value(employer_FixLevel_db) + 1
      prg.Max = p11d32.TemplateFixlevel
      prg.Indicator = ValueOfMax
      prg.Caption = "Run structural fix"
      prg.value = i
      Call RunFixes(i, False)
    Next i
    Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_ID), dbOpenDynaset, dbFailOnError)
    rs.Edit
    rs.Fields("Application").value = p11d32.Rates.ApplicationType
    rs.Fields("FixLevel").value = p11d32.TemplateFixlevel
    ben.value(employer_FixLevel_db) = p11d32.TemplateFixlevel
    rs.Update
    Set ibf = F_Employers
    Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ben)
  End If
End Sub
Private Function SetBenefitDirtyStatus(benefit As IBenefitClass) As Boolean
  On Error GoTo SetBenefitDirtyStatus_ERR
  
  xSet ("SetBenefitDirtyStatus")
  
  If Not benefit Is Nothing Then
    If benefit.InvalidFields Then
      Call SetPanel2(S_NOSAVE)
      Call MDIMain.SetUndo
      Call MDIMain.SetConfirmUndo
    ElseIf benefit.Dirty Then
      Call SetPanel2("")
      Call MDIMain.SetConfirmUndo
      benefit.Dirty = True
    End If
    SetBenefitDirtyStatus = True
  End If
  
SetBenefitDirtyStatus_END:
  xReturn ("SetBenefitDirtyStatus")
  Exit Function
SetBenefitDirtyStatus_ERR:
  SetBenefitDirtyStatus = False
  Call ErrorMessage(ERR_ERROR, Err, "SetBenefitDirtyStatus", "Set Benefit Status", "Error setting the status for a benefit")
  Resume SetBenefitDirtyStatus_END
End Function
'MP DB change - not called from anywhere
'Public Function ReadMileage(mileage As ObjectList, ByVal lCarKey As Long) As Long
'  Dim Miles As MileDetail
'  Dim rs As Recordset
'  Dim lMiles As Long
'
'  On Error GoTo ReadMileage_Err
'  Call xSet("ReadMileage")
'
'  Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_CARMILES, lCarKey), dbOpenForwardOnly)
'
'  Do While Not rs.EOF
'    Set Miles = New MileDetail
'    Miles.MileItem = "" & rs.Fields("Item").value
'    If IsNull(rs.Fields("MilesDate").value) Then
'      Miles.MileDate = UNDATED
'    Else
'      Miles.MileDate = rs.Fields("MilesDate").value
'    End If
'    Miles.MileAmount = IIf(IsNull(rs.Fields("Miles").value), 0, rs.Fields("Miles").value)
'    lMiles = lMiles + Miles.MileAmount
'    Call mileage.Add(Miles)
'    rs.MoveNext
'  Loop
'
'ReadMileage_End:
'  Set Miles = Nothing
'  ReadMileage = lMiles
'  Call xReturn("ReadMileage")
'  Exit Function
'ReadMileage_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "ReadMileage", "Read Mileage", "Error reading the car mileage.")
'  ReadMileage = 0
'  Resume ReadMileage_End
'End Function

'MP DB ToDo - chk ok to move to EmployeeCar class and change to Private as that's only place where its called from
'Public Function ReadEEMileage(mileage As ObjectList, ByVal lCarKey As Long) As Long
'  Dim Miles As MileDetail
'  Dim rs As Recordset
'  Dim lMiles As Long
'
'  On Error GoTo ReadEEMileage_Err
'  Call xSet("ReadEEMileage")
'
'  Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_EECARMILES, lCarKey), dbOpenForwardOnly)
'
'  Do While Not rs.EOF
'    Set Miles = New MileDetail
'    Miles.MileItem = "" & rs.Fields("MilesItem").value
'    If IsNull(rs.Fields("MilesDate").value) Then
'      Miles.MileDate = UNDATED
'    Else
'      Miles.MileDate = rs.Fields("MilesDate").value
'    End If
'    Miles.MileAmount = IIf(IsNull(rs.Fields("Miles").value), 0, rs.Fields("Miles").value)
'    lMiles = lMiles + Miles.MileAmount
'    Call mileage.Add(Miles)
'    rs.MoveNext
'  Loop
'
'ReadEEMileage_End:
'  Set Miles = Nothing
'  ReadEEMileage = lMiles
'  Call xReturn("ReadEEMileage")
'  Exit Function
'ReadEEMileage_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "ReadEEMileage", "Read Mileage", "Error reading the car mileage.")
'  ReadEEMileage = 0
'  Resume ReadEEMileage_End
'End Function
Public Sub SplitFix(ByVal QueryID As QUERY_NAMES)
  Call F_SplitNames.InitSplit(m_dbfile, QueryID)
  If Not F_SplitNames.Complete Then
    Set F_SplitNames = Nothing
    Call Err.Raise(ERR_NO_SPLIT_NAMES, "Run Fixes", "Unable to run the Split Name fix.")
  Else
    Set F_SplitNames = Nothing
  End If
End Sub
Public Sub RunFixes(ByVal FixToRun As Long, ByVal UpDateOnly As Boolean)
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim s As String
  Dim sqlimp As SQLQUERIES_IMP
  If Not RunFixesYearCheck(FixToRun) Then GoTo RunFixes_END
  
  Call xSet("RunFixes")
  Select Case FixToRun
    Case 0 To 66
    
    Case 68
      Call m_dbfile.Execute(sql.Queries(FIX_O_OTHER_EXPENSE))
      Call m_dbfile.Execute(sql.Queries(FIX_SHARES))
      Call m_dbfile.Execute(sql.Queries(FIX_MEDICAL))
      Call m_dbfile.Execute(sql.Queries(FIX_CHAUFFEUR))
      Call m_dbfile.Execute(sql.Queries(FIX_GENERAL))
      Call m_dbfile.Execute(sql.Queries(FIX_ENTERTAINMENT))
      Call m_dbfile.Execute(sql.Queries(FIX_TRAVEL))
      Call m_dbfile.Execute(sql.Queries(FIX_TAX_PAID))
      Call m_dbfile.Execute(sql.Queries(FIX_EDUCATION))
      Call m_dbfile.Execute(sql.Queries(FIX_NURSERY))
      Call m_dbfile.Execute(sql.Queries(FIX_SUBSCRIPTIONS))
      Call m_dbfile.Execute(sql.Queries(FIX_O_OTHER))
      Call m_dbfile.Execute(sql.Queries(FIX_SERVICES_PROVIDED))
      Call m_dbfile.Execute(sql.Queries(FIX_ASSETS_DISPOSAL))
  
    Case 69
      Call m_dbfile.Execute(sql.Queries(FIX_FREQUENCY_OF_PAYMENTS))
    Case 70
      'force sync db
    Case 71
    Case 72
      'force sync db
    Case 73
    Case 74
      'force sync db
    Case 75
      Call m_dbfile.Execute(sql.Queries(FIX_MOBILE_PHONES))
    Case 76
      'force sync db - KM
    Case 87
      'remove validation rule on FuelType as if it ever changes mid year befire final release they will not be able to save the new value
      If (TablePresent(m_dbfile.TableDefs, "T_BenCar")) Then
        Call RemoveValidationRule(m_dbfile, "T_BenCar", "P46FuelType")
      End If
    Case 88 'last for 2003
    Case 89 ' first for 2004
    Case 90
      Call m_dbfile.Execute(sql.Queries(FIX_SHARES_REMOVE))
      'move n to M, and o to M
    Case 91
    Case 92
    Case 93
      Call m_dbfile.Execute(sql.Queries(FIX_N_2004))
      Call m_dbfile.Execute(sql.Queries(FIX_N_OTHER_2004))
      Call m_dbfile.Execute(sql.Queries(FIX_M_2004))
    Case 94 'this was not released with march 2005, so can add to this
      Set sqlimp = New SQLQUERIES_IMP
      m_dbfile.Execute (sqlimp.Queries(UPDATE_COMPANY_CAR_FUEL_AVAILABLE_TO))
  End Select
    
  Set rs = Nothing
  
RunFixes_END:
  Call xReturn("RunFixes")
  
End Sub
Private Sub RemoveValidationRule(db As Database, TABLE As String, Field As String)
  Dim td As TableDef
  Dim f As Field
  
  Set td = db.TableDefs(TABLE)
  Set f = td.Fields(Field)
  f.ValidationRule = ""
  
End Sub
Private Function IsTemplate() As Boolean
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  IsTemplate = True
  If StrComp(ben.value(employer_FileName), "Blank" & p11d32.Rates.FileExtensionCurrent, vbTextCompare) = 0 Then Exit Function
  If StrComp(Trim(ben.value(employer_ApplicationName_db)), p11d32.Rates.TemplateApp, vbTextCompare) = 0 Then Exit Function
  IsTemplate = False
  
End Function

Public Property Set db(NewValue As Database)
  Set m_dbfile = NewValue
End Property

Public Property Get db() As Database
  Set db = m_dbfile
End Property

Private Function CheckVer() As DBVERSION
  Dim db As Database
  Dim ben As IBenefitClass
  Dim rs As Recordset
  
  On Error GoTo checkver_err
  Call xSet("Checkver")
  CheckVer = DBVER_DIFFERENT_APP
  Set ben = Me
  If StrComp(ben.value(employer_ApplicationName_db), p11d32.Rates.ApplicationType, vbTextCompare) <> 0 Then
    If StrComp(ben.value(employer_ApplicationName_db), p11d32.Rates.AppOldType, vbTextCompare) = 0 Then
      CheckVer = DBVER_LOWER_FIX
    ElseIf StrComp(ben.value(employer_ApplicationName_db), p11d32.Rates.TemplateApp, vbTextCompare) = 0 Then
      CheckVer = DBVER_TEMPLATE_DB
    End If
  Else
    Select Case ben.value(employer_FixLevel_db)
      Case Is = p11d32.TemplateFixlevel
        CheckVer = DBVER_OK
      Case Is < p11d32.TemplateFixlevel
        CheckVer = DBVER_LOWER_FIX
      Case Else
        CheckVer = DBVER_HIGHER_FIX
    End Select
  End If
  
checkver_end:
  Call xReturn("Checkver")
  Exit Function
checkver_err:
  Resume checkver_end
End Function

Private Sub SyncDB()
  Dim dbName  As String, sFileName As String
  Dim TemplateDb As Database, qd As QueryDef
  Dim i As Long
  Dim ben As IBenefitClass
  
  On Error GoTo SyncDB_err
  Call xSet("SyncDB")
  
  Call SetCursor
  
  Set ben = Me

  sFileName = ben.value(employer_PathAndFile)
  Set TemplateDb = InitDB(p11d32.P11DWS, p11d32.TemplateFile(Me), "P11D template")
  If TemplateDb Is Nothing Then Call Err.Raise(ERR_SYNCDB, "SyncDB", "Unable to open template database")
  If (Not m_dbfile Is Nothing) Then
    Call m_dbfile.Close
  End If
  Set m_dbfile = InitDB(p11d32.P11DWS, sFileName, sFileName, , True)     'open exclusive
  dbName = m_dbfile.Name
  If m_dbfile Is Nothing Then Call Err.Raise(ERR_DB_IS_NOTHING, "SyncDB", "Unable to open " & dbName)
  
  Call SetPanel2("Updating the database structure of " & sFileName)
  
  If Not UpdateDatabaseEx(p11d32.P11DWS, TemplateDb, m_dbfile, DB_ADD + DB_SYNC, PROP_ALLOWZEROLENGTH + PROP_DEFAULTVALUE + PROP_REQUIRED, True) Then Call Err.Raise(ERR_SYNCDB, "SyncDB", "Unable to update Employer database")
  Call DBEngine.Idle(dbFreeLocks + dbRefreshCache)
  m_dbfile.Close
  
  Set m_dbfile = Nothing
  Set m_dbfile = InitDB(p11d32.P11DWS, dbName, dbName)
  
SyncDB_end:
  Call SetPanel2("")
  Call ClearCursor
  Call xReturn("SyncDB")
  Exit Sub
SyncDB_err:
  Call ClearCursor
  Call Err.Raise(ERR_SYNCDB, "Employer SyncDB", "Error synchronisning Employer database " & sFileName & " with template database " & p11d32.TemplateFile(Nothing) & vbCr & Err.Description)
  Resume
End Sub
Public Property Get rsBenTables(i As BENEFIT_TABLES) As Recordset
  Set rsBenTables = m_rsBenTables(i)
End Property
Public Function EmployeeScreen() As Boolean
  Dim ibf As IBenefitForm2
  Dim lEmployeeIndex As Long
  Dim ee As Employee
  Dim li As ListItem
  
  On Error GoTo EmployeeScreen_Err
  Call xSet("EmployeeScreen")
  
  If Not (CurrentForm Is F_Employees) Then
    'create an empoyee so setbenitems information is called
    'this is required so the group code aliass can be updated on the employee
    Set ee = New Employee
    Set ee = Nothing
    Call UpdateBenefitFromTags
    If p11d32.CurrentEmployer.CurrentEmployee Is p11d32.CurrentEmployer.CDBEmployee Then
      If F_Employees.lb.listitems.Count > 0 Then
        Set ee = employees(GetEmployeeIndexFromSelectedEmployee)
      End If
    End If
    lEmployeeIndex = -1
    If Not LoadEmployeeEx(ee) Then GoTo EmployeeScreen_End
    Call ShowMaximized(F_Employees, Nothing, D_EMPLOYEES)
    Set ibf = F_Employees
    
    If CurrentForm Is F_Employers Then
      ibf.benclass = BC_EMPLOYEE
      If ibf.BenefitsToListView Then lEmployeeIndex = 1
    Else
      If Not ibf.lv.SelectedItem Is Nothing Then lEmployeeIndex = ibf.lv.SelectedItem.Index
    End If
    Set CurrentForm = F_Employees
    
    Call DisplayEx(D_EMPLOYEES)  ' dup. above check dependancies
    If lEmployeeIndex > -1 Then
      Call SelectBenefitByListItem(ibf, ibf.lv.listitems(lEmployeeIndex))
    Else
      Call SelectBenefitByListItem(ibf, Nothing)
    End If
    Call p11d32.CurrentEmployer.SelectedPanel
  End If
  Call F_Employees.UpdateGroupCodeLables
  Call UpdateMDICaption(p11d32.CurrentEmployer)
  
  
EmployeeScreen_End:
  Call xReturn("EmployeeScreen")
  Exit Function

EmployeeScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeScreen", "Employee Screen", "Error setting up the employee screen.")
  Resume EmployeeScreen_End
  Resume
End Function

Public Property Set rsBenTables(i As BENEFIT_TABLES, Newrs As Recordset)
  Set m_rsBenTables(i) = Newrs
End Property

Private Property Let IBenefitClass_BenefitClass(NewValue As Long)
  ECASE ("Ben class let invalid for employer.")
End Property

Private Property Get IBenefitClass_BenefitClass() As Long
  IBenefitClass_BenefitClass = BC_EMPLOYER
End Property

Private Function IBenefitClass_Calculate() As Variant
  Dim i As Long
  Dim ee As IBenefitClass
  Dim ey As IBenefitClass

  Call xSet("Employer_Calculate")
  
  On Error GoTo Employer_Calculate_ERR
  
  Set ey = Me
  
  Call ey.SetCalcDefaults
  IBenefitClass_Calculate = 0
  
  For i = 1 To employees.Count
    Set ee = employees(i)
    If Not ee Is Nothing Then
       IBenefitClass_Calculate = IBenefitClass_Calculate + ee.Calculate
       ey.value(ITEM_BENEFIT) = ey.value(ITEM_BENEFIT) + ee.value(ITEM_BENEFIT)
       ey.value(ITEM_NIC_CLASS1A_BENEFIT) = ey.value(ITEM_NIC_CLASS1A_BENEFIT) + ee.value(ITEM_NIC_CLASS1A_BENEFIT)
       ey.value(ITEM_CLASS1A_ADJUSTMENT) = ey.value(ITEM_CLASS1A_ADJUSTMENT) + ee.value(ITEM_CLASS1A_ADJUSTMENT)
       ey.value(employer_NIC_AdjustmentAdd) = ey.value(ITEM_CLASS1A_ADJUSTMENT) + ey.value(employer_AddClass1AAmounts_db)
       ey.value(employer_NIC_AdjustmentDeduct) = ey.value(employer_deductClass1AAmounts_db)
    End If
  Next
  
  ' EK added to calculate ey.value(employer_Class1A_due) 1/04 TTP#51
  If (ey.value(ITEM_NIC_CLASS1A_BENEFIT) + ey.value(employer_NIC_AdjustmentAdd) - ey.value(employer_NIC_AdjustmentDeduct)) > 0 Then
    ey.value(employer_Class1A_due) = p11d32.Rates.value(carNICRate) * (ey.value(ITEM_NIC_CLASS1A_BENEFIT) + ey.value(employer_NIC_AdjustmentAdd) - ey.value(employer_NIC_AdjustmentDeduct))
  End If
    
Employer_Calculate_END:
  Call xReturn("Employer_Calculate")
  Exit Function
Employer_Calculate_ERR:
  IBenefitClass_Calculate = S_ERROR
  Resume Employer_Calculate_END
End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal RHS As Boolean)
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  
End Property

Private Function IBenefitClass_DeleteDB() As Boolean
  ECASE ("Delete DB invalid for employer.")
End Function

Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  
End Property

Private Property Get IBenefitClass_HasBookMark() As Boolean
  ECASE ("BookMark invalid for employer.")
End Property

Private Sub Class_Initialize()
  Set employees = New ObjectList
  employees.Increment = 500
  Set FPCSchemes = New ObjectList
  Set SharedVans = New SharedVans
  Set SharedVans.IBenefitClass_Parent = Me
  Set CDCs = New ObjectList
  Call IBenefitClass_SetBenItemsInformation
End Sub
Public Function ValidatePersonnelNumber(vNewNumber As Variant) As Boolean
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo ValidatePersonnelNumber_Err
  Call xSet("ValidatePersonnelNumber")
  
  If IsCBDEmployee(CStr(vNewNumber)) Then
    Call ErrorMessage(ERR_INFO, Err, "Employee_AddBenefit", "Employee AddBenefit", "This employee personnel number cannot be used because the start contains " & S_CDB_EMPLOYEE_NUMBER_PREFIX & ".")
    GoTo ValidatePersonnelNumber_End
  End If
  
  For i = 1 To employees.Count
    Set ben = employees(i)
    If Not ben Is Nothing Then
      If vNewNumber = ben.value(ee_PersonnelNumber_db) Then
        Call ErrorMessage(ERR_INFO, Err, "Employee_AddBenefit", "Employee AddBenefit", "This pesonnel reference has already been used." & vbCrLf & "Please use a unique personnel reference for each employee.")
        GoTo ValidatePersonnelNumber_End
      End If
    End If
  Next
  
  ValidatePersonnelNumber = True
  
ValidatePersonnelNumber_End:
  Set ben = Nothing
  Call xReturn("ValidatePersonnelNumber")
  Exit Function

ValidatePersonnelNumber_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ValidatePersonnelNumber", "Validate Personnel Number", "Error validating the personnel number.")
  Resume ValidatePersonnelNumber_End
  Resume
End Function


Public Function LoadEmployeeEx(NewEE As Employee, Optional ByVal ForceLoad As Boolean = False) As Boolean
  Dim ben As IBenefitClass
   
  On Error GoTo LoadEmployeeEx_Err
  Call xSet("LoadEmployeeEx")
  
  If Not SaveCurrentEmployee() Then
    If Not ForceLoad Then GoTo LoadEmployeeEx_End
  End If

  If CurrentEmployee Is CDBEmployee Then
    Call p11d32.CurrentEmployer.EditingCompanyDefinedBenefits(ECDB_SETVALUE, False)
  End If

  If Not NewEE Is Nothing Then
    Set CurrentEmployee = NewEE
  
  Else
    Set CurrentEmployee = Nothing
  End If
  
  LoadEmployeeEx = True
  
LoadEmployeeEx_End:
  Set ben = Nothing
  Call xReturn("LoadEmployeeEx")
  Exit Function

LoadEmployeeEx_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadEmployeeEx", "Load EmployeeEx", "Error loading the employee.")
  Resume LoadEmployeeEx_End
End Function


Public Function SaveCurrentEmployee() As Boolean
  Dim l As Long
  Dim ben As IBenefitClass, benEmployee As IBenefitClass
  Dim bAnythingDirty As Boolean, bAnybenefitInvalidFields As Boolean, bAnyInvalidFields As Boolean
  Dim bAllOK As Boolean
  Dim ee As Employee
  Dim ibf As IBenefitForm2
  Dim sErrorTitle As String, s As String
  Dim sPreFix As String, bCBDEmployee As Boolean
  
  On Error GoTo SaveCurrentEmployee_Err
  Call xSet("SaveCurrentEmployee")

  If Not CurrentEmployee Is Nothing Then
    Set benEmployee = CurrentEmployee
    If benEmployee.Dirty Or CurrentEmployee.IterateBenefits(AnyDirty) Then
      Set ben = CurrentEmployee
      Call SetPanel2("Attempting to save " & CurrentEmployee.FullName)
      bAnybenefitInvalidFields = CurrentEmployee.IterateBenefits(AnyInvalid)
      bAnyInvalidFields = bAnybenefitInvalidFields Or CurrentEmployee.InvalidFields
      If CurrentEmployee.InvalidFields And bAnybenefitInvalidFields Then
        s = "The employee " & CurrentEmployee.FullName & " has invalid employee information and benefit information." & vbCrLf & vbCrLf
      ElseIf CurrentEmployee.InvalidFields Then
        s = "The employee " & CurrentEmployee.FullName & " has invalid employee information." & vbCrLf & vbCrLf
      ElseIf bAnybenefitInvalidFields Then
        If ben Is p11d32.CurrentEmployer.CDBEmployee Then
          s = "There as company defined benefits with invalid information." & vbCrLf & vbCrLf
        Else
          s = "The employee " & CurrentEmployee.FullName & " has invalid benefit information." & vbCrLf & vbCrLf
        End If
      Else
        bAllOK = True
      End If
      sErrorTitle = "Warning "
      
      bAnyInvalidFields = False
      bAnybenefitInvalidFields = False
      CurrentEmployee.InvalidFields = False     'RH - reset invalid field flag
      
      If bAllOK Then
        Call ben.WriteDB
        Call MDIMain.ClearConfirmUndo
      Else
        s = s & "To save any valid information please press Continue and then amend the information for the invalid benefits. " & vbCrLf & vbCrLf
        s = s & "Selecting Discard will save the valid information but will throw away any invalid amendments to existing benefits. Any new items which contain invalid information will also be lost." & vbCrLf & vbCrLf
        s = s & "Cancel will not save any data. "
        l = MultiDialog(sErrorTitle & " is invalid.", s, "&Continue", "&Discard", "&Cancel")
        Select Case l
          Case 1
            ben.WriteDB 'writes any information it can
            
            GoTo SaveCurrentEmployee_End 'do nothing
          Case 2
            Call ReloadCurrentEmployee 'writes what it can and reloads
            'clears it own confirm undo
          Case 3
            GoTo SaveCurrentEmployee_End 'do nothing
        End Select
      End If
    End If
  End If
  SaveCurrentEmployee = True

SaveCurrentEmployee_End:
  Call SetPanel2("")
  Call xReturn("SaveCurrentEmployee")
  Exit Function

SaveCurrentEmployee_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SaveCurrentEmployee", "Save Current Employee", "Undefined error.")
  Resume SaveCurrentEmployee_End
  Resume
End Function

Public Function EditCompanyDefinedBenefits() As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo EditCompanyDefinedBenefits_Err
  
  Call xSet("EditCompanyDefinedBenefits")
  
  
  If Not MoveMenuUpdateEmployee Then GoTo EditCompanyDefinedBenefits_End
  
  Set ben = Me
  
  ben.value(employer_SomeoneEditingCDBS_db) = EditingCompanyDefinedBenefits(ECDB_Value)
  
  If ben.value(employer_SomeoneEditingCDBS_db) And Not ben.value(employer_MeEditingCDBS_db) Then
    Call Err.Raise(ERR_NO_EDIT_CDB, "EditCompanyDefinedBenefits", "Unable to edit the company defined benefits because another user is editing them.")
  Else
    
    Call EditingCompanyDefinedBenefits(ECDB_SETVALUE, True)
    Call BenScreenSwitch(BC_CDB, p11d32.CurrentEmployer.CDBEmployee)
  End If
  
EditCompanyDefinedBenefits_End:
  Call xReturn("EditCompanyDefinedBenefits")
  Exit Function
EditCompanyDefinedBenefits_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EditCompanyDefinedBenefits", "Edit Company Defined Benefits", "Error edting the company defined benefits.")
  Resume EditCompanyDefinedBenefits_End
  Resume
End Function


Public Function GetEmployees(OL As ObjectList) As Boolean
  Dim i As Long
  Dim li As ListItem
  
  On Error GoTo GetEmployees_Err
  Call xSet("GetEmployees")

  Set OL = New ObjectList
  
  With F_Employees.lb
    For i = 1 To .listitems.Count
      Set li = .listitems(i)
      If p11d32.EmployeeSelectedForPrint(employees(li.Tag), Me) Then
        Call OL.Add(employees(li.Tag))
      End If
    Next
  End With
  
  GetEmployees = OL.Count
  
GetEmployees_End:
  Call xReturn("GetEmployees")
  Exit Function

GetEmployees_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetEmployees", "Get Employees", "Error obtaining object list of selected employees.")
  Resume GetEmployees_End
  Resume
End Function
Public Property Get NoOfEmployees() As Long
  Dim l As Long
  
  If Not employees Is Nothing Then
    l = employees.CountValid
    NoOfEmployees = l
  End If
  
End Property

Public Function Validate(ByVal sPathAndFile As String) As Boolean
  Validate = ValidateEx(sPathAndFile, True, False, False)
End Function

Public Function ValidateEx(ByVal sPathAndFile As String, ByVal bCloseDB As Boolean, ByVal bOpenExclusive As Boolean, ByVal bShowErrors As Boolean) As Boolean

  Dim ben As IBenefitClass
  
  On Error GoTo Validate_err
  Call xSet("Validate")
  Set ben = Me
  
  ben.value(employer_PathAndFile) = sPathAndFile
  
  If FileExists(ben.value(employer_PathAndFile)) Then
    ben.value(employer_FileIsOpenBySomeOneElse) = IsFileOpen(ben.value(employer_PathAndFile))
    Set m_dbfile = InitDB(p11d32.P11DWS, ben.value(employer_PathAndFile), "P11D Employer Database", , bOpenExclusive, bShowErrors)
  End If
  
  If m_dbfile Is Nothing Then
    If Not bShowErrors Then
      Call Err.Raise(ERR_EMPLOYER_DB, "Initialize", "Unable to create the employer file " & ben.value(employer_FileName))
    Else
      GoTo Validate_end
    End If
  End If
  
  Call ben.ReadDB
  If IsTemplate() Then GoTo Validate_end
  
  ValidateEx = True
  
Validate_end:
  If bCloseDB Then
    Set m_dbfile = Nothing
    ben.value(employer_FileIsOpenBySomeOneElse) = False
  End If
  Call xReturn("Validate")
  Exit Function

Validate_err:
  ValidateEx = False
  If bShowErrors Then
    Call ErrorMessage(ERR_ERROR, Err, "Initialize", "Read Employer database", "Error reading employer database " & ben.value(employer_FileName))
  End If
  Resume Validate_end
  Resume
End Function

Public Function CreateEmployer(ByVal sPathAndFile As String) As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo CreateEmployer_Err
  Call xSet("CreateEmployer")
  
  
  If FileExists(sPathAndFile) Then Err.Raise 58, "CreateEmployer", "File already exists in this directory."
  If Not FileCopyEx(p11d32.TemplateFile(Nothing), sPathAndFile) Then Err.Raise 70, "CreateEmployer", "Unable to copy master database " & p11d32.TemplateFile(Nothing) & "."
  Set m_dbfile = InitDB(p11d32.P11DWS, sPathAndFile, "P11D Employer Database", sql.Queries(SELECT_ID), , False)
  If m_dbfile Is Nothing Then Err.Raise ERR_CREATEEMPLOYER, "CreateEmployer", "Unable to open the employer file " & sPathAndFile
  Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_ID), dbOpenDynaset, dbFailOnError)
  rs.Edit ' new file
   rs.Fields("Application").value = p11d32.Rates.ApplicationType
   rs.Fields("AppVersion").value = GetStatic("Version")
  rs.Update
  Set ben = Me
  ben.value(employer_PathAndFile) = sPathAndFile
  CreateEmployer = True
  
CreateEmployer_End:
  Set rs = Nothing
  Set m_dbfile = Nothing
  Call RegisterDB(m_dbfile)
  Call xReturn("CreateEmployer")
  Exit Function

CreateEmployer_Err:
  CreateEmployer = False
  Call ErrorMessage(ERR_ERROR, Err, "CreateEmployer", "Creating new employer file", "Unable to create the employer (" & sPathAndFile & ").")
  Resume CreateEmployer_End
  Resume
End Function
Private Function RunRepairAndCompanctOnLoad(ByVal benEmployer As IBenefitClass)
  Dim lFixLevel As Long
  Dim b As Boolean
  
  On Error GoTo RunRepairAndCompanctOnLoad_ERR
  
  If benEmployer Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "RunRepairAndCompanctOnLoad", "The employer is nothing.")
  lFixLevel = benEmployer.value(employer_FixLevel_db)
  
  Select Case p11d32.AppYear
    Case 2000
      b = lFixLevel < 75
    Case 2001
      b = lFixLevel < 76 'km
    Case 2002
      b = lFixLevel < 76 'km
    Case 2003            'case 2003 added by IK 23/05/2003
      b = lFixLevel < 77
  End Select
  
  If b Then
    If Not FileExists(benEmployer.value(employer_PathAndFile)) Then Call Err.Raise(ERR_FILE_EXISTS, "RunRepairAndCompanctOnLoad", "The file " & benEmployer.value(employer_PathAndFile) & " does not exist.")
    If FileLen(benEmployer.value(employer_PathAndFile)) < 5000000 Then Call RepairAndCompact
  End If
  
  
RunRepairAndCompanctOnLoad_END:
  Exit Function
RunRepairAndCompanctOnLoad_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RunRepairAndCompanctOnLoad", "Run Repair And Companct OnLoad", "Error running repair and compact on load of employer.")
  Resume RunRepairAndCompanctOnLoad_ERR
  
End Function

Public Function OpenEmployer(ByVal bOpenExclusive As Boolean) As Boolean
  Dim RunUpdateFixes As Boolean
  Dim ben As IBenefitClass
  Dim s As String
  
  On Error GoTo OpenEmployer_Err
  
  OpenEmployer = True
  Set ben = Me
  If m_dbfile Is Nothing Then
    Call RunRepairAndCompanctOnLoad(ben)
    If Not ValidateEx(ben.value(employer_PathAndFile), False, bOpenExclusive, True) Then Call Err.Raise(ERR_OPENEMPLOYER, "OpenEmployer", "Unable to open Employer database as the database is not valid")
  End If
  
  If Not m_dbfile Is Nothing Then
    If m_dbfile Is Nothing Then Call Err.Raise(ERR_OPENEMPLOYER, "OpenEmployer", "Unable to open Employer database")
    
    
    RunUpdateFixes = StrComp(ben.value(employer_LastApplicationName_db), p11d32.Rates.ApplicationType)
    If EmployerYear <> p11d32.AppYearShort Then
      If p11d32.BringForward.Yes Then
        GoTo OpenEmployer_End
      Else
        Call Err.Raise(ERR_EMPLOYER_INVALID, "OpenEmployer", "The employer year, " & EmployerYear & " is not equal to the app year ," & p11d32.AppYear & " cannot syncdb and run fixes.")
      End If
    End If
    
    Select Case CheckVer()
      Case DBVER_LOWER_FIX
        Call SyncDB
        Call RunAllFixes(RunUpdateFixes)
      Case DBVER_HIGHER_FIX
        Call ECASE(" Database " & ben.value(employer_FileName) & " has a higher fix level than the template Database")
      Case DBVER_TEMPLATE_DB
        GoTo OpenEmployer_End
      Case DBVER_DIFFERENT_APP
        GoTo OpenEmployer_End
      Case DBVER_OK
        If RunUpdateFixes Then Call RunAllFixes(RunUpdateFixes)
      End Select
        
      m_dbfile.Execute (sql.Queries(UPDATE_LAST_APPLICATION))
      ben.value(employer_LastApplicationName_db) = p11d32.Rates.ApplicationType
      If (Not ben.value(employer_FileIsOpenBySomeOneElse)) Then
        Call EditingCompanyDefinedBenefits(ECDB_SETVALUE, False, True)
      End If
      
      Call UpdateGuid
      
    End If

OpenEmployer_End:
  Call PrgStop
  Call RegisterDB(m_dbfile)
  Exit Function
OpenEmployer_Err:
  OpenEmployer = False
  Call Err.Raise(ERR_OPENEMPLOYER, "Open Employer", "Cannot open employer " & ben.value(employer_Name_db) & " (" & ben.value(employer_PathAndFile) & ") ")
  Resume OpenEmployer_End
  Resume
End Function
Private Sub UpdateGuid()
  Dim ben As IBenefitClass
  Dim s As String
  
  On Error GoTo err_Err
  
  Set ben = Me
  If (Len(ben.value(employer_GUID_db)) = 0) Then
    s = GuidNewEmployer
    Call m_dbfile.Execute(sql.Queries(FIX_UPDATE_GUID, s))
    ben.value(employer_GUID_db) = s
  End If

  
err_End:
  Exit Sub
err_Err:
  Resume err_End
End Sub
  
Public Function EditingCompanyDefinedBenefits(ECDB As EDIT_CDBS, Optional bNewValue As Boolean = False, Optional bOveride As Boolean = False) As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo EditingCompanyDefinedBenefits_Err
  Call xSet("EditingCompanyDefinedBenefits")
  
  Set ben = Me
  If ECDB = ECDB_Value Then
    Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_ID))
    EditingCompanyDefinedBenefits = IIf(IsNull(rs.Fields(S_EDIT_CDB_FIELD_db).value), False, rs.Fields(S_EDIT_CDB_FIELD_db).value)
  Else
    If (Not ben.value(employer_SomeoneEditingCDBS_db)) Or bOveride Then
      If Not bOveride Then
        If ben.value(employer_MeEditingCDBS_db) = bNewValue Then GoTo EditingCompanyDefinedBenefits_End
      End If
      
      Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_ID))
      rs.Edit
      rs.Fields(S_EDIT_CDB_FIELD_db) = bNewValue
      rs.Update
      ben.value(employer_MeEditingCDBS_db) = bNewValue
      EditingCompanyDefinedBenefits = bNewValue
    End If
  End If
  
EditingCompanyDefinedBenefits_End:
  Call xReturn("EditingCompanyDefinedBenefits")
  Exit Function

EditingCompanyDefinedBenefits_Err:
  EditingCompanyDefinedBenefits = False
  Call ErrorMessage(ERR_ERROR, Err, "EditingCompanyDefinedBenefits", "Editing Company Defined Benefits", "Error editing the company defined benefits.")
  Resume EditingCompanyDefinedBenefits_End
  Resume
End Function

Private Property Let IBenefitClass_InvalidFields(ByVal RHS As Long)

End Property

Private Property Get IBenefitClass_InvalidFields() As Long

End Property

Private Sub IBenefitClass_Kill()
  Dim emp As IBenefitClass
  Dim ben As IBenefitClass
  Dim i As Long
  Dim FPCS As FPCS
  
  On Error GoTo CloseEmployer_Err
  Call xSet("CloseEmployer")
  
  For i = 1 To employees.Count
    Set emp = employees(i)
    If Not (emp Is Nothing) Then emp.Kill
  Next i
  
  Set CurrentEmployee = Nothing
  Set m_rsEmpployeesRecordset = Nothing
  Set CDBEmployee = Nothing
  
  Call employees.RemoveAll
  
  For i = 1 To TBL_LASTITEM
    Set m_rsBenTables(i) = Nothing
  Next i
  If Not m_dbfile Is Nothing Then
    For i = 1 To FPCSchemes.Count
      Set FPCS = FPCSchemes(i)
      If Not FPCS Is Nothing Then Call FPCS.Kill
    Next
    FPCSchemes.RemoveAll
    ' shared vans
    Call SharedVans.Vans.RemoveAll
    m_dbfile.Close
    Set m_dbfile = Nothing  ' close current employer when loading next
  End If
  Call RegisterDB(Nothing)
  
CloseEmployer_End:
  Call xReturn("CloseEmployer")
  Call DBEngine.Idle(dbFreeLocks)
  Exit Sub
  
CloseEmployer_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CloseEmployer", "Unload Employer", "Error unloading employer.")
  Resume CloseEmployer_End
  Resume

End Sub

Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_BenItems(employer_Name_db)
End Property

Private Property Let IBenefitClass_NeedToCalculate(ByVal RHS As Boolean)
  m_NeedToCalculate = NeedToCalculateHelper(Me, RHS)
End Property

Private Property Get IBenefitClass_NeedToCalculate() As Boolean
  IBenefitClass_NeedToCalculate = m_NeedToCalculate
End Property

Private Property Set IBenefitClass_Parent(RHS As Object)
  Set m_Parent = RHS
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  Call ECASE("PrintWk invalid for employer.")
End Function

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean
  Call ECASE("PrintWkBody invalid for employer.")
End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim ben As IBenefitClass
  Dim rs As Recordset
  
  Set ben = Me
  
  On Error GoTo ReadDB_Err
  
  Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_ID), dbOpenDynaset, dbFailOnError)
  
  'check for ID entries
  If rs.EOF And rs.BOF Then Err.Raise ERR_EMPLOYER_DB, "Initialize", "Unable to open the employer file as there are no records in the ID table of " & ben.value(employer_FileName)
  
  ben.value(employer_EmployeesWithSharedVan) = 0
  
  ben.value(employer_ApplicationName_db) = rs.Fields("Application").value
  ben.value(employer_ApplicationVersion_db) = rs.Fields("AppVersion").value
  ben.value(employer_MagneticMediaSelected) = False
  ben.value(employer_PAYEOnlineSelected) = False
  ben.value(employer_hasP46Cars) = False
  ben.value(employer_IntranetSelected) = False
  ben.value(employer_FixLevel_db) = rs.Fields("FixLevel").value
  Call FieldPresentToBen(ben, employer_GUID_db, rs, "guid")
  Call FieldPresentToBen(ben, employer_LastApplicationName_db, rs, "LastApplication")
  
  Set rs = m_dbfile.OpenRecordset("T_Employer", dbOpenSnapshot, dbFailOnError)
  If rs.EOF And rs.BOF Then Err.Raise ERR_EMPLOYER_DB, "Initialize", "Unable to open the employer file as there are no records in the employer table of " & ben.value(employer_FileName)
  
  ben.value(employer_NoOfSelectedEmployees) = 0
  ben.value(employer_Contact_db) = "" & rs.Fields("CONTACT_NAME").value
  ben.value(employer_contactnumber_db) = "" & rs.Fields("CONTACT_EXT").value
  ben.value(employer_CT_db) = "" & rs.Fields("TRADINGCOENT").value
  ben.value(employer_Name_db) = "" & rs.Fields("Name").value
  ben.value(employer_Payeref_db) = "" & rs.Fields("PAYE").value
  ben.value(employer_Signatory_db) = "" & rs.Fields("SIGNATORY").value
  ben.value(employer_Year) = EmployerYear
  
  Call FieldPresentToBen(ben, employer_CheapTaxableLoansAllNon_db, rs, "CheapTaxableLoansAllNon", True)
  
  Call FieldPresentToBen(ben, employer_BroughtForward_db, rs, "BroughtForward", True)
  Call FieldPresentToBen(ben, employer_AddressLine1_db, rs, "AddressLine1")
  Call FieldPresentToBen(ben, employer_AddressLine2_db, rs, "AddressLine2")
  Call FieldPresentToBen(ben, employer_AddressLine3_db, rs, "AddressLine3")
  Call FieldPresentToBen(ben, employer_AddressLine4_db, rs, "AddressLine4")
  
  If p11d32.BringForward.Yes Then
  ' New fields for 2005 so default to false on bring forward
    ben.value(employer_EerDeclSect2Chk1_db) = False
    ben.value(employer_EerDeclSect2Chk2_db) = False
    ben.value(employer_EerDeclSect2Chk3_db) = False
  Else
    Call FieldPresentToBen(ben, employer_EerDeclSect2Chk1_db, rs, "EerDeclSect2Chk1")
    Call FieldPresentToBen(ben, employer_EerDeclSect2Chk2_db, rs, "EerDeclSect2Chk2")
    Call FieldPresentToBen(ben, employer_EerDeclSect2Chk3_db, rs, "EerDeclSect2Chk3")
  End If
    
  Call FieldPresentToBen(ben, employer_AddressPostCode_db, rs, "AddressPostCode")
  Call FieldPresentToBen(ben, employer_SubmitterRef_db, rs, "SubmitterRef")
  Call FieldPresentToBen(ben, employer_SubmitterName_db, rs, "SubmitterName")
  Call FieldPresentToBen(ben, employer_PayeOnlineID_db, rs, "PayeOnlineID")
  Call FieldPresentToBen(ben, employer_PayeOnlinePwd_db, rs, "PayeOnlinePwd")
  Call FieldPresentToBen(ben, employer_PayeOnlineUser_db, rs, "PayeOnlineUser")
  Call FieldPresentToBen(ben, employer_PayeOnlineSubID_db, rs, "PayeOnlineSubID")
  
  
  Call FieldPresentToBenDate(ben, employer_PayeOnlineSubmissionDate_db, rs, "PayeOnlineSubmissionDate")
  Call FieldPresentToBen(ben, employer_PayeOnlineSubmissionType_db, rs, "PayeOnlineSubmissionType")  'p11d or p46 car
  
  Call FieldPresentToBen(ben, employer_PayeOnlineEmail_db, rs, "PayeOnlineEmail")
  Call FieldPresentToBen(ben, employer_IRTaxOffice_db, rs, "IRTaxOffice")
  Call FieldPresentToBen(ben, employer_IRAddressLine1_db, rs, "IRAddressLine1")
  Call FieldPresentToBen(ben, employer_IRAddressLine2_db, rs, "IRAddressLine2")
  Call FieldPresentToBen(ben, employer_IRAddressLine3_db, rs, "IRAddressLine3")
  Call FieldPresentToBen(ben, employer_IRAddressLine4_db, rs, "IRAddressLine4")
  Call FieldPresentToBen(ben, employer_IRPostcode_db, rs, "IRPostcode")
  Call FieldPresentToBen(ben, employer_IRTelephoneNumber_db, rs, "IRTelephoneNumber")
  Call FieldPresentToBen(ben, employer_IRYourReference_db, rs, "IRYourReference")
  Call FieldPresentToBen(ben, employer_IRAccountsOfficeReference_db, rs, "IRAccountsOfficeReference")
  Call FieldPresentToBen(ben, employer_AddClass1AAmounts_db, rs, "AddClass1AAmount")
  Call FieldPresentToBen(ben, employer_deductClass1AAmounts_db, rs, "DeductClass1AAmount")
  Call FieldPresentToBen(ben, employer_AddClass1ADescription_db, rs, "AddClass1ADescription")
  Call FieldPresentToBen(ben, employer_deductClass1ADescription_db, rs, "DeductClass1ADescription")
  
  Call FieldPresentToBen(ben, employer_PassWord_db, rs, "PassWord")
  Call FieldPresentToBen(ben, employer_EmployeeResponseDate_db, rs, "EmployeeResponseDate")
  Call FieldPresentToBen(ben, employer_TreatEmployeeCarsAsUnderApprovedFPCS_db, rs, "TreatEmployeeCarsAsUnderApprovedFPCS", True)
  If Len(ben.value(employer_EmployeeResponseDate_db)) = 0 Then ben.value(employer_EmployeeResponseDate_db) = p11d32.Rates.value(EmpLetEmployeeResponseDate)
  If Len(ben.value(employer_AddClass1AAmounts_db)) = 0 Then ben.value(employer_AddClass1AAmounts_db) = 0
  If Len(ben.value(employer_deductClass1AAmounts_db)) = 0 Then ben.value(employer_deductClass1AAmounts_db) = 0
  
  Call FieldPresentToBenEx(ben, employer_GroupCode1Alias_db, rs, "GroupCode1Alias", S_GROUP_CODE1)
  Call FieldPresentToBenEx(ben, employer_GroupCode2Alias_db, rs, "GroupCode2Alias", S_GROUP_CODE2)
  Call FieldPresentToBenEx(ben, employer_GroupCode3Alias_db, rs, "GroupCode3Alias", S_GROUP_CODE3)
  
  
  Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_EMPLOYEES_COUNT))
  
  If Not (rs.EOF And rs.BOF) Then
    ben.value(employer_EmployeesCount) = rs.Fields("Count").value
    If ben.value(employer_EmployeesCount) < 0 Then ben.value(employer_EmployeesCount) = 0
  Else
    ben.value(employer_EmployeesCount) = 0
  End If
  
ReadDB_End:
  Exit Function
  
ReadDB_Err:
  Err.Raise ERR_EMPLOYER_DB, "Initialize", "Unable to open the employer file as the ID table of " & ben.value(employer_FileName) & " could not be opened."
  Resume ReadDB_End
  Resume
End Function

Public Function MagneticMediaValid(ByVal bShowErrors As Boolean) As Boolean
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  On Error GoTo MagneticMediaValid_ERR
  
  Call xSet("MagneticMediaValid")
  
  If (Not PayeRefValid) Then Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateMMData", "Some employers chosen do not have valid PAYE references, they should be of the format 992/E12345.")
  If Len(ben.value(employer_SubmitterRef_db)) = 0 Then Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateMMData", "Submitter ref is zero length for employer " & ben.value(employer_Name_db) & ".")
  If Len(ben.value(employer_SubmitterName_db)) = 0 Then Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateMMData", "Submitter name is zero length for employer " & ben.Name & ".")
  If Len(ben.value(employer_EmployeesCount)) = 0 Then Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateMMData", "No employees for" & ben.Name & ".")
  
  MagneticMediaValid = True
  
MagneticMediaValid_END:
  Call xReturn("MagneticMediaValid")
  Exit Function
MagneticMediaValid_ERR:
  If bShowErrors Then Call ErrorMessage(ERR_ERROR, Err, "MagneticMediaValid", "Magnetic Media Valid", "The employer is invalid for magnetic media.")
  Resume MagneticMediaValid_END
  Resume
End Function

Public Function PAYEOnlineValid(ByVal bShowErrors As Boolean) As Boolean
  Dim benEmployer As IBenefitClass
  'This function applies employer level validations for PAYEOnline required fields only
  Set benEmployer = Me
  
  On Error GoTo PAYEOnlineValid_ERR
  Call xSet("PAYEOnlineValid")
  
  'Check PAYE Refs
  If (Not PayeRefValid) Then Call Err.Raise(ERR_EMPLOYER_INVALID, "PAYEOnlineValid", "Employer '" & benEmployer.Name & "' does not have a valid PAYE reference, format should be 992/E12345.")
    
  'Check SubmitterID
  If Not Len(benEmployer.value(employer_PayeOnlineID_db)) > 0 Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "PAYEOnlineValid", "PAYE online ID for employer " & benEmployer.value(employer_Name_db) & " is missing.")
  End If
  
  'Check SubmitterPwd
  If Not Len(benEmployer.value(employer_PayeOnlinePwd_db)) > 0 Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "PAYEOnlineValid", "PAYE online password for employer " & benEmployer.value(employer_Name_db) & " is missing.")
  End If
  
  'Check SubmitterUsername
  If Not Len(benEmployer.value(employer_PayeOnlineUser_db)) > 0 Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "PAYEOnlineValid", "The PAYE username for employer " & benEmployer.value(employer_Name_db) & " is missing.")
  End If
  
  'IR Tax Office Number which is part of PAYE ref
  If Not Len(benEmployer.value(employer_Payeref_db)) > 0 Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "PAYEOnlineValid", "The PAYE reference for employer " & benEmployer.value(employer_Name_db) & " is missing.")
  End If
  
  'IR Tax Office Reference
  If Not Len(CStr(EmployerRefNo)) > 0 Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "PAYEOnlineValid", "Tax office reference for employer " & benEmployer.value(employer_Name_db) & " is missing.")
  End If

  'Employee Count
  If Len(benEmployer.value(employer_EmployeesCount)) = 0 Then
    Call Err.Raise(ERR_EMPLOYER_INVALID, "PAYEOnlineValid", "No employees for" & benEmployer.Name & ".")
  End If
  
  PAYEOnlineValid = True
  
PAYEOnlineValid_END:
  Call xReturn("PAYEOnlineValid")
  Exit Function
PAYEOnlineValid_ERR:
  If bShowErrors Then Call ErrorMessage(ERR_ERROR, Err, "PAYEOnlineValid", benEmployer.Name, "")
  Resume PAYEOnlineValid_END
  Resume
End Function

Public Function IntranetValid(ByVal bShowErrors As Boolean) As Boolean
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  On Error GoTo IntranetValid_ERR
  
  Call xSet("IntranetValid")
  
  If ben.value(employer_EmployeesCount) = 0 Then Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateMMData", "No employees for" & ben.Name & ".")
  
  IntranetValid = True
  
IntranetValid_END:
  Call xReturn("IntranetValid")
  Exit Function
IntranetValid_ERR:
  If bShowErrors Then Call ErrorMessage(ERR_ERROR, Err, "IntranetValid", "Intranet Valid", "The employer is invalid for intranet.")
  Resume IntranetValid_END
  Resume
End Function


Public Function PassWordOK() As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo PassWordOK_ERR
  
  Call xSet("PassWordOK")
  
  Set ben = Me
  
  If Len(ben.value(employer_PassWord_db)) > 0 Then
    PassWordOK = F_PassWord.PassWord(Me, PWM_CHECK_CURRENT)
  Else
    PassWordOK = True
  End If
  
PassWordOK_END:
  Set F_PassWord = Nothing
  Call xReturn("PassWordOK")
  Exit Function
PassWordOK_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PassWordOK", "Pass Word OK", "Error determining password protection for employer.")
  Resume PassWordOK_END
End Function

Private Function EmployerYear() As String
  Dim ben As IBenefitClass
  Dim s As String
  
  On Error GoTo EmployerYear_ERR
  
  Call xSet("EmployerYear")
  Set ben = Me
  s = Left$(ben.value(employer_FileName), Len(ben.value(employer_FileName)) - Len(S_DB_FILE_EXTENSION))
  If Len(s) < 2 Then Call Err.Raise(ERR_STRING_TOO_SHORT, "EmployerYear", "The filename after removeing the extension is not greater than 2 characterd, = " & s)
  
  s = Right$(s, 2)
  If Not IsNumeric(s) Then Call Err.Raise(ERR_NOT_NUMERIC, "EmployerYear", "The expected year definer is not numeric, = " & s)
  EmployerYear = s
  
EmployerYear_END:
  Call xReturn("EmployerYear")
  Exit Function
EmployerYear_ERR:
  EmployerYear = p11d32.AppYearShort
  Call ErrorMessage(ERR_ERROR, Err, "EmployerYear", "Employer Year", "Error getting the year of the employer file.")
  Resume EmployerYear_END
End Function
Private Property Let IBenefitClass_ReadFromDB(ByVal RHS As Boolean)
  
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean

End Property


Private Property Let IBenefitClass_RSBookMark(RHS As String)
  ECASE ("RSBookMark invalid for employer.")
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  ECASE ("RSBookMark invalid for employer.")
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  
  On Error GoTo SetBenItemsInformation_err
  bc = BC_EMPLOYER
  m_NeedToCalculate = True

  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  p11d32.BenDataLinkDataType(bc, employer_PathAndFile) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_PathAndFileDEL) = TYPE_STR
  
  p11d32.BenDataLinkDataType(bc, employer_Name_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_Payeref_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_Contact_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_Signatory_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_contactnumber_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_CT_db) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_PassWord_db) = TYPE_STR
  
  p11d32.BenDataLinkDataType(bc, employer_NoOfSelectedEmployees) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_FileName) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_FixLevel_db) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_ApplicationName_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_LastApplicationName_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_ApplicationVersion_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_EmployeesWithSharedVan) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_SomeoneEditingCDBS_db) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_MeEditingCDBS_db) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_FileIsOpenBySomeOneElse) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_SubmitterRef_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_SubmitterName_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_PayeOnlineID_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_PayeOnlinePwd_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_PayeOnlineUser_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_PayeOnlineSubID_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_PayeOnlineEmail_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_Benefit) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_NIC_Class1A_Benefit) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_Class1A_Adjustment) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_AddressLine1_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_AddressLine2_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_AddressLine3_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_AddressLine4_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_AddressPostCode_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_MagneticMediaSelected) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_IntranetSelected) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_EmployeesCount) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_BroughtForward_db) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_Year) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_EmployeeResponseDate_db) = TYPE_DATE
  p11d32.BenDataLinkDataType(bc, employer_TreatEmployeeCarsAsUnderApprovedFPCS_db) = TYPE_BOOL
    
  p11d32.BenDataLinkDataType(bc, employer_PAYEOnlineSelected) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, employer_hasP46Cars) = TYPE_BOOL
  
  p11d32.BenDataLinkDataType(bc, employer_IRTaxOffice_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRAddressLine1_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRAddressLine2_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRAddressLine3_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRAddressLine4_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRPostcode_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRTelephoneNumber_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRYourReference_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_IRAccountsOfficeReference_db) = TYPE_STR

  p11d32.BenDataLinkDataType(bc, employer_AddClass1AAmounts_db) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_AddClass1ADescription_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_deductClass1AAmounts_db) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_deductClass1ADescription_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_contactnumber_db) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, employer_Benefit) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_NIC_Class1A_Benefit) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_Class1A_Adjustment) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_NIC_AdjustmentAdd) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_NIC_AdjustmentDeduct) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, employer_Class1A_due) = TYPE_DOUBLE 'EK added 1/04 TTP#51
  
  p11d32.BenDataLinkUDMDescription(bc, employer_Name_db) = "Name"
  p11d32.BenDataLinkUDMDescription(bc, employer_Payeref_db) = "Paye Ref"
  p11d32.BenDataLinkUDMDescription(bc, employer_Contact_db) = "Contact"
  p11d32.BenDataLinkUDMDescription(bc, employer_Signatory_db) = "Signatory"
  p11d32.BenDataLinkUDMDescription(bc, employer_contactnumber_db) = "Extension"
  
  'SO
  p11d32.BenDataLinkUDMRequiresCalculate(bc, employer_Benefit) = True
  p11d32.BenDataLinkUDMRequiresCalculate(bc, employer_NIC_Class1A_Benefit) = True
  
  p11d32.BenDataLinkUDMDescription(bc, employer_Benefit) = "Benefits"
  p11d32.BenDataLinkUDMDescription(bc, employer_NIC_Class1A_Benefit) = "Benefits subject to Class 1A"
  p11d32.BenDataLinkUDMDescription(bc, employer_NIC_AdjustmentAdd) = "Adjustments to add"
  p11d32.BenDataLinkUDMDescription(bc, employer_NIC_AdjustmentDeduct) = "Adjustments to deduct"
  p11d32.BenDataLinkUDMDescription(bc, employer_Class1A_due) = "Class1A NIC" 'EK added 1/04 TTP#51

    
SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information")
  Resume SetBenItemsInformation_end
End Sub

Private Sub IBenefitClass_SetCalcDefaults()
  Dim ben As IBenefitClass
  
  Set ben = Me
'
  ben.value(ITEM_BENEFIT) = 0
  ben.value(ITEM_NIC_CLASS1A_BENEFIT) = 0
  ben.value(ITEM_CLASS1A_ADJUSTMENT) = 0
  ben.value(employer_Class1A_due) = 0 'EK added 1/04 TTP#51
  
End Sub

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  ECASE ("Table invalid for employer.")
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  If Item = employer_PathAndFile Then Call CreatePathAndFile(RHS)
  
  m_BenItems(Item) = CorrectBenValue(BC_EMPLOYER, Item, RHS)
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
  
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim ben As IBenefitClass
  
  On Error GoTo WriteDB_Err
  Call xSet("WriteDB")
  
  Set ben = Me
  
  With ben
    Set rs = m_dbfile.OpenRecordset("SELECT * from T_Employer", dbOpenDynaset, dbFailOnError)
    
    rs.Edit
    rs.Fields("Name").value = .value(employer_Name_db)
    rs.Fields("PAYE").value = .value(employer_Payeref_db)
    rs.Fields("SIGNATORY").value = .value(employer_Signatory_db)
    rs.Fields("CONTACT_NAME").value = .value(employer_Contact_db)
    rs.Fields("CONTACT_EXT").value = .value(employer_contactnumber_db)
    rs.Fields("TRADINGCOENT").value = .value(employer_CT_db)
    rs.Fields("AddressLine1").value = .value(employer_AddressLine1_db)
    rs.Fields("AddressLine2").value = .value(employer_AddressLine2_db)
    rs.Fields("AddressLine3").value = .value(employer_AddressLine3_db)
    rs.Fields("AddressLine4").value = .value(employer_AddressLine4_db)
    rs.Fields("AddressPostCode").value = .value(employer_AddressPostCode_db)
    
    
    If Not p11d32.BringForward.Yes Then
    'new fields for 2005 - remove if statement for subsequent years
      rs.Fields("EerDeclSect2Chk1").value = .value(employer_EerDeclSect2Chk1_db)
      rs.Fields("EerDeclSect2Chk2").value = .value(employer_EerDeclSect2Chk2_db)
      rs.Fields("EerDeclSect2Chk3").value = .value(employer_EerDeclSect2Chk3_db)
    End If
    
    rs.Fields("SubmitterRef").value = .value(employer_SubmitterRef_db)
    rs.Fields("SubmitterName").value = .value(employer_SubmitterName_db)
      
    rs.Fields("PayeOnlineID").value = .value(employer_PayeOnlineID_db)
    rs.Fields("PayeOnlinePwd").value = .value(employer_PayeOnlinePwd_db)
    rs.Fields("PayeOnlineUser").value = .value(employer_PayeOnlineUser_db)
    rs.Fields("PayeOnlineSubID").value = .value(employer_PayeOnlineSubID_db)
    rs.Fields("PayeOnlineSubmissionDate").value = .value(employer_PayeOnlineSubmissionDate_db)
    rs.Fields("PayeOnlineSubmissionType").value = .value(employer_PayeOnlineSubmissionType_db)
    rs.Fields("PayeOnlineEmail").value = .value(employer_PayeOnlineEmail_db)
    
    rs.Fields("BroughtForward").value = .value(employer_BroughtForward_db)
    rs.Fields("PassWord").value = .value(employer_PassWord_db)
    
    
    'AM These now need to be written normally
    rs.Fields("IRTaxOffice").value = .value(employer_IRTaxOffice_db)
    rs.Fields("IRAddressLine1").value = .value(employer_IRAddressLine1_db)
    rs.Fields("IRAddressLine2").value = .value(employer_IRAddressLine2_db)
    rs.Fields("IRAddressLine3").value = .value(employer_IRAddressLine3_db)
    rs.Fields("IRAddressLine4").value = .value(employer_IRAddressLine4_db)
    rs.Fields("IRPostcode").value = .value(employer_IRPostcode_db)
    rs.Fields("IRTelephoneNumber").value = .value(employer_IRTelephoneNumber_db)
    rs.Fields("IRYourReference").value = .value(employer_IRYourReference_db)
    rs.Fields("IRAccountsOfficeReference").value = .value(employer_IRAccountsOfficeReference_db)
            
    rs.Fields("TreatEmployeeCarsAsUnderApprovedFPCS").value = .value(employer_TreatEmployeeCarsAsUnderApprovedFPCS_db)
    
    If p11d32.BringForward.Yes Then rs.Fields("EmployeeResponseDate").value = DateAdd("yyyy", 1, .value(employer_EmployeeResponseDate_db))
    
    'only use this function when new fields have to into the employer table
    If p11d32.BringForward.Yes Then
      rs.Fields("AddClass1AAmount").value = 0
      rs.Fields("AddClass1ADescription").value = ""
      rs.Fields("DeductClass1AAmount").value = 0
      rs.Fields("DeductClass1ADescription").value = ""
    Else
      rs.Fields("EmployeeResponseDate").value = ben.value(employer_EmployeeResponseDate_db)
      rs.Fields("AddClass1AAmount").value = .value(employer_AddClass1AAmounts_db)
      rs.Fields("AddClass1ADescription").value = .value(employer_AddClass1ADescription_db)
      rs.Fields("DeductClass1AAmount").value = .value(employer_deductClass1AAmounts_db)
      rs.Fields("DeductClass1ADescription").value = .value(employer_deductClass1ADescription_db)
      rs.Fields("CheapTaxableLoansAllNon").value = .value(employer_CheapTaxableLoansAllNon_db)
    End If
    
    Call BenValueToFieldPresent(ben, employer_GroupCode1Alias_db, rs, "GroupCode1Alias")
    Call BenValueToFieldPresent(ben, employer_GroupCode2Alias_db, rs, "GroupCode2Alias")
    Call BenValueToFieldPresent(ben, employer_GroupCode3Alias_db, rs, "GroupCode3Alias")
    
  End With
  
  
  rs.Update
    
  IBenefitClass_WriteDB = True
WriteDB_End:
  Set rs = Nothing
  Call xReturn("WriteDB")
  Exit Function

WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  If Err.Number = 3265 Then 'item not found in this collection
    'they may not have opened the employer thus no sync db so tell them to
    Call ErrorMessage(ERR_ERROR, Err, "WriteDB", "Employer write error", "Unable to write the employer details, please open the employer first")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "WriteDB", "Employer write error", "Unable to write the employer details.")
  End If
  
  
  Resume WriteDB_End
  Resume
End Function


Public Sub FindCompanyCar()
  Dim rs As Recordset
  Dim sCarReg As String
  Dim sEmployees As String
  Dim sFirstEmployee As String
  Dim b As Boolean
  Dim i As Long
  Dim ben As IBenefitClass
  Dim lv As ListView
  Dim li As ListItem
  
  On Error GoTo FindCompanyCar_ERR
  
  Call xSet("FindCompanyCar")
  
  F_Input.ValText.TypeOfData = VT_STRING
  If F_Input.Start("Find Company Car", "Enter the registration of the car to find", "") Then
    sCarReg = Trim$(F_Input.ValText.Text)
    Set rs = db.OpenRecordset(sql.Queries(FIND_COMPANY_CAR, sCarReg))
    If Not (rs.EOF And rs.BOF) Then
      Do While Not rs.EOF
        If Not b Then
          sFirstEmployee = rs!P_Num
          b = True
        End If
        sEmployees = sEmployees & rs!P_Num & vbCrLf
        rs.MoveNext
      Loop
      If MsgBox(rs.RecordCount & " employee(s) found with company car registration of " & sCarReg & vbCrLf & vbCrLf & sEmployees & vbCrLf & vbCrLf & "Do you wish to go to " & sFirstEmployee & "?", vbYesNo, "Results") = vbYes Then
        Set lv = F_Employees.lb
        For i = 1 To lv.listitems.Count
          Set li = lv.listitems(i)
          Set ben = employees(li.Tag)
          If StrComp(ben.value(ee_PersonnelNumber_db), sFirstEmployee, vbTextCompare) = 0 Then
            Set lv.SelectedItem = li
            Call BenScreenSwitch(BC_COMPANY_CARS_F)
            GoTo FindCompanyCar_END
          End If
        Next
        Call Err.Raise(ERR_NO_FIND_EMPLOYEE, "FindCompanyCar", "Unable to find the employee " & sFirstEmployee & ".")
      End If
    Else
      Call Err.Raise(ERR_NO_FIND_EMPLOYEE, "FindCompanyCar", "Unable to find any employees with a company car having registration " & sCarReg & ".")
    End If
    
  End If
  
FindCompanyCar_END:
  Set F_Input = Nothing
  Call xReturn("FindCompanyCar")
  Exit Sub
FindCompanyCar_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FindCompanyCar", "Find Company Car", "Error finding a company car.")
  Resume FindCompanyCar_END
  Resume
End Sub
Public Sub CompanyCarChecker()
  On Error GoTo CompanyCarChecker_ERR
  
  Call F_CompanyCarChecker.Start(Me)
  'Call F_CompanyCarCheckerWizard.Start(Me)
CompanyCarChecker_END:
  Unload F_CompanyCarChecker
  Set F_CompanyCarChecker = Nothing
  Exit Sub
CompanyCarChecker_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CompanyCarChecker", "Company Car Checker", "Error in F_CompanyCarChecker.")
  Resume CompanyCarChecker_END
  Resume
End Sub

Public Function TransferEmployees() As Long
  Dim i As Long
  Dim ben As IBenefitClass
  Dim li As ListItem
  
  On Error GoTo TransferEmployees_ERR
  
  Call xSet("TransferEmployees")
    
  F_Transfer.lv.listitems.Clear
  Set F_Transfer.LastListItem = Nothing
  Set F_Transfer.lv.SmallIcons = MDIMain.imlTree
  F_Transfer.prg.Caption = ""
  F_Transfer.prg.value = F_Transfer.prg.Min
  F_Transfer.prg.Indicator = None
  
  For i = 1 To p11d32.Employers.Count
    Set ben = p11d32.Employers(i)
      If Not ben Is Nothing Then
        If Not ben Is Me Then
          Set li = F_Transfer.lv.listitems.Add(, , ben.Name, , IMG_UNSELECTED)
          li.SubItems(1) = ben.value(employer_FileName)
          li.SubItems(2) = ben.value(employer_FixLevel_db)
          li.Tag = i
        End If
      End If
  Next
  
  Call p11d32.Help.ShowForm(F_Transfer, vbModal)
  
TransferEmployees_END:
  Call xReturn("TransferEmployees")
  Exit Function
TransferEmployees_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "TransferEmployees", "Transfer Employees", "Error transferring employees.")
  Resume TransferEmployees_END
  Resume
End Function

'MP DB confirm removal ok - does not need for any compatability?
'Public Sub FindAccomExpenses()
'  Dim rs As Recordset
'  Dim sNames As String
'  On Error GoTo FindAccomExpenses_Err
'
'  Call xSet("FindAccomExpenses")
'
'  Set rs = db.OpenRecordset(sql.Queries(FIND_ACCOM_EXPENSES))
'  If Not rs Is Nothing And rs.RecordCount Then
'      rs.MoveFirst
'      Do While Not rs.EOF
'        If Len(sNames) = 0 Then
'          sNames = sNames + rs.Fields("P_NUM").value
'        Else
'          sNames = sNames + ", " + rs.Fields("P_NUM").value
'        End If
'        rs.MoveNext
'      Loop
'      Call MsgBox("Extra accomodation expenses are now to be entered separately under section N." & vbCrLf & _
'              "Please re-enter any accomodation expenses for employee(s) " & sNames & " under section N.", vbOKOnly, "Accomodation Expenses")
'      Call Err.Raise(ERR_ACCOMODATION_EXPENSES, , "Extra accomodation expenses are now to be entered separately under section N.  Please re-enter any accomodation expenses for employee(s) " & sNames & ".")
'      Call db.Execute(sql.Queries(DELETE_ACCOM_EXPENSES))
'  End If
'
'  Set rs = Nothing
'FindAccomExpenses_End:
'  Call xReturn("FindAccomExpenses")
'  Exit Sub
'FindAccomExpenses_Err:
'  If Err.Number = ERR_ACCOMODATION_EXPENSES Then
'    Call ErrorMessage(ERR_ACCOMODATION_EXPENSES, Err, "FindAccomExpenses", "Accomodation expenses change.", Err.Description)
'    Resume Next
'  Else
'    Call ErrorMessage(ERR_ERROR, Err, "FindAccomExpenses", "FindAccomExpenses", "Error finding where there are any accomodation expenses.")
'    Resume FindAccomExpenses_End
'  End If
'End Sub

Public Property Get GroupCode1Alias() As String
  GroupCode1Alias = m_BenItems(employer_GroupCode1Alias_db)
End Property
Public Property Get GroupCode2Alias() As String
  GroupCode2Alias = m_BenItems(employer_GroupCode2Alias_db)
End Property
Public Property Get GroupCode3Alias() As String
  GroupCode3Alias = m_BenItems(employer_GroupCode3Alias_db)
End Property

