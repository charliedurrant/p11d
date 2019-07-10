Attribute VB_Name = "BenCTRL"

Option Explicit

Public Enum GPFB_TYPE
  GPBF_EMPLOYEE = 1
  GPBF_EMPLOYER
End Enum
Public Sub IRDescDB(ByVal ben As IBenefitClass, ByVal rs As Recordset, bRead As Boolean)
  Dim b As Boolean
  Dim iBenITem As Long
  On Error GoTo err_err
  
  b = HasIRDescription(ben.BenefitClass)
  If (Not b) Then
    Call Err.Raise(ERR_BENCLASS_INVALID, "IRDescDB", "The benefit class does not have an IR description")
  End If
  iBenITem = IRDescriptionBenItem(ben.BenefitClass)
  If bRead Then
    If p11d32.BringForward.Yes And p11d32.AppYear = 2003 Then
      ben.value(iBenITem) = S_IR_DESC_OTHER
    Else
      If Not FieldPresent(rs.Fields, "IRDesc") Then
        Call Err.Raise(ERR_FILE_NOT_EXIST, "IRDescDB", "No IRDesc field")
      End If
      ben.value(iBenITem) = IIf(IsNull(rs.Fields("IRDesc").value), S_IR_DESC_OTHER, rs.Fields("IRDesc").value)
    End If
  Else
    If Not FieldPresent(rs.Fields, "IRDesc") Then
      Call Err.Raise(ERR_FILE_NOT_EXIST, "IRDescDB", "No IRDesc field")
    End If
    rs.Fields("IRDesc") = ben.value(iBenITem)
  End If
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(Err.Number, Err, "IRDescDB", "IRDescDB", "Error getting the IR description from the DB")
  Resume err_end
  Resume
End Sub
Public Function BenefitsOfType(ee As Employee, ByVal bc As BEN_CLASS) As ObjectList
  Dim i As Long
  Dim o As ObjectList
  Dim ben As IBenefitClass
  Dim benefits As ObjectList
  
  If ee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenfitsOfType", "The employee is nothing.")
  
  Set o = New ObjectList
  
  If BenefitIsLoan(bc) Then
    i = ee.GetLoansBenefitIndex()
    If i = 0 Then Call Err.Raise(ERR_IS_NOTHING, "BenefitsOfType", "BenefitsOfType", "The loans collection is nothing for " & ee.FullName & ".")
    Set benefits = ee.benefits(i)
  Else
    Set benefits = ee.benefits
  End If
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = bc Then Call o.Add(ben)
    End If
  Next
   
  Set BenefitsOfType = o
End Function
Public Function BenefitNewProperties(ByVal ben As IBenefitClass)
  
End Function

Public Function NeedToCalculatePropogate(ben As IBenefitClass)
  Dim benParent As IBenefitClass
  
  Set benParent = ben.Parent
  Do While Not benParent Is Nothing
    benParent.NeedToCalculate = benParent.NeedToCalculate Or True
    Set benParent = benParent.Parent
  Loop
  
End Function
Public Function IsCBDEmployee(sPersonnelNumber As String) As Boolean
  If Len(sPersonnelNumber) >= Len(S_CDB_EMPLOYEE_NUMBER_PREFIX) Then
    If StrComp(Left$(sPersonnelNumber, Len(S_CDB_EMPLOYEE_NUMBER_PREFIX)), S_CDB_EMPLOYEE_NUMBER_PREFIX, vbTextCompare) = 0 Then
      IsCBDEmployee = True
    End If
  End If
End Function
Public Function DirtyHelper(ben As IBenefitClass, ByVal NewValue As Boolean) As Boolean
  Dim benParent As IBenefitClass
  DirtyHelper = ben.Dirty Or NewValue
  If (NewValue) Then
    ben.NeedToCalculate = True
  End If
End Function
'cad cdb stuff
Public Function NeedToCalculateHelper(ben As IBenefitClass, ByVal NewValue As Boolean) As Boolean
  Dim i As Long, j As Long
  Dim employees As ObjectList
  Dim Employer As Employer
  Dim benEE As IBenefitClass
  Dim benOther As other
  Dim ee As Employee
  Dim bOriginal As Boolean
  Dim lCDBEmployeeBenefitIndex As Long
On Error GoTo err_err

  Dim benParent As IBenefitClass
      
  NeedToCalculateHelper = NewValue

  If Not (ben.Parent Is Nothing) Then
    Set benParent = ben.Parent
    If (Not benParent.NeedToCalculate) And NewValue Then
      benParent.NeedToCalculate = NewValue
    End If
    If (NewValue) Then
      If StrComp(TypeName(benParent), "Employee") = 0 Then
        If (IsCBDEmployee(benParent.value(ee_PersonnelNumber_db))) Then
          'loop thrugh all the employees that have benefits loaded
          Set Employer = benParent.Parent
          Set employees = Employer.employees
          For i = 1 To employees.Count
            Set ee = employees(i)
            If Not ee Is Nothing Then
              If Not IsCBDEmployee(ee.PersonnelNumber) Then
                If ee.BenefitsLoaded Then
                  Set benEE = ee
                  benEE.NeedToCalculate = ee.HasCDBBenefit(ben)
                End If
              End If
            End If
          Next i
        End If
      End If
    End If
    
  End If
err_end:
  Exit Function
err_err:
  Call ECASE("parent is not a benefit class:" & Err.Description)
  Resume err_end
  Resume
End Function

Public Function CalculateHelper(ben As IBenefitClass) As Variant
  On Error GoTo CalculateHelper_ERR
  
  If ben.NeedToCalculate Then
    CalculateHelper = ben.CalculateBody
    ben.NeedToCalculate = False
  Else
    CalculateHelper = ben.value(ITEM_BENEFIT)
  End If
  
CalculateHelper_END:
  Exit Function
CalculateHelper_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CalculateHelper", "Calculate Helper", "Error calculating a benefit.")
  Resume CalculateHelper_END
  Resume
End Function
Public Function LastFixLevel(lYear As Long) As Long

  Select Case lYear
    Case 0
      LastFixLevel = 75
    Case 1
      LastFixLevel = 76
    Case 2
      LastFixLevel = 77
    Case 3
      LastFixLevel = 88
    Case 4
      LastFixLevel = 93
    Case 5
      LastFixLevel = 94
    Case 6
      LastFixLevel = 95
    Case 7
      LastFixLevel = 98
    Case 8
      LastFixLevel = 101
    Case 9
      LastFixLevel = 103 'CAD 20
    Case 10
      LastFixLevel = 105
    Case 11
      LastFixLevel = 106
    Case 12
      LastFixLevel = 107
    Case 13
      LastFixLevel = 108
    Case 14
      LastFixLevel = 109
    Case 15
      LastFixLevel = 110
    Case 16
      LastFixLevel = 111
    Case 17
      LastFixLevel = 113
    Case Else
      Call ECASE("Invalid year in LastFixLevel.")
  End Select
  
End Function
Public Function EnumEmployerFiles(ByVal sExtension As String, IENUME As IEnumEmployers) As Long
  Dim q As String
  Dim s As String
  Dim empr As Employer
  
  On Error GoTo EnumEmployerFiles_ERR
  Call xSet("EnumEmployerFiles")
  
  If IENUME Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "EnumEmployerFiles", "The IBaseNotofy is nothing.")
  
  s = p11d32.WorkingDirectory & "*" & sExtension
  
  Call IENUME.Count(CountFiles(s))  ' display number of employers
    
  q = Dir$(s)
  Do While Len(q) > 0
    IENUME.CurrentFile (q)
    Set empr = New Employer
    If empr.Validate(p11d32.WorkingDirectory & q) Then
      Call IENUME.Employer(empr)
      EnumEmployerFiles = EnumEmployerFiles + 1
    End If
    Set empr = Nothing
    q = Dir$()
  Loop
  
EnumEmployerFiles_END:
  Call xReturn("EnumEmployerFiles")
  Exit Function
EnumEmployerFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EnumEmployerFiles", "Enum Employer Files", "Error enumeration the employer files.")
  Resume EnumEmployerFiles_END
  Resume
End Function

Public Function CopyBenData(benDst As IBenefitClass, benSrc As IBenefitClass) As Boolean
  Dim bc As BEN_CLASS
  Dim i As Long
  On Error GoTo CopyBenData_ERR
  Call xSet("CopyBenData")
  
  If benDst Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "CopyBenData", "The destination ben is nothing.")
  If benSrc Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "CopyBenData", "The source ben is nothing.")
  
  bc = benSrc.BenefitClass
  
  For i = 1 To p11d32.Rates.BenClassTo(bc, BCT_BENITEMS_LAST_ITEM)
    benDst.value(i) = benSrc.value(i)
  Next
  benDst.CompanyDefined = benSrc.CompanyDefined
  If Not BenefitIsLoan(bc) Then benDst.Dirty = True
  CopyBenData = True
  
CopyBenData_END:
  Call xReturn("CopyBenData")
  Exit Function
CopyBenData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CopyBenData", "Copy Ben Data", "Error copying a benefits data.")
  Resume CopyBenData_END
  Resume
End Function
Private Function ObjectListIndexExist(OL As ObjectList, ByVal lIndex As Long) As Boolean
  Dim o As Object
  
  On Error GoTo ObjectListIndexExist_ERR
  
  Set o = OL(lIndex)
  ObjectListIndexExist = True
ObjectListIndexExist_END:
  Exit Function
ObjectListIndexExist_ERR:
  Resume ObjectListIndexExist_END
End Function
Private Function CopyDataCheck(ey As Employer) As Boolean
  If ey.CopyBenClass < BC_FIRST_ITEM Then Exit Function
  If ey.CopyEmployeeIndex < 1 Then Exit Function
  If ey.CopyBenIndex < 1 Then Exit Function
  CopyDataCheck = True
End Function
Private Function CopySetData() As Boolean
  Dim ibf As IBenefitForm2
  
  On Error GoTo CopySetData_ERR
  Call xSet("CopySetData")
    
  If IsBenefitForm(CurrentForm) Then Else Call Err.Raise(ERR_NOT_BENEFIT_FORM, "CopySetData", "The form is not a benefit form.")
  
  Set ibf = CurrentForm
  
  If ibf.lv.SelectedItem Is Nothing Then GoTo CopySetData_END
  p11d32.CurrentEmployer.CopyBenIndex = ibf.lv.SelectedItem.Tag
  
  p11d32.CurrentEmployer.CopyBenClass = ibf.benclass
  p11d32.CurrentEmployer.CopyEmployeeIndex = p11d32.CurrentEmployer.employees.ItemIndex(p11d32.CurrentEmployer.CurrentEmployee)
  
  CopySetData = True
CopySetData_END:
  Call xReturn("CopySetData")
  Exit Function
CopySetData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CopySetData", "Copy Set Data", "Error setting the data needed to copy.")
  Resume CopySetData_END
End Function
Private Sub ListItemSelectedGrey(ibf As IBenefitForm2)
  Dim i As Long
  Dim li As ListItem
  
  If ibf Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "ListItemGrey", "The IBenefitForm is nothing.")
  If Not ibf.lv.SelectedItem Is Nothing Then
    Set li = ibf.lv.SelectedItem
    li.ForeColor = vbGrayText
    For i = 1 To li.ListSubItems.Count
      li.ListSubItems(i).ForeColor = vbGrayText
    Next
  End If
End Sub
Public Sub LVKeyDown(KeyCode As Integer, Shift As Integer, Optional frm As Form)
  Dim ben As IBenefitClass
  Dim i As Long
  Dim ee As Employee
  Dim employees As ObjectList
  Dim ibf As IBenefitForm2
  Dim bInFunc As Boolean
  Dim f As Form
  
  On Error GoTo LVKeyDown_ERR
  Call xSet("LVKeyDown")

  If bInFunc Then GoTo LVKeyDown_END
  bInFunc = True
  
  If Not frm Is Nothing Then
    Set f = frm
  Else
    Set f = Screen.ActiveForm
  End If
  
  If Not IsBenefitFormReal(f) Then GoTo LVKeyDown_END
  Set ibf = f
  
  If (Shift And vbCtrlMask) = vbCtrlMask Then
    Select Case KeyCode
      Case vbKeyC, vbKeyX
        If Not CopySetData Then GoTo LVKeyDown_END_PROC
        If KeyCode = vbKeyX Then
          p11d32.CurrentEmployer.CopyBenCut = True
          Call ListItemSelectedGrey(ibf)
        End If
        KeyCode = 0
      Case vbKeyV
        If CopyBenHelper(CurrentForm) Then
          If p11d32.CurrentEmployer.CopyBenCut Then
            'ie I was in a cut
            Set employees = p11d32.CurrentEmployer.employees
            Set ee = employees(p11d32.CurrentEmployer.CopyEmployeeIndex)
            Set ibf = CurrentForm
            i = p11d32.CurrentEmployer.CopyBenIndex
            If BenefitIsLoan(p11d32.CurrentEmployer.CopyBenClass) Then
              Set ben = ee.GetLoan(i)
            Else
              Set ben = ee.benefits(i)
            End If
            Call ee.RemoveBenefit(ibf, ben, i, p11d32.CurrentEmployer.CurrentEmployee Is ee)
            Call p11d32.CurrentEmployer.CopyReset
          End If
          Call MDIMain.SetConfirmUndo
        End If
        KeyCode = 0
    End Select
  End If
  
LVKeyDown_END_PROC:
  bInFunc = False
LVKeyDown_END:
  Call xReturn("LVKeyDown")
  Exit Sub
LVKeyDown_ERR:
  bInFunc = False
  Call ErrorMessage(ERR_ERROR, Err, "LVKeyDown", "Copy Paste Ben", "Error copying, cuting or pasting a benefit.")
  Resume LVKeyDown_END
  Resume
End Sub
Public Function CopyBenStandard(Parent As Object, benDst As IBenefitClass, benSrc As IBenefitClass) As IBenefitClass
  Dim ee As Employee
  
  On Error GoTo CopyBenStandard_ERR
  Call xSet("CopyBenStandard")
  
  If benSrc Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CopyBenStandard", "Source ben is nothing.")
  If benDst Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CopyBenStandard", "Destination ben is nothing.")
  
  If IsEmployee(Parent) Then
    benDst.BenefitClass = benSrc.BenefitClass
    If CopyBenData(benDst, benSrc) Then
      Set ee = Parent
      Set benDst.Parent = ee
      Call ee.benefits.Add(benDst)
      Call CopyBenStandardEnd(benDst)
      Set CopyBenStandard = benDst
    End If
  End If
  
CopyBenStandard_END:
  Call xReturn("CopyBenStandard")
  Exit Function
CopyBenStandard_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CopyBenStandard", "Copy Ben Standard", "Error copying a standard benefit.")
  Resume CopyBenStandard_END
  Resume
End Function
Public Function LoansAddToCurrentEmployee(ee As Employee) As loans
  Dim loans As loans
  Dim lLoansIndex As Long
  Dim ben As IBenefitClass
  
  On Error GoTo LoansAddToCurrentEmployee_ERR
  
  Call xSet("LoansAddToCurrentEmployee")
  
  If ee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "LoansAddToCurrentEmployee", "The employee is nothing.")
  
  lLoansIndex = ee.GetLoansBenefitIndex
  If lLoansIndex = 0 Then
    Set loans = New loans
    Set ben = loans
    Set ben.Parent = p11d32.CurrentEmployer.CurrentEmployee
    Call p11d32.CurrentEmployer.CurrentEmployee.benefits.Add(loans)
    Set LoansAddToCurrentEmployee = loans
  Else
    Set LoansAddToCurrentEmployee = ee.benefits(lLoansIndex)
  End If
  
  
LoansAddToCurrentEmployee_END:
  Call xReturn("LoansAddToCurrentEmployee")
  Exit Function
LoansAddToCurrentEmployee_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LoansAddToCurrentEmployee", "Loans Add To Current Employee", "Error adding the loans collection to the current employee.")
  Resume LoansAddToCurrentEmployee_END
End Function

Public Sub WriteDBDates(ben As IBenefitClass, ByVal FromID As Long, ByVal ToID As Long, rs As Recordset, sFieldFrom As String, sFieldTo As String)

End Sub
Public Sub BackupEmployer(ey As Employer)
  Dim ben As IBenefitClass
  
  On Error GoTo BackupEmployer_ERR

  Call xSet("BackupEmployer")

  If ey Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BackupEmployer", "The employer is nothing.")
  Set ben = ey
  
  Call FileCopyEx(ben.value(employer_PathAndFile), ben.value(employer_PathAndFileDEL))
  Call xKill(ben.value(employer_PathAndFile))
  
BackupEmployer_END:
  Call xReturn("BackupEmployer")
  Exit Sub
BackupEmployer_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BackupEmployer", "Backup Employer", "Error backing up an employer to .del file.")
  Resume BackupEmployer_END
End Sub
Public Function BringForward(l As Long) As Boolean
  If l > -2 Then BringForward = True
End Function
Public Sub BenefitAddNewRecord(ben As IBenefitClass, rs As Recordset)
  Dim ey As Employer
  
  If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenefitAddNewRecord", "Benefit is nothing")
  
  If ben.TABLE = 0 Then Call Err.Raise(ERR_IS_NOTHING, "BenefitAddNewRecord", "Table for benefit is 0 length string")
  
  Set ey = GetParentFromBenefit(ben, GPBF_EMPLOYER)
  Set rs = ey.rsBenTables(ben.TABLE)
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenefitAddNewRecord", "Recordset is nothing")
  
  If (p11d32.BringForward.Yes And Not ben.ReadFromDB) Or (Len(ben.RSBookMark) = 0) Then
    rs.AddNew
    rs.Fields(S_FIELD_PERSONEL_NUMBER) = GetEmployeeNumber(ben)
  Else
    rs.Bookmark = ben.RSBookMark
    rs.Edit
  End If
  
End Sub
Public Function BenefitCloseRecord(ben As IBenefitClass, rs As Recordset, Optional bUpdate As Boolean = True) As Boolean
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenefitCloseRecord", "Recordset is nothing.")
  If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenefitCloseRecord", "The benefit is nothing.")
  
  If bUpdate Then Call rs.Update
  ben.RSBookMark = rs.LastModified
  rs.Bookmark = rs.LastModified
  ben.Dirty = False
  BenefitCloseRecord = True
End Function
Public Sub CopyBenStandardEnd(ben As IBenefitClass)
  If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CopyBenStandardEnd", "The ben is nothing.")
  ben.ReadFromDB = True
  ben.Dirty = True
End Sub
Public Function CopyBenHelper(ibf As IBenefitForm2) As Boolean
  Dim bc As BEN_CLASS
  Dim ey As Employer
  Dim ee As Employee
  Dim ben As IBenefitClass
  Dim loans As loans
  Dim bIsLoan As Boolean
  
  On Error GoTo CopyBenHelper_ERR
  Call xSet("CopyBenHelper")
  
  If ibf Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CopyBenHelper", "Ibenefit form is nothing.")
  'is the ben class correct for the form to copy to?
  
  Set ey = p11d32.CurrentEmployer
  If Not CopyDataCheck(ey) Then GoTo CopyBenHelper_CLEAR
  If ibf.benclass <> ey.CopyBenClass Then Call Err.Raise(ERR_BEN_CLASS_NOT_EQUAL, "CopyBenHelper", "The current paste table benefit is of type " & p11d32.Rates.BenClassTo(ey.CopyBenClass, BCT_FORM_CAPTION))
  If Not ObjectListIndexExist(p11d32.CurrentEmployer.employees, ey.CopyEmployeeIndex) Then GoTo CopyBenHelper_END
  Set ee = ey.employees(ey.CopyEmployeeIndex)
  If ee Is Nothing Then GoTo CopyBenHelper_CLEAR
  'get valid benefit
  
  bIsLoan = BenefitIsLoan(ey.CopyBenClass)
  
  If bIsLoan Then
    If Not ee.AnyLoanBenefit Then GoTo CopyBenHelper_END
    Set loans = ee.benefits(ee.GetLoansBenefitIndex)
    If Not ObjectListIndexExist(loans.loans, ey.CopyBenIndex) Then GoTo CopyBenHelper_END
    Set ben = loans.loans(ey.CopyBenIndex)
  Else
    If Not ObjectListIndexExist(ee.benefits, ey.CopyBenIndex) Then GoTo CopyBenHelper_END
    Set ben = ee.benefits(ey.CopyBenIndex)
    If ben.CompanyDefined Then GoTo CopyBenHelper_END
  End If
  
  If ben Is Nothing Then GoTo CopyBenHelper_CLEAR
  Set ee = p11d32.CurrentEmployer.CurrentEmployee
  Set ben = ben.Copy(ee)
  
  If ben Is Nothing Then Call Err.Raise(ERR_COPY_FAILED, "CopyBenHelper", "The benefit " & ben.Name & " failed to copy itself, please check your data.")
  
  If bIsLoan Then
    Set loans = ee.benefits(ee.GetLoansBenefitIndex)
    If loans Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CopyBenHelper", "Copied a loan and the loans collection is nothing.")
    Call AddBenefitHelperSub(ben, ibf, loans.loans.ItemIndex(ben))
  Else
    Call AddBenefitHelperSub(ben, ibf, p11d32.CurrentEmployer.CurrentEmployee.benefits.ItemIndex(ben))
  End If
  CopyBenHelper = True
  
  
  
CopyBenHelper_END:
  Call xReturn("CopyBenHelper")
  Exit Function
CopyBenHelper_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CopyBenHelper", "Copy Ben Helper", "Error copying the benefit.")
  Resume CopyBenHelper_END
  Resume
CopyBenHelper_CLEAR:
  If Not ey Is Nothing Then Call ey.CopyReset
  GoTo CopyBenHelper_END
End Function
Public Function GetEmployeeNumber(ben As IBenefitClass) As String
  Dim ee As Employee
  Dim benX As IBenefitClass
  
  On Error GoTo GetEmployeeNumber_ERR
  
    
  If ben.BenefitClass > BC_UDM_BENEFITS_LAST_ITEM Or ben.BenefitClass < 1 Then Call ErrorMessage(ERR_BENCLASS_INVALID, Err, "GetEmployeeNumber", "Get Employee Number", "Error getting the personnel number from a benefits parent employee.")
   
  Set benX = ben
  
  Do Until (benX Is Nothing)
    Set benX = benX.Parent
    If IsEmployee(benX) Then
      Set ee = benX
      Exit Do
    End If
  Loop
  
  If Not ee Is Nothing Then GetEmployeeNumber = ee.PersonnelNumber
  
GetEmployeeNumber_END:
  Exit Function
GetEmployeeNumber_ERR:
  Resume GetEmployeeNumber_END
End Function
Public Function IsEmployee(ben As IBenefitClass) As Boolean
  Dim ee As Employee
  On Error GoTo IsEmployee_ERR
  
  Set ee = ben
  If Not ee Is Nothing Then IsEmployee = True
  
IsEmployee_END:
  Exit Function
IsEmployee_ERR:
  Resume IsEmployee_END
End Function
Public Function IsEmployer(ben As IBenefitClass) As Boolean
  Dim ey As Employer
  On Error GoTo IsEmployer_ERR
  
  Set ey = ben
  If Not ey Is Nothing Then IsEmployer = True
  
IsEmployer_END:
  Exit Function
IsEmployer_ERR:
  Resume IsEmployer_END
End Function
Public Function PassWordWrite(ey As Employer, ByVal sNewPassWord) As Boolean
  Dim ben As IBenefitClass
  
On Error GoTo PassWordWrite_ERR

  Call xSet("PassWordWrite")
  
  If ey Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "PassWordWrite", "The employer is nothing.")
  
  Set ben = ey
  ben.value(employer_PassWord_db) = sNewPassWord
  Call ben.writeDB
  PassWordWrite = True
PassWordWrite_END:
  Call xReturn("PassWordWrite")
  Exit Function
PassWordWrite_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PassWordWrite", "Pass Word Write", "Error writing a password.")
  Resume PassWordWrite_END
End Function
Public Sub InitFindFiles()
      
  On Error GoTo InitFindFiles_ERR
  
  Call xSet("InitFindFiles")
  
  F_FindFiles.chkShowSubDirs.value = BoolToChkBox(p11d32.FindFilesSearchSubDirs)
'  F_FindFiles.Show vbModal
  Call p11d32.Help.ShowForm(F_FindFiles, vbModal)
  Set F_FindFiles = Nothing
  
InitFindFiles_END:
  Call xReturn("InitFindFiles")
  Set F_Print = Nothing
  Exit Sub
  
InitFindFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "InitFindFiles", "Find Files", "Error Displaying Dialogue")
  Resume InitFindFiles_END
    
End Sub

  
  
  
  
  

Public Sub RepairCompactEmployer()
  Dim ibf As IBenefitForm2
  Dim ey As Employer
  
  On Error GoTo RepairCompactEmployer_ERR
  
  Call xSet("RepairCompactEmployer")
  
  Call OnlyFromForm(F_Employers)
  Set ibf = CurrentForm
  If ibf.lv.listitems.Count = 0 Then Call Err.Raise(ERR_REPAIR_COMPACT, "RepairCompactEmployer", "No employers.")
  If Not ibf.lv.SelectedItem Is Nothing Then
    Set ey = p11d32.Employers(ibf.lv.SelectedItem.Tag)
    If ey Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "RepairCompactEmployer", "Employer is nothing.")
    Call ey.RepairAndCompact
  End If
    
  
  
RepairCompactEmployer_END:
  Call xSet("RepairCompactEmployer")
  Exit Sub
RepairCompactEmployer_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RepairCompactEmployer", "Repair Compact Employer", "Error repairing and compacting file")
  Resume RepairCompactEmployer_END
  Resume
End Sub
Public Sub BenValueToFieldPresent(ben As IBenefitClass, Item As Long, rs As Recordset, ByVal sField As String)
  If FieldPresent(rs.Fields, sField) Then
    rs.Fields(sField) = ben.value(Item)
  End If
End Sub
Public Sub FieldPresentToBen(ben As IBenefitClass, Item As Long, rs As Recordset, ByVal sField As String, Optional bFalseASDefault As Boolean = False)

  On Error GoTo FieldPresentToBen_ERR
  
  Call xSet("FieldPresentToBen")
  
  If ben Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "FieldPresentToBen", "The benefit passed is nothing.")
  
  If FieldPresent(rs.Fields, sField) Then
    ben.value(Item) = "" & rs.Fields(sField).value
  Else
    If bFalseASDefault Then
      ben.value(Item) = False
    Else
      ben.value(Item) = ""
    End If
  End If
  
FieldPresentToBen_END:
  Call xReturn("FieldPresentToBen")
  Exit Sub
FieldPresentToBen_ERR:
  If Err.Number = ERR_BEN_IS_NOTHING Then
    Call ErrorMessage(ERR_ERROR, Err, "FieldPresentToBen", "Field Present To Ben", "Error setting a ben value when checking if db field is present. Field = " & sField & ".")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "FieldPresentToBen", "Field Present To Ben", "Error setting a ben value when checking if db field is present. Field = " & sField & ", Ben = " & ben.Name)
  End If
End Sub
Public Sub FieldPresentToBenDate(ben As IBenefitClass, Item As Long, rs As Recordset, ByVal sField As String)

  On Error GoTo FieldPresentToBen_ERR
  
  Call xSet("FieldPresentToBen")
  
  If ben Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "FieldPresentToBen", "The benefit passed is nothing.")
  
  If FieldPresent(rs.Fields, sField) Then
    ben.value(Item) = IsNullEx(rs.Fields(sField).value, UNDATED)
  Else
    ben.value(Item) = UNDATED
  End If
  
FieldPresentToBen_END:
  Call xReturn("FieldPresentToBen")
  Exit Sub
FieldPresentToBen_ERR:
  If Err.Number = ERR_BEN_IS_NOTHING Then
    Call ErrorMessage(ERR_ERROR, Err, "FieldPresentToBenDate", "Field Present To Ben Date", "Error setting a ben value when checking if db field is present. Field = " & sField & ".")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "FieldPresentToBenDate", "Field Present To Ben Date", "Error setting a ben value when checking if db field is present. Field = " & sField & ", Ben = " & ben.Name)
  End If
End Sub
Public Sub FieldPresentToBenEx(ben As IBenefitClass, Item As Long, rs As Recordset, ByVal sField As String, ByVal vDefault As Variant)

  On Error GoTo FieldPresentToBenEx_ERR
  
  Call xSet("FieldPresentToBen")
  
  If ben Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "FieldPresentToBen", "The benefit passed is nothing.")
  
  If FieldPresent(rs.Fields, sField) Then
    ben.value(Item) = IsNullEx(rs.Fields(sField).value, vDefault)
  Else
    ben.value(Item) = vDefault
  End If
  
FieldPresentToBenEx_END:
  Call xReturn("FieldPresentToBen")
  Exit Sub
FieldPresentToBenEx_ERR:
  If Err.Number = ERR_BEN_IS_NOTHING Then
    Call ErrorMessage(ERR_ERROR, Err, "FieldPresentToBenEx", "Field Present To Ben Ex", "Error setting a ben value when checking if db field is present. Field = " & sField & ".")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "FieldPresentToBenEx", "Field Present To Ben Ex", "Error setting a ben value when checking if db field is present. Field = " & sField & ", Ben = " & ben.Name)
  End If
End Sub

Public Function PrintWKHelper(rep As Reporter, ben As IBenefitClass, Optional bPrintDescription As Boolean = True) As Boolean
  On Error GoTo PrintWKHelper_ERR
  
  Call xSet("PrintWKHelper")
  
  If ben Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "PrintWKHelper", "The benefit is nothing.")
  If rep Is Nothing Then Call Err.Raise(ERR_REPORTER_IS_NOTHING, "PrintWKHelper", "The reporter is nothing.")
  rep.PageFooter = HMITFooter("Working Papers", GetParentFromBenefit(ben, GPBF_EMPLOYEE))
  Call rep.Out("{BEGINSECTION}")
  
  Call WKBenefitHeader(rep, ben, bPrintDescription)
  Call ben.PrintWkBody(rep)
  Call WKOut(rep, WK_SECTION_BREAK)
  Call rep.Out("{ENDSECTION}")
  
  PrintWKHelper = True
PrintWKHelper_END:
  Call xReturn("PrintWKHelper")
  Exit Function
PrintWKHelper_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PrintWKHelper", "Print WK Helper", "Error printing WK paper.")
  Resume PrintWKHelper_END
End Function
Public Sub SetAvaialbleRange(ben As IBenefitClass, ee As Employee, ByVal AvailableFrom_ITEM As Long, ByVal AvailableTo_ITEM As Long)
  Dim benEE As IBenefitClass
  
  On Error GoTo SetAvaialbleRange_ERR
  
  If ben Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "SetAvaialbleRange", "Error setting the available from/to range.")
  
  Set benEE = ee
  
  If IsDate(benEE.value(ee_joined_db)) And (benEE.value(ee_joined_db) <> UNDATED) Then
    If benEE.value(ee_joined_db) < p11d32.Rates.value(TaxYearStart) Then
      ben.value(AvailableFrom_ITEM) = p11d32.Rates.value(TaxYearStart)
    Else
      ben.value(AvailableFrom_ITEM) = benEE.value(ee_joined_db)
    End If
  Else
    ben.value(AvailableFrom_ITEM) = p11d32.Rates.value(TaxYearStart)
  End If
  
  If IsDate(benEE.value(ee_left_db)) And (benEE.value(ee_left_db) <> UNDATED) Then
    ben.value(AvailableTo_ITEM) = benEE.value(ee_left_db)
  Else
    ben.value(AvailableTo_ITEM) = p11d32.Rates.value(TaxYearEnd)
  End If
  
  Call xSet("SetAvaialbleRange")
  
SetAvaialbleRange_END:
  Call xReturn("SetAvaialbleRange")
  Exit Sub
SetAvaialbleRange_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetAvaialbleRange", "Set Avaialble Range", "Error setting the available from/to for a benefit")
  Resume SetAvaialbleRange_END

End Sub
Public Function IsBenEmployee(ben As IBenefitClass) As Boolean
  Dim er As Employee
  
  On Error GoTo IsBenEmployee_END
  Set er = ben
  IsBenEmployee = True
  
IsBenEmployee_END:
  Exit Function
End Function
Private Function IsCDBEmployeeEx(ByVal ben As IBenefitClass) As Boolean
  If ben.BenefitClass <> BC_EMPLOYEE Then Exit Function
  If Not IsCBDEmployee(ben.value(ee_PersonnelNumber_db)) Then Exit Function
  IsCDBEmployeeEx = True
  
  
End Function

Public Function GetParentFromBenefit(ben As IBenefitClass, GPFB As GPFB_TYPE) As IBenefitClass
  Dim ee As Employee
  Dim er As Employer
  Dim l As Long
  Dim benX As IBenefitClass
  Dim i As Long, j As Long
  Dim ees As ObjectList
  Dim bFoundEmployeeForCompanyDefined As Boolean
  Dim bx As IBenefitClass
     
  On Error GoTo GetEmployeeFromBenefit_ERR
  
  Call xSet("GetEmployeeFromBenefit")
  
  l = 0
  Set benX = ben
TRY_AGAIN:
  l = l + 1
  
  'If l > 1 Then
  '
  '  If benX.CompanyDefined And Not CurrentForm Is F_CompanyDefined Then
  '    If Not benX.Parent Is Nothing Then
  '      If (TypeOf benX.Parent Is Employee) And p11d32.BringForward.Yes Then
  '        If IsCDBEmployeeEx(benX.Parent) Then
  '          GoTo STANDARD:
  '        End If
  '      End If
  '    End If
  '
  '    If (Not p11d32.CurrentEmployeeIsNothing) And GPFB = GPFB_TYPE.GPBF_EMPLOYEE Then
  '      Set benX = p11d32.CurrentEmployer.CurrentEmployee
  '      bFoundEmployeeForCompanyDefined = True
  '    End If
  '    'Set ees = p11d32.CurrentEmployer.employees
  '   'Set bx = benX.Parent
  '    'For i = 1 To ees.Count
  '    '  Set ee = ees(i)
  '    '  If (Not ee Is Nothing) Then
  '    '    For j = 1 To ee.benefits.Count
  '    '      If benX Is ee.benefits(j) Then
  '    '        Set benX = ee
  '   '        bFoundEmployeeForCompanyDefined = True
  '    '        Exit For
  '    '      End If
  '    '    Next
  '    '  End If
  '    'Next
  '    'need to loop through the employees and find the person that has this benefits
  '    If Not bFoundEmployeeForCompanyDefined Then Set benX = Nothing
  '  Else
'STANDARD:
  '    Set benX = benX.Parent
  '  End If
  'End If
  
  Set benX = benX.Parent
  
  If benX Is Nothing Then Call Err.Raise(ERR_PARENT_IS_NOTHING, "GetParentFromBenefit", "The benefits parent is nothing at " & l & " levels up the tree.")
  Select Case GPFB
    Case GPBF_EMPLOYEE
      Set ee = benX
      Set GetParentFromBenefit = benX
    Case GPBF_EMPLOYER
      Set er = benX
      Set GetParentFromBenefit = benX
    Case Else
      Call ECASE("The get parent from benefit type is invalid.")
  End Select
  
GetEmployeeFromBenefit_END:
  Call xReturn("GetEmployeeFromBenefit")
  Exit Function
GetEmployeeFromBenefit_ERR:
  Set GetParentFromBenefit = Nothing
  If l = 5 Or Err.Number = ERR_PARENT_IS_NOTHING Then
    Call ErrorMessage(ERR_ERROR, Err, "Get Parent From Benefit", "Get Parent From Benefit", "Error getting the parent from a benefit, iterated more than 5 times.")
    Resume GetEmployeeFromBenefit_END
  End If
  Resume TRY_AGAIN
  Resume
End Function
Public Function BenefitIsLoan(bc As BEN_CLASS) As Boolean
  BenefitIsLoan = (bc = BC_LOAN_OTHER_H)
End Function

Public Function IsBenOtherClass(bc As BEN_CLASS) As Boolean

  On Error GoTo IsBenOtherClass_Err
  Call xSet("IsBenOtherClass")

  If (bc = BC_CHAUFFEUR_OTHERO_N Or bc = BC_ENTERTAINMENT_N Or bc = BC_GENERAL_EXPENSES_BUSINESS_N Or bc = BC_INCOME_TAX_PAID_NOT_DEDUCTED_M Or bc = BC_PAYMENTS_ON_BEFALF_B Or bc = BC_OOTHER_N Or bc = BC_PRIVATE_MEDICAL_I Or bc = BC_CLASS_1A_M Or bc = BC_TAX_NOTIONAL_PAYMENTS_B Or bc = BC_TRAVEL_AND_SUBSISTENCE_N Or bc = BC_VOUCHERS_AND_CREDITCARDS_C Or bc = BC_NON_CLASS_1A_M) Then ' Or bc = BC_SHARES_M
    IsBenOtherClass = True
  End If

IsBenOtherClass_End:
  Call xReturn("IsBenOtherClass")
  Exit Function

IsBenOtherClass_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsBenOtherClass", "Is CDB Ben Class", "Error determining if the benefit can be company defined. Benefit class = " & bc)
  Resume IsBenOtherClass_End
End Function
Public Sub SetIRDesriptionInformation(bc As BEN_CLASS)
  Dim iFieldLen As Long
  Dim iBenITem As Long
  On Error GoTo err_err
  
  Call xSet("SetStandardBenItemsInformation")
  
  If (Not HasIRDescription(bc)) Then
    Call Err.Raise(ERR_BENCLASS_INVALID, "SetIRDesriptionInformation", "Ben class has no IR Description, can not set Standard Information")
  End If
  iBenITem = IRDescriptionBenItem(bc)
      
  'see IR book on MM submission
  iFieldLen = -1
  Select Case bc
    Case BC_ASSETSATDISPOSAL_L, BC_OOTHER_N, BC_PAYMENTS_ON_BEFALF_B, BC_ASSETSTRANSFERRED_A, BC_CHAUFFEUR_OTHERO_N
      iFieldLen = 30
    Case BC_NON_CLASS_1A_M
      iFieldLen = 33
    Case BC_CLASS_1A_M
      iFieldLen = 34
  End Select
  If iFieldLen = -1 Then
    Call Err.Raise(ERR_INVALID_BENEFIT_INDEX, "SetIRDesriptionInformation", "No Field Length was obtained for the IR Description as the BenClass was invalid")
  End If
  
  p11d32.BenDataLinkDataType(bc, iBenITem) = TYPE_STR
  p11d32.BenDataLinkMMRequired(bc, iBenITem) = True
  
  p11d32.BenDataLinkMMFieldSize(bc, iBenITem) = iFieldLen
  p11d32.BenDataLinkUDMDescription(bc, iBenITem) = S_UDM_IR_DESCRIPTION
  
  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  Call SetStandardBenItemsMMRequired(bc)
  
err_end:
  Call xReturn("SetStandardBenItemsInformation")
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetIRDesriptionInformation", "Set IR description information", "Error setting the IR description information.")
  Resume err_end
  Resume
End Sub
Public Sub SetStandardBenItemsInformation(bc As BEN_CLASS, ben As IBenefitClass)
  On Error GoTo SetStandardBenItemsInformation_ERR
  
  Call xSet("SetStandardBenItemsInformation")
  p11d32.BenDataLinkBenfitTable(bc) = ben.TABLE
  Call SetStandardBenItemsDataTypes(bc)
  Call SetStandardBenItemsMMFieldSize(bc)
  Call SetStandardBenItemsMMRequired(bc)
  Call SetStandardBenItemsUDMData(bc)
  
SetStandardBenItemsInformation_END:
  Call xReturn("SetStandardBenItemsInformation")
  Exit Sub
SetStandardBenItemsInformation_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetStandardBenItemsInformation", "Set Standard BenItems Information", "Error setting the standard benitems information.")
  Resume SetStandardBenItemsInformation_END
  Resume
End Sub
Public Function GetAlignment(dt As DATABASE_FIELD_TYPES) As AlignConstants
  If dt = TYPE_STR Then
    GetAlignment = vbAlignLeft
  Else
    GetAlignment = vbAlignRight
  End If
End Function

Public Sub SetCalcDefaultsStandard(ben As IBenefitClass)
  If ben Is Nothing Then
    ECASE ("ben is nothing in SetCalcDefaultsStandard")
  Else
    ben.value(ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE) = False
    ben.value(ITEM_VALUE_NON_OPRA) = S_ERROR
    ben.value(ITEM_BENEFIT) = S_ERROR
    ben.value(ITEM_BENEFIT_REPORTABLE) = False
    ben.value(ITEM_ERROR) = ""
  End If
  
End Sub

Public Sub CalculateOpRAValue(ben As IBenefitClass, Optional forceIgnoreAmountForgone As Boolean = False, Optional iITEM_VALUE_NON_OPRA As Integer = ITEM_VALUE_NON_OPRA, Optional iITEM_OPRA_AMOUNT_FOREGONE As Integer = ITEM_OPRA_AMOUNT_FOREGONE, Optional iITEM_VALUE As Integer = ITEM_VALUE, Optional iITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE As Integer = ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE)
  If (ben.value(iITEM_VALUE_NON_OPRA) < ben.value(iITEM_OPRA_AMOUNT_FOREGONE)) And Not forceIgnoreAmountForgone Then
    ben.value(iITEM_VALUE) = ben.value(iITEM_OPRA_AMOUNT_FOREGONE)
    ben.value(iITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE) = True
  Else
    ben.value(iITEM_VALUE) = ben.value(iITEM_VALUE_NON_OPRA)
    ben.value(iITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE) = False
  End If
End Sub

Public Function GetBenListItem(ibf As IBenefitForm2, ByVal lBenIndex As Long) As ListItem
  Dim lsti As listitems
  Dim li As ListItem
  
  On Error GoTo GetBenListItem_ERR
  Call xSet("GetBenListItem")
  
  If ibf Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "GetBenListItem", "The IBenefitForm is nothing")
  
  Set lsti = ibf.lv.listitems
  For Each li In lsti
    If li.Tag = lBenIndex Then
      Set GetBenListItem = li
      Exit For
    End If
  Next
  
GetBenListItem_END:
  Call xReturn("GetBenListItem")
  Exit Function
GetBenListItem_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetBenListItem", "Get Ben List Item", "Error getting the benefits list item, benindex = " & lBenIndex & ".")
  Resume GetBenListItem_END
End Function
Public Function CorrectBenValue(bc As BEN_CLASS, Item As Long, value As Variant) As Variant
  Dim v As Variant
  
  On Error GoTo CorrectBenValue_ERR
    
  If (value = UNDATED) Then
    CorrectBenValue = value
  Else
    v = GetTypedValue(value, p11d32.BenDataLinkDataType(bc, Item))
    If (Item = ITEM_MADEGOOD) Then
      If IsNumeric(v) Then
        If (v < 0) Then
          v = v * -1
        End If
      End If
    End If
    CorrectBenValue = v
  End If
  
  
CorrectBenValue_END:
   Exit Function
CorrectBenValue_ERR:
  CorrectBenValue = value
  Resume CorrectBenValue_END
End Function
Public Function DateValReadToScreen(ByVal v As Variant) As String
  Dim s As String
  Dim sDay As String
  Dim sMonth As String
    
  If (IsDate(v)) Then
    sDay = DatePart("d", v)
    If (Len(sDay) < 2) Then sDay = "0" + sDay
    sMonth = DatePart("m", v)
    If (Len(sMonth) < 2) Then sMonth = "0" + sMonth
    s = sDay & "/" & sMonth & "/" & DatePart("yyyy", v)
  Else
    s = DateStringEx(v, v)
  End If
  
  DateValReadToScreen = s
End Function
Public Function DateValReadToScreenOnlyValidDates(ByVal v As Variant) As String
  Dim s As String
    
  s = ""

  If (IsDate(v)) Then
    If (v <> UNDATED) Then
      s = DateValReadToScreen(v)
    End If
  End If
  
  DateValReadToScreenOnlyValidDates = s
End Function
Public Sub SetStandardBenItemsMMFieldSize(bc As BEN_CLASS)
  On Error GoTo SetStandardBenItemsMMFieldSize_ERR
  
  Call xSet("SetStandardBenItemsMMFieldSize")

  p11d32.BenDataLinkMMFieldSize(bc, ITEM_BENEFIT) = MM_SFS_DATA
  p11d32.BenDataLinkMMFieldSize(bc, ITEM_MADEGOOD_NET) = MM_SFS_DATA
  p11d32.BenDataLinkMMFieldSize(bc, ITEM_VALUE) = MM_SFS_DATA
  p11d32.BenDataLinkMMFieldSize(bc, ITEM_DESC) = MM_SFS_DESCRIPTION
  
  
SetStandardBenItemsMMFieldSize_END:
  Call xReturn("SetStandardBenItemsMMFieldSize")
  Exit Sub
SetStandardBenItemsMMFieldSize_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetStandardBenItemsMMFieldSize", "Set Standard BenItems Information", "Error setting the standard benitems information.")
  Resume SetStandardBenItemsMMFieldSize_END
  
End Sub
Public Sub SetStandardBenItemsMMRequired(bc As BEN_CLASS)
  On Error GoTo SetStandardBenItemsMMFieldSize_ERR
  
  Call xSet("SetStandardBenItemsMMFieldSize")

  p11d32.BenDataLinkMMRequired(bc, ITEM_DESC) = True
  
SetStandardBenItemsMMFieldSize_END:
  Call xReturn("SetStandardBenItemsMMFieldSize")
  Exit Sub
SetStandardBenItemsMMFieldSize_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetStandardBenItemsMMFieldSize", "Set Standard BenItems Information", "Error setting the standard benitems information.")
  Resume SetStandardBenItemsMMFieldSize_END
  
End Sub

Public Sub SetStandardBenItemsDataTypes(ByVal bc As BEN_CLASS, Optional bOPRAFields As Boolean = True)
  
  On Error GoTo SetStandardBenItemsDataTypes_ERR
  
  Call xSet("SetStandardBenItemsDataTypes")

  p11d32.BenDataLinkDataType(bc, ITEM_VALUE) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, ITEM_MADEGOOD) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, ITEM_DESC) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, ITEM_BENEFIT) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, ITEM_MADEGOOD_NET) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, ITEM_ACTUALAMOUNTMADEGOOD) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, ITEM_BENEFIT_REPORTABLE) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, ITEM_UDM_BENEFIT_TITLE) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, ITEM_BOX_NUMBER) = TYPE_STR
  p11d32.BenDataLinkDataType(bc, ITEM_NIC_CLASS1A_BENEFIT) = TYPE_DOUBLE
  p11d32.BenDataLinkDataType(bc, ITEM_BENEFIT_SUBJECT_TO_CLASS1A) = TYPE_LONG
  p11d32.BenDataLinkDataType(bc, ITEM_CLASS1A_ADJUSTMENT) = TYPE_LONG
  
  p11d32.BenDataLinkDataType(bc, ITEM_MADEGOOD_IS_TAXDEDUCTED) = TYPE_BOOL
  p11d32.BenDataLinkDataType(bc, ITEM_NIC_CLASS1A_ABLE) = TYPE_BOOL
  
  p11d32.BenDataLinkDataType(bc, ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE) = TYPE_BOOL
  
  'certain benefits do not have OpRa
  If IsOpRABenefitClass(bc) Then
    p11d32.BenDataLinkDataType(bc, ITEM_OPRA_AMOUNT_FOREGONE) = TYPE_LONG
    p11d32.BenDataLinkDataType(bc, ITEM_VALUE_NON_OPRA) = TYPE_LONG
    p11d32.BenDataLinkDataType(bc, ITEM_ERROR) = TYPE_STR
  End If
  
  
  
SetStandardBenItemsDataTypes_END:
  Call xReturn("SetStandardBenItemsDataTypes")
  Exit Sub
SetStandardBenItemsDataTypes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetStandardBenItemsDataTypes", "Set Standard BenItems Information", "Error setting the standard benitems information.")
  Resume SetStandardBenItemsDataTypes_END
  Resume

End Sub

Public Function IsOpRABenefitClassDataBase(ByVal bc As BEN_CLASS) As Boolean
  IsOpRABenefitClassDataBase = IsOpRABenefitClass(bc) And bc <> BEN_CLASS.BC_SHAREDVAN_G And bc <> BEN_CLASS.BC_SHAREDVANs_G And bc <> BEN_CLASS.BC_nonSHAREDVAN_G And bc <> BEN_CLASS.BC_NONSHAREDVANS_G
  
  
End Function


Public Function IsOpRABenefitClass(ByVal bc As BEN_CLASS) As Boolean
  IsOpRABenefitClass = Not ((bc = BEN_CLASS.BC_EMPLOYEE_CAR_E Or bc = BEN_CLASS.BC_QUALIFYING_RELOCATION_J Or bc = BEN_CLASS.BC_NON_QUALIFYING_RELOCATION_N Or bc = BEN_CLASS.BC_PAYMENTS_ON_BEFALF_B Or bc = BEN_CLASS.BC_TAX_NOTIONAL_PAYMENTS_B))
  
End Function
Public Sub Class1AAdjustment(ByVal ben As IBenefitClass)
  'error handler
  'if ben is nothing
  If ben.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) Then ben.value(ITEM_CLASS1A_ADJUSTMENT) = ben.value(ITEM_MADEGOOD_NET)
  
End Sub

Public Sub StandardWriteData(ben As IBenefitClass, rs As Recordset, Optional opraFields As Boolean = True)

  rs.Fields("MadeGoodIsTaxDeducted").value = ben.value(ITEM_MADEGOOD_IS_TAXDEDUCTED)
  If (Not p11d32.BringForward.Yes) Then
    If (opraFields) Then
      rs.Fields(S_DB_FIELD_OPRA_AMOUNT_FOREGONE).value = ben.value(ITEM_OPRA_AMOUNT_FOREGONE)
    End If
  End If
  
  
  
End Sub
Public Function IRDescription(ben As IBenefitClass) As String
  Dim iBenITem As Long
  iBenITem = IRDescriptionBenItem(ben.BenefitClass)
  If iBenITem <> ITEM_NONE Then
    IRDescription = ben.value(iBenITem)
  Else
    Call Err.Raise(ERR_BEN_INCORRECT, "IRDescription", "Can not get the IR description as the benefit does not have one")
  End If

End Function
Public Function IRDescriptionBenItem(bc As BEN_CLASS) As Long
  Select Case bc
    Case BC_NON_CLASS_1A_M, BC_CLASS_1A_M, BC_OOTHER_N, BC_PAYMENTS_ON_BEFALF_B, BC_CHAUFFEUR_OTHERO_N
      IRDescriptionBenItem = oth_IRDesc_db
    Case BC_ASSETSTRANSFERRED_A
      IRDescriptionBenItem = trans_IRDesc_db
    Case BC_ASSETSATDISPOSAL_L
      IRDescriptionBenItem = AssetsAtDisposal_IRDesc_db
    Case Else
      IRDescriptionBenItem = ITEM_NONE
  End Select
End Function
Public Function HasIRDescription(bc As BEN_CLASS) As Boolean
  HasIRDescription = IRDescriptionBenItem(bc) <> ITEM_NONE
End Function

Public Sub StandardReadData(ben As IBenefitClass, Optional rs As Recordset = Nothing, Optional ByVal opraFields As Boolean = True)
  Dim Rates As Rates
  Dim bc As BEN_CLASS
  
  If ben Is Nothing Then
    ECASE ("ben is nothing in StandardReadData")
  Else
    Set Rates = p11d32.Rates
    bc = ben.BenefitClass
    If bc > 0 Then
      ben.value(ITEM_BOX_NUMBER) = Rates.BenClassTo(bc, BCT_HMIT_BOX_NUMBER)
      ben.value(ITEM_UDM_BENEFIT_TITLE) = Rates.BenClassTo(bc, BCT_UDM_BENEFIT_TITLE)
      ben.value(ITEM_NIC_CLASS1A_ABLE) = Rates.BenClassTo(bc, BCT_CLASS1A_ABLE)
      
      If (IsOpRABenefitClassDataBase(ben.BenefitClass)) And opraFields Then
        
        Call OPRAReadDB(ben, rs)
      End If
      
      
      If p11d32.BringForward.Yes Then
        ben.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) = False
      Else
        If Not rs Is Nothing Then
          ben.value(ITEM_MADEGOOD_IS_TAXDEDUCTED) = rs.Fields("MadeGoodIsTaxDeducted")
        End If
      End If
    Else
      Call Err.Raise(ERR_BENCLASS_INVALID, ErrorSource(Err, "StandardReadData"), "Benefit class is 0")
    End If
  End If
End Sub
'cad loan fix
Public Sub BenCalcNIC(ben As IBenefitClass, Optional INCV As BASE_ITEMS = ITEM_BENEFIT_SUBJECT_TO_CLASS1A, Optional INCB As BASE_ITEMS = ITEM_NIC_CLASS1A_BENEFIT, Optional IV As BASE_ITEMS = ITEM_VALUE, Optional IB As BASE_ITEMS = ITEM_BENEFIT, Optional IC1AA As BASE_ITEMS = ITEM_CLASS1A_ADJUSTMENT, Optional IMGITD As BASE_ITEMS = ITEM_MADEGOOD_IS_TAXDEDUCTED, Optional IMGN As BASE_ITEMS = ITEM_MADEGOOD_NET, Optional INCA As BASE_ITEMS = ITEM_NIC_CLASS1A_ABLE)
  Dim d As Double
  Dim benEmployee As IBenefitClass
  
  On Error GoTo BenCalcNIC_ERR
  
  If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenCalcNIC", "The benefit is nothing.")
  
  ben.value(IC1AA) = 0
  ben.value(INCB) = 0
  ben.value(INCV) = 0
  If CheckIfNotSubjectToClass1A(ben, INCA) Then GoTo BenCalcNIC_END
  
  ben.value(INCV) = ben.value(IB) 'benefit subject to class 1A = benefit
  
  Set benEmployee = GetParentFromBenefit(ben, GPBF_EMPLOYEE)
  If benEmployee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenCalcNIC", "No Employee exists")
  If Not benEmployee.value(ee_Class1AEmployeeIsNotSubjectTo_db) Then
    'this is assigning the whole benefit to the be that which is subject to class 1A
    ben.value(INCV) = ben.value(IB)
    d = p11d32.Rates.value(carNICRate)
    'ben.value(INCB) = ben.value(IB) * d
    If ben.value(IMGITD) Then
      ben.value(IC1AA) = ben.value(IMGN)
    End If
    ben.value(INCV) = ben.value(INCV) + ben.value(IC1AA)
    ben.value(INCB) = ben.value(INCV) * d
  Else
    'if expat the actual benefit subject to class 1A is zero
    ben.value(INCV) = 0
  End If

BenCalcNIC_END:
  Exit Sub
BenCalcNIC_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BenCalcNIC", "Ben Calc NIC", "Error in calculating NIC for a benefit.")
  Resume BenCalcNIC_END
End Sub


Public Sub SetStandardBenItemsUDMData(ByVal bc As BEN_CLASS, Optional sClass1ANICFieldNameSuffix As String = "") ', Optional bOPRAFields = True)
  
  On Error GoTo SetStandardBenItemsDataTypes_ERR
  
  Call xSet("SetStandardBenItemsDataTypes")

  With p11d32
    .BenDataLinkUDMDescription(bc, ITEM_VALUE) = S_UDM_VALUE
    .BenDataLinkUDMDescription(bc, ITEM_MADEGOOD_NET) = S_UDM_MADE_GOOD_NET
    .BenDataLinkUDMDescription(bc, ITEM_DESC) = S_UDM_DESCRIPTION
    .BenDataLinkUDMDescription(bc, ITEM_BENEFIT) = S_UDM_BENEFIT
    .BenDataLinkUDMDescription(bc, ITEM_UDM_BENEFIT_TITLE) = S_UDM_BENEFIT_TITLE
    .BenDataLinkUDMDescription(bc, ITEM_BOX_NUMBER) = S_UDM_BOX_NUMBER
    .BenDataLinkUDMDescription(bc, ITEM_NIC_CLASS1A_ABLE) = S_UDM_NIC_CLASS1A_ABLE
    .BenDataLinkUDMDescription(bc, ITEM_NIC_CLASS1A_BENEFIT) = S_UDM_NIC_CLASS1A_BENEFIT & sClass1ANICFieldNameSuffix
    .BenDataLinkUDMDescription(bc, ITEM_MADEGOOD_IS_TAXDEDUCTED) = S_UDM_NIC_AMOUNT_MADEGOOD_TAXDEDUCTED
    .BenDataLinkUDMDescription(bc, ITEM_BENEFIT_SUBJECT_TO_CLASS1A) = S_UDM_IR_BENEFIT_SUBJECT_TO_CLASS1A
    
    
    If IsOpRABenefitClass(bc) Then
      .BenDataLinkUDMDescription(bc, ITEM_VALUE_NON_OPRA) = S_UDM_VALUE_NON_OPRA
      .BenDataLinkUDMDescription(bc, ITEM_OPRA_AMOUNT_FOREGONE) = S_UDM_OPRA_AMOUNT_FOREGONE
      .BenDataLinkUDMDescription(bc, ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE) = S_UDM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
      
    End If
  
    
             
  
  End With
  
  Call AbacusUDMData(bc)

SetStandardBenItemsDataTypes_END:
  Call xReturn("SetStandardBenItemsDataTypes")
  Exit Sub
SetStandardBenItemsDataTypes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetStandardBenItemsDataTypes", "Set Standard BenItems Information", "Error setting the standard benitems information.")
  Resume SetStandardBenItemsDataTypes_END
  Resume

End Sub

Public Sub AbacusUDMData(ByVal bc As BEN_CLASS)
  
  On Error GoTo SetStandardBenItemsDataTypes_ERR
  Call xSet("AbacusUDMData")
  If p11d32.ReportPrint.AbacusUDM Then
    p11d32.BenDataLinkUDMDescription(bc, ITEM_MADEGOOD) = S_UDM_MADE_GOOD_GROSS
  End If
  
SetStandardBenItemsDataTypes_END:
  Call xReturn("AbacusUDMData")
  Exit Sub
  
SetStandardBenItemsDataTypes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AbacusUDMData", "Set Standard BenItems Information", "Adding Abacus specific UDM data.")
  Resume SetStandardBenItemsDataTypes_END
  Resume
End Sub

Public Function InvalidFields(ifrm As IFrmGeneral, Optional ctlContainer As Object = Nothing) As Long
  Dim bFoundFirstInvalidField As Boolean
  Dim frm As Form, tcsvt As Control
  
  On Error GoTo InvalidFields_Err
  Call xSet("InvalidFields")
  
  Set ifrm.InvalidVT = Nothing
  Set frm = ifrm
  
  For Each tcsvt In frm.Controls
    If TypeOf tcsvt Is ValText Or TypeOf tcsvt Is ValCombo Then
      If Not (ctlContainer Is Nothing) Then
        If Not tcsvt.Container Is ctlContainer Then GoTo skip_field  ' no check
      End If
      If tcsvt.FieldInvalid Then
        If p11d32.DisplayInvalidFields Then
          MsgBox ("Invalid Field" & vbCrLf & vbCrLf & "Control = " & tcsvt.Name & vbCrLf & "Index = " & IIf(ControlIndex(tcsvt) = -1, "None", ControlIndex(tcsvt)) & vbCrLf & "Parent = " & tcsvt.Parent.Name)
        End If
        Call GetFirstInvalidVT(ifrm, tcsvt, bFoundFirstInvalidField)
        InvalidFields = InvalidFields + 1
      End If
    End If
    
skip_field:
  Next tcsvt
    
InvalidFields_End:
  Call xReturn("InvalidFields")
  Exit Function

InvalidFields_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InvalidFields", "Invalid Fields", "Error finding the invalid fields on a form.")
  Resume InvalidFields_End
  Resume
End Function
Private Function ControlIndex(c As Control) As Long
  On Error GoTo ControlIndex_ERR
  
  ControlIndex = -1
  ControlIndex = c.Index
  
ControlIndex_END:
  Exit Function
ControlIndex_ERR:
  Resume ControlIndex_END
End Function
Private Function GetFirstInvalidVT(ifrm As IFrmGeneral, vt As Control, FoundFirst As Boolean) As Long
  If vt.FieldInvalid And Not FoundFirst Then
    Set ifrm.InvalidVT = vt
  End If
End Function
Public Sub AddBenefitHelperSub(ben As IBenefitClass, ibf As IBenefitForm2, ByVal lBenefitIndex As Long)
  
  Call SelectBenefitByListItem(ibf, ibf.lv.listitems(ibf.BenefitToListView(ben, lBenefitIndex)))
End Sub
Public Function AddBenefitHelper(ibf As IBenefitForm2, ben As IBenefitClass, Optional bSetDefaults As Boolean = True) As Boolean

  On Error GoTo AddBenefitHelper_Err
  
  Call xSet("AddBenefitHelper")
  
  ben.BenefitClass = ibf.benclass
  Set ben.Parent = p11d32.CurrentEmployer.CurrentEmployee
  If bSetDefaults Then Call ibf.AddBenefitSetDefaults(ben)
  
  ben.ReadFromDB = True
  ben.Dirty = True
  
  Call AddBenefitHelperSub(ben, ibf, p11d32.CurrentEmployer.CurrentEmployee.benefits.Add(ben))
AddBenefitHelper_End:
  Call xReturn("AddBenefitHelper")
  Exit Function
AddBenefitHelper_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefitHelper", "Add Benefit Helper", "Unable to add a benefit") ' cd SENSIBLE ERROR
  Resume AddBenefitHelper_End
  Resume
End Function
Public Sub RemoveCDBAssignment(rs As Recordset, ee As Employee, sBenCode As String)
  On Error GoTo RemoveCDBAssignment_ERR
  
  If ee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "RemoveCDBAssignment", "Employee is nothing.")
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "RemoveCDBAssignment", "Recordset is ntohing.")
  
  rs.FindFirst ("P_Num = " & StrSQL(ee.PersonnelNumber) & " AND BenCode = " & StrSQL(sBenCode))
  If Not rs.NoMatch Then
    rs.Delete
  Else
    Call Err.Raise(ERR_BEN_IS_NOTHING, "RemoveCDBAddignment", "Failed to find the Company defined benefit with code=" & sBenCode & " and employee number=" & ee.PersonnelNumber)
  End If


      
RemoveCDBAssignment_END:
  Exit Sub
RemoveCDBAssignment_ERR:
  Call Err.Raise(ERR_BEN_IS_NOTHING, ErrorSource(Err, "RemoveCDBAssignment"), Err.Description)
  Resume RemoveCDBAssignment_END
End Sub
Public Function BenefitToScreenHelper(ibf As IBenefitForm2, ByVal BenefitIndex As Long, ByVal UpdateBenefit As Boolean) As Boolean
  Dim ben As IBenefitClass

  On Error GoTo BenefitToScreenHelper_Err
  Call xSet("BenefitToScreenHelper")
  
  If BenefitIndex <> -1 Then
    Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(BenefitIndex)
    If Not ibf.ValididateBenefit(ben) Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreenHelper", "Benefit To Screen Helper", "Invalid benefit type.")
    Set ibf.benefit = ben
    Call ibf.BenefitOn
  Else
    
    Set ibf.benefit = Nothing
    Call ibf.BenefitOff
  End If
  
  
  Call SetBenefitFormState(ibf)
  BenefitToScreenHelper = True
  
BenefitToScreenHelper_End:
  Set ben = Nothing
  Call xReturn("BenefitToScreenHelper")
  Exit Function
BenefitToScreenHelper_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreenHelper", "Benefit To Screen Helper", "Unable to place the chosen benefit to the screen. Benefit index = " & BenefitIndex & ".")
  Resume BenefitToScreenHelper_End
  Resume
End Function
Public Function BenOtherWhereClause(bc As BEN_CLASS) As String
  BenOtherWhereClause = " WHERE " & BenOtherWhereClauseSub(bc)
End Function
Public Function BenOtherWhereClauseSub(bc As BEN_CLASS) As String
  BenOtherWhereClauseSub = "(T_BenOther.UDBCode=" & StrSQL(p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING)) & ")" & _
                        " AND (T_BenOther.Category=" & StrSQL(p11d32.Rates.BenClassTo(bc, BCT_DBCATEGORY)) & ") AND (T_BenOther.Class=" & StrSQL(p11d32.Rates.BenClassTo(bc, BCT_DBCLASS)) & ")"
End Function
Public Sub BringForwardDatesWrite(ben As IBenefitClass, itemFrom As Long, itemTo As Long, rs As Recordset, sFromField As String, sToField As String)
  On Error GoTo BringForwardDates_ERR
  
  'simple function to set the dates correctly for bring forward Availablefrom/to
  
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BringForwardDatesWrite", "Recordset is nothing.")
  If ben Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BringForwardDatesWrite", "Benefit is nothing.")
  
  
  If p11d32.BringForward.Yes Then
   
    rs.Fields(sFromField) = p11d32.Rates.value(TaxYearStart)
    rs.Fields(sToField) = p11d32.Rates.value(TaxYearEnd)
  Else
    rs.Fields(sFromField) = ben.value(itemFrom)
    rs.Fields(sToField) = ben.value(itemTo)
  End If
  
  
  
BringForwardDates_END:
  Exit Sub
BringForwardDates_ERR:
  If Not ben Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "BringForwardDates", "Bring Forward Dates", "Error setting date fields for " & ben.Name & ", ben class = " & ben.BenefitClass & ".")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "BringForwardDates", "Bring Forward Dates", "Error setting date fields for a benefit")
  End If
  Resume BringForwardDates_END
  Resume
End Sub
Public Function StandardCanBringForward(ben As IBenefitClass, ByVal itemAvailableTo As Long)
  On Error GoTo StandardCanBringForward_ERR
  
  StandardCanBringForward = (ben.value(itemAvailableTo) = p11d32.Rates.value(LastTaxYearEnd))
  
StandardCanBringForward_END:
  Exit Function
StandardCanBringForward_ERR:
  If Not ben Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "StandardCanBringForward", "Standard Can Bring Forward", "Error determining if can bring benefit forward, benefit = " & ben.Name & ".")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "StandardCanBringForward", "Standard Can Bring Forward", "Error determining if can bring benefit forward.")
  End If
  Resume StandardCanBringForward_END
End Function

'cad loan fix
Public Function CheckIfNotSubjectToClass1A(ByVal ben As IBenefitClass, Optional ByVal INCA As BASE_ITEMS = ITEM_NIC_CLASS1A_ABLE) As Boolean
  Dim bc As BEN_CLASS

  On Error GoTo CheckIfNotSubjectToClass1A_ERR

  CheckIfNotSubjectToClass1A = True

  If Not ben.value(INCA) Then GoTo CheckIfNotSubjectToClass1A_END
  
  CheckIfNotSubjectToClass1A = False

CheckIfNotSubjectToClass1A_END:
  Exit Function
CheckIfNotSubjectToClass1A_ERR:
  If Not ben Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "CheckIfNotSubjectToClass1A", "Check If Subject To Class1A", "Error determining if employee is subject to class1A, benefit = " & ben.Name)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "CheckIfNotSubjectToClass1A", "Check If Subject To Class1A", "Error determining if employee is subject to class1A benefit")
  End If
  Resume CheckIfNotSubjectToClass1A_END
End Function

Public Function OPRAReadDB(ben As IBenefitClass, rs As Recordset, Optional iITEM_OPRA_AMOUNT_FOREGONE As Integer = ITEM_OPRA_AMOUNT_FOREGONE, Optional fieldNameAddition As String = "")
  Call OPRAReadWriteDB(False, ben, rs, iITEM_OPRA_AMOUNT_FOREGONE, fieldNameAddition)
End Function
Public Function OPRAWriteDB(ben As IBenefitClass, rs As Recordset, Optional iITEM_OPRA_AMOUNT_FOREGONE As Integer = ITEM_OPRA_AMOUNT_FOREGONE, Optional fieldNameAddition As String = "")
  Call OPRAReadWriteDB(True, ben, rs, iITEM_OPRA_AMOUNT_FOREGONE, fieldNameAddition)
End Function


Private Function OPRAReadWriteDB(ByVal writeDB As Boolean, ByVal ben As IBenefitClass, ByVal rs As Recordset, Optional ByVal iITEM_OPRA_AMOUNT_FOREGONE As Integer = ITEM_OPRA_AMOUNT_FOREGONE, Optional ByVal fieldNameAddition As String = "")
  Dim sFieldNameFull As String
  
  sFieldNameFull = S_DB_FIELD_OPRA_AMOUNT_FOREGONE & fieldNameAddition
  
  If (writeDB) Then
    If (p11d32.BringForward.Yes) Then
      rs.Fields(sFieldNameFull) = 0
    Else
      rs.Fields(sFieldNameFull) = ben.value(iITEM_OPRA_AMOUNT_FOREGONE)
    End If
  Else
    If (p11d32.BringForward.Yes) Or rs Is Nothing Then
      ben.value(iITEM_OPRA_AMOUNT_FOREGONE) = 0
    Else
      ben.value(iITEM_OPRA_AMOUNT_FOREGONE) = IIf(IsNull(rs.Fields(sFieldNameFull).value), 0, rs.Fields(sFieldNameFull).value)
    End If
  End If
End Function


