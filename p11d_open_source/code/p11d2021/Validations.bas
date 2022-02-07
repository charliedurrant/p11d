Attribute VB_Name = "Validations"
Option Explicit
Public Const L_MAX_LONG_FIELD_LENGTH = 35
Public Const L_MAX_LONG_STANDARD_FIELD_LENGTH = 30
Public Enum PAYE_FIELD_TYPES
  PAYE_TITLE
  PAYE_SURNAME
  PAYE_FORENAME
  PAYE_INDICATOR
  PAYE_DATE
  PAYE_AMOUNT
  
  'P46 Specific
  PAYE_P46_MAKEANDMODEL
  PAYE_P46_ENGINESIZE
  
End Enum
Public Sub CompanyCarAvailableDatesValid(ByVal ben As IBenefitClass)
   If ben.value(Car_AvailableFrom_db) > ben.value(Car_AvailableTo_db) Then
      Call Err.Raise(ERR_CAR_DATES_INVALID, "CompanyCarAvailableDatesValid", "'Available from' is greater than 'available to'")
   End If
End Sub


Public Function PAYEFieldValid(ByVal FieldData As String, FieldType As Integer) As Boolean
  Dim sData As String
  On Error GoTo PAYEFieldValid_ERR
  
  Call xSet("PAYEFieldValid")
  sData = FieldData
  'If FieldLen = 0 Then Call Err.Raise(ERR_FIELD_LEN_0, "PAYEFieldValid", "The fieldsize passed can not be 0.")
  
  Select Case FieldType
    Case PAYE_TITLE
        PAYEFieldValid = PAYETitleValid(sData)
    Case PAYE_SURNAME
        PAYEFieldValid = PAYESurnameValid(sData)
        If Not PAYEFieldValid Then p11d32.PAYEonline.AtLeastOneError = True
    Case PAYE_FORENAME
        PAYEFieldValid = PAYEForenameValid(sData)
        If Not PAYEFieldValid Then p11d32.PAYEonline.AtLeastOneError = True
    Case PAYE_INDICATOR
        'can be zero in some cases, therefore only look for greater than 3
        PAYEFieldValid = (Len(sData) <= 3)
    Case PAYE_DATE
    Case PAYE_AMOUNT
        PAYEFieldValid = PAYEAmountValid(sData)
        If Not PAYEFieldValid Then p11d32.PAYEonline.AtLeastOneError = True
    Case PAYE_P46_MAKEANDMODEL
        PAYEFieldValid = PAYEP46MakeAndModelValid(sData)
    Case PAYE_P46_ENGINESIZE
        PAYEFieldValid = PAYEP46EngineSizeValid(sData)
        If Not PAYEFieldValid Then p11d32.PAYEonline.AtLeastOneError = True
  End Select
  
    'IK 02/04/2004 - Do not set here. Title & MakeAndModel are warnings
    'If Not PAYEFieldValid Then p11d32.PAYEonline.AtLeastOneError = True
    
PAYEFieldValid_END:
  Call xReturn("PAYEFieldValid")
  Exit Function
PAYEFieldValid_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PAYEFieldValid", "Paye Field", "Error in Paye Fields Validation")
  Resume PAYEFieldValid_END
  Resume
End Function


Public Function PAYETitleValid(Title As String) As Boolean
  Dim sTxt As String
  Dim i As Long
  Dim lLen As Long
  
  On Error GoTo PAYETitleValid_Err
  
  sTxt = UCASE$(Title)
  lLen = Len(sTxt)
  
  PAYETitleValid = False
  
  
  If lLen > 4 Then GoTo PAYETitleValid_END
  
  For i = 1 To lLen
    'must be alpha or hyphen or apostrophe
        If Not ((IsAlphaStrEx(sTxt, i)) Or (IsHyphenStrEx(sTxt, i)) Or _
                (IsApostropheStrEx(sTxt, i))) Then
                    PAYETitleValid = False
                    GoTo PAYETitleValid_END
        End If
  Next
  
  PAYETitleValid = True
  

PAYETitleValid_END:
  Call xReturn("PAYETitleValid")
  Exit Function
PAYETitleValid_Err:
  Resume PAYETitleValid_END
  Resume

End Function

Public Function PAYESurnameValid(surname As String) As Boolean
  Dim sTxt As String
  Dim i As Long
  Dim lLen As Long
  
  On Error GoTo PAYESurnameValid_Err
  
  sTxt = UCASE$(surname)
  PAYESurnameValid = False

  lLen = Len(sTxt)
  'invalid field length
  If (lLen > L_MAX_LONG_FIELD_LENGTH) Or (lLen < 1) Then GoTo PAYESurnameValid_END
  
  
  For i = 1 To lLen
    If i < 2 Then
    'first char must be alpha
            If Not IsAlphaStrEx(sTxt, i) Then
                    PAYESurnameValid = False
                    GoTo PAYESurnameValid_END
            End If
    Else
    'other must be alpha or space or hyphen or apostrophe
        If Not ((IsAlphaStrEx(sTxt, i)) Or (IsSpaceStrEx(sTxt, i)) Or _
                (IsHyphenStrEx(sTxt, i)) Or (IsApostropheStrEx(sTxt, i))) Then
                    PAYESurnameValid = False
                    GoTo PAYESurnameValid_END
        End If
     End If
  Next
  
  PAYESurnameValid = True

PAYESurnameValid_END:
  Call xReturn("PAYESurnameValid")
  Exit Function
PAYESurnameValid_Err:
  Resume PAYESurnameValid_END
  Resume

End Function


Public Function PAYEForenameValid(Name As String) As Boolean
  Dim sTxt As String
  Dim i As Long
  Dim lLen As Long
  
  On Error GoTo PAYEForenameValid_Err
  
  sTxt = UCASE$(Name)
  lLen = Len(sTxt)
  
  PAYEForenameValid = False
  
  
  If lLen > L_MAX_LONG_FIELD_LENGTH Or lLen = 0 Then GoTo PAYEForenameValid_END
  
  For i = 1 To lLen
    'must be alpha or hyphen or apostrophe
        If Not ((IsAlphaStrEx(sTxt, i)) Or (IsHyphenStrEx(sTxt, i)) Or _
                (IsApostropheStrEx(sTxt, i))) Then
                    PAYEForenameValid = False
                    GoTo PAYEForenameValid_END
        End If
  Next
  
  PAYEForenameValid = True
  

PAYEForenameValid_END:
  Call xReturn("PAYEForenameValid")
  Exit Function
PAYEForenameValid_Err:
  Resume PAYEForenameValid_END
  Resume

End Function

Public Function PAYEAmountValid(amount As String) As Boolean
  Dim sTxt As String
  Dim i As Long
  Dim lLen As Long
  
  On Error GoTo PAYEAmountValid_Err
  
  sTxt = amount
  lLen = Len(sTxt)
  PAYEAmountValid = False
  
  'length can not be greater than 9 includes pence as we don't
  'and must be round figures,
  'and no negative numers,
  'hence every char numeric
  
  If lLen > 9 Then GoTo PAYEAmountValid_END
  
  For i = 1 To lLen
    'must be numeric
        If Not (IsNumericStrEx(sTxt, i)) Then
                    PAYEAmountValid = False
                    GoTo PAYEAmountValid_END
        End If
  Next
  
  PAYEAmountValid = True
  
PAYEAmountValid_END:
  Call xReturn("PAYEAmountValid")
  Exit Function
PAYEAmountValid_Err:
  Resume PAYEAmountValid_END
  Resume

End Function

Public Function PAYEP46MakeAndModelValid(amount As String) As Boolean
  Dim sTxt As String
  Dim i As Long
  Dim lLen As Long
  
  On Error GoTo PAYEP46MakeAndModelValid_Err
  
  sTxt = UCASE$(amount)
  lLen = Len(sTxt)
  PAYEP46MakeAndModelValid = False
  
  'length can not be greater than 35
  
  If lLen > 35 Then GoTo PAYEP46MakeAndModelValid_END
  
  PAYEP46MakeAndModelValid = True
  
PAYEP46MakeAndModelValid_END:
  Call xReturn("PAYEP46MakeAndModelValid")
  Exit Function
PAYEP46MakeAndModelValid_Err:
  Resume PAYEP46MakeAndModelValid_END
  Resume

End Function

Public Sub CheckCarListPrice(ByVal ben As IBenefitClass)
  If ben.value(car_ListPrice_db) < 1000 Or ben.value(car_ListPrice_db) > 999999999 Then
    Call Err.Raise(ERR_CAR_LIST_PRICE, , "Car list price (" & ben.value(car_ListPrice_db) & ") <1000 or >999,999,999")
  End If
End Sub
Public Sub CompanyCarCheckDateFuelAvaialbleTo(ByVal ben As IBenefitClass)
  If (ben.value(Car_FuelAvailableTo_calc) <> UNDATED) Then
    If ben.value(Car_FuelAvailableTo_calc) < ben.value(Car_AvailableFrom_db) Or (ben.value(Car_FuelAvailableTo_calc) > ben.value(Car_AvailableTo_db)) Then
      Call Err.Raise(ERR_CAR_FUEL_AVAILABLE_TO_INVALID, , "'Fuel avaiable to' is greater than the 'car available to' or less than the 'car avaialble from' for car " & ben.value(ITEM_DESC))
    End If
  End If
End Sub
Public Function PAYEP46EngineSizeValid(amount As String) As Boolean
  Dim sTxt As String
  Dim i As Long
  Dim lLen As Long
  
  On Error GoTo PAYEP46EngineSizeValid_Err
  
  sTxt = UCASE$(amount)
  lLen = Len(sTxt)
  PAYEP46EngineSizeValid = False
  
  'Must be atleast 1 char long and at most 4 chars long
  If (lLen > 4) Or (lLen < 1) Then GoTo PAYEP46EngineSizeValid_END
  
  For i = 1 To lLen
    'must be numeric
    If Not (IsNumericStrEx(sTxt, i)) Then
      PAYEP46EngineSizeValid = False
      GoTo PAYEP46EngineSizeValid_END
    End If
  Next i
  
  PAYEP46EngineSizeValid = True
  
PAYEP46EngineSizeValid_END:
  Call xReturn("PAYEP46EngineSizeValid")
  Exit Function
PAYEP46EngineSizeValid_Err:
  Resume PAYEP46EngineSizeValid_END
  Resume

End Function


Public Sub CheckEmployeePersonnelNumberForSpaces(ee As Employee)
  If InStr(1, ee.PersonnelNumber, " ", vbTextCompare) Then
    Call Err.Raise(ERR_SPACE_IN_PAYE_NUM, , "Space found in employee PAYE reference " & ee.PersonnelNumber)
    p11d32.PAYEonline.AtLeastOneError = True
  End If
End Sub

Public Sub CheckEmployeePersonnelNumberValidForEDI(ee As Employee)
  If Not ValidForEDI(ee.PersonnelNumber) Then
    Call Err.Raise(ERR_PERSONEL_NUMBER_CHANGED, "Personnel Number Valid for EDI", "Personnel number '" & ee.PersonnelNumber & "' contains invalid characters.")
    p11d32.PAYEonline.AtLeastOneError = True
  End If
End Sub

Public Function ValidForEDI(ByVal s As String) As Boolean
  Dim r As String
  
  ValidForEDI = Not (StrComp(s, MMStr(s), vbTextCompare) <> 0)
   
End Function

Public Function MMStr(v As Variant) As String
  Dim j As Long, k As Long
  Dim r As String
  
  On Error GoTo MMStr_ERR
  
  Call xSet("MMStr")
  
  MMStr = UCASE$(Trim$(v))
  MMStr = ReplaceString(MMStr, r, Chr(13))
  MMStr = ReplaceString(MMStr, r, Chr(10))
  
  Do While j < Len(MMStr)
    j = j + 1
    r = Mid$(MMStr, j, 1)
    k = AscB(r)
    If k < vbKeyA Or k > vbKeyZ Then
      If k < vbKey0 Or k > vbKey9 Then
        If InStr(1, "/-,.'&() ", r) = 0 Then
          MMStr = ReplaceString(MMStr, r, "")
        End If
      End If
    End If
  Loop
  
  MMStr = Trim$(MMStr)
  
MMStr_END:
  Call xReturn("MMStr")
  Exit Function
  
MMStr_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "MMStr", "MMStr", "Error removing the invalid magnetic media characters from " & v & ".")
  Resume MMStr_END
  Resume
  
End Function


Public Function ValidatePAYE(ByVal sPAYE As String) As Boolean
  'function which validates whether the PAYE data input is of the right format
  Dim istring_length As Long
  Dim icurrent_asc As Long, ilast_asc As Long
  Dim p0 As Long
  Dim s As String
  
  On Error GoTo err_err
  
  
  If (p11d32.PayeReferenceAnyFormat) Then
    ValidatePAYE = True
    GoTo err_end
  End If
  
  
  'trim and make lowercase
  sPAYE = LCase$(sPAYE)
  sPAYE = Trim$(sPAYE)
  
  'check that we had a /
  p0 = InStr(1, sPAYE, "/")
  If p0 < 1 Then GoTo err_end
    
  'validate tax office
  s = Trim$(Left$(sPAYE, p0 - 1))
  istring_length = Len(s)
  If istring_length > 3 Or istring_length = 0 Then GoTo err_end
  'check the left part is numeric
  For p0 = 1 To istring_length
    If Not IsDigit(Asc(Mid$(s, p0, 1))) Then GoTo err_end
  Next
  p0 = CInt(s)
  If (p0 = 0) Then GoTo err_end
  
  
  s = Trim$(Right$(sPAYE, Len(sPAYE) - InStr(sPAYE, "/")))
  istring_length = Len(s)
  If istring_length > 10 Or istring_length = 0 Then GoTo err_end
  
  For p0 = 1 To istring_length
    icurrent_asc = Asc(Mid$(s, p0, 1))
    'ensure first or last letter is not a / or -
    If (p0 = 1 Or p0 = istring_length) Then
      If IsInvalidChar(icurrent_asc) Then GoTo err_end
    End If
    'ensure the character is one of the permitted characters
    If (Not (IsDigit(icurrent_asc) Or IsAlpha(icurrent_asc) Or IsInvalidChar(icurrent_asc))) Then GoTo err_end
    
    'ensure backslashes or hyphons not next to each other
    If InvalidCharRepeat(icurrent_asc, ilast_asc) Then GoTo err_end
    ilast_asc = icurrent_asc
  Next
  
  ValidatePAYE = True
  
err_end:
  Exit Function
err_err:
  'Call MsgBox(Err.Description)
  Resume err_end
  Resume
End Function
Private Function InvalidCharRepeat(ByVal chCurrent As Long, ByVal chLast As Long) As Boolean
  If (IsInvalidChar(chCurrent)) Then
    InvalidCharRepeat = IsInvalidChar(chLast)
  End If
End Function
Private Function IsInvalidChar(ByVal ch As Long) As Boolean
  IsInvalidChar = (ch = L_CH_FORWARD_SLASH) Or (ch = L_CH_HYPHON) Or (ch = L_CH_SPACE)
End Function
Public Function TaxOfficeNumeral(ByVal sPAYEref As String) As String
  Dim i As Long
  
  i = InStr(1, sPAYEref, "/")
  
  If i > 1 Then
    TaxOfficeNumeral = Trim$(Left$(sPAYEref, i - 1))
  End If

End Function
Public Function EmployerRefNo(ByVal sPAYEref As String) As String
  Dim i As Long
  
  i = InStr(1, sPAYEref, "/")
  
  If i > 0 Then
    EmployerRefNo = Trim$(Mid$(sPAYEref, i + 1, Len(sPAYEref) - i))
  End If
  
End Function
Public Sub CheckDuplicateNINumbers(ey As Employer)
  
  Dim rs As Recordset
  Dim ben As IBenefitClass
  Dim ee As Employee
  Dim sPnumsToExclude As String
  Dim sAnd As String
  On Error GoTo CheckDuplicateNINumbers_ERR
  
  If ey Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CheckDuplicateNINumbers", "The employer is nothing.")
  Set ben = ey
  
  Set rs = ey.db.OpenRecordset("select P_Num from t_employees where NI in (SELECT NI  From T_Employees  GROUP BY NI  HAVING (Count(NI)>1)) and P_NUM <> " & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX), dbOpenSnapshot)
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CheckDuplicateNINumbers", "Recordset is nothing.")

  
  sPnumsToExclude = S_FIELD_PERSONEL_NUMBER & "<>" & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX)

  'this can fail as if there are too many people with no benefits then the access query will be too big
  'therefore removed from code
  'Do While Not rs.EOF
  '
  '  Set ee = ey.FindEmployee(rs.Fields("P_Num"))
  '  Call ee.LoadBenefits(TBL_ALLBENEFITS, False)
  '
  '  If Not ee.IterateBenefits(AnyReportable) Then
  '    If Len(sPnumsToExclude) Then sPnumsToExclude = sPnumsToExclude & " AND "
  '    sPnumsToExclude = sPnumsToExclude & "P_NUM<>" & StrSQL(rs.Fields("P_Num"))
  '  End If
  '  'Call Err.Raise(ERR_DUPLICATE_NI_NUMBERS, "CheckDuplicateNINumbers", rs.Fields("CountOfNI").value & " duplicates, NI number=" & rs.Fields("NI").value)
  '  rs.MoveNext
  'Loop
  
  Set rs = ey.db.OpenRecordset(sql.Queries(SELECT_DUPLICATE_NI_NUMBERS, sPnumsToExclude), dbOpenSnapshot)
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CheckDuplicateNINumbers", "Recordset is nothing.")
    
  Do While Not rs.EOF
    Call Err.Raise(ERR_DUPLICATE_NI_NUMBERS, "CheckDuplicateNINumbers", rs.Fields("CountOfNI").value & " duplicates, NI number=" & rs.Fields("NI").value)
    rs.MoveNext
  Loop
  
    
CheckDuplicateNINumbers_END:
  Exit Sub
CheckDuplicateNINumbers_ERR:
  If Err.Number = ERR_DUPLICATE_NI_NUMBERS Then
    Call ErrorMessage(ERR_ERROR, Err, "CheckDuplicateNINumbers", FilterMessageTitle(ben.value(employer_Name_db)), Err.Description)
    Resume Next
  Else
    Call ErrorMessage(ERR_ERROR, Err, "CheckDuplicateNINumbers", FilterMessageTitle(), "Error determining if too many invalid NI numbers.")
  End If
  Resume CheckDuplicateNINumbers_END
End Sub

Public Function FilterMessageTitle(Optional EmployerName As Variant, Optional EmployeeName As Variant, Optional HMITSectionString As Variant, Optional BenefitName As Variant, Optional BenfitFormCaption As Variant) As String
  If Not IsMissing(EmployerName) Then FilterMessageTitle = EmployerName
  FilterMessageTitle = FilterMessageTitle & ","
  If Not IsMissing(EmployeeName) Then FilterMessageTitle = FilterMessageTitle & EmployeeName
  FilterMessageTitle = FilterMessageTitle & ","
  If Not IsMissing(HMITSectionString) Then FilterMessageTitle = FilterMessageTitle & HMITSectionString
  FilterMessageTitle = FilterMessageTitle & ","
  If Not IsMissing(BenefitName) Then FilterMessageTitle = FilterMessageTitle & BenefitName
  FilterMessageTitle = FilterMessageTitle & ","
  If Not IsMissing(BenfitFormCaption) Then FilterMessageTitle = FilterMessageTitle & BenfitFormCaption
End Function


Public Sub CheckNIRatio(ByVal lEmployeesWithInvalidNI, ByVal lEmployeesAdded As Long)
  On Error GoTo CheckNIRatio_ERR
  
  Call xSet("CheckNIRatio")
  
  'if no employees added then exit
  If Not lEmployeesAdded > 0 Then GoTo CheckNIRatio_END
  
  If (CSng(lEmployeesWithInvalidNI) / CSng(lEmployeesAdded) * 100) > 5 Then
    Call Err.Raise(ERR_INVALID_NI_RATIO, "CheckNIRatio", "More than 5% of employees have invalid NI numbers.")
    p11d32.PAYEonline.AtLeastOneError = True
  End If
  
CheckNIRatio_END:
  Call xReturn("CheckNIRatio")
  Exit Sub
CheckNIRatio_ERR:
  
  If Err.Number = ERR_INVALID_NI_RATIO Then
    Call ErrorMessage(ERR_ERROR, Err, "CheckNIRatio", FilterMessageTitle(), Err.Description)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "CheckNIRatio", FilterMessageTitle(), "Error determining if too many invalid NI numbers.")
  End If
  Resume Next
End Sub





'Standard checking functions

Public Function IsAlphaStrEx(s As String, ByVal CharPos As Long) As Boolean
  Dim l As Long
  
  l = Asc(Mid$(UCASE(s), CharPos, 1))
  IsAlphaStrEx = (l >= 65) And (l <= 90)
  
End Function

Public Function IsNumericStrEx(s As String, ByVal CharPos As Long) As Boolean
  Dim l As Long
  
  l = Asc(Mid$(UCASE(s), CharPos, 1))
  IsNumericStrEx = (l >= 48) And (l <= 57)
  
End Function

Public Function IsSpaceStrEx(s As String, ByVal CharPos As Long) As Boolean
  Dim l As Long
  
  l = Asc(Mid$(UCASE(s), CharPos, 1))
  IsSpaceStrEx = (l = 32)
  
End Function

Public Function IsApostropheStrEx(s As String, ByVal CharPos As Long) As Boolean
  Dim l As Long
  
  l = Asc(Mid$(UCASE(s), CharPos, 1))
  IsApostropheStrEx = (l = 39)
  
End Function

Public Function IsHyphenStrEx(s As String, ByVal CharPos As Long) As Boolean
  Dim l As Long
  
  l = Asc(Mid$(UCASE(s), CharPos, 1))
  IsHyphenStrEx = (l = 45)
  
End Function
Public Function SumBenefit(ByRef Description As Variant, ByRef value As Variant, ByRef MadeGood As Variant, ByRef benefit As Variant, ey As IBenefitClass, ee As Employee, benefits As ObjectList, BenArr() As BEN_CLASS, Optional IRDesc As Variant) As Boolean
  Dim ben As IBenefitClass
  Dim i As Long, j As Long
  Dim lError As Long
  Dim sHMITSectionString As String
  Dim sBenefitFormCaption As String
  Dim bc As BEN_CLASS
  
  On Error GoTo SumBenefit_Err
  Call xSet("SumBenefit")
  
  j = 0
  value = 0
  MadeGood = 0
  benefit = 0
  Description = ""
  IRDesc = ""
  
  For i = 1 To benefits.Count
    Set ben = benefits(i)
    If Not (ben Is Nothing) Then
      bc = ben.BenefitClass
      If BenClassInArray(BenArr, bc) Then
        sHMITSectionString = p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING)
        sBenefitFormCaption = p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION)
        If CheckBen(ey, ee, ben) Then
          
          Call SumBenefitEx(j, value, MadeGood, benefit, Description, IRDesc, ben, OT_MAGENTIC_MEDIA)
        End If
      End If
     Set ben = Nothing
    End If
  Next i
  
  SumBenefit = CBool(j)
  
SumBenefit_End:
  Call xReturn("SumBenefit")
  Exit Function
SumBenefit_Err:
  lError = Err.Number
  If Not ben Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "SumBenefit", FilterMessageTitle(ey.Name, ee.PersonnelNumber, sHMITSectionString, ben.Name, sBenefitFormCaption), "Unable to sum benefits.")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "SumBenefit", FilterMessageTitle(ey.Name, ee.PersonnelNumber, , "Unknown"), "Unable to sum benefits.")
  End If
  SumBenefit = False
  Description = S_ERROR
  value = S_ERROR
  MadeGood = S_ERROR
  benefit = S_ERROR
  Resume SumBenefit_End
  Resume
End Function

Private Sub SumBenefitEx(jCount As Long, value As Variant, MadeGood As Variant, benefit As Variant, Description As Variant, IRDescription As Variant, ByVal ben As IBenefitClass, ByVal ot As OUTPUT_TYPE)
  Dim iBenITem As Long
  jCount = jCount + 1
  
  value = value + ben.value(ITEM_VALUE)
  MadeGood = MadeGood + ben.value(ITEM_MADEGOOD_NET)
  benefit = benefit + ben.value(ITEM_BENEFIT)
          
  If jCount = 2 Then
    Description = "Multiple items"
  ElseIf jCount = 1 Then
    Description = ben.value(ITEM_DESC)
  End If
          
  If HasIRDescription(ben.BenefitClass) Then
    If jCount = 2 Then
      IRDescription = "MULTIPLE"
      If (ot = OT_MAGENTIC_MEDIA Or ot = OT_PAYE_ONLINE) Then
        Description = ""
      End If
    ElseIf jCount = 1 Then
      iBenITem = IRDescriptionBenItem(ben.BenefitClass)
      IRDescription = ben.value(iBenITem)
      If (ot = OT_MAGENTIC_MEDIA Or ot = OT_PAYE_ONLINE) Then
        If (StrComp(IRDescription, S_IR_DESC_OTHER) <> 0) Then
          Description = ""
        End If
      End If
      
    End If
  End If
          
End Sub


Public Function ListViewAnyChecked(lv As ListView) As Boolean
  Dim i As Long
  For i = 1 To lv.listitems.Count
    If (lv.listitems(i).Checked) Then
      ListViewAnyChecked = True
      Exit Function
    End If
    
  Next
End Function
Private Function FirstTwoLettersValid(ByRef sNI As String) As Boolean
   Dim s As String
   Dim i As Long
   Const VALID_CODES As String = "AA, AB, AE, AH, AK, AL, AM, AP, AR, AS, AT, AW, AX, AY, AZ, BA , BB, BE, BH, BK, BL, BM, BT, CA , CB, CE, CH, CK, CL, CR, EA , EB, EE, EH, EK, EL, EM, EP, ER, ES, ET, EW, EX, EY, EZ, GY, HA , HB, HE, HH, HK, HL, HM, HP, HR, HS, HT, HW, HX, HY, HZ, JA , JB, JC, JE, JG, JH, JJ, JK, JL, JM, JN, JP, JR, JS, JT, JW, JX, JY, JZ, KA , KB, KC, KE, KH, KK, KL, KM, KP, KR, KS, KT, KW, KX, KY, KZ, LA , LB, LE, LH, LK, LL, LM, LP, LR, LS, LT, LW, LX, LY, LZ, MA , MW, MX, NA , NB, NE, NH, NL, NM, NP, NR, NS, NW, NX, NY, NZ, OA, OB, OE, OH, OK, OL, OM, OP, OR, OS, OX, PA , PB, PC, PE, PG, PH, PJ, PK, PL, PM, PN, PP, PR, PS, PT, PW, PX, PY, RA , RB, RE, RH, RK, RM, RP, RR, RS, RT, RW, RX, RY, RZ, SA , SB, SC, SE, SG, SH, SJ, SK, SL, SM, SN, SP, SR, SS, ST, SW, SX, SY, SZ, TA , TB, TE, TH, TK, TL, TM, TP, TR, TS, TT, TW, TX, TY, TZ, WA , WB, WE, WK, WL, WM, WP, YA , YB, YE, YH, YK, YL, YM, YP, YR, YS, YT, YW, YX, YY, YZ, ZA , ZB, ZE, ZH, ZK, ZL, ZM, ZP, ZR, ZS, ZT, ZW, ZX, ZY"
   
   s = Trim$(sNI)
   
   If (Len(s) < 2) Then
    Exit Function
   End If
  
   s = Left$(s, 2)
   s = UCASE$(s)
   'list found at in appendix at the end of the exb payeonline documentation
   i = InStr(1, VALID_CODES, s, vbBinaryCompare)
   FirstTwoLettersValid = i > 0
End Function
Public Function ValidateNI(ByVal sNI As String, ByVal bAllowTemporaryNumbers As Boolean) As NI_VALID
  Dim i As Long
  Dim lLen As Long
  Dim bIsTemporary As Boolean
  
  sNI = Trim$(UCASE$(sNI))
   'CAD check NI formats
  lLen = Len(sNI)
  ValidateNI = INVALID
  
  If lLen < 8 Then GoTo NINumberValid_END
  bIsTemporary = (StrComp("TN", Left$(sNI, 2), vbTextCompare) = 0)
  'temporary numbers invalid
  If Not bAllowTemporaryNumbers Then
    If bIsTemporary Then GoTo NINumberValid_END
  End If
  
  If IsAlphaStrEx(sNI, 1) Then
    If Len(sNI) = 8 And bIsTemporary Then
      GoTo NINumberValid_END
    End If
    'two types of vailid NI
    '2 alpha, 6 numeric, l alpha
    If IsNumeric(Mid$(sNI, 3, 6)) Then
      If Not bIsTemporary And Not (FirstTwoLettersValid(sNI)) Then
        GoTo NINumberValid_END
      End If
      For i = 1 To 9
        If i < 3 Then
          If Not IsAlphaStrEx(sNI, i) Then
            GoTo NINumberValid_END
          End If
        ElseIf i > 2 And i < 9 Then
          If Not IsNumericStrEx(sNI, i) Then
            GoTo NINumberValid_END
          End If
        ElseIf i > 8 And Len(sNI) > 8 Then
          If bIsTemporary Then
            If InStr(1, "FM", Mid(sNI, i, 1)) = 0 Then
              GoTo NINumberValid_END
            End If
          Else
            If InStr(1, "ABCD", Mid(sNI, i, 1)) = 0 Then
              GoTo NINumberValid_END
            End If
          End If
        End If
      Next
      ValidateNI = STANDARD
    End If
    '2 numeric, 1 alpha, 5 numeric  - expat
    'cad CORRECT FOR edIcAHR SET ON MIDLLE CHAR
  ElseIf lLen = 8 Then
    'not allowed for PAYE online therefore valid musdt not be TWO_number
    For i = 1 To 8
      Select Case i
        Case Is < 3
          If Not IsNumericStrEx(sNI, i) Then GoTo NINumberValid_END
        Case 3
          If Not IsAlphaStrEx(sNI, i) Then GoTo NINumberValid_END
        Case Is > 3
          If Not IsNumericStrEx(sNI, i) Then GoTo NINumberValid_END
      End Select
    Next
    ValidateNI = TWO_NUMBER
  End If

NINumberValid_END:
  

End Function

Private Function IsAtoD(ByVal ch As Long) As Boolean
  IsAtoD = (ch = Asc("A")) Or (ch = Asc("B")) Or (ch = Asc("C")) Or (ch = Asc("D"))
End Function

