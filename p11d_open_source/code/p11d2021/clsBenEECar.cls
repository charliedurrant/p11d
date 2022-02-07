VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBenEECar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Implements IBenefitClass

Public Enum EECarItems
  eecar_Item = ITEM_DESC
  eecar_GrossBenefit = ITEM_VALUE
  eecar_MadeGood = ITEM_MADEGOOD
  eecar_Benefit = ITEM_BENEFIT
  'Strings
  
  'off screen
  eecar_EmployeeNumber
  
  'combo boxes
  eecar_FPCS
  'textboxes
  eecar_EngineSize
  eecar_totalExtras
  eecar_AmountReceived
  'check boxes
  eecar_AlternativeMethod
  eecar_AverageIRRate
  eecar_UseCompanyCarScheme
  'extras dialog /text boxes
  'eecar_MadeGood 'standard item
  eecar_LumpSum
  eecar_HireCost
  eecar_HireCostMadeGood
  
  eecar_TotalMiles
    
    
  eecar_CCbandKey
 
  'backwards compatibility
  
  eecar_CompanyFPCSValue
  eecar_IRFPCSValue
  eecar_DifferanceBetweenCompanyAndIRFPCS
  eecar_SumOfIncome
  
  eecar_ercontrib
  eecar_carkey
  eecar_PreErBenefit
  eecar_OLDEECarKey
  eecar_PYValue
  eecar_TotalRelief
  eecar_TotalValue
  eecar_UDBCode
  'Dates
  eecar_calculated
  'Lastitem
  eecar_lastitem
End Enum

Public mileage As ObjectList

Private m_BenefitClass As benClass
Private m_readfromdb As Boolean
Private m_sbookmark As String
Private m_dirty As Boolean
Private m_dateedit As Date
Private m_parent As clsEmployee
Private m_InvalidFields As Long
Private m_EECarItems(0 To car_lastitem) As Variant
Private Const S_SCHEME1 = "Scheme 1"

Private Sub Class_Initialize()
  Set mileage = New ObjectList
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As benClass)
  m_BenefitClass = NewValue
End Property

Private Property Get IBenefitClass_BenefitClass() As benClass
  IBenefitClass_BenefitClass = BC_EECAR
End Property

Private Property Get IBenefitClass_BenefitSubClass() As Long
  IBenefitClass_BenefitSubClass = 0
End Property

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE "clsBenEECar_CompanyDefined"
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
'  ECASE "clsBenCar_CompanyDefined"
End Property

Private Property Get IBenefitClass_DateEdit() As Date
  IBenefitClass_DateEdit = m_dateedit
End Property

Private Property Let IBenefitClass_Dirty(RHS As Boolean)
  m_dirty = RHS
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_GetItem(ByVal Item As Long) As Variant
  IBenefitClass_GetItem = m_EECarItems(Item)
End Function

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property

Private Sub IBenefitClass_MakeDirty()
  m_dirty = True
End Sub

Private Property Get IBenefitClass_name() As String
  IBenefitClass_name = m_EECarItems(eecar_Item)
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_parent
End Property

Private Property Get IBenefitClass_PrintHeader() As String
  IBenefitClass_PrintHeader = "{x=0}{Arial=10,n}Employee Owned Car Allowance" & vbCrLf
End Property

Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_readfromdb = NewValue
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_readfromdb
End Property

Private Property Get IBenefitClass_Reference() As String
  ECASE "clsBenCar_Reference"
End Property

Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  m_sbookmark = NewValue
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = m_sbookmark
End Property

Private Function IBenefitClass_SetItem(ByVal Item As Long, value As Variant) As Boolean
  m_EECarItems(Item) = value
End Function

Private Sub IBenefitClass_Kill()
  Set m_parent = Nothing
End Sub


Private Function GetIrValue(Miles, EngineSize, AveIRRate) As Double
Dim I As Long
Dim icounter As Long
Dim DLowRate As Double, DHighrate As Double
Dim IRFPCSBandHighRate As Double
Dim IRFPCSBandLowRate As Double

'On Error GoTo clsBenEECar_getIRValue_Err
'Call xSet("clsBenEECar_GetIRValue")

'First select correct rates

'If AveIRRate Then
'  For i = 1 To 4
'    IRFPCSBandHighRate = rates.GetIRFPCSHighRate(i)
'    IRFPCSBandLowRate = rates.GetIRFPCSLowRate(i)
'    If i = L_FPCS_BAND2KEY Or i = L_FPCS_BAND3KEY Then
'      icounter = icounter + 1
'      DLowRate = DLowRate + IRFPCSBandLowRate
'      DHighrate = DHighrate + IRFPCSBandHighRate
'    End If
'  Next i
'  If icounter > 0 Then
'    DLowRate = DLowRate / icounter
'    DHighrate = DHighrate / icounter
'  End If
'Else
'  Select Case EngineSize
'    Case Is < L_FPCS_BAND1
'      i = L_FPCS_BAND1KEY
'    Case L_FPCS_BAND1 To (L_FPCS_BAND2 - 1)
'      i = L_FPCS_BAND2KEY
'    Case L_FPCS_BAND2 To (L_FPCS_BAND3 - 1)
'      i = L_FPCS_BAND3KEY
'    Case Else
'      i = L_FPCS_BAND4KEY
'  End Select
'  DLowRate = rates.GetIRFPCSLowRate(i)
'  DHighrate = rates.GetIRFPCSHighRate(i)
'End If
'
'
''Now work out Revenue Allowed figure
'If Miles <= 4000 Then
'  GetIrValue = Miles * DLowRate
'Else
'  GetIrValue = (4000 * DLowRate) + ((Miles - 4000) * DHighrate)
'End If
'
'clsBenEECar_GetIRValue_End:
'  Call xReturn("clsBenEECar_GetIRValue")
'  Exit Function
'clsBenEECar_getIRValue_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "clsBenEECar_GetIRValue", "ERR_CAR_CALCULATE", "Unable to correctly calculate the employee owned car benefit on the car.")
'  Resume clsBenEECar_GetIRValue_End

End Function


Private Function IBenefitClass_Calculate() As Variant
  Dim vMileageValue As Variant

  On Error GoTo clsBenEECar_Calculate_Err
  Call xSet("clsBenEECar_Calculate")
  
  m_EECarItems(eecar_CompanyFPCSValue) = S_ERROR
  m_EECarItems(eecar_IRFPCSValue) = S_ERROR
  m_EECarItems(eecar_SumOfIncome) = S_ERROR
  m_EECarItems(eecar_Benefit) = S_ERROR

  'simple case
  'get the Company FPCS value
  
  Call GetFPCSMileageValue(vMileageValue, m_EECarItems(eecar_TotalMiles), m_EECarItems(eecar_EngineSize), CLng(m_EECarItems(eecar_FPCS)), m_EECarItems(eecar_AverageIRRate))
  m_EECarItems(eecar_CompanyFPCSValue) = CLng((vMileageValue + 0.5))
  'get IR FPCS value
  vMileageValue = 0
  'get the FPCS IR value
  
  'has the Alternative method been ticked
  Call GetIRFPCSValue(vMileageValue)
  m_EECarItems(eecar_IRFPCSValue) = vMileageValue
  
  If m_EECarItems(eecar_UseCompanyCarScheme) Then
    m_EECarItems(eecar_AmountReceived) = m_EECarItems(eecar_CompanyFPCSValue)
    m_EECarItems(eecar_DifferanceBetweenCompanyAndIRFPCS) = m_EECarItems(eecar_IRFPCSValue) - m_EECarItems(eecar_CompanyFPCSValue)
  Else
    m_EECarItems(eecar_DifferanceBetweenCompanyAndIRFPCS) = 0
  End If
  
  m_EECarItems(eecar_SumOfIncome) = Max(0, m_EECarItems(eecar_AmountReceived) + m_EECarItems(eecar_totalExtras))
  m_EECarItems(eecar_Benefit) = Max(0, m_EECarItems(eecar_SumOfIncome) - m_EECarItems(eecar_IRFPCSValue))
  
  'calculate the other ones
  'If m_EECarItems(eecar_AlternativeMethod) = 0 Then
  Call CalcEmployeeCarsWithAlternativMethodFalse(Me)
    
  'End If
  
  IBenefitClass_Calculate = m_EECarItems(eecar_Benefit)
  
clsBenEECar_Calculate_End:
  Call xReturn("clsBenEECar_Calculate")
  Exit Function
clsBenEECar_Calculate_Err:
  'Call ErrorMessage(ERR_ERROR, Err, "clsBenCar_Calculate", "ERR_CAR_CALCULATE", "Unable to correctly calculate the company car benefit on the car.")
  IBenefitClass_Calculate = S_ERROR
  Resume clsBenEECar_Calculate_End
  Resume
End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim rs As Recordset
  Dim eecar As IBenefitClass
  Dim s As String
  Dim I As Long
  Dim FPCS As FPCS
  
  On Error GoTo clsBenEECar_readDB_Err
  Call xSet("clsBenEECar_readDB")
  
  If m_readfromdb Then GoTo clsBenEECar_readDB_End
  Set rs = m_parent.Parent.rsBenTables(TBL_EECAR)
  If Len(m_sbookmark) = 0 Then
    rs.FindFirst ("P_NUM = '" & m_parent.PersonelNo & "'")
    If Not rs.NoMatch Then
      m_sbookmark = rs.Bookmark
      I = I + 1
      rs.FindNext ("P_NUM = '" & m_parent.PersonelNo & "'")
      Do While Not rs.NoMatch
        I = I + 1
        Set eecar = New clsBenEECar
        eecar.RSBookMark = rs.Bookmark
        Set eecar.Parent = m_parent
        Call m_parent.benefits.Add(eecar)
        Set eecar = Nothing
        rs.FindNext ("P_NUM = '" & m_parent.PersonelNo & "'")
      Loop
    Else
      m_readfromdb = True
    End If
  End If
  
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    
    'general
    m_EECarItems(eecar_EmployeeNumber) = IIf(IsNull(rs.Fields(S_EMPLOYEE_NUM_FIELD)), "", rs.Fields(S_EMPLOYEE_NUM_FIELD))
    m_EECarItems(eecar_PYValue) = IIf(IsNull(rs.Fields("PYValue")), 0, rs.Fields("PYValue"))
    
    m_EECarItems(eecar_TotalRelief) = IIf(IsNull(rs.Fields("TotalRelief")), 0, rs.Fields("TotalRelief"))
    m_EECarItems(eecar_TotalValue) = IIf(IsNull(rs.Fields("TotalValue")), 0, rs.Fields("TotalValue"))
    m_EECarItems(eecar_UDBCode) = IIf(IsNull(rs.Fields("UDBCode")), 0, rs.Fields("UDBCode"))
    m_EECarItems(eecar_Benefit) = IIf(IsNull(rs.Fields("Benefit")), 0, rs.Fields("Benefit"))
    
    'SCREEN
    m_EECarItems(eecar_Item) = IIf(IsNull(rs.Fields("item")), "", rs.Fields("item"))
    m_EECarItems(eecar_EngineSize) = IIf(IsNull(rs.Fields("enginesize")), 0, rs.Fields("enginesize"))
    m_EECarItems(eecar_AmountReceived) = IIf(IsNull(rs.Fields("EECarAllowance")), 0, rs.Fields("EECarAllowance"))
    m_EECarItems(eecar_GrossBenefit) = IIf(IsNull(rs.Fields("GrossBenefit")), 0, rs.Fields("GrossBenefit"))
    m_EECarItems(eecar_AlternativeMethod) = rs.Fields("EECarAltMethod")
    m_EECarItems(eecar_carkey) = rs.Fields("EECarKey")
    m_EECarItems(eecar_AverageIRRate) = rs.Fields("AveIRRate")
    m_EECarItems(eecar_UseCompanyCarScheme) = rs.Fields("EECarAltMethod")
    
    s = IIf(IsNull(rs.Fields("FPCS")), "", rs.Fields("FPCS"))
    
    If Len(s) > 0 Then
      For I = 1 To CurrentEmployer.FPCSchemes.count
        Set FPCS = CurrentEmployer.FPCSchemes(I)
        If Not FPCS Is Nothing Then
          If StrComp(FPCS.name, s, vbTextCompare) = 0 Then
            m_EECarItems(eecar_FPCS) = I
          End If
        End If
      Next
      If I > CurrentEmployer.FPCSchemes.count Then
        'must be old default in p11d16 ie SCHEME1
        m_EECarItems(eecar_FPCS) = 1 'ir FPCS so default to IR
      End If
    Else
      m_EECarItems(eecar_FPCS) = 1 'ir FPCS so default to IR
    End If
     
    'dialog for extras
    m_EECarItems(eecar_HireCostMadeGood) = IIf(IsNull(rs.Fields("HCmadegood")), 0, rs.Fields("HCmadegood"))
    m_EECarItems(eecar_HireCost) = IIf(IsNull(rs.Fields("HireCost")), 0, rs.Fields("HireCost"))
    m_EECarItems(eecar_LumpSum) = IIf(IsNull(rs.Fields("Lumpsum")), 0, rs.Fields("Lumpsum"))
    m_EECarItems(eecar_MadeGood) = IIf(IsNull(rs.Fields("madegood")), 0, rs.Fields("madegood"))
    'sum of extras
    
    m_EECarItems(eecar_totalExtras) = GetTotalOfExtras ''m_EECarItems(eecar_HireCostMadeGood) + m_EECarItems(eecar_HireCost) + m_EECarItems(eecar_LumpSum) + m_EECarItems(eecar_MadeGood)
    m_EECarItems(eecar_TotalMiles) = ReadMileage(m_parent.Parent.db, mileage, m_EECarItems(eecar_carkey))
    
    
    m_readfromdb = True
  End If
  
clsBenEECar_readDB_End:
  Set eecar = Nothing
  Set FPCS = Nothing
  Set rs = Nothing
  IBenefitClass_ReadDB = I
  Call xReturn("clsBenEECar_readDB")
  Exit Function
clsBenEECar_readDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenEECar_readDB", "ERR_EECAR_READDB", "Unable to read the employee owned car details.")
  Resume clsBenEECar_readDB_End
  Resume
End Function

Public Function GetTotalOfExtras() As Long
  GetTotalOfExtras = CLng(Max(0, (m_EECarItems(eecar_HireCostMadeGood)) * -1) + CLng(m_EECarItems(eecar_HireCost))) + CLng(m_EECarItems(eecar_LumpSum)) + (CLng(m_EECarItems(eecar_MadeGood) * -1))
End Function

Private Function IBenefitClass_PrintWK(rep As Reporter) As Boolean
  Dim ben  As Long
  Dim subirvalue As Long
  Dim total As Long
  Dim totalexcess As Long
  Dim SubMadeGood As Long
  Dim SubMiles As Long
  Dim SubCoValue As Long
  Dim SubHireCost As Long
  Dim subhiremg As Long
  Dim sublumpsum As Long
  
  Dim I As Long
  
  On Error GoTo EECAR_PRINTWK_Err
  Call xSet("EECAR_PRINTWK")
  If rep.InitReport("EECarBen", gtarget, PORTRAIT, False, False, False) Then
    
    
  
  rep.Out STR_CR & STR_CR
  rep.Out "{Arial=8,n}{x=36}Total"
  rep.Out "{Arial=8,n}{x=16}Company"
  rep.Out "{Arial=8,n}{x=44}Lump sum " & STR_CR
  rep.Out "{Arial=8,n}{x=10}Engine"
  rep.Out "{Arial=8,n}{x=16}scheme"
  rep.Out "{Arial=8,n}{x=28}Business"
  rep.Out "{Arial=8,n}{x=36}mileage"
  rep.Out "{Arial=8,n}{x=44}and other"
  rep.Out "{Arial=8,n}{x=52}FPCS"
  rep.Out "{Arial=8,n}{x=60}Mileage"
  rep.Out "{Arial=8,n}{x=68}"
  rep.Out "{Arial=8,n}{x=76}Hire cost"
  rep.Out "{Arial=8,n}{x=0}" & STR_CR & "Car"
  rep.Out "{Arial=8,n}{x=10}size"
  rep.Out "{Arial=8,n}{x=16}name"
  rep.Out "{Arial=8,n}{x=28}miles"
  rep.Out "{Arial=8,n}{x=36}allowance"
  rep.Out "{Arial=8,n}{x=44}payments"
  rep.Out "{Arial=8,n}{x=52}rate"
  rep.Out "{Arial=8,n}{x=60}made good"
  rep.Out "{Arial=8,n}{x=68}Hire cost"
  rep.Out "{Arial=8,n}{x=76}made good"
  rep.Out "{Arial=8,n}{x=84}Benefit"
  rep.Out "{Arial=8,n}{x=92}Excess"
  rep.Out "{Arial=8,n}" & STR_CR & "{x=12}cc{Arial=8,n}{x=43}£   "
  rep.Out "{Arial=8,n}{x=51}£   "
  rep.Out "{Arial=8,n}{x=59}£   "
  rep.Out "{Arial=8,n}{x=67}£   "
  rep.Out "{Arial=8,n}{x=75}£   "
  rep.Out "{Arial=8,n}{x=83}£   "
  rep.Out "{Arial=8,n}{x=91}£   "
  rep.Out "{Arial=8,n}{x=99}£   "
  rep.Out STR_CR & STR_CR

  
  
  rep.Out "{Arial=8,n}{x=0}" & UCase$(IIf(IsNull(m_EECarItems(eecar_Item)), "Not available", (m_EECarItems(eecar_Item))))
  rep.Out "{Arial=8,rb}{x=15}" & Format$(IIf(IsNull(m_EECarItems(eecar_EngineSize)), 0, (m_EECarItems(eecar_EngineSize))))
  rep.Out "{Arial=8,n}{x=16}" & UCase$(IIf(IsNull(m_EECarItems(eecar_FPCS)), "Not available", Left$((m_EECarItems(eecar_FPCS)), 12)))
  rep.Out "{Arial=8,rb}{x=35}" & Format$(IIf(IsNull(m_EECarItems(eecar_TotalMiles)), 0, (m_EECarItems(eecar_TotalMiles))))
  rep.Out "{Arial=8,rb}{x=43}" & formatworkingnumber(IIf(IsNull(m_EECarItems(eecar_CompanyFPCSValue)), 0, (m_EECarItems(eecar_CompanyFPCSValue))), "")
  rep.Out "{Arial=8,rb}{x=75}" & formatworkingnumber(IIf(IsNull(m_EECarItems(eecar_HireCost)), 0, (m_EECarItems(eecar_HireCost))), "")
  rep.Out "{Arial=8,rb}{x=83}" & formatworkingnumber(IIf(IsNull(m_EECarItems(eecar_HireCost)), 0, -(m_EECarItems(eecar_HireCostMadeGood))), "")
  rep.Out "{Arial=8,rb}{x=51}" & formatworkingnumber(IIf(IsNull(m_EECarItems(eecar_LumpSum)), 0, (m_EECarItems(eecar_LumpSum))), "")
  rep.Out "{Arial=8,rb}{x=67}" & formatworkingnumber(IIf(IsNull(m_EECarItems(eecar_MadeGood)), 0, -(m_EECarItems(eecar_MadeGood))), "")

  rep.Out "{Arial=8,rb}{x=59}" & formatworkingnumber(IIf(IsNull(m_EECarItems(eecar_IRFPCSValue)), 0, (m_EECarItems(eecar_IRFPCSValue))), "")
  rep.Out "{Arial=8,rb}{x=91}" & formatworkingnumber(IIf(m_EECarItems(eecar_Benefit) < 0, 0, (m_EECarItems(eecar_Benefit))), "")
  ben = m_EECarItems(eecar_CompanyFPCSValue) + m_EECarItems(eecar_LumpSum) - m_EECarItems(eecar_IRFPCSValue) - m_EECarItems(eecar_MadeGood)
  If ben < 0 Then
    rep.Out "{Arial=8,rb}{x=99}" & formatworkingnumber(ben, "")
  Else
    rep.Out "{Arial=8,rb}{x=99}0"
  End If
  subirvalue = subirvalue + CDbl(IIf(IsNull(m_EECarItems(eecar_IRFPCSValue)), 0, (m_EECarItems(eecar_IRFPCSValue))))
  total = total + CDbl(IIf(m_EECarItems(eecar_Benefit) < 0, 0, (m_EECarItems(eecar_Benefit))))
  totalexcess = totalexcess + IIf(ben < 0, ben, 0)

  SubMadeGood = SubMadeGood + CDbl(IIf(IsNull(m_EECarItems(eecar_MadeGood)), 0, (m_EECarItems(eecar_MadeGood))))
  SubMiles = SubMiles + CLng(IIf(IsNull(m_EECarItems(eecar_TotalMiles)), 0, (m_EECarItems(eecar_TotalMiles))))
  SubCoValue = SubCoValue + CDbl(IIf(IsNull(m_EECarItems(eecar_CompanyFPCSValue)), 0, (m_EECarItems(eecar_CompanyFPCSValue))))
  SubHireCost = SubHireCost + CDbl(IIf(IsNull(m_EECarItems(eecar_HireCost)), 0, (m_EECarItems(eecar_HireCost))))
  subhiremg = subhiremg + CDbl(IIf(IsNull(m_EECarItems(eecar_HireCostMadeGood)), 0, (m_EECarItems(eecar_HireCostMadeGood))))
  sublumpsum = sublumpsum + CDbl(IIf(IsNull(m_EECarItems(eecar_LumpSum)), 0, (m_EECarItems(eecar_LumpSum))))
  rep.Out STR_CR

  rep.Out "{x=35}{LINE=-7}"
  rep.Out "{x=43}{LINE=-7}"
  rep.Out "{x=51}{LINE=-7}"
  rep.Out "{x=59}{LINE=-7}"
  rep.Out "{x=67}{LINE=-7}"
  rep.Out "{x=75}{LINE=-7}"
  rep.Out "{x=83}{LINE=-7}"
  rep.Out "{x=91}{LINE=-7}"
  rep.Out "{x=99}{LINE=-7}"
  rep.Out STR_CR

  rep.Out "{Arial=8,n}{x=0}"
  rep.Out "{Arial=8,rb}{x=35}" & Format$(SubMiles)
  rep.Out "{Arial=8,rb}{x=43}" & formatworkingnumber(CLng(SubCoValue), "£")

  rep.Out "{Arial=8,rb}{x=51}" & formatworkingnumber(CLng(sublumpsum), "£")
  rep.Out "{Arial=8,rb}{x=67}" & formatworkingnumber(-CLng(SubMadeGood), "£")
  rep.Out "{Arial=8,rb}{x=59}" & formatworkingnumber(CLng(subirvalue), "£")
  rep.Out "{Arial=8,rb}{x=75}" & formatworkingnumber(CLng(SubHireCost), "£")
  rep.Out "{Arial=8,rb}{x=83}" & formatworkingnumber(-CLng(subhiremg), "£")
  rep.Out "{Arial=8,rb}{x=91}" & formatworkingnumber(CLng(total), "£")
  rep.Out "{Arial=8,rb}{x=99}" & formatworkingnumber(CLng(totalexcess), "£")
  rep.Out STR_CR & "{x=35}{LINE=-7}"
  rep.Out "{x=43}{LINE=-7}"
  rep.Out "{x=51}{LINE=-7}"
  rep.Out "{x=59}{LINE=-7}"
  rep.Out "{x=67}{LINE=-7}"
  rep.Out "{x=75}{LINE=-7}"
  rep.Out "{x=83}{LINE=-7}"
  rep.Out "{x=91}{LINE=-7}"
  rep.Out "{x=99}{LINE=-7}"
  rep.Out STR_CR & "{x=0}{LINE=100}" & STR_CR
  rep.Out "{Arial=8,n}{x=0}Total Net Benefit"
  rep.Out "{Arial=8,rb}{x=91}" & formatworkingnumber(CLng(total), "£") & STR_CR
  rep.Out "{Arial=8,n}{x=0}{LINE=100}" & STR_CR & STR_CR



'If sublumpsum <> 0 Then
'  wkeecarlump y, dyee
'End If

  rep.Out STR_CR & STR_CR

'Summary of FPCS details
  rep.Out "{x=0}{LINE=100}" & STR_CR
  rep.Out "{Arial=8,n}{x=0}" & "Details of the employers mileage scheme(s) used." & STR_CR & STR_CR
  rep.Out "{Arial=8,n}{x=0}Name of scheme"
  rep.Out "{x=40}Engine Size" & STR_CR
  rep.Out "{x=40}greater than"
  rep.Out "{x=60}Mileage exceeding"
  rep.Out "{x=80}Cost per mile" & STR_CR


  rep.Out "{x=0}{LINE=100}" & STR_CR
' message about more than 1 car in FPCS apf
'If iCarCount > 1 Then
  'rep.out "{Arial8i}The calculation of the benefit or excess assumes that all the above cars were owned concurrently by the employee in the tax year." & str_cr
'End If


    rep.Out "{x=0}{Arial=8,n}Where payment of mileage allowance results in an excess over the Inland Revenue's Fixed Profit Car Scheme (FPCS)" & STR_CR
    rep.Out "{x=0}{Arial=8,n}rates, the figure should be enterend in box 1.32 of the employment pages of the self assessment tax return." & STR_CR & STR_CR

    
    
    'rep.Out "{Arial=8,n}{x=0}Make and model{Arial=8,b} " & m_EECarItems(car_make) & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,b}Price of the car for tax purposes" & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}{x=0}(a) Price of the car"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_ListPrice), "£") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}{x=0}(b) Price of optional accessories available when the car was first made available"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_OriginalAccessories), "") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}{x=0}(c) Price of accessories added after car was first made available"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_NewAccessories), "") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}{x=97}{LINE=-7}" & vbCrLf
    'rep.Out "{Arial=8,b}{x=0}Total car plus accessories (a + b + c)"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_GrossPrice), "£") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Less: capital contribution by the employee towards the cost of the car (" & formatworkingnumber(m_EECarItems(car_capitalcontribution), "£") & " restricted to £5000)"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(Min(m_EECarItems(car_capitalcontribution), rates.GetItem(carMAXCAPCONTRIB)), "", True) & vbCrLf & vbCrLf
    'rep.Out "{x=97}{line=-7}" & vbCrLf
    'rep.Out "{Arial=8,n}{x=0}Adjusted price of the car for " & rates.GetItem(T_STRING Or taxFormYear) & " (restricted to £80,000)"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_Price), "£") & vbCrLf
    'rep.Out "{x=97}{line=-7,d}" & vbCrLf
    'rep.Out "{Arial=8,n}Gross benefit for " & rates.GetItem(T_STRING Or taxFormYear) & " (Adjusted price x 35%)"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_Scaled), "£") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Less: reduction for business mileage (" & m_EECarItems(car_businessmiles) & ")"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_BusMilereduction), "", True) & vbCrLf & vbCrLf
   '
    'rep.Out "{Arial=8,n}Less: reduction for age of car (first registered " & m_EECarItems(car_Registrationdate) & ")"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_Agereduction), "", True) & vbCrLf & vbCrLf
    
    'rep.Out "{Arial=8,n}Made available from {Arial=8,b}" & m_EECarItems(car_Availablefrom) & " to " & "{Arial=8,b}" & m_EECarItems(car_Availableto) & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Qualifying Days unavailable {Arial=8,b}" & m_EECarItems(car_Unavailabledays) & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Less: reduction for days when the car was unavailable "
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_cardaysreduction), "", True) & vbCrLf & vbCrLf
   '
    'rep.Out "{Arial=8,n}Less: Payment from the employee for private use of the car"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_usagecontribution), "", True) & vbCrLf & vbCrLf
   '
    ''rep.Out "{x=97}{LINE=-7}" & vbCrLf
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_carbenefit), "£") & vbCrLf
    'rep.Out "{x=97}{LINE=-7,d}" & vbCrLf
    'rep.Out "{Arial=8,ni}{LINE=100}" & vbCrLf
    'rep.Out "{Arial=8,nb}Car fuel benefit charge" & vbCrLf & vbCrLf
    'rep.Out "{Arial=8}Was fuel provided other than for business travel?"
    'rep.Out "{Arial=8,b}{x=80}" & IIf(m_EECarItems(car_privatefuel), "Yes", "No") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Was the employee required to make good the cost of all fuel so provided"
    'rep.Out "{Arial=8,b}{x=80}" & IIf(m_EECarItems(car_requiredmakegood), "Yes", "No") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}If so required, did the employee do so ?"
    'rep.Out "{Arial=8,b}{x=80}" & IIf(m_EECarItems(car_actualmadegood), "Yes", "No") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Cylinder Capacity"
    'rep.Out "{Arial=8,b}{x=80}" & m_EECarItems(car_Enginesize) & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Type of Engine"
    'rep.Out "{Arial=8,b}{x=80}" & IIf(m_EECarItems(car_Diesel), "Diesel", "Petrol") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Car fuel scale charge"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_GrossFuel), "£") & vbCrLf & vbCrLf
    'rep.Out "{Arial=8,n}Less reduction for days when the car was unavailable (refer to availability above)"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_fueldaysreduction), "", True) & vbCrLf & vbCrLf
    'rep.Out "{x=97}{LINE=-7}" & vbCrLf
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_FuelBenefit), "£") & vbCrLf
    'rep.Out "{x=97}{LINE=-7,d}" & vbCrLf
    'rep.Out "{Arial=8,i}{Line=100}" & vbCrLf
    'rep.Out "{Arial=8,n}Assesable car benefit"
    'rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_EECarItems(car_Benefit), "£") & vbCrLf
    'rep.Out "{Arial=8,i}{Line=100}" & vbCrLf
  End If
  Call rep.EndReport

EECAR_PRINTWK_End:
  Call xReturn("EECAR_PRINTWK")
  Exit Function

EECAR_PRINTWK_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EECAR_PRINTWK", "ERR_EECAR_PRINTWK", "Error printing the employee owned car working paper")
  Resume EECAR_PRINTWK_End
End Function


'Public Function ReadEEMileage() As Long
'  Dim miles As clsBenEECarMiles
'  Dim rs As Recordset
'  Dim lMiles As Long
'
'  On Error GoTo ReadEEMileage_Err
'  Call xSet("ReadEEMileage")
'
'  Set rs = m_parent.Parent.db.OpenRecordset(Queries.Queries(SELECT_EECARMILES, m_EECarItems(eecar_carkey), 0), dbOpenForwardOnly)
'  Do While Not rs.EOF
'    Set miles = New clsBenEECarMiles
'    If IsNull(rs.Fields("MilesItem")) Then
'      miles.MileItem = ""
'    Else
'      miles.MileItem = rs.Fields("MilesItem")
'    End If
'    If IsNull(rs.Fields("MilesDate")) Then
'      miles.MileDate = UNDATED
'    Else
'      miles.MileDate = rs.Fields("MilesDate")
'    End If
'    miles.MileAmount = rs.Fields("Miles")
'    lMiles = lMiles + miles.MileAmount
'    Call Mileage.Add(miles)
'    Set miles = Nothing
'    rs.MoveNext
'  Loop
'
'ReadEEMileage_End:
'  ReadEEMileage = lMiles
'  Call xReturn("ReadEEMileage")
'  Exit Function
'ReadEEMileage_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "ReadEEMileage", "ERR_UNDEFINED", "Undefined error.")
'  ReadEEMileage = 0
'  Resume ReadEEMileage_End
'End Function


'Public Sub WriteEEMileage()
'  Dim miles As clsBenEECarMiles
'  Dim rs As Recordset
'  Dim I As Long
'  Dim lMiles As Long
'
'  On Error GoTo WriteEEMileage_Err
'  Call xSet("WriteEEMileage")
'
'  If Not (m_EECarItems(eecar_carkey) = Empty) Then _
'    Call m_parent.Parent.db.Execute(Queries.Queries(DELETE_EECARMILES, m_EECarItems(eecar_carkey)))
'  Set rs = m_parent.Parent.db.OpenRecordset(Queries.Queries(SELECT_EECARMILES, m_EECarItems(eecar_carkey), 0), dbOpenDynaset)
'
'  For I = 0 To Mileage.count - 1
'    Set miles = Mileage(I)
'    rs.AddNew
'    If Len(miles.MileItem) > 0 Then
'      rs.Fields("MilesItem") = miles.MileItem
'    Else
'      rs.Fields("MilesItem") = Null
'    End If
'    If miles.MileDate <> UNDATED Then
'      rs.Fields("MilesDate") = miles.MileDate
'    Else
'      rs.Fields("MilesDate") = Null
'    End If
'    rs.Fields("Miles") = miles.MileAmount
'    rs.Fields("eeCarKey") = m_EECarItems(eecar_carkey)
'    rs.Update
'  Next I
'
'WriteEEMileage_End:
'  Set rs = Nothing
'  Call xReturn("WriteEEMileage")
'  Exit Sub
'
'WriteEEMileage_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "WriteEEMileage", "ERR_UNDEFINED", "Undefined error.")
'  Resume WriteEEMileage_End
'End Sub

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_EECAR
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim FPCS As FPCS
  
  On Error GoTo clsBenEECar_WriteDB_Err
  Call xSet("clsBenEECar_WriteDB")
  Set rs = m_parent.Parent.rsBenTables(TBL_EECAR)
  
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Edit
  Else
    rs.AddNew
    rs.Fields(S_EMPLOYEE_NUM_FIELD) = m_EECarItems(car_employeeref)
  End If
  
'  rs.Fields("Item") = m_EECarItems(eecar_Item)
'
'
'  rs.Fields("CCBandKey") = m_EECarItems(eecar_CCbandKey)
'  rs.Fields("CCBand") = m_EECarItems(eecar_CCband)
'  rs.Fields("Eligible") = m_EECarItems(eecar_Eligible)
'  rs.Fields("TotalMiles") = m_EECarItems(eecar_TotalMiles)
'  rs.Fields("FPCS") = m_EECarItems(eecar_FPCS)
'  rs.Fields("ercontrib") = m_EECarItems(eecar_ercontrib)
'  rs.Fields("madegood") = m_EECarItems(eecar_MadeGood)
'  rs.Fields("enginesize") = m_EECarItems(eecar_EngineSize)
'  rs.Fields("EECarAllowance") = m_EECarItems(eecar_AmountReceived)
'  rs.Fields("GrossBenefit") = m_EECarItems(eecar_GrossBenefit)
'  rs.Fields("TwoCars") = m_EECarItems(eecar_AlternativeMethod)
'  'rs.Fields("EECarKey") = m_EECarItems(eecar_carkey)
'  rs.Fields("AveIRRate") = m_EECarItems(eecar_AverageIRRate)
'  rs.Fields("EECarAltMethod") = m_EECarItems(eecar_UseCompanyCarScheme)
'  rs.Fields("HCmadegood") = m_EECarItems(eecar_HireCostMadeGood)
'  rs.Fields("HireCost") = m_EECarItems(eecar_HireCost)
'  rs.Fields("IRValue") = m_EECarItems(eecar_IRFPCSValue)
'  rs.Fields("CoValue") = m_EECarItems(eecar_CompanyFPCSValue)
'  rs.Fields("PreERBenefit") = m_EECarItems(eecar_PreErBenefit)
'  rs.Fields("Lumpsum") = m_EECarItems(eecar_LumpSum)
'  rs.Fields("OldEECarKey") = m_EECarItems(eecar_OLDEECarKey)
'  rs.Fields("PYValue") = m_EECarItems(eecar_PYValue)
'  rs.Fields("TotalRelief") = m_EECarItems(eecar_TotalRelief)
'  rs.Fields("TotalValue") = m_EECarItems(eecar_TotalValue)
'  rs.Fields("UDBCode") = m_EECarItems(eecar_UDBCode)
'  rs.Fields("Benefit") = m_EECarItems(eecar_Benefit)
  
  
  'general
  rs.Fields(S_EMPLOYEE_NUM_FIELD) = m_EECarItems(eecar_EmployeeNumber)
  rs.Fields("PYValue") = m_EECarItems(eecar_PYValue)
  
  rs.Fields("TotalRelief") = m_EECarItems(eecar_TotalRelief)
  rs.Fields("TotalValue") = m_EECarItems(eecar_TotalValue)
  rs.Fields("UDBCode") = m_EECarItems(eecar_UDBCode)
  rs.Fields("Benefit") = m_EECarItems(eecar_Benefit)
  
  'SCREEN
  rs.Fields("item") = m_EECarItems(eecar_Item)
  rs.Fields("enginesize") = m_EECarItems(eecar_EngineSize)
  rs.Fields("EECarAllowance") = m_EECarItems(eecar_AmountReceived)
  rs.Fields("GrossBenefit") = m_EECarItems(eecar_GrossBenefit)
  rs.Fields("EECarAltMethod") = m_EECarItems(eecar_AlternativeMethod)
  
  rs.Fields("AveIRRate") = m_EECarItems(eecar_AverageIRRate)
  rs.Fields("EECarAltMethod") = m_EECarItems(eecar_UseCompanyCarScheme)
  
  If m_EECarItems(eecar_FPCS) = 1 Then
    'ir scheme , place old default for backward compatibility
    rs.Fields("FPCS") = S_SCHEME1
  Else
    Set FPCS = CurrentEmployer.FPCSchemes(m_EECarItems(eecar_FPCS))
    rs.Fields("FPCS") = FPCS.name
  End If
  
  'dialog for extras
  rs.Fields("HCmadegood") = m_EECarItems(eecar_HireCostMadeGood)
  rs.Fields("HireCost") = m_EECarItems(eecar_HireCost)
  rs.Fields("Lumpsum") = m_EECarItems(eecar_LumpSum)
  rs.Fields("madegood") = m_EECarItems(eecar_MadeGood)
  
  rs.Update
    
  If Len(m_sbookmark) = 0 Then
    m_sbookmark = rs.LastModified 'zzzz check APF
    rs.Bookmark = rs.LastModified
  End If
  
  'Fill out the car key value for the actual miles details
  m_EECarItems(eecar_carkey) = rs.Fields("EECarKey")
  
  'zzzz RMc We should not always call this - only when necessary.
  Call WriteMileage(m_parent.Parent.db, mileage, m_EECarItems(eecar_carkey))
  
  m_dirty = False
  IBenefitClass_WriteDB = True
  
clsBenEECar_WriteDB_End:
  Set FPCS = Nothing
  Set rs = Nothing
  Call xReturn("clsBenEECar_WriteDB")
  Exit Function
clsBenEECar_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "clsBenEECar_WriteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume clsBenEECar_WriteDB_End
  Resume
End Function

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  On Error GoTo clsBenEECar_DeleteDB_Err
  
  Call xSet("clsBenEECar_DeleteDB")
  Set rs = m_parent.Parent.rsBenTables(TBL_EECAR)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  'zzzz delete miles
  IBenefitClass_DeleteDB = True
  
clsBenEECar_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("clsBenEECar_DeleteDB")
  Exit Function
clsBenEECar_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenEECar_DeleteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume clsBenEECar_DeleteDB_End
  Resume
End Function

Public Sub Ibenefitclass_Cleardirty()
  m_dirty = False
End Sub


Public Function GetFPCSBand(CSB As FPCSBand, ByVal Miles As Long, ByVal cc As Long, FPCSIndex As Long) As Boolean
  Dim CS As FPCS
  Dim l As Long
  Dim b As FPCSBand
  
  On Error GoTo GetFPCSBand_Err
  Call xSet("GetFPCSBand")

  Set CS = CurrentEmployer.FPCSchemes(FPCSIndex)
  
  If CS Is Nothing Then
    'err.Raise ZZZZZ
  Else
    'Find the CC
    For l = 1 To CS.Bands.count
      Set b = CS.Bands(l)
      If Not b Is Nothing Then
        If cc > b.GreaterThanCC And cc <= b.LessThanOrEqualToCC Then
          'now check the miles
          If Miles > b.GreaterThanMiles And Miles <= b.LessThanOrEqualToMiles Then
            'we have the right one
            Set CSB = b
            GetFPCSBand = True
            Exit For
          End If
        End If
      End If
    Next
    'Check the miles
  End If


GetFPCSBand_End:
  Set CS = Nothing
  Set b = Nothing
  Call xReturn("GetFPCSBand")
  Exit Function
GetFPCSBand_Err:
  GetFPCSBand = False
  Call ErrorMessage(ERR_ERROR, Err, "GetFPCSBand", "ERR_UNDEFINED", "Undefined error.")
  Resume GetFPCSBand_End
End Function


Private Function GetFPCSMileageValue(value As Variant, ByVal Miles As Long, ByVal cc As Long, FPCSIndex As Long, ByVal UseAverageIRRates As Boolean) As Boolean
  Dim CSB As FPCSBand
  Dim Bands As ObjectList
  Dim CSB2 As FPCSBand
  Dim l As Long, m As Long, lTempMiles As Long
  Dim lLastMileageLessThanOrEqualTo As Long
  Dim CS As FPCS
  Dim dblRate0to4000 As Double, dblRate4000andAbove As Double 'only used for AverageIR
  
  On Error GoTo GetFPCSMileageValue_Err
  Call xSet("GetFPCSMileageValue")

  Set Bands = New ObjectList
  
  If UseAverageIRRates And m_EECarItems(eecar_FPCS) = L_IRFPCS Then
    'we use ir scheme regardless of FPCSIndex
    'we want the 2 middle bands for cc purposes and and average of the 2 rates at each cc mileage level
    'there are only two levels 0 to 4000 and > 4000 miles
    Set CS = CurrentEmployer.FPCSchemes(L_IRFPCS)
    If Not CS Is Nothing Then
      If CS.Bands.count = 8 Then '4 cc bands with 2 mileages
        Set CSB = CS.Bands(3) 'see REadIRF??? in TaxRates.cls for order of reading in IR fixedprofit car schemes
        Set CSB2 = CS.Bands(5)
        dblRate0to4000 = (CSB.Rate + CSB2.Rate) / 2
        Set CSB = CS.Bands(4)
        Set CSB2 = CS.Bands(6)
        dblRate4000andAbove = (CSB.Rate + CSB2.Rate) / 2
        lTempMiles = Miles - rates.GetItem(fpcsMILEAGELIMIT)
        If lTempMiles > 0 Then
          value = lTempMiles * dblRate4000andAbove
          value = value + (rates.GetItem(fpcsMILEAGELIMIT) * dblRate0to4000)
        Else
          value = Miles * dblRate0to4000
        End If
        GetFPCSMileageValue = True
      Else
        'err.raise ZZZZ
      End If
    Else
      'err.raise ZZZZZ
    End If
  Else
    If GetFPCSBand(CSB, Miles, cc, FPCSIndex) Then
      Set CS = CurrentEmployer.FPCSchemes(FPCSIndex)
      'get the lower bands with the same CC range and miles less than and in ascending order
      Call GetBandsInOrder(Bands, CS, CSB.ObjectListIndex)
      
      'we now have the bands in order so now use our Miles and determine the money/Value
      m = Bands.count
      For l = 1 To m
        Set CSB = Bands(l)
        If l = m Then
          'last one and thus no taking away
          value = value + (Miles * CSB.Rate)
        Else
          lTempMiles = CSB.LessThanOrEqualToMiles - lLastMileageLessThanOrEqualTo
          value = value + (lTempMiles * CSB.Rate)
          Miles = Miles - lTempMiles
        End If
        If CSB.LessThanOrEqualToMiles = DBL_LESSTHANINFINITY Then Exit For
        lLastMileageLessThanOrEqualTo = CSB.LessThanOrEqualToMiles
      Next
      GetFPCSMileageValue = True
    End If
  End If
  
  If Not CS Is Nothing Then CS.ClearAdded
  
GetFPCSMileageValue_End:
  Set CSB = Nothing
  Set Bands = Nothing
  Set CSB2 = Nothing
  Set CS = Nothing
  Call xReturn("GetFPCSMileageValue")
  Exit Function

GetFPCSMileageValue_Err:
  GetFPCSMileageValue = False
  Call ErrorMessage(ERR_ERROR, Err, "GetFPCSMileageValue", "ERR_UNDEFINED", "Undefined error.")
  Resume GetFPCSMileageValue_End
  Resume
End Function


Private Function GetBandsInOrder(Bands As ObjectList, CS As FPCS, lTopBandIndex As Long) As Boolean
  Dim dblLastTopMileage As Double
  Dim band As FPCSBand, CurrentTopBand As FPCSBand
  Dim l As Long, m As Long
  Dim lNextTopBandIndex As Long
  Dim lNumberAdded As Long
  Dim IndexesOfBands() As Long
  
  On Error GoTo GetBandsInOrder_Err
  Call xSet("GetBandsInOrder")
  
  Bands.RemoveAll
  Bands.Compact
  
  Set CurrentTopBand = CS.Bands(lTopBandIndex)
  
  lNumberAdded = 1
  
  ReDim Preserve IndexesOfBands(1 To lNumberAdded)
  
  CurrentTopBand.Added = True
  
  IndexesOfBands(lNumberAdded) = CurrentTopBand.ObjectListIndex
  
  For l = 1 To CS.Bands.count
    For m = 1 To CS.Bands.count
      Set band = CS.Bands(m)
      'check the CC range
      If Not band Is Nothing Then
        If (band.GreaterThanCC = CurrentTopBand.GreaterThanCC) And (band.GreaterThanCC = CurrentTopBand.GreaterThanCC) Then
          'check for mileage less
          If (band.LessThanOrEqualToMiles > dblLastTopMileage) And (band.LessThanOrEqualToMiles <= CurrentTopBand.GreaterThanMiles) Then
            If Not band.Added Then
              
              dblLastTopMileage = band.LessThanOrEqualToMiles
              lNextTopBandIndex = m
            End If
          End If
        End If
      End If
    Next
    
    If lNextTopBandIndex <> 0 Then
      Set CurrentTopBand = CS.Bands(lNextTopBandIndex)
      CurrentTopBand.Added = True
      lNumberAdded = lNumberAdded + 1
      ReDim Preserve IndexesOfBands(1 To lNumberAdded)
      IndexesOfBands(lNumberAdded) = CurrentTopBand.ObjectListIndex
      Set CurrentTopBand = CS.Bands(IndexesOfBands(lNumberAdded))
      lNextTopBandIndex = 0
    Else
      'none found on last go
      
      Exit For
    End If
  Next
  
  For l = UBound(IndexesOfBands) To 1 Step -1
    Call Bands.Add(CS.Bands(IndexesOfBands(l)))
  Next
  
GetBandsInOrder_End:
  Call xReturn("GetBandsInOrder")
  Exit Function

GetBandsInOrder_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetBandsInOrder", "ERR_UNDEFINED", "Undefined error.")
  Resume GetBandsInOrder_End
  Resume
End Function


Private Function GetIRFPCSValue(vMileageValue As Variant) As Boolean

  Dim lTotalMilesAllNonAlternativeCars As Long
  On Error GoTo GetIRFPCSValue_Err
  Call xSet("GetIRFPCSValue")

  
  vMileageValue = 0
  
  If m_EECarItems(eecar_AlternativeMethod) <> 0 Then
    Call GetFPCSMileageValue(vMileageValue, m_EECarItems(eecar_TotalMiles), m_EECarItems(eecar_EngineSize), L_IRFPCS, m_EECarItems(eecar_AverageIRRate))
  Else
    'get the miles for all the cars where m_EECarItem(eecar_AlternativeMethod) = 0 and add together
      lTotalMilesAllNonAlternativeCars = GetNonAlternativeCarMiles()
      Call GetFPCSMileageValue(vMileageValue, lTotalMilesAllNonAlternativeCars, m_EECarItems(eecar_EngineSize), L_IRFPCS, m_EECarItems(eecar_AverageIRRate))
      vMileageValue = vMileageValue * (m_EECarItems(eecar_TotalMiles) / IIf(lTotalMilesAllNonAlternativeCars = 0, 1, lTotalMilesAllNonAlternativeCars))
  End If

  



GetIRFPCSValue_End:
  Call xReturn("GetIRFPCSValue")
  Exit Function

GetIRFPCSValue_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetIRFPCSValue", "ERR_UNDEFINED", "Undefined error.")
  Resume GetIRFPCSValue_End
  Resume
End Function


Public Function GetNonAlternativeCarMiles() As Long
  Dim benefit As IBenefitClass
  Dim l As Long
  
  On Error GoTo GetNonAlternativeCarMiles_Err
  Call xSet("GetNonAlternativeCarMiles")

  'have to go through all current as can not reference forms list view for a shortcut
  With m_parent
    For l = 1 To .benefits.count
      Set benefit = .benefits(l)
      If Not benefit Is Nothing Then
        If (benefit.BenefitClass = IBenefitClass_BenefitClass) Then
          If benefit.GetItem(eecar_AlternativeMethod) = 0 Then
            GetNonAlternativeCarMiles = GetNonAlternativeCarMiles + benefit.GetItem(eecar_TotalMiles)
          End If
        End If
      End If
    Next
  End With
  


GetNonAlternativeCarMiles_End:
  Call xReturn("GetNonAlternativeCarMiles")
  Exit Function

GetNonAlternativeCarMiles_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetNonAlternativeCarMiles", "ERR_UNDEFINED", "Undefined error.")
  Resume GetNonAlternativeCarMiles_End
End Function


Public Function CalcEmployeeCarsWithAlternativMethodFalse(benefitX As IBenefitClass) As Boolean
  Dim benefit As IBenefitClass
  Dim l As Long, m As Long
  Dim ibf As IBenefitForm2
  Dim vBenefit As Variant
  
  On Error GoTo CalcEmployeeCarsWithAlternativMethodFalse_Err
  Call xSet("CalcEmployeeCarsWithAlternativMethodFalse")

  Call benefitX.SetItem(eecar_calculated, True)
  
  For l = 1 To m_parent.benefits.count
    Set benefit = m_parent.benefits(l)
    If Not (benefit Is Nothing) And Not (benefit Is benefitX) Then
      If benefit.BenefitClass = benefitX.BenefitClass Then
        If benefit.GetItem(eecar_calculated) <> True Then
          vBenefit = S_ERROR
          Call benefitX.SetItem(eecar_calculated, True)
          vBenefit = benefit.Calculate
          If CurrentForm Is F_BenEECar Then
            Set ibf = CurrentForm
            For m = 1 To ibf.lv.ListItems.count
              If l = ibf.lv.ListItems(m).Tag Then
                ibf.lv.ListItems(m).Text = benefit.name
                ibf.lv.ListItems(m).SubItems(1) = formatworkingnumber(vBenefit, "£")
              End If
            Next
          End If
        End If
      End If
    End If
  Next l
  
  
  For l = 1 To m_parent.benefits.count
    If Not benefit Is Nothing Then
      If benefit.BenefitClass = benefitX.BenefitClass Then
        Call benefitX.SetItem(eecar_calculated, False)
      End If
    End If
  Next
  
  
  
CalcEmployeeCarsWithAlternativMethodFalse_End:
  Call xReturn("CalcEmployeeCarsWithAlternativMethodFalse")
  Exit Function

CalcEmployeeCarsWithAlternativMethodFalse_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CalcEmployeeCarsWithAlternativMethodFalse", "ERR_UNDEFINED", "Undefined error.")
  Resume CalcEmployeeCarsWithAlternativMethodFalse_End
  Resume
End Function

