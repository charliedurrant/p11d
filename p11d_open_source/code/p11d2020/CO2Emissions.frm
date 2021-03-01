VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form F_CompanyCarCO2Emissions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Make - CO2 Emissions"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dbMakes 
      Caption         =   "Model Records"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   4635
      Visible         =   0   'False
      Width           =   6180
   End
   Begin TrueDBGrid60.TDBGrid CO2TDBGrid 
      Bindings        =   "CO2Emissions.frx":0000
      Height          =   3930
      Left            =   90
      OleObjectBlob   =   "CO2Emissions.frx":0016
      TabIndex        =   4
      Top             =   585
      Width           =   8835
   End
   Begin VB.CommandButton B_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7785
      TabIndex        =   3
      Top             =   4635
      Width           =   1140
   End
   Begin VB.CommandButton B_Ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6435
      TabIndex        =   2
      Top             =   4650
      Width           =   1185
   End
   Begin VB.ComboBox CB_Make 
      DataSource      =   "DB"
      Height          =   315
      ItemData        =   "CO2Emissions.frx":271A
      Left            =   1125
      List            =   "CO2Emissions.frx":271C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "This information is provided by Crown Copyright.  These details can also be found on their website at www.vcacarfueldata.org.uk"
      Height          =   420
      Left            =   120
      TabIndex        =   5
      Top             =   5085
      Width           =   8790
   End
   Begin VB.Label lblMake 
      Alignment       =   1  'Right Justify
      Caption         =   "Make:"
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   870
   End
End
Attribute VB_Name = "F_CompanyCarCO2Emissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Looks like the source has changed ****
'email fuel@vca.gov.uk and request the data

'Dear Sir / Madam,

'Would it be possible to receive a copy of the car fuel / Co2 data as per

'http://carfueldata.direct.gov.uk/downloads/default.aspx

'in MS Access format by email return.

'Yours,

'Charles Durrant
'P11D Developer - Thomson Reuters



'new car data'

'download from http://www.vcacarfueldata.org.uk/downloads/latest.asp
'get the mdb file
'run

'SELECT m.Manufacturer, Replace(Replace(Model,"½",""),"é","e") AS CarModel, Vehicles.Transmission, Vehicles.EngineCapacity, Vehicles.FuelType,Vehicles.CO2
'FROM (Manufacturers AS m INNER JOIN Models AS mo ON m.ManufacturerID = mo.ManufacturerID) INNER JOIN Vehicles ON mo.ModelID = Vehicles.ModelID   order by m.manufacturer, Model

'copy the results ctrl c
'delete all from T_CO2Emissions in the pd file
'ctr v the new items !

'*** ALSO get any new fuel types and map to internal types
'SELECT Vehicles.FuelType from vehicles group by fueltype
'and map text in the B_OK_Click event

Option Explicit
Implements IBenefitForm2

Public Parentibf As IBenefitForm2
Public benefit As IBenefitClass

Private m_bDirty As Boolean
Private m_Miles As ObjectList
 
Private Sub B_OK_Click()
  
  On Error GoTo B_OK_Click_ERR

  Call xSet("B_OK_Click")

  'KM - populate text boxes, Combo boxes and check box on company car dialog
  benefit.value(car_Make_db) = CB_Make.Text
  
  benefit.value(car_Model_db) = CO2TDBGrid.Columns(0).value
  benefit.value(car_enginesize_db) = CO2TDBGrid.Columns(2).value
  'CASE statement to determine fuel type
  Select Case UCASE$(Me.CO2TDBGrid.Columns(3).value)
  Case "PETROL"
    benefit.value(car_p46FuelType_db) = CCFT_PETROL
  Case "DIESEL"
    benefit.value(car_p46FuelType_db) = CCFT_EUROIVDIESEL
  Case "PETROL ELEC"
    'If LPG, which is a propane mix fuel, set fueltype to bi-fuel
    benefit.value(car_p46FuelType_db) = CCFT_HYBRID
  Case "PETROL HYBRID"
    'If LPG, which is a propane mix fuel, set fueltype to bi-fuel
    benefit.value(car_p46FuelType_db) = CCFT_BIFUEL_WITH_CO2_FOR_GAS
  Case Else
    'set to petrol
    benefit.value(car_p46FuelType_db) = CCFT_PETROL
  End Select
  
  benefit.value(car_p46CarbonDioxide_db) = Me.CO2TDBGrid.Columns(4).value
  
  benefit.Dirty = True
  
  Me.Hide
  
B_OK_Click_END:
  Call xReturn("B_OK_Click")
  Exit Sub
B_OK_Click_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "B_OK_Click", "B_OK_Click", "Unable to populate fields on Company Car dialogue.")
  Resume B_OK_Click_END
  Resume
End Sub

Private Sub B_Cancel_Click()
  Me.Hide
End Sub

Private Sub CB_Make_Change()
  dbMakes.RecordSource = "SELECT Model, Transmission, EngineCapacity, FuelType, CO2 FROM T_CO2EMISSIONS WHERE MAKE = " & StrSQL(CB_Make)
  
  dbMakes.Refresh
  CO2TDBGrid.ReBind
End Sub

Private Sub CB_Make_Click()
  dbMakes.RecordSource = "SELECT Model, Transmission, EngineCapacity, FuelType, CO2 FROM T_CO2EMISSIONS WHERE MAKE = " & StrSQL(CB_Make)
  dbMakes.Refresh
  CO2TDBGrid.ReBind
End Sub

Private Sub CO2TDBGrid_Click()
  B_Ok.Enabled = True
End Sub

'cad review 20/02 can this be tidied
Private Sub Form_Load()
  
  Dim grd As Object
  Dim rsMakes As Recordset
  Dim i As Long
  Dim makeIndex As Long
  Dim k As Integer  'AM
  
  On Error GoTo F_CompanyCarCO2Emissions_Load_ERR

  Call xSet("F_CompanyCarCO2Emissions_Load")
    
  Set rsMakes = p11d32.Rates.IRCO2Makes
  
  Do While Not rsMakes.EOF
    CB_Make.AddItem (rsMakes.Fields("Make").value)
    rsMakes.MoveNext
  Loop
  
  'km - initialise dbMakes control datasource to appropriate .pd file
  dbMakes.DatabaseName = p11d32.PDDBPath
  'km - set the MAKE Combo box default to the first item in the list
  CB_Make.ListIndex = 0
  'km - disable the OK button initially, until a model is chosen
  B_Ok.Enabled = False
  
  For k = 0 To CO2TDBGrid.Columns.Count - 1 'AM
    CO2TDBGrid.Columns(k).width = (CO2TDBGrid.width - 610) / CO2TDBGrid.Columns.Count
  Next k
  
F_CompanyCarCO2Emissions_Load_END:
  Set grd = Nothing
  Call xReturn("F_CompanyCarCO2Emissions_Load")
  Exit Sub
F_CompanyCarCO2Emissions_Load_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "F_CompanyCarCO2Emissions_Load", "F_CompanyCarCO2Emissions Load", "Unable to load the Make - CO2 Emissions form.")
  Resume F_CompanyCarCO2Emissions_Load_END
  Resume
  
End Sub
Private Sub IBenefitForm2_AddBenefit()
End Sub
Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
End Function
Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property
Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property
Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  'not used
End Function
Private Function IBenefitForm2_BenefitOff() As Boolean
  'not used
End Function
Private Function IBenefitForm2_BenefitOn() As Boolean
  'NOT USED
End Function
Private Function IBenefitForm2_BenefitsToListView() As Long
  'not used
End Function
Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  'not used

End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
On Error GoTo BenefitToScreen_Err:
  
  Call xSet("BenefitToScreen")
  
  'ensure OK button is disabled
  B_Ok.Enabled = False
  
  If BenefitIndex <> -1 Then Call p11d32.Help.ShowForm(Me, vbModal)     ' Me.Show vbModal
  
BenefitToScreen_End:
  Set ibf = Nothing
  Set ben = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unable to place the CO2Emissions to the screen.")
  Resume BenefitToScreen_End
  Resume
End Function
Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property
Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_COMPANY_CARS_F
End Property
Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  'not used
End Property
Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  'not used
End Function
Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  'not used
End Function
Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  'not used
End Function
'Private Sub ubgrd_DeleteData(ObjectList As ObjectList, ObjectListIndex As Long)
'  Call UBGRD.ObjectList.Remove(ObjectListIndex)
'End Sub

Private Sub UBGRD_ValidateTCS(FirstColIndexInError As Long, ValidateMessage As String, ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ByVal ObjectListIndex As Long)
  'Not used
End Sub
Private Sub ubgrd_WriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long)
  'Not used
End Sub

