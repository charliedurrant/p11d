VERSION 5.00
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "ATC2STAT.OCX"
Begin VB.Form F_CompanyCarChecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Car Checker"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin atc2stat.TCSStatus sts 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   4605
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstChecks 
      Height          =   2790
      Left            =   90
      TabIndex        =   8
      Top             =   270
      Width           =   3705
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "P&review"
      Height          =   375
      Left            =   3915
      TabIndex        =   7
      Top             =   1485
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3915
      TabIndex        =   6
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3915
      TabIndex        =   5
      Top             =   675
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   4
      Top             =   270
      Width           =   1200
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1140
      Left            =   3870
      TabIndex        =   3
      Top             =   1890
      Width           =   1230
      Begin VB.PictureBox pctFrame 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   45
         ScaleHeight     =   825
         ScaleWidth      =   1140
         TabIndex        =   11
         Top             =   225
         Width           =   1140
         Begin VB.OptionButton optOptions 
            Caption         =   "Check"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   13
            Top             =   45
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optOptions 
            Caption         =   "Change"
            Height          =   285
            Index           =   1
            Left            =   45
            TabIndex        =   12
            Top             =   360
            Width           =   1005
         End
      End
   End
   Begin VB.Frame fraPicture 
      Height          =   1500
      Left            =   90
      TabIndex        =   1
      Top             =   3060
      Width           =   5010
      Begin VB.Label lblCO2Warning 
         Caption         =   "Please also use the report wizard to check for zero value or missing data e.g. Company cars with no CO2 emissions figure entered. "
         ForeColor       =   &H000000FF&
         Height          =   690
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   3390
      End
      Begin VB.Image imgWizard 
         Height          =   885
         Left            =   135
         Picture         =   "F_COCARCh.frx":0000
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label lblCaption 
         Caption         =   "The wizard will check your company car data for inconsistencies."
         Height          =   450
         Left            =   1440
         TabIndex        =   2
         Top             =   225
         Width           =   3390
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Checks"
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1950
   End
End
Attribute VB_Name = "F_CompanyCarChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Prg As Object 'RK previously TCSProgressBar
Private m_Panel As TCSPANEL
Private m_ey As Employer
Private Enum CHECK_DATA
  CD_CHECK
  CD_CHANGE
End Enum
'km 13/06/02
Private m_tmpReg As String
Private m_tmpEe As String
Private Sub cmdCancel_Click()
  Me.Hide
End Sub
Public Sub Start(ey As Employer)
  On Error GoTo Start_ERR
  Set m_Prg = sts.prg
  Set m_Panel = sts.AddPanel(20, "", Down3D, , MDIMain.imlTree.ListImages(IMG_INFO).Picture)
  If ey Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "Start", "The employer is nothing.")
  Set m_ey = ey
  Call SettingsToScreen
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
Start_END:
  Set m_Prg = Nothing
  Set m_Panel = Nothing
  Exit Sub
Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Start", "Start", "Error in Start of CompanyCarChecker.")
  Resume Start_END
  Resume
End Sub

Private Sub CompanyCarCheckerDates(rsRecordChanges As Recordset, rsdata As Recordset)

  Dim cars() As COMPANY_CAR_CHECK
  Dim i As Long

  
  On Error GoTo CompanyCarCheckerDates_ERR
  
  Call xSet("CompanyCarCheckerDates")
  
  
    Do While Not rsdata.EOF
      i = GetCompanyCarsForChecker(cars, rsdata)
      If i = 0 Then
        GoTo CompanyCarCheckerDates_END
      ElseIf i > 1 Then
        If AnalyseCompanyCarsForCheckerDates(rsdata, rsRecordChanges, cars, i) Then
          Call WriteCompanyCarsForChecker(rsdata, rsRecordChanges, cars, i)
        End If
      End If
    Loop

  
CompanyCarCheckerDates_END:
  
  Call xReturn("CompanyCarCheckerDates")
  Exit Sub
CompanyCarCheckerDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ICompanyCarCheckerDates", "Company Car Checker Dates", "Error in CompanyCarCheckerDates.")
  Resume CompanyCarCheckerDates_END
  Resume
End Sub
Private Function CompanyCarCheckerFuncStart(rsRecordChanges As Recordset, rsdata As Recordset, ByVal QN As QUERY_NAMES) As Boolean
  On Error GoTo CompanyCarCheckerFuncStart_ERR
  
  If Not m_ey.MoveMenuUpdateEmployee Then Call Err.Raise(ERR_MOVE_MENU_UPDATE_EMPLOYEE, "CompanyCarCheckerFuncStart", "Failed to update employee.")
  
  Call m_ey.KillEmployeesBenefits
  
  If m_ey.db Is Nothing Then Call Err.Raise(ERR_DB_IS_NOTHING, "IProgress_Progress", "Employer db is nothing when trying company car checker.")
    
  Set rsdata = m_ey.db.OpenRecordset(sql.Queries(QN), dbOpenDynaset)

  If rsdata Is Nothing Then Call Err.Raise(ERR_NORECORDS, "CompanyCarCheckerFuncStart", "Error getting company car checker cars.")

  Call m_ey.db.Execute(sql.Queries(DELETE_COMPANYCAR_CHECKER_CARS_LOG))
  'to go into master function
  ' SELECT_COMPANYCAR_CHECKER_CARS_LOG_ALL is simply the whole recordset so is safe

  Set rsRecordChanges = m_ey.db.OpenRecordset(sql.Queries(SELECT_COMPANYCAR_CHECKER_CARS_LOG_ALL), dbOpenDynaset)
  If rsRecordChanges Is Nothing Then Call Err.Raise(ERR_RS_IS_NOTHING, "IProgress_Progress", "The rsRecordChanges is nothing.")

  If Not (rsdata.EOF And rsdata.BOF) Then
    rsdata.MoveLast
    m_Prg.Max = rsdata.RecordCount
    rsdata.MoveFirst
  End If

  CompanyCarCheckerFuncStart = True
  
CompanyCarCheckerFuncStart_END:
  Exit Function
CompanyCarCheckerFuncStart_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CompanyCarCheckerFuncStart"), Err.Description)
End Function
Private Sub CompanyCarCheckerFuncEnd()
  On Error GoTo CompanyCarCheckerFuncEnd_END
  
  If Not m_ey.CurrentEmployee Is Nothing Then
    m_ey.CurrentEmployee.LoadBenefits (TBL_ALLBENEFITS)
    If IsBenefitForm(CurrentForm) And Not CurrentForm Is F_Employees Then
      Call BenScreenSwitchEnd(CurrentForm)
    End If
  End If
    
CompanyCarCheckerFuncEnd_END:
  Exit Sub
CompanyCarCheckerFuncEnd_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CompanyCarCheckerFuncEnd"), Err.Description)
End Sub
Private Sub AddRecordToCar(cars() As COMPANY_CAR_CHECK, ByVal iUBound As Long, rsdata As Recordset)
    
    Dim EmpNumbStr As String
    Dim EmpRS As Recordset
    
    ReDim Preserve cars(1 To iUBound)
    
    cars(iUBound).Employee_db = IsNullEx(rsdata.Fields("employee").value, "") 'km
    cars(iUBound).PersonnelNumber_db = rsdata.Fields("Personnel number").value
    cars(iUBound).Replaced_db = rsdata.Fields("Replaced").value
    cars(iUBound).Replacement_db = rsdata.Fields("Replacement").value
    cars(iUBound).SecondCar_db = rsdata.Fields("SecondCar").value
    cars(iUBound).Registration_db = rsdata.Fields("Registration").value
    cars(iUBound).AvailableFrom_db = rsdata.Fields("Available from").value
    cars(iUBound).OldAvailableTo_db = rsdata.Fields("Available To").value
    cars(iUBound).RegistrationReplaced_db = IIf(IsNull(rsdata.Fields("Registration Replaced").value), "", rsdata.Fields("Registration Replaced").value)
    cars(iUBound).MakeModelReplaced_db = IIf(IsNull(rsdata.Fields("MakeModelReplaced").value), "", rsdata.Fields("MakeModelReplaced").value)
    cars(iUBound).DateCarReplaced_db = IIf(IsNull(rsdata.Fields("DateCarReplaced").value), UNDATED, rsdata.Fields("DateCarReplaced").value)
    cars(iUBound).MakeAndModel_db = rsdata.Fields("MakeAndModel").value
    'IK, this is wring CAD, prior to release for June/July 2004
    cars(iUBound).FuelAvailableFrom_db = cars(iUBound).AvailableFrom_db
    cars(iUBound).FuelOldAvailableTo_db = rsdata.Fields("FuelTo").value
'MP DB ToDo - below line in use?
    cars(iUBound).DaysUnavailable_db = rsdata.Fields("Days unavailable").value
    
'MP DB (not used)      cars(iUBound).FuelNewAvailableTo = UNDATED
    cars(iUBound).NewAvailableTo_db = UNDATED
    If rsdata.Fields("NumberOfUsers").value Then cars(iUBound).OldNumberOfUsers_db = rsdata.Fields("NumberOfUsers").value
    cars(iUBound).NewNumberOfUsers_db = -1
    '/IK
  
    cars(iUBound).OldDateRegistered_db = rsdata.Fields("regdate").value
    cars(iUBound).NewDateRegistered_db = UNDATED
    cars(iUBound).EmployeeStartDate_db = IsNullEx(rsdata.Fields("joined").value, UNDATED)
    cars(iUBound).EmployeeLeaveDate_db = IsNullEx(rsdata.Fields("Left").value, UNDATED)
    
           
           
     ' r.d.c
     ' Need to speak with Charlie
'    If p11d32.AppYear > 2000 Then ' Fetch Employee Join/Start date for Company Car Checker comparison
      'CAD review 20/02 join to employess and add start/ed date
'      EmpNumbStr = "select * from T_Employees where P_Num = '" & cars(iUBound).PersonnelNumber & "';"
'      Set EmpRS = m_ey.db.OpenRecordset(EmpNumbStr, dbOpenDynaset)
'      cars(iUBound).EmployeeStartDate = IsNullEx(EmpRS("joined"), UNDATED)
'      cars(iUBound).EmployeeLeaveDate = IsNullEx(EmpRS("left"), UNDATED)
  '  End If

End Sub
Private Sub CompanyCarCheckerOverlaps(rsRecordSetChanges As Recordset, rsdata As Recordset)
  Dim i As Long, j As Long
  Dim sCurrentReg As String
  Dim sCurrentEe As String          'RH
  Dim cars() As COMPANY_CAR_CHECK
      
  On Error GoTo CompanyCarCheckerOverlaps_ERR
  
  Call xSet("CompanyCarCheckerOverlaps")
  
  ReDim cars(1 To 1)
  
  Do While Not rsdata.EOF
      sCurrentReg = rsdata.Fields("Registration")
      sCurrentEe = rsdata.Fields("personnel number")  'RH
      Call StepCaptionChecker(rsdata)
      i = 1
      rsdata.MoveNext
      Do While Not rsdata.EOF
        If StrComp(rsdata.Fields("Registration"), sCurrentReg, vbTextCompare) = 0 Then
          If StrComp(rsdata.Fields("personnel number"), sCurrentEe, vbTextCompare) <> 0 Then   'RH
            i = i + 1
          Else
            Exit Do
           End If
        Else
          Exit Do
        End If
        Call StepCaptionChecker(rsdata)
        rsdata.MoveNext
      Loop
      
      If i > 1 Then
        For j = 1 To i
          rsdata.MovePrevious
        Next
        For j = 1 To i
          Call AddRecordToCar(cars, j, rsdata)
          rsdata.MoveNext
        Next
        For j = 1 To i
          rsdata.MovePrevious
        Next
        
        
        'km 13/06/02 - take temporary storage of current record
        'cad review 28/06/2002 Why modular, why not checked for normal cases of duPlicates....no testing !!!!
        Call AnalyseCompanyCarsForCheckerOverlaps(cars, i, rsRecordSetChanges)
        Call WriteCompanyCarsForCheckerOverlaps(cars, i, rsRecordSetChanges, rsdata)
        'km 13/06/02 - reset db to current record
        'rsdata.MoveFirst
        'Do While (StrComp(rsdata.Fields("personnel number") <> sCurrentEe, vbTextCompare) <> 0) Or (StrComp(mprsdata.Fields("registration"), sCurrentReg, vbTextCompare) <> 0)
        '  rsdata.MoveNext
        'Loop
      End If
    Loop
  
CompanyCarCheckerOverlaps_END:
  Call xReturn("CompanyCarCheckerDates")
  Exit Sub
CompanyCarCheckerOverlaps_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "CompanyCarCheckerDates"), Err.Description)
  Resume
End Sub

Private Function WriteCompanyCarsForChecker(rsdata As Recordset, rsRecordChanges As Recordset, cars() As COMPANY_CAR_CHECK, lNoCars As Long)
  Dim i As Long
  
  On Error GoTo AnalyseCompanyCarsForCheckerDates_ERR
  
  Call xSet("AnalyseCompanyCarsForCheckerDates")
  
  For i = 1 To lNoCars
    rsdata.MovePrevious
  Next
  
  For i = 1 To lNoCars
    If cars(i).Amended Then
      'write out new car
      rsdata.Edit
      rsRecordChanges.AddNew
      
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      
      rsRecordChanges!From = cars(i).AvailableFrom_db
      rsRecordChanges!OldTo = cars(i).OldAvailableTo_db
      rsRecordChanges!OldDateReg = cars(i).OldDateRegistered_db
      
      rsRecordChanges!Reg = cars(i).Registration_db
      rsRecordChanges!RegReplaced = cars(i).RegistrationReplaced_db
      rsRecordChanges!Replaced = cars(i).Replaced_db
      rsRecordChanges!Replacement = cars(i).Replacement_db
      
      If cars(i).AvailableToAmended Then
        rsdata![Available To] = cars(i).NewAvailableTo_db
        rsRecordChanges!NewTo = cars(i).NewAvailableTo_db
        rsRecordChanges!ToAmended = True
      End If
      
      If cars(i).DateRegisteredAmended Then
        rsdata![regdate] = cars(i).NewDateRegistered_db
        rsRecordChanges!NewDateReg = cars(i).NewDateRegistered_db
        rsRecordChanges!DateRegAmended = True
      End If
      
      rsdata!Replaced = cars(i).Replaced_db
      rsdata!Replacement = cars(i).Replacement_db
      rsdata![Registration Replaced] = cars(i).RegistrationReplaced_db
      rsdata!MakeModelReplaced = cars(i).MakeModelReplaced_db
      
      If cars(i).DateCarReplaced_db <> UNDATED Then rsdata!DateCarReplaced = cars(i).DateCarReplaced_db
      
      If optOptions(CD_CHANGE).value = True Then rsdata.Update
      rsRecordChanges.Update
      
    End If
    rsdata.MoveNext
  Next
  
AnalyseCompanyCarsForCheckerDates_END:
  Call xReturn("AnalyseCompanyCarsForCheckerDates")
  Exit Function
  
AnalyseCompanyCarsForCheckerDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseCompanyCarsForCheckerDates", "Analyse Company Cars For Checker", "Error analysing company cars for company car checker")
  If Not rsdata.EOF Then
    rsdata.MoveLast
    rsdata.MoveNext
  End If
  Resume AnalyseCompanyCarsForCheckerDates_END
  Resume
End Function
Private Function AddCompanyCArCheckRecord(rsRecordChanges As Recordset, sEmployee As String, sPersonnelNumber As String)
  'error checker
  On Error GoTo AddCompanyCArCheckRecord_ERR
  
  Call xSet("AddCompanyCArCheckRecord")
  
  
  rsRecordChanges.AddNew
  rsRecordChanges.Fields("Employee") = sEmployee
  rsRecordChanges.Fields("PNum") = sPersonnelNumber
  AddCompanyCArCheckRecord = True
  
  
AddCompanyCArCheckRecord_END:
  Call xReturn("AddCompanyCArCheckRecord")
  Exit Function
AddCompanyCArCheckRecord_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddCompanyCArCheckRecord", "Add Company Car Check Record", "Error adding a record to company car check.")
  Resume AddCompanyCArCheckRecord_END
  Resume
End Function
'CAD REVIEW 28/06/2002 REMOVED , rsdata As Recordset AS NOT USING IT !!!
Private Function AnalyseCompanyCarsForCheckerOverlaps(cars() As COMPANY_CAR_CHECK, lNoCars As Long, rsRecordChanges As Recordset) As Boolean
  Dim i As Long, j As Long, lCarWithMaxDate As Long
  Dim d As Date
  
  Dim s As String
  On Error GoTo AnalyseCompanyCarsForCheckerOverlaps_ERR
  
  Call xSet("AnalyseCompanyCarsForCheckerDates")
  
  d = UNDATED
  
  For i = 1 To lNoCars
    s = ""
    For j = 1 To lNoCars
      If j <> i Then
        If DateInRange(cars(i).AvailableFrom_db, cars(j).AvailableFrom_db, cars(j).OldAvailableTo_db) Or DateInRange(cars(i).OldAvailableTo_db, cars(j).AvailableFrom_db, cars(j).OldAvailableTo_db) Then
          If Len(s) Then
            s = s & ", " & cars(j).PersonnelNumber_db
          Else
            s = "Overlaps with " & cars(j).PersonnelNumber_db
          End If
          'km 13/06/02
          cars(i).Amended = True
        End If
        If cars(i).OldDateRegistered_db <> cars(j).OldDateRegistered_db Then
          If d < cars(i).OldDateRegistered_db Then
            lCarWithMaxDate = i
            d = cars(i).OldDateRegistered_db
          End If
          'km 13/06/02
          cars(i).Amended = True
        End If
        'km 13/06/02 - this should only be set if there is an overlap / inconsistency
        'cars(i).Amended = True
      End If
    Next
  
    'cehcking number of employees sharing the car     IK 20/06/2003
    If cars(i).OldNumberOfUsers_db <> lNoCars Then
      cars(i).NewNumberOfUsers_db = lNoCars
    End If
    cars(i).Comments_db = s
  Next
  
  If d <> UNDATED Then
    For i = 1 To lNoCars
      If i <> lCarWithMaxDate Then
        cars(i).NewDateRegistered_db = d
        cars(i).DateRegisteredAmended = True
      End If
    Next
  End If
AnalyseCompanyCarsForCheckerOverlaps_END:
  Call xReturn("AnalyseCompanyCarsForCheckerOverlaps")
  Exit Function
AnalyseCompanyCarsForCheckerOverlaps_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseCompanyCarsForCheckerOverlaps", "Analyse Company Cars For Checker Overlaps", "Error analysing company cars for company car checker")
  Resume AnalyseCompanyCarsForCheckerOverlaps_END
  End Function
Private Sub WriteCompanyCarsForCheckerOverlaps(cars() As COMPANY_CAR_CHECK, ByVal lNoCars As Long, rsRecordChanges As Recordset, rsdata As Recordset)
  Dim i As Long
  
  On Error GoTo WriteCompanyCarsForCheckerOverlaps_ERR
  
  For i = 1 To lNoCars
    If cars(i).Amended Then
'MP DB ToDo new record added in AddCompanyCArCheckRecord - should add another one here?
      rsRecordChanges.AddNew
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      rsRecordChanges.Fields("Employee") = cars(i).Employee_db
      rsRecordChanges.Fields("PNum") = cars(i).PersonnelNumber_db
      rsRecordChanges!From = cars(i).AvailableFrom_db
      rsRecordChanges!OldTo = cars(i).OldAvailableTo_db
      rsRecordChanges!Reg = cars(i).Registration_db
      rsRecordChanges!Comments = cars(i).Comments_db
      rsRecordChanges!OldDateReg = cars(i).OldDateRegistered_db
      
      If cars(i).NewDateRegistered_db <> UNDATED Then
        rsRecordChanges!NewDateReg = cars(i).NewDateRegistered_db
      End If
      If cars(i).DateRegisteredAmended Then
        If optOptions(CD_CHANGE).value = True Then
          rsdata.Edit
          rsdata![regdate] = cars(i).NewDateRegistered_db
          rsdata.Update
        End If
      End If
    
      'displaying number of users fields   IK 20/06/2003
      rsRecordChanges!OldNumOfUsers = cars(i).OldNumberOfUsers_db
      If cars(i).NewNumberOfUsers_db >= 0 Then
        rsRecordChanges!NewNumOfUsers = cars(i).NewNumberOfUsers_db
        'correcting if selected by user
        If optOptions(CD_CHANGE).value = True Then
          rsdata.Edit
          rsdata![NumberOfUsers] = cars(i).NewNumberOfUsers_db
          rsdata.Update
        End If
      End If
      
      'only update changes if Number of users is wrong OR registration date is wrong
      If (cars(i).NewNumberOfUsers_db >= 0) Or (cars(i).NewDateRegistered_db <> UNDATED) Then
        rsRecordChanges.Update
      End If
      
    End If
    rsdata.MoveNext
    'km 13/06/02
    cars(i).Amended = False
  Next
  
WriteCompanyCarsForCheckerOverlaps_END:
  Exit Sub
WriteCompanyCarsForCheckerOverlaps_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteCompanyCarsForCheckerOverlaps", "Write Company Cars For Checker Overlaps", "Error in WriteCompanyCarsForCheckerOverlaps")
  Resume WriteCompanyCarsForCheckerOverlaps_END
End Sub
Private Function AnalyseCompanyCarsForCheckerDates(rsdata As Recordset, rsRecordChanges As Recordset, cars() As COMPANY_CAR_CHECK, lNoCars As Long) As Boolean
  Dim i As Long
  Dim j As Long
  
    
  On Error GoTo AnalyseCompanyCarsForCheckerDates_ERR
  
  Call xSet("AnalyseCompanyCarsForCheckerDates")
   
  'CHECK FOR ANY TWO CARS having same Available From date
  For i = 1 To lNoCars - 1
    For j = i + 1 To lNoCars
       If cars(i).AvailableFrom_db = cars(j).AvailableFrom_db Then
        Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
        rsRecordChanges.Fields("Comments") = "2 or more cars same to date."
        rsRecordChanges.Update
        'reject
        GoTo AnalyseCompanyCarsForCheckerDates_END
      End If
    Next
  Next
  'CHECK FOR ANY SECOND CARS
  For i = 1 To lNoCars
    If cars(i).SecondCar_db Then
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      rsRecordChanges.Fields("Comments") = "Second car ticked."
      rsRecordChanges.Update
      'reject
      GoTo AnalyseCompanyCarsForCheckerDates_END
    End If
  Next
  
  For i = 1 To lNoCars - 1
  
    If cars(i).OldAvailableTo_db >= cars(i + 1).AvailableFrom_db Then
      cars(i).NewAvailableTo_db = DateAdd("d", -1, cars(i + 1).AvailableFrom_db)
      cars(i).AvailableToAmended = True
      cars(i).Amended = True
      AnalyseCompanyCarsForCheckerDates = AnalyseCompanyCarsForCheckerDates Or True
    End If
    
    If Not cars(i).Replaced_db Then
      cars(i).Replaced_db = True
      cars(i).Amended = True
      AnalyseCompanyCarsForCheckerDates = AnalyseCompanyCarsForCheckerDates Or True
    End If
    
    'km 11/06/02 If (Not cars(i + 1).Replacement) Or (StrComp(cars(i).MakeAndModel, cars(i + 1).MakeModelReplaced) <> 0) Or (StrComp(cars(i).Registration, cars(i + 1).RegistrationReplaced) <> 0) Then
    'km 11/06/02
    If Not ((cars(i + 1).Replacement_db) And (StrComp(cars(i).MakeAndModel_db, cars(i + 1).MakeModelReplaced_db) = 0) And (StrComp(cars(i).Registration_db, cars(i + 1).RegistrationReplaced_db) = 0)) Then
      cars(i + 1).Replacement_db = True
      cars(i + 1).Amended = True
      AnalyseCompanyCarsForCheckerDates = True
      cars(i + 1).RegistrationReplaced_db = cars(i).Registration_db
      cars(i + 1).MakeModelReplaced_db = cars(i).MakeAndModel_db
      cars(i + 1).RegistrationReplaced_db = cars(i).Registration_db
      If cars(i).AvailableToAmended Then
        cars(i + 1).DateCarReplaced_db = cars(i).NewAvailableTo_db
      Else
        cars(i + 1).DateCarReplaced_db = cars(i).OldAvailableTo_db
      End If
    End If
    
    AnalyseCompanyCarsForCheckerDates = AnalyseCompanyCarsForCheckerDates Or CompanyCarCheckerRegDate(cars(i))
    If i = lNoCars - 1 Then AnalyseCompanyCarsForCheckerDates = AnalyseCompanyCarsForCheckerDates Or CompanyCarCheckerRegDate(cars(i + 1))
  Next
  
  
  
AnalyseCompanyCarsForCheckerDates_END:
  Call xReturn("AnalyseCompanyCarsForCheckerDates")
  Exit Function
AnalyseCompanyCarsForCheckerDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseCompanyCarsForCheckerDates", "Analyse Company Cars For Checker", "Error analysing company cars for company car checker.")
  Resume AnalyseCompanyCarsForCheckerDates_END
  Resume
End Function
Private Function CompanyCarCheckerRegDate(car As COMPANY_CAR_CHECK)
  If car.OldDateRegistered_db > car.AvailableFrom_db Then
    car.DateRegisteredAmended = True
    car.NewDateRegistered_db = car.AvailableFrom_db
    car.Amended = True
    CompanyCarCheckerRegDate = True
  End If
End Function
Private Sub StepCaptionChecker(rs As Recordset)
  If m_Prg Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "StepCaptionChecker", "The progress bar is nothing.")
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "StepCaptionChecker", "The data recordset is nothing.")
  m_Prg.StepCaption ("Personnel number: " & rs.Fields("Personnel number").value & " ,Car: " & rs![Registration])
End Sub
Private Function GetCompanyCarsForChecker(cars() As COMPANY_CAR_CHECK, rs As Recordset) As Long
  Dim sPNum As String
  
  On Error GoTo GetCompanyCarsForChecker_ERR
  
  Call xSet("GetCompanyCarsForChecker")
  
  ReDim cars(1 To 1) 'erase the previous
  
  'get the cars with same employee ref
  If rs.EOF Then GoTo GetCompanyCarsForChecker_END
  sPNum = rs.Fields("Personnel number").value
  Do While Not rs.EOF
    If StrComp(sPNum, rs.Fields("Personnel number").value, vbBinaryCompare) = 0 Then
      Call StepCaptionChecker(rs)
      GetCompanyCarsForChecker = GetCompanyCarsForChecker + 1
      Call AddRecordToCar(cars, GetCompanyCarsForChecker, rs)
    Else
      Exit Do
    End If
    rs.MoveNext
  Loop
  
  
GetCompanyCarsForChecker_END:
  Call xReturn("GetCompanyCarsForChecker")
  Exit Function
  
GetCompanyCarsForChecker_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetCompanyCarsForChecker", "Get Company Cars For Checker", "Error getting the company cars for company car checker.")
  Resume GetCompanyCarsForChecker_END
  Resume
End Function
Private Sub CompanyCarCheckerLog(ByVal RD As REPORT_TARGET)
  Dim QN As QUERY_NAMES
  Dim sTitle As String
  
  On Error GoTo CompanyCarCheckerLog_ERR
  
  Select Case p11d32.CompanyCarCheckerCheck
    Case CCCC_DATES
      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_DATES
    Case CCCC_OVERLAPS
      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_OVERLAPS
    Case CCCC_NOCARS
      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_NOCARS
'    Case CCCC_FuelAvailDates
'      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_FUELAVAILDATES
    Case CCCC_AvailDates
      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_AVAILDATES
    Case CCCC_RegDates
      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_REGDATES
    Case Else
      Call ECASE("Ivalid CompanyCarCheckerCheck = " & p11d32.CompanyCarCheckerCheck)
  End Select
  Call StartAutoSTD(sql.Queries(QN), m_ey.db, "Notes from company car checker.", RD, "Altered cars:")
  
CompanyCarCheckerLog_END:
  Exit Sub
CompanyCarCheckerLog_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CompanyCarCheckerLog", "Company Car Checker Log", "Error in CompanyCarCheckerLog.")
  Resume CompanyCarCheckerLog_END
End Sub

Private Sub cmdOK_Click()
  Call CompanyCarChecker
End Sub
Private Sub CompanyCarChecker()
  Dim rsdata As Recordset
  Dim rsRecordChanges As Recordset
  Dim CCCMT As COMPANY_CAR_CHECKER_MESSAGE_TYPE
  On Error GoTo CompanyCarChecker_ERR
  If (p11d32.CompanyCarCheckerCheck < [_CCCC_FIRST_ITEM]) Or (p11d32.CompanyCarCheckerCheck > [_CCCC_LAST_ITEM]) Then GoTo CompanyCarChecker_END
  If optOptions(CD_CHANGE).value = True Then
    CCCMT = CCCMT_ALERT_MESSAGE_CHANGE
  Else
    CCCMT = CCCMT_ALERT_MESSAGE_CHECK
  End If
  
  'Apply company car checker
  If MultiDialog(CCCListCaption(p11d32.CompanyCarCheckerCheck, cccmt_list_item), CCCListCaption(p11d32.CompanyCarCheckerCheck, CCCMT), "&OK", "&Cancel") = 2 Then GoTo CompanyCarChecker_END
  Select Case p11d32.CompanyCarCheckerCheck
    Case CCCC_DATES
      Call CompanyCarCheckerFuncStart(rsRecordChanges, rsdata, SELECT_COMPANYCAR_CHECKER_CARS_DATES)
      Call CompanyCarCheckerDates(rsRecordChanges, rsdata)
    Case CCCC_OVERLAPS
      'Call CompanyCarCheckerFuncStart(rsRecordChanges, rsdata, SELECT_COMPANYCAR_CHECKER_CARS_OVERLAPS)
      'km 11/06/02 - needs to use regdates query (as ordered by registration)
      Call CompanyCarCheckerFuncStart(rsRecordChanges, rsdata, SELECT_COMPANYCAR_CHECKER_CARS_REGDATES)
      Call CompanyCarCheckerOverlaps(rsRecordChanges, rsdata)
    Case CCCC_NOCARS
      Call CompanyCarCheckerFuncStart(rsRecordChanges, rsdata, SELECT_COMPANYCAR_CHECKER_CARS_NOCARS)
      Call CompanyCarCheckerNoCars(rsRecordChanges, rsdata)
'Fuel available dates are no longer used, so have removed EK.
'    Case CCCC_FuelAvailDates
'      Call CompanyCarCheckerFuncStart(rsRecordChanges, rsdata, SELECT_COMPANYCAR_CHECKER_CARS_FUELAVAILDATES)
'      Call CompanyCarCheckerFuelAvailDates(rsRecordChanges, rsdata)
    Case CCCC_AvailDates
      Call CompanyCarCheckerFuncStart(rsRecordChanges, rsdata, SELECT_COMPANYCAR_CHECKER_CARS_AVAILDATES)
      Call CompanyCarCheckerAvailDates(rsRecordChanges, rsdata)
    Case CCCC_RegDates
      Call CompanyCarCheckerFuncStart(rsRecordChanges, rsdata, SELECT_COMPANYCAR_CHECKER_CARS_REGDATES)
      Call CompanyCarCheckerRegDates(rsRecordChanges, rsdata)
    'no case else as dealt with above
  End Select
  Call CompanyCarCheckerFuncEnd
CompanyCarChecker_END:
  Exit Sub
CompanyCarChecker_ERR:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "CompanyCarChecker"), "Company Car Checker", "Error in company car checker, choice = " & p11d32.CompanyCarCheckerCheck)
  Resume CompanyCarChecker_END
  Resume
End Sub

Private Sub cmdPreview_Click()
  Call CompanyCarCheckerLog(PREPARE_REPORT)
''  Dim frmComCarChange As F_CompanyCarCheckerChange
''  Dim QN As QUERY_NAMES
''  Dim sTitle As String
''  Dim rs As Recordset
''  Dim ac As AutoClass
''  Select Case p11d32.CompanyCarCheckerCheck
''    Case CCCC_DATES
''      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_DATES
''    Case CCCC_OVERLAPS
''      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_OVERLAPS
''    Case CCCC_NOCARS
''      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_NOCARS
'''    Case CCCC_FuelAvailDates
'''      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_FUELAVAILDATES
''    Case CCCC_AvailDates
''      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_AVAILDATES
''    Case CCCC_RegDates
''      QN = SELECT_COMPANYCAR_CHECKER_CARS_LOG_REGDATES
''    Case Else
''      Call ECASE("Ivalid CompanyCarCheckerCheck = " & p11d32.CompanyCarCheckerCheck)
''  End Select
''
'  Set frmComCarChange = New F_CompanyCarCheckerChange
'  Set rs = m_ey.db.OpenRecordset(sql.Queries(QN), dbOpenSnapshot)
'  Set ac = New AutoClass
'
'  'frmComCarChange.grdCarChecker.AllowAddNew = True
'  'frmComCarChange.grdCarChecker.AllowDelete = True
'  'frmComCarChange.grdCarChecker.AllowUpdate = True
'
'  'Call ac.InitAutoData("CompanyCarChecker", rs)
'
'  Call ac.ShowGrid(frmComCarChange.grdCarChecker)
'
'  frmComCarChange.Show vbModal
  
End Sub

Private Sub cmdPrint_Click()
  Call CompanyCarCheckerLog(PRINT_REPORT)
End Sub
Private Function CCCListCaption(ByVal CCCC As COMPANY_CAR_CHECKER_CHECKS, ByVal CCCMT As COMPANY_CAR_CHECKER_MESSAGE_TYPE) As String
  Dim s As String
  On Error GoTo CCCListCaption_ERR
  
  
  Select Case CCCC
    Case CCCC_DATES
      Select Case CCCMT
        Case cccmt_list_item
          s = "Employees with two or more overlapping cars"
        Case CCCMT_ALERT_MESSAGE_CHANGE, CCCMT_ALERT_MESSAGE_CHECK
          s = "Employees who have two or more cars available to them on one day." & vbCrLf & vbCrLf & _
              "It will indicate P46(Car) flags that should be set whether the cars have registration dates after the date of first use."
        Case Else
          ECASE ("Invalid Company car check message type = " & CCCMT)
      End Select
      
    Case CCCC_OVERLAPS
      Select Case CCCMT
        Case cccmt_list_item
          s = "Cars in use by more than one employee"
        Case CCCMT_ALERT_MESSAGE_CHANGE, CCCMT_ALERT_MESSAGE_CHECK
          s = "Cars that are in use by two or more employees at the same time"
        Case Else
          ECASE ("Invalid Company car check message type = " & CCCMT)
      End Select
    
    Case CCCC_NOCARS
      Select Case CCCMT
        Case cccmt_list_item
          s = "Employees with gaps between car usage"
        Case CCCMT_ALERT_MESSAGE_CHANGE, CCCMT_ALERT_MESSAGE_CHECK
          s = "Employees not having a car available for all days in the tax year"
        Case Else
          ECASE ("Invalid Company car check message type = " & CCCMT)
      End Select
    
    Case CCCC_AvailDates
      Select Case CCCMT
        Case cccmt_list_item
          s = "Cars start date occurring after end date"
        Case CCCMT_ALERT_MESSAGE_CHANGE, CCCMT_ALERT_MESSAGE_CHECK
          s = "Any cars that have their 'Available to' date before their 'Available from' date"
        Case Else
          ECASE ("Invalid Company car check message type = " & CCCMT)
      End Select
      
      'EK removed as no longer store fuel available dates separately.
'    Case CCCC_FuelAvailDates
'      Select Case CCCMT
'        Case cccmt_list_item
'          s = "Fuel dates are inconsistent with car dates"
'        Case CCCMT_ALERT_MESSAGE_CHANGE, CCCMT_ALERT_MESSAGE_CHECK
'          s = "Fuel dates that are inconsistent with car availability dates"
'        Case Else
'          ECASE ("Invalid Company car check message type = " & CCCMT)
'      End Select
            
    Case CCCC_RegDates
      Select Case CCCMT
        Case cccmt_list_item
          s = "Registration dates of cars inconsistent"
        Case CCCMT_ALERT_MESSAGE_CHANGE, CCCMT_ALERT_MESSAGE_CHECK
          s = "Registration dates that are inconsistent with availability dates"
        Case Else
          ECASE ("Invalid Company car check message type = " & CCCMT)
      End Select
      
    Case Else
      Call ECASE("Invalid Company car check = " & CCCC)
     
  End Select
  
  
  Select Case CCCMT
    Case CCCMT_ALERT_MESSAGE_CHANGE
      s = S_CCCC_MESSAGE_PREFIX_CHANGE & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf & S_CCCC_MESSAGE_SUFFIX
    Case CCCMT_ALERT_MESSAGE_CHECK
      s = S_CCCC_MESSAGE_PREFIX_CHECK & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf & S_CCCC_MESSAGE_SUFFIX
  End Select

  CCCListCaption = s

CCCListCaption_END:
  Exit Function
CCCListCaption_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CCCListCaption", "CCC List Caption", "Error in CCCListCaption.")
  Resume CCCListCaption_END
End Function

Private Sub SettingsToScreen()
  Dim i As Long
  
  On Error GoTo SettingsToScreen_Err
  
  For i = [_CCCC_FIRST_ITEM] To [_CCCC_LAST_ITEM]
    Call lstChecks.AddItem(CCCListCaption(i, cccmt_list_item))
  Next i

  If lstChecks.ListCount > ([_CCCC_LAST_ITEM] + 1) Then Call Err.Raise(ERR_INVALID_COMPANY_CAR_CHECKER_CHECK, "SettingsToScreen", "Number of options in list should be " & [_CCCC_LAST_ITEM] + 1 & " count = " & lstChecks.ListCount)
  lstChecks.ListIndex = p11d32.CompanyCarCheckerCheck
  
  
SettingsToScreen_End:
  Exit Sub
SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error in SettingsToScreen.")
  Resume SettingsToScreen_End
  Resume
End Sub


Private Sub lstChecks_Click()
  If lstChecks.ListIndex <> -1 Then
    p11d32.CompanyCarCheckerCheck = lstChecks.ListIndex
  End If
End Sub


Private Sub CompanyCarCheckerNoCars(rsRecordChanges As Recordset, rsdata As Recordset)

  Dim cars() As COMPANY_CAR_CHECK
  Dim i As Long

  
  On Error GoTo CompanyCarCheckerNoCars_ERR
  
  Call xSet("CompanyCarCheckerNoCars")
  
    Do While Not rsdata.EOF
      i = GetCompanyCarsForChecker(cars, rsdata)
      If i = 0 Then
        GoTo CompanyCarCheckerNoCars_END
      ElseIf i >= 1 Then
        'crap, crap , crap we were passing rsdata in ....why ?????? TAKEN OUT NOW
        If AnalyseCompanyCarsForCheckerNoCars(rsRecordChanges, cars, i) Then
          Call WriteCompanyCarsForCheckerNoCars(rsRecordChanges, cars, i)
        End If
      End If
    Loop

  
CompanyCarCheckerNoCars_END:
  
  Call xReturn("CompanyCarCheckerNoCars")
  Exit Sub
CompanyCarCheckerNoCars_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ICompanyCarCheckerNoCars", "Company Car Checker No Cars", "Error in CompanyCarCheckerNoCars.")
  Resume CompanyCarCheckerNoCars_END
  Resume
  
End Sub

Private Function AnalyseCompanyCarsForCheckerNoCars(rsRecordChanges As Recordset, cars() As COMPANY_CAR_CHECK, lNoCars As Long) As Boolean
 
  Dim i As Long
  Dim j As Long
  Dim k As Long
      
  Dim EmployeeStartDate As Date
  Dim EmployeeLeaveDate As Date
  
On Error GoTo AnalyseCompanyCarsForCheckerNoCars_ERR
  
  Call xSet("AnalyseCompanyCarsForCheckerNoCars")
   
  'Check for any employee who had a car at some time but not for the whole year
    
  For i = 1 To lNoCars
    If cars(i).SecondCar_db = False Then
      j = DateDiff("d", cars(i).AvailableFrom_db - 1, cars(i).OldAvailableTo_db)
      k = k + j
    End If
  Next

'  If p11d32.AppYear > 2000 Then
  ' RDC Set date for "Employees with gaps in car useage" to earlier of joining date OR start of tax year
     If IsDate(cars(i - 1).EmployeeStartDate_db) And ((cars(i - 1).EmployeeStartDate_db) <> UNDATED) Then
      If cars(i - 1).EmployeeStartDate_db < p11d32.Rates.value(TaxYearStart) Then
        EmployeeStartDate = p11d32.Rates.value(TaxYearStart)
      Else
        EmployeeStartDate = cars(i - 1).EmployeeStartDate_db
      End If
    Else
      EmployeeStartDate = p11d32.Rates.value(TaxYearStart)
    End If
  
    If IsDate(cars(i - 1).EmployeeLeaveDate_db) And ((cars(i - 1).EmployeeLeaveDate_db) <> UNDATED) Then
      EmployeeLeaveDate = cars(i - 1).EmployeeLeaveDate_db
    Else
      EmployeeLeaveDate = p11d32.Rates.value(TaxYearEnd)
    End If
  
'rdc compare period of having a car to employee join date and/or the tax year

    'CAD review 20/02
    'AnalyseCompanyCarsForCheckerNoCars =  AFunction(rsRecordChanges, cars(i-1),StartDate,EndDate, Comment) the goto AnalyseCompanyCarsForCheckerNoCars_END
    'If AnalyseCompanyCarsForCheckerNoCars Then GoTo AnalyseCompanyCarsForCheckerNoCars
    
    If k < DateDiff("d", EmployeeStartDate - 1, EmployeeLeaveDate) Then
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i - 1).Employee_db, cars(i - 1).PersonnelNumber_db)
      cars(i - 1).Comments_db = "Employee does not have a car available for all days in the tax year"
      'CAD review 20/02
      AnalyseCompanyCarsForCheckerNoCars = AnalyseCompanyCarsForCheckerNoCars Or True
      GoTo AnalyseCompanyCarsForCheckerNoCars_END
    End If
 ' Else
'    If k < DateDiff("d", p11d32.Rates.value(TaxYearStart) - 1, p11d32.Rates.value(TaxYearEnd)) Then
'      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i - 1).employee, cars(i - 1).PersonnelNumber)
'      cars(i - 1).Comments = "Employee does not have a car available for all days in the tax year"
'      'CAD review 20/02
'      AnalyseCompanyCarsForCheckerNoCars = AnalyseCompanyCarsForCheckerNoCars Or True
'      GoTo AnalyseCompanyCarsForCheckerNoCars_END
'    End If
'  End If
  
AnalyseCompanyCarsForCheckerNoCars_END:
  Call xReturn("AnalyseCompanyCarsForCheckerNoCars")
  Exit Function
AnalyseCompanyCarsForCheckerNoCars_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseCompanyCarsForCheckerNoCars", "Analyse Company Cars For Checker", "Error analysing company cars for company car checker")
  Resume AnalyseCompanyCarsForCheckerNoCars_END
  Resume
End Function


Private Function WriteCompanyCarsForCheckerNoCars(rsRecordChanges As Recordset, cars() As COMPANY_CAR_CHECK, lNoCars As Long)
  Dim i As Long

  On Error GoTo WriteCompanyCarsForCheckerNoCars_ERR
 
 
  For i = 1 To lNoCars
    'CAD this is crap !!!! why do we MOveFirst then find then loop forward again !!!!!!
    
    'Do While rsdata.Fields("personnel number") <> cars(i).PersonnelNumber Or rsdata.Fields("registration") <> cars(i).Registration
    '  rsdata.MoveNext
    'Loop
    'rsdata.Edit
    rsRecordChanges.AddNew
    Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
'MP DB ToDo - why add new record again?
    rsRecordChanges.AddNew
    rsRecordChanges.Fields("Employee") = cars(i).Employee_db
    rsRecordChanges.Fields("PNum") = cars(i).PersonnelNumber_db
    rsRecordChanges!From = cars(i).AvailableFrom_db
    rsRecordChanges!OldTo = cars(i).OldAvailableTo_db
    rsRecordChanges!OldDateReg = cars(i).OldDateRegistered_db
    rsRecordChanges!Reg = cars(i).Registration_db
    rsRecordChanges!Comments = cars(i).Comments_db
    rsRecordChanges.Update
    'rsdata.MoveNext
    'end this is crap
  Next
  
WriteCompanyCarsForCheckerNoCars_END:
  Call xReturn("AnalyseCompanyCarsForCheckerNoCars")
  Exit Function
WriteCompanyCarsForCheckerNoCars_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteCompanyCarsForCheckerNoCars", "Write Company Cars For Checker", "Error analysing company cars for company car checker.")
  Resume WriteCompanyCarsForCheckerNoCars_END
  Resume
End Function



Private Sub CompanyCarCheckerRegDates(rsRecordChanges As Recordset, rsdata As Recordset)
  Dim i As Long
  Dim cars() As COMPANY_CAR_CHECK
      
  On Error GoTo CompanyCarCheckerRegDates_ERR
  
  Call xSet("CompanyCarCheckerRegDates")
  
  Do While Not rsdata.EOF
      i = GetCompanyCarsForChecker(cars, rsdata)
      If i = 0 Then
        GoTo CompanyCarCheckerRegDates_END
      ElseIf i > 0 Then
        If AnalyseCompanyCarsForCheckerRegDates(cars, i, rsRecordChanges) Then
          Call WriteCompanyCarsForCheckerRegDates(cars, i, rsRecordChanges, rsdata)
        End If
      End If
    Loop
  
CompanyCarCheckerRegDates_END:
  
  Call xReturn("CompanyCarCheckerRegDates")
  Exit Sub
CompanyCarCheckerRegDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ICompanyCarCheckerRegDates", "Company Car Checker Registration Dates", "Error in CompanyCarCheckerRegDates.")
  Resume CompanyCarCheckerRegDates_END
  Resume
  
End Sub

Private Function AnalyseCompanyCarsForCheckerRegDates(cars() As COMPANY_CAR_CHECK, lNoCars As Long, rsRecordChanges As Recordset) As Boolean
'MP DB removed rsData as not in use in this proc
' Private Function AnalyseCompanyCarsForCheckerRegDates(cars() As COMPANY_CAR_CHECK, lNoCars As Long, rsRecordChanges As Recordset, rsdata As Recordset) As Boolean
  Dim i As Long
  Dim s As String
  Dim retValue As Boolean
  On Error GoTo AnalyseCompanyCarsForCheckerRegDates_ERR
  
  Call xSet("AnalyseCompanyCarsForCheckerRegDates")
  
  For i = 1 To lNoCars
    s = ""
    If cars(i).AvailableFrom_db < cars(i).OldDateRegistered_db Then
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      s = "Registration date is after car is first made available"
      cars(i).NewDateRegistered_db = cars(i).AvailableFrom_db
      cars(i).Amended = True
      cars(i).DateRegisteredAmended = True
      retValue = True
    End If
    cars(i).Comments_db = s
  Next
  
  AnalyseCompanyCarsForCheckerRegDates = retValue

AnalyseCompanyCarsForCheckerRegDates_END:
  Call xReturn("AnalyseCompanyCarsForCheckerRegDates")
  Exit Function
AnalyseCompanyCarsForCheckerRegDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseCompanyCarsForCheckerRegDates", "Analyse Company Cars For Registration Dates", "Error analysing company cars for company car checker")
  Resume AnalyseCompanyCarsForCheckerRegDates_END

End Function

Private Sub WriteCompanyCarsForCheckerRegDates(cars() As COMPANY_CAR_CHECK, ByVal lNoCars As Long, rsRecordChanges As Recordset, rsdata As Recordset)
  Dim i As Long
  
  On Error GoTo WriteCompanyCarsForCheckerRegDates_ERR
   
  'loop back by the number of cars
  For i = 1 To lNoCars
    Call rsdata.MovePrevious
  Next
  For i = 1 To lNoCars
    If cars(i).Amended Then
      
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      rsRecordChanges.Fields("Employee") = cars(i).Employee_db
      rsRecordChanges.Fields("PNum") = cars(i).PersonnelNumber_db
      rsRecordChanges!From = cars(i).AvailableFrom_db
      rsRecordChanges!OldTo = cars(i).OldAvailableTo_db
      rsRecordChanges!Reg = cars(i).Registration_db
      rsRecordChanges!Comments = cars(i).Comments_db
      rsRecordChanges!OldDateReg = cars(i).OldDateRegistered_db
      If cars(i).AvailableFrom_db <> UNDATED Then
        rsRecordChanges!NewDateReg = cars(i).NewDateRegistered_db
      End If
      
      If cars(i).DateRegisteredAmended Then
        If optOptions(CD_CHANGE).value = True Then
          rsdata.Edit
          rsdata![regdate] = cars(i).NewDateRegistered_db
          rsdata.Update
        End If
      End If
      rsRecordChanges.Update
    End If
    rsdata.MoveNext
  Next

WriteCompanyCarsForCheckerRegDates_END:
  Exit Sub
WriteCompanyCarsForCheckerRegDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteCompanyCarsForCheckerRegDates", "Write Company Cars For Checker Registration Dates", "Error in WriteCompanyCarsForCheckerRegDates")
  Resume WriteCompanyCarsForCheckerRegDates_END
  Resume
End Sub


'IK ...

Private Sub CompanyCarCheckerAvailDates(rsRecordChanges As Recordset, rsdata As Recordset)
  Dim i As Long
  Dim cars() As COMPANY_CAR_CHECK
      
  On Error GoTo CompanyCarCheckerAvailDates_ERR
  
  Call xSet("CompanyCarCheckerAvailDates")
  
  Do While Not rsdata.EOF
      i = GetCompanyCarsForChecker(cars, rsdata)
      If i = 0 Then
        GoTo CompanyCarCheckerAvailDates_END
      ElseIf i > 0 Then
        If AnalyseCompanyCarsForCheckerAvailDates(cars, i, rsRecordChanges, rsdata) Then
          Call WriteCompanyCarsForCheckerAvailDates(cars, i, rsRecordChanges, rsdata)
        End If
      End If
    Loop
  
CompanyCarCheckerAvailDates_END:
  
  Call xReturn("CompanyCarCheckerAvailDates")
  Exit Sub
CompanyCarCheckerAvailDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ICompanyCarCheckerAvailDates", "Company Car Checker Registration Dates", "Error in CompanyCarCheckerAvailDates.")
  Resume CompanyCarCheckerAvailDates_END
  Resume
  
End Sub

Private Function AnalyseCompanyCarsForCheckerAvailDates(cars() As COMPANY_CAR_CHECK, lNoCars As Long, rsRecordChanges As Recordset, rsdata As Recordset) As Boolean
  Dim i As Long
  Dim s As String
  Dim retValue As Boolean
  On Error GoTo AnalyseCompanyCarsForCheckerAvailDates_ERR
  
  Call xSet("AnalyseCompanyCarsForCheckerAvailDates")
   
  retValue = False
  
  For i = 1 To lNoCars
    s = ""
    If cars(i).OldAvailableTo_db <= cars(i).AvailableFrom_db Then
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      If cars(i).OldAvailableTo_db < cars(i).AvailableFrom_db Then
        s = "To date is before car is first made available."
      Else
        s = "To date equals car is first made available."
      End If
      
      cars(i).AvailableToAmended = True
      cars(i).Amended = True
      retValue = True
    End If
    cars(i).Comments_db = s
  
 Next
 
 AnalyseCompanyCarsForCheckerAvailDates = retValue

AnalyseCompanyCarsForCheckerAvailDates_END:
  Call xReturn("AnalyseCompanyCarsForCheckerAvailDates")
  Exit Function
AnalyseCompanyCarsForCheckerAvailDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseCompanyCarsForCheckerAvailDates", "Analyse Company Cars For Registration Dates", "Error analysing company cars for company car checker")
  Resume AnalyseCompanyCarsForCheckerAvailDates_END

End Function

Private Sub WriteCompanyCarsForCheckerAvailDates(cars() As COMPANY_CAR_CHECK, ByVal lNoCars As Long, rsRecordChanges As Recordset, rsdata As Recordset)
  Dim i As Long
  
  On Error GoTo WriteCompanyCarsForCheckerAvailDates_ERR
   
  For i = 1 To lNoCars
    If cars(i).Amended Then
      rsRecordChanges.AddNew
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      rsRecordChanges.Fields("Employee") = cars(i).Employee_db
      rsRecordChanges.Fields("PNum") = cars(i).PersonnelNumber_db
      rsRecordChanges!Reg = cars(i).Registration_db
      rsRecordChanges!From = cars(i).AvailableFrom_db
      rsRecordChanges!OldTo = cars(i).OldAvailableTo_db
      rsRecordChanges!Comments = cars(i).Comments_db
      rsRecordChanges.Update
    End If
  Next

WriteCompanyCarsForCheckerAvailDates_END:
  Exit Sub
WriteCompanyCarsForCheckerAvailDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteCompanyCarsForCheckerAvailDates", "Write Company Cars For Checker Registration Dates", "Error in WriteCompanyCarsForCheckerAvailDates")
  Resume WriteCompanyCarsForCheckerAvailDates_END
  Resume
End Sub

'IK 17/04/2003
'Check fuel availability dates. The following 3 functions are different from the other check's functions. IK 17/06/2003
Private Sub CompanyCarCheckerFuelAvailDates(rsRecordChanges As Recordset, rsdata As Recordset)
  Dim i As Long
  Dim cars() As COMPANY_CAR_CHECK
      
  On Error GoTo CompanyCarCheckerFuelAvailDates_ERR
  
  Call xSet("CompanyCarCheckerFuelAvailDates")
  
  Do While Not rsdata.EOF
      i = GetCompanyCarsForChecker(cars, rsdata)
      If i = 0 Then
        GoTo CompanyCarCheckerFuelAvailDates_END
      ElseIf i > 0 Then
        Call AnalyseCompanyCarsForCheckerFuelAvailDates(cars, i, rsRecordChanges, rsdata)
      End If
  Loop
  
CompanyCarCheckerFuelAvailDates_END:
  
  Call xReturn("CompanyCarCheckerFuelAvailDates")
  Exit Sub
CompanyCarCheckerFuelAvailDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ICompanyCarCheckerFuelAvailDates", "Company Car Checker Registration Dates", "Error in CompanyCarCheckerFuelAvailDates.")
  Resume CompanyCarCheckerFuelAvailDates_END
  Resume
  
End Sub

Private Function AnalyseCompanyCarsForCheckerFuelAvailDates(cars() As COMPANY_CAR_CHECK, lNoCars As Long, rsRecordChanges As Recordset, rsdata As Recordset) As Boolean
  Dim i As Long
  Dim s As String
  Dim retValue As Boolean
  On Error GoTo AnalyseCompanyCarsForCheckerFuelAvailDates_ERR
  
  Call xSet("AnalyseCompanyCarsForCheckerFuelAvailDates")
   
  retValue = False
  
  For i = 1 To lNoCars
    s = ""
    
    'check if fuel dates exist
    If (cars(i).FuelAvailableFrom_db And cars(i).FuelOldAvailableTo_db) Then
      'FuelTo < FuelFrom
      If cars(i).FuelOldAvailableTo_db <= cars(i).FuelAvailableFrom_db Then
        Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
        s = "'Fuel Available To' date is before fuel was first made available"
        cars(i).Comments_db = s
'MP DB (not used)          cars(i).FuelAvailableToAmended = True
        cars(i).Amended = True
        retValue = True
        'add entry for this error
        Call WriteCompanyCarsForCheckerFuelAvailDates(cars, lNoCars, rsRecordChanges, rsdata, i)
      End If
      
      'FuelTo > CarTo
      If cars(i).FuelOldAvailableTo_db > cars(i).OldAvailableTo_db Then
        Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
        s = "Fuel was being provided after the car was made unavailable"
        cars(i).Comments_db = s
'MP DB (not used)          cars(i).FuelAvailableToAmended = True
        cars(i).Amended = True
        retValue = True
        'add entry for this error
        Call WriteCompanyCarsForCheckerFuelAvailDates(cars, lNoCars, rsRecordChanges, rsdata, i)
      End If
      
      'FuelFrom < CarFrom
      If cars(i).FuelAvailableFrom_db < cars(i).AvailableFrom_db Then
        Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
        s = "Fuel is being provided before the car is first made available"
        cars(i).Comments_db = s
'MP DB (not used)          cars(i).FuelAvailableToAmended = True
        cars(i).Amended = True
        retValue = True
        'add entry for this error
        Call WriteCompanyCarsForCheckerFuelAvailDates(cars, lNoCars, rsRecordChanges, rsdata, i)
      End If
      
      If cars(i).Amended Then
        rsdata.MoveNext
      End If
      
    End If
    
    cars(i).Comments_db = s
  
 Next
 
 AnalyseCompanyCarsForCheckerFuelAvailDates = retValue

AnalyseCompanyCarsForCheckerFuelAvailDates_END:
  Call xReturn("AnalyseCompanyCarsForCheckerFuelAvailDates")
  Exit Function
AnalyseCompanyCarsForCheckerFuelAvailDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseCompanyCarsForCheckerFuelAvailDates", "Analyse Company Cars For Registration Dates", "Error analysing company cars for company car checker")
  Resume AnalyseCompanyCarsForCheckerFuelAvailDates_END

End Function

Private Sub WriteCompanyCarsForCheckerFuelAvailDates(cars() As COMPANY_CAR_CHECK, ByVal lNoCars As Long, rsRecordChanges As Recordset, rsdata As Recordset, i As Long)
  
  On Error GoTo WriteCompanyCarsForCheckerFuelAvailDates_ERR

    If cars(i).Amended Then
'MP DB ToDo - keep loop below? Not doing anything?
      rsdata.MoveFirst
      Do While rsdata.Fields("personnel number") <> cars(i).PersonnelNumber_db Or rsdata.Fields("registration") <> cars(i).Registration_db
        rsdata.MoveNext
      Loop
      rsRecordChanges.AddNew
      Call AddCompanyCArCheckRecord(rsRecordChanges, cars(i).Employee_db, cars(i).PersonnelNumber_db)
      rsRecordChanges.Fields("Employee") = cars(i).Employee_db
      rsRecordChanges.Fields("PNum") = cars(i).PersonnelNumber_db
      rsRecordChanges!Reg = cars(i).Registration_db
      rsRecordChanges!From = cars(i).AvailableFrom_db
      rsRecordChanges!OldTo = cars(i).OldAvailableTo_db
      rsRecordChanges!FuelFrom = cars(i).FuelAvailableFrom_db
      rsRecordChanges!OldFuelTo = cars(i).FuelOldAvailableTo_db
      rsRecordChanges!Comments = cars(i).Comments_db
      
'      fixing the values
'      If cars(i).NewAvailableTo <> UNDATED Then
'        rsRecordChanges!NewTo = cars(i).NewAvailableTo
'      End If
'      If cars(i).FuelAvailableToAmended Then
'        If optOptions(CD_CHANGE).value = True Then
'          rsdata.Edit
'          rsdata!FuelTo = cars(i).FuelNewAvailableTo
'          rsdata.Update
'        End If
'      End If

      rsRecordChanges.Update
    End If

WriteCompanyCarsForCheckerFuelAvailDates_END:
  Exit Sub
WriteCompanyCarsForCheckerFuelAvailDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "WriteCompanyCarsForCheckerFuelAvailDates", "Write Company Cars For Checker Registration Dates", "Error in WriteCompanyCarsForCheckerFuelAvailDates")
  Resume WriteCompanyCarsForCheckerFuelAvailDates_END
  Resume
End Sub



'   NOTES on company car checker - IK 17/06/2003
'
'   Sql statements
'           Need two for each check. One for rsdata and another for rsRecordChanges
'           Both the statements are stored in SQLQUERIES CLASS
'
'   rsdata
'           contains all the necessary data
'
'   rsRecordChanges
'           contains information of cars which do not pass the check
'           if a car does not pass the test then add an entry and write to this recordset
'
'   Every check uses three functions as follows
'
'   1. CompanyCarCheckerCHECKNAME
'         standard function and calls fucntion 2 (see below)
'
'   2. AnalyseCompanyCarsForCheckerCHECKNAME
'         Analyses every car using IF statements. if an inconsistency is
'         detected then it calls the write function (see next function)
'
'   3. WriteCompanyCarsForCheckerFuelAvailDates
'         Add entries to rsrecordChanges in this function



