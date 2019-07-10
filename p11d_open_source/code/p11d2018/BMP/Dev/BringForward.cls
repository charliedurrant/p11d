VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BringForward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements IErrorFilter

Public Enum YEAR_OF_EMPLOYER
  YOE_CURRENT = 1
  YOE_PRIOR
End Enum

'read in from ini file
Public BringForwardType As Long
Public HMITSChosen As Long
Public CDCs As Boolean
Public CDBs As Boolean
Public Comments As Boolean
Public SharedVans As Boolean
Public NewFilesForAll As Boolean
'end read in from ini file
Public Yes As Boolean

Public OLCurrentEmployers As ObjectList
Public OLPriorEmployers As ObjectList

Public YOE As Long

Private Function IErrorFilter_FilterErrorMessage(ByVal UserName As String, ByVal DateTime As Date, ByVal ErrorNumber As Long, ByVal ErrorName As String, ByVal ErrorText As String, ByVal SourceFunction As String) As Boolean

End Function

Private Function IErrorFilter_PrintAvailable() As Boolean

End Function

Private Function IErrorFilter_PrintErrors() As Boolean

End Function

Private Function IErrorFilter_ViewAvailable() As Boolean

End Function

Private Function IErrorFilter_ViewErrors() As Boolean

End Function

Private Function GetDestination(SourceEmployer As Employer, DestinationEmployer As Employer) As Boolean
  Dim i As Long
  Dim li As ListItem
  Dim eySrc As IBenefitClass, eyDst As IBenefitClass
  Dim tempEy As IBenefitClass
  Dim sNewFileName As String
  
  On Error GoTo GetDestination_ERR
  
  Call xSet("GetDestination")
    
  If SourceEmployer Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "GetDestination", "The source employer is nothing.")
  
  Set p11d32.CurrentEmployer = Nothing
  
  For i = 1 To F_BringForward.lvCurrentYearFiles.listitems.Count
    Set li = F_BringForward.lvCurrentYearFiles.listitems(i)
    If li.SmallIcon = IMG_SELECTED Then
      Set DestinationEmployer = OLCurrentEmployers(li.Tag)
      Exit For
    End If
  Next
      
  If DestinationEmployer Is Nothing Then
    Call ClearCursor
    'create a new file
    Set eySrc = SourceEmployer
TRY_AGAIN:
    F_Input.ValText.MaxLength = 6
    F_Input.ValText.AllowEmpty = False
    F_Input.ValText.TypeOfData = VT_STRING
    
    
    sNewFileName = Left$(eySrc.value(employer_FileName), Len(eySrc.value(employer_FileName)) - Len(p11d32.Rates.FileExtensionPrior))
    
    If Not F_Input.Start("New file", "Enter a new file name", sNewFileName) Then GoTo GetDestination_END
    
    sNewFileName = F_Input.ValText.Text & p11d32.Rates.FileExtensionCurrent
    
    'check that it does not already exist
    If FileExists(FullPath(p11d32.WorkingDirectory) & sNewFileName) Then Call Err.Raise(ERR_FILE_EXISTS, "GetFiles", "The file " & sNewFileName & " already exists, please choose another name or, choose the file from the current year files.")
    
    Set DestinationEmployer = New Employer
    'create a new file
    F_BringForward.P.Caption = "Creating file " & sNewFileName
    Call SetCursor
    If Not DestinationEmployer.CreateEmployer(p11d32.WorkingDirectory & sNewFileName) Then Call Err.Raise(ERR_NO_EMPLOYER, "GetFiles", "Failed to create employer.")
    
    Call p11d32.LoadEmployer(DestinationEmployer, False, True, True)
    Set eyDst = DestinationEmployer
    Call CopyBenData(DestinationEmployer, eySrc)
    eyDst.value(employer_PathAndFile) = p11d32.WorkingDirectory & sNewFileName
    'if change here change below
    eyDst.value(employer_Year) = p11d32.AppYearShort
    Call eyDst.WriteDB
    'add to the list
    
    
    Call EmployerToListView(OLCurrentEmployers.Add(DestinationEmployer), DestinationEmployer, YOE_CURRENT)
  Else
    'we have selected a current file?
    Call p11d32.LoadEmployer(DestinationEmployer, False, False, True)
    
    If BringForwardType = BFT_OVERWRITE Then
      'we are simply deleting the employer and recreating it with the old employer data, saves queries to delete data from tables
      Set eyDst = DestinationEmployer
      F_BringForward.P.Caption = "Recreating file " & eyDst.value(employer_FileName)
      Set tempEy = New Employer
      Call CopyBenData(tempEy, DestinationEmployer)
      
      Call eyDst.Kill
      Call BackupEmployer(DestinationEmployer)
      Call DestinationEmployer.CreateEmployer(tempEy.value(employer_PathAndFile))
      Call p11d32.LoadEmployer(DestinationEmployer, False, True, True)
      Call CopyBenData(DestinationEmployer, tempEy)
      eyDst.value(employer_Year) = p11d32.AppYearShort
      Call eyDst.WriteDB
    End If
  End If
  
  
  
  GetDestination = True
  
GetDestination_END:
  Set p11d32.CurrentEmployer = SourceEmployer
  Call p11d32.LoadEmployer(SourceEmployer, False, False, True)
  Call ClearCursor
  Set F_Input = Nothing
  F_BringForward.P.Caption = ""
  Call xReturn("GetDestination")
  Exit Function
GetDestination_ERR:
  If Err.Number = ERR_FILE_EXISTS Then
    Call ErrorMessage(ERR_ERROR, Err, "GetDestination", "Get Destination", "Error getting the destination employer.")
    Resume TRY_AGAIN
  Else
    Call ErrorMessage(ERR_ERROR, Err, "GetDestination", "Get Destination", "Error getting the destination employer.")
    Resume GetDestination_END
  End If
  Resume
End Function
Private Sub AllApartFromEmployees(DestinationEmployer As Employer, SourceEmployer As Employer)
  Dim dbTemp As Database
  
  Dim ben As IBenefitClass
  
  On Error GoTo AllApartFromEmployees_ERR
  
  Call xSet("AllApartFromEmployees")
  
  Set dbTemp = SourceEmployer.db
  Set SourceEmployer.db = DestinationEmployer.db
  Call SourceEmployer.WriteFPCS
  
  ' we will have problems with keys if combining two files
  If Me.CDCs Then Call SourceEmployer.WriteCDCs
  
  If Me.SharedVans Then
    'apf Call DestinationEmployer.db.Execute(sql.Queries(DELETE_CDB_EMPLOYEE_BENEFITS))
    Set ben = SourceEmployer.SharedVans
    Set ben.Parent = DestinationEmployer
    Call ben.WriteDB
  End If
  
  Set SourceEmployer.db = dbTemp
    
  If Me.CDBs Then
    Call DestinationEmployer.db.Execute(sql.Queries(DELETE_CDB_EMPLOYEE_BENEFITS))
    Call WriteEmployee(DestinationEmployer, SourceEmployer.CDBEmployee)
    Set ben = DestinationEmployer
    Call SourceEmployer.db.Execute(sql.Queries(COPY_CDB_LINKS, ben.value(employer_PathAndFile)))
  End If
  
AllApartFromEmployees_END:
  Call xReturn("AllApartFromEmployees")
  Exit Sub
AllApartFromEmployees_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AllApartFromEmployees", "All Apart From Empoyees", "Error bringing forward data that is not employees.")
  Resume AllApartFromEmployees_END
  Resume
End Sub
Public Function GetSourceFiles(SourceEmployers As ObjectList) As Boolean
  Dim i As Long
  Dim li As ListItem
  Dim benEY As IBenefitClass
  Dim ey As Employer
  
  On Error GoTo GetSourceFiles_ERR
  
  Call xSet("GetSourceFiles")
  
  Set SourceEmployers = New ObjectList
  
  For i = 1 To F_BringForward.lvPriorYearFiles.listitems.Count
    Set li = F_BringForward.lvPriorYearFiles.listitems(i)
    If li.SmallIcon = IMG_SELECTED Then
      Set ey = OLPriorEmployers(li.Tag)
      Set benEY = ey
      If Not FixLevelCheck(benEY) Then
        If (p11d32.AppYear > 1999) Then
          Call Err.Raise(ERR_NOT_UPTODATE, "GetSourceFiles", "Before you bring forward the employer, " & benEY.Name & ", please open the required employer file in last years P11D software, close the employer and then return to this screen.")
        End If
      End If
      Call SourceEmployers.Add(OLPriorEmployers(li.Tag))
    End If
  Next
  
  If SourceEmployers.Count = 0 Then Call Err.Raise(ERR_NO_EMPLOYER, "GetFiles", "No source employers selected.")
  
  GetSourceFiles = True

GetSourceFiles_END:
  Call xReturn("GetSourceFiles")
  Exit Function
GetSourceFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetSourceFiles", "Get Source Files", "Error getting the source files for roll forward.")
  Resume GetSourceFiles_END
  Resume
End Function
Private Function FixLevelCheck(benEY As IBenefitClass) As Boolean
  On Error GoTo FixLevelCheck_Err
  
  If benEY Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "FixLevelCheck", "The employer is nothing.")
  FixLevelCheck = benEY.value(employer_FixLevel_db) = LastFixLevel(benEY.value(employer_Year))
  
FixLevelCheck_End:
  Exit Function
FixLevelCheck_Err:
  Call ErrorMessage(ERR_ERROR, Err, "FixLevelCheck", "Fix Level Check", "Error checking the fix level of the file to bring forward.")
  Resume FixLevelCheck_End
End Function

Public Sub ProcessFiles()
  Dim SourceEmployers As ObjectList
  Dim DestinationEmployer As Employer
  Dim SourceEmployer As Employer
  Dim ey As IBenefitClass
  Dim i As Long, j As Long
  Dim Employee As Employee
  Dim ben As IBenefitClass
  Dim prg As Object
  Dim bGotDestination As Boolean, bIsCDBEmployee As Boolean
  
  On Error GoTo ProcessFiles_ERR

  Call xSet("ProcessFiles")
  'get source files
  
  With F_BringForward
    If .optType(1).value = True And p11d32.BringForward.Yes Then
      With F_BringForwardOptions
        If .chkCDBs Or .chkCDCs Or .chkSharedVans Then
         If MsgBox("By running an update, you may get duplicate company defined benefits, company defined categories and shared vans." & vbCrLf & "Press cancel to go back and untick these sections, or OK to continue.", vbOKCancel + vbExclamation + vbDefaultButton2, "Update") = vbCancel Then GoTo ProcessFiles_END
        End If
      End With
    End If
  End With
  
  F_BringForward.sts.StopPrg
  
  If Not GetSourceFiles(SourceEmployers) Then GoTo ProcessFiles_END
  
  'we now have a destination employer
  
  Set prg = F_BringForward.sts.prg
  prg.TextAlignment = TCSPROG.ALIGN_RIGHT
  
  Me.Yes = True 'controls writeDb see write db of employee etc
  
  For i = 1 To SourceEmployers.Count
    Call SetCursor
    Set SourceEmployer = SourceEmployers(i)
    If Not p11d32.LoadEmployer(SourceEmployer, False, False, True) Then GoTo ProcessFiles_END
    Set ey = SourceEmployer
    Call ClearCursor
    'has it been rolled forward before
    If ey.value(employer_BroughtForward_db) Then
      Select Case MultiDialog("Warning", "The employer " & ey.Name & " has already been brought forward, do you wish to continue.", "Continue", "Ignore this file", "Abort")
        Case 2
          Call CleanUpEmployer(SourceEmployer)
          GoTo NEXT_EMPLOYER
        Case 3
          GoTo ProcessFiles_END
      End Select
    End If
    
    If (Not bGotDestination) Or NewFilesForAll Then
      If Not GetDestination(SourceEmployers(i), DestinationEmployer) Then GoTo ProcessFiles_END
      bGotDestination = True
    End If
    
    Call SetCursor
    Call AllApartFromEmployees(DestinationEmployer, SourceEmployer)
    
    F_BringForward.P.Caption = "Analysing file " & ey.value(employer_FileName)
    Call F_BringForward.sts.StartPrg(SourceEmployer.employees.Count, "Employee", ValueOfMax)
    
    For j = 1 To SourceEmployer.employees.Count
      Set Employee = SourceEmployer.employees(j) 'will always be something as just loaded
      Call F_BringForward.sts.StepCaption("Employee no " & Employee.PersonnelNumber)
      
      Call WriteEmployee(DestinationEmployer, Employee)
    Next
    
    
    Set ben = SourceEmployer
    ben.value(employer_BroughtForward_db) = True
    Call ben.WriteDB
    
    Call CleanUpEmployer(SourceEmployer)
    Call UpdateEmployerListViewItem(ben, YOE_PRIOR)
     
    Set ben = DestinationEmployer
    
    'this was commeneted out, but seems OK CAD fixes ttp 241
    ben.value(employer_FixLevel_db) = p11d32.TemplateFixlevel
    Call UpdateEmployerListViewItem(ben, YOE_CURRENT)
    
    
    If i = SourceEmployers.Count Then
      Call F_BringForward.sts.StopPrg
    Else
      If NewFilesForAll Then Call CleanUpEmployer(DestinationEmployer)
      F_BringForward.P.Caption = ""
      prg.Indicator = TCSPROG.None
      prg.TextAlignment = Align_Centre
      prg.Caption = "Finished bring forward"
    End If
    
    
    Call ClearCursor
    
NEXT_EMPLOYER:
  Next
        
  
          
  Call p11d32.LoadEmployer(Nothing, False)
  Call ToolBarButton(TBR_REFRESH_EMPLOYERS, 0)
    
ProcessFiles_END:
  Call CleanUpEmployer(DestinationEmployer)
  Me.Yes = False 'controls writeDb see write db of employee etc
  Call ClearCursor
  Call xReturn("ProcessFiles")
  Exit Sub
ProcessFiles_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ProcessFiles", "Process Files", "Error processing files to bring forward.")
  Resume ProcessFiles_END
  Resume
End Sub
Private Sub WriteEmployee(DestinationEmployer As Employer, ee As Employee)
  Dim eeBen As IBenefitClass
  
  On Error GoTo WriteEmployee_ERR
  
  Call xSet("WriteEmployee")
  
  If ee Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "WriteEmployee", "The employee is nothing.")
  Set eeBen = ee
  If eeBen.CanBringForward Then
    Call ee.LoadBenefits(TBL_ALLBENEFITS, False)
    Set eeBen.Parent = DestinationEmployer
    'prepare my destination employee if update etc...
    Call ee.BringForwardPrepare(p11d32.BringForward.HMITSChosen, BringForwardType)
    eeBen.ReadFromDB = False
    eeBen.Dirty = True
    Call eeBen.WriteDB
    Call ee.KillBenefits
  End If
  
WriteEmployee_END:
  Call xReturn("WriteEmployee")
  Exit Sub
WriteEmployee_ERR:
  If ee Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "WriteEmployee", "Write Employee", "Error writing an employee during bring forward.")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "WriteEmployee", "Write Employee", "Error writing an employee during bring forward, employee personnel number = " & ee.PersonnelNumber)
  End If
  'Resume
End Sub
Private Sub UpdateEmployerListViewItem(benEY As IBenefitClass, YOE As YEAR_OF_EMPLOYER, Optional li As ListItem)
  Dim OL As ObjectList
  Dim lv As ListView
  Dim benTemp As IBenefitClass
On Error GoTo UpdateEmployerListViewItem_ERR

  Call xSet("UpdateEmployerListViewItem")
  
  If benEY Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "UpdateEmployerListViewItem", "The employer is nothing")
  
  If li Is Nothing Then
    Select Case YOE
      Case YEAR_OF_EMPLOYER.YOE_CURRENT
        Set OL = Me.OLCurrentEmployers
        Set lv = F_BringForward.lvCurrentYearFiles
      Case YEAR_OF_EMPLOYER.YOE_PRIOR
        Set OL = Me.OLPriorEmployers
        Set lv = F_BringForward.lvPriorYearFiles
    End Select
    
    Set li = lv.listitems(OL.ItemIndex(benEY))
    If li Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "UpdateEmployerListViewItem", "The list item is nothing.")
    GoTo UPDATE_LI
  Else
UPDATE_LI:
    li.SmallIcon = IMG_UNSELECTED
    
    li.Text = benEY.Name
    li.SubItems(1) = benEY.value(employer_FileName)
    If YOE = YOE_PRIOR Then
      li.SubItems(2) = p11d32.MagneticMedia.MMBool(benEY.value(employer_BroughtForward_db))
      li.SubItems(3) = benEY.value(employer_FixLevel_db)
    Else
      li.SubItems(2) = benEY.value(employer_FixLevel_db)
    End If
  End If
  
    
UpdateEmployerListViewItem_END:
  Call xReturn("UpdateEmployerListViewItem")
  Exit Sub
UpdateEmployerListViewItem_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateEmployerListViewItem", "Update Employer List View Item", "Error updating an employer list item for bring forward.")
  Resume UpdateEmployerListViewItem_END
  Resume
End Sub
Private Sub CleanUpEmployer(ey As Employer)
  Dim ben As IBenefitClass
  
  On Error GoTo CleanUpEmployer_ERR
  
  Call xSet("CleanUpEmployer")
  
  If Not ey Is Nothing Then
    Set ben = ey
    Call ben.Kill
    Set ey = Nothing
  End If
  
  
CleanUpEmployer_END:
  Call xReturn("CleanUpEmployer")
  Exit Sub
CleanUpEmployer_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CleanUpEmployer", "Clean Up Employer", "Error terminating an employer in bring forward.")
  Resume CleanUpEmployer_END
End Sub
Public Sub Initialise()
  On Error GoTo Initialise_ERR
  
  Call xSet("Initialise")
  
' F_BringForward.Show vbModal
  Call p11d32.Help.ShowForm(F_BringForward, vbModal)
  Set F_BringForward = Nothing
  
Initialise_END:
  Call xReturn("Initialise")
  Exit Sub
Initialise_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Initialise", "Initialise", "Error initialising the BringForward.")
  Resume Initialise_END
End Sub

Public Sub EmployerToListView(ByVal lBenefitIndex As Long, benEY As Object, YOEmployer As YEAR_OF_EMPLOYER)
  Dim li As ListItem
  Dim lv As ListView
  
  On Error GoTo EmplyoyerToListView_ERR
  
  Call xSet("EmplyoyerToListView")
  
  If Not IsEmployer(benEY) Then Call Err.Raise(ERR_NO_EMPLOYER, "EmployerToListView", "The object sent is not an emploter.")
    
  Select Case YOEmployer
    Case YEAR_OF_EMPLOYER.YOE_CURRENT
      Set lv = F_BringForward.lvCurrentYearFiles
    Case YEAR_OF_EMPLOYER.YOE_PRIOR
      Set lv = F_BringForward.lvPriorYearFiles
    Case Else
      ECASE ("Invalid YEAR_OF_EMPLOYER in EmployerToListView")
  End Select
    
  Set li = lv.listitems.Add()
  li.Tag = lBenefitIndex
  Call UpdateEmployerListViewItem(benEY, YOEmployer, li)
  
  
EmplyoyerToListView_END:
  Call xReturn("EmplyoyerToListView")
  Exit Sub
EmplyoyerToListView_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployersToListView", "Employers To List View", "Error placing the employer to the list view in bring forward.")
  
  Resume EmplyoyerToListView_END
  Resume
End Sub

Public Sub EmployersToListView()
  Dim i As Long
  Dim ben As IBenefitClass
  
On Error GoTo EmployersToListView_ERR

Call xSet("EmployersToListView")
  
  Call F_BringForward.lvCurrentYearFiles.listitems.Clear
  Call F_BringForward.lvPriorYearFiles.listitems.Clear
  
  For i = 1 To p11d32.BringForward.OLCurrentEmployers.Count
    Set ben = p11d32.BringForward.OLCurrentEmployers(i)
    If Not ben Is Nothing Then Call p11d32.BringForward.EmployerToListView(i, ben, YOE_CURRENT)
  Next
  
  For i = 1 To p11d32.BringForward.OLPriorEmployers.Count
    Set ben = p11d32.BringForward.OLPriorEmployers(i)
    If Not ben Is Nothing Then Call EmployerToListView(i, ben, YOE_PRIOR)
  Next
  
EmployersToListView_END:
  Call xReturn("EmployersToListView")
  Exit Sub
EmployersToListView_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "EmployersToListView", "Employers To List View", "Error placing the employers to the list view in bring forward.")
  Resume EmployersToListView_END
End Sub

Public Sub Options()
  On Error GoTo Options_ERR
  
  Call xSet("Options")
  
'  F_BringForwardOptions.Show vbModal
  Call p11d32.Help.ShowForm(F_BringForwardOptions, vbModal)
  Set F_BringForwardOptions = Nothing
  
Options_END:
  Call xReturn("Options")
  Exit Sub
Options_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Options", "Options", "Error showing the Bring Forward Options.")
  Resume Options_END
End Sub
