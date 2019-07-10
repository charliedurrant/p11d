VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form F_SelectEmployeesByReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Employees"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4185
      TabIndex        =   3
      Top             =   5940
      Width           =   1095
   End
   Begin MSComctlLib.TreeView tvwReports 
      Height          =   3435
      Left            =   90
      TabIndex        =   2
      Top             =   2385
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   6059
      _Version        =   393217
      Indentation     =   441
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2970
      TabIndex        =   1
      Top             =   5940
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "To select employees select a report from the list below. If an employee appears in the report he/she will be selected. "
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5145
   End
End
Attribute VB_Name = "F_SelectEmployeesByReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBaseNotify
Private m_NodeLastSelected As Node
Private m_hashEmployees As ObjectHash
Private m_NumberOfEmployeesSelected As Long
Private Sub cmdCancel_Click(Index As Integer)
  Me.Hide
  Unload Me
End Sub

Private Sub SettingsToScreen()
  On Error GoTo err_Err
  
  lblInfo.Caption = "To select employees select a report from the list below and click select. " & _
                    "If an employee appears in the report he/she will be selected." & vbCrLf & vbCrLf & _
                    "If you wish to have a different selection criteria then simply design a new report " & _
                    "in the 'Report Wizard' on the priint dialog, save it and return to this screen." & vbCrLf & vbCrLf & _
                    "Please note any report must include the 'Personnel number' and also " & _
                    "any report will only analyse those employees avaialble acording to the print range " & _
                    "on the print dialog."
  
  Call ReportsImageListSet(tvwReports)
  Call ManagementReportsToTree(tvwReports)
  Call ReportsToTreeEndSelectDialog
  
err_End:
  Exit Sub
err_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "SettingsToScreen"), "Failed to init dialog")
  Resume
End Sub
Private Sub ReportsToTreeEndSelectDialog()
  If ReportsToTreeEnd(tvwReports, m_NodeLastSelected) Then
    Call tvwReports_NodeClick(tvwReports.SelectedItem)
  End If
End Sub
  
Private Sub cmdOK_Click(Index As Integer)
  Dim rt As RPT_TYPE
  Dim sPathAndFile As String, sAdditional As String
  Dim ibnLast As IBaseNotify
  Dim udm As ITCSUDM
  Dim ees As ObjectList
  Dim benEE As IBenefitClass
  Dim bClearSelection As Boolean
  Dim ibf As IBenefitForm2
  Dim LVI As ListItem
  Dim i As Long
  
  Dim es As EMPLOYEE_SELECTION
On Error GoTo err_Err
  
  es = p11d32.ReportPrint.EmployeeSelection
  p11d32.ReportPrint.EmployeeSelection = ES_ALL
  Set m_hashEmployees = New ObjectHash
  rt = p11d32.ReportPrint.ReportPrintPrapare(p11d32.ReportPrint.DefaultSelectEmployeeReportIndex)
  If (rt <> RPTT_USER) And (rt <> RPTT_MANAGEMENT) Then
    Call Err.Raise(ERR_REPORT_INVALID, "Select", "Invalid report selected, select another report")
  End If
  Set udm = p11d32.udm
  
  Set ibnLast = udm.Notify
  Set udm.Notify = Me
  m_NumberOfEmployeesSelected = 0
  sPathAndFile = p11d32.ReportPrint.ManagementReportPathAndFile(p11d32.ReportPrint.DefaultSelectEmployeeReportIndex, p11d32.ReportPrint.UserReportSelectEmployeeFileLessExtension)
  
  bClearSelection = MsgBox("Do you wish to clear the currently selected employees?", vbQuestion Or vbYesNo, "Clear selected") = vbYes
  
  Set ees = p11d32.CurrentEmployer.employees
  For i = 1 To ees.Count
    Set benEE = ees.Item(i)
    Call m_hashEmployees.Add(benEE, benEE.value(ee_PersonnelNumber_db))
  Next
  Set ibf = F_Employees
  If (bClearSelection) Then
    For Each LVI In ibf.lv.listitems
      If LVI.Checked Then
        LVI.Checked = False
        Call F_Employees.LB_ItemCheck(LVI)
      End If
    Next
  End If
  
  Call p11d32.ReportPrint.StartReportWizard(sPathAndFile, REPORTW_PREPARE_EXPORT, p11d32.TempPath & "P11D-Selection_report.csv")
  If (Not bClearSelection) Then sAdditional = " additional"
  Call MsgBox(m_NumberOfEmployeesSelected & sAdditional & " employee(s) selected", vbInformation, "Employees Selected")
  
  For Each LVI In ibf.lv.listitems
    Set benEE = ees.Item(LVI.Tag)
    If (benEE.value(ee_Selected)) And Not LVI.Checked Then
      LVI.Checked = benEE.value(ee_Selected)
      Call F_Employees.LB_ItemCheck(LVI)
    End If
    'bug in comctl, setting the property does not cause the event
  Next
  
  Set m_hashEmployees = Nothing
err_End:
  p11d32.ReportPrint.EmployeeSelection = es
  If (Not udm Is Nothing) Then
    If Not ibnLast Is Nothing Then
      Set udm.Notify = ibnLast
    End If
  End If
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Select", "Select", "Failed to select employees")
  Resume err_End
  Resume
End Sub

Private Sub Form_Load()
  Call SettingsToScreen
End Sub


Private Sub IBaseNotify_Notify(ByVal Current As Long, ByVal Max As Long, ByVal Message As String)
  Dim p0 As Long
  Dim p1 As Long
  Dim sPNum As String
  Dim benEE As IBenefitClass
  If (Current = -1 And Max = -1) Then
    If InStr(1, Message, "OUTPUT_LINE:") > 0 Then
      p0 = InStr(1, Message, S_PNUM)
      If (p0 > 0) Then
        p0 = InStr(p0, Message, ";")
        If (p0 > 0) Then
          p0 = p0 + 1
          p1 = InStr(p0, Message, ",")
          sPNum = Mid$(Message, p0, (p1 - p0))
          sPNum = Trim$(Replace$(sPNum, """", ""))
          If (Len(sPNum)) Then
            Set benEE = m_hashEmployees.Item(sPNum, False)
            If Not benEE Is Nothing Then
              m_NumberOfEmployeesSelected = m_NumberOfEmployeesSelected + 1
              benEE.value(ee_Selected) = True
            End If
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub tvwReports_Collapse(ByVal Node As MSComctlLib.Node)
  Node.Image = IMG_FOLDER_CLOSED
End Sub
Private Sub tvwReports_Expand(ByVal Node As MSComctlLib.Node)
  Node.Image = IMG_FOLDER_OPEN
End Sub
Public Sub NodeClick(ByVal Node As Node)
On Error GoTo NodeClick_ERR
  
  
    If Len(Node.Tag) > 0 Then
    'valid report not header
      Call ReportsSelectNodeImage(m_NodeLastSelected, Node)
      Set tvwReports.SelectedItem = Node
      p11d32.ReportPrint.DefaultSelectEmployeeReportIndex = Node.Tag
      If p11d32.ReportPrint.ReportType(Node.Tag, True) = RPTT_USER Then
        p11d32.ReportPrint.UserReportSelectEmployeeFileLessExtension = Node.Text
      End If
    End If


NodeClick_END:
  Exit Sub
NodeClick_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NodeClick", "NodeClick", "Error clicking a node on the reports list.")
  Resume NodeClick_END
  Resume

End Sub

Private Sub tvwReports_NodeClick(ByVal Node As MSComctlLib.Node)
  Call NodeClick(Node)
End Sub
