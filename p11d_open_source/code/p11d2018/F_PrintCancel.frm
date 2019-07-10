VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Begin VB.Form F_PrintCancel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printing"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TCSPROG.TCSProgressBar prg 
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   675
      Width           =   4470
      _cx             =   7885
      _cy             =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   100
      Value           =   50
      BarBackColor    =   -2147483633
      BarForeColor    =   12937777
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   1
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   2
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   3285
      TabIndex        =   0
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label lbl 
      Caption         =   "lbl"
      Height          =   465
      Left            =   45
      TabIndex        =   1
      Top             =   135
      Width           =   4470
   End
End
Attribute VB_Name = "F_PrintCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Cancel As Boolean
Private m_rep As Reporter
Private m_pr As P11D_REPORTS
Private m_ReportOrient As REPORT_ORIENTATION
Private m_sReportName As String
Private m_ReportDest As REPORT_TARGET
Public Function DoReportStandardSub(ByVal rep As Reporter, ByVal pr As P11D_REPORTS, ByVal ReportDest As REPORT_TARGET, ByVal ReportOrient As REPORT_ORIENTATION, ByVal sReportName As String) As Boolean
  On Error GoTo err_Err
  Set m_rep = rep
  m_pr = pr
  m_ReportOrient = ReportOrient
  m_sReportName = sReportName
  m_ReportDest = ReportDest
  Me.Show 1
  DoReportStandardSub = m_Cancel <> True
  

err_End:
  m_Cancel = False
  Exit Function
err_Err:
  Call Err.Raise(Err.Number, "F_PrintCancel.DoReportStandardSub", Err.Description)
End Function
Private Sub DoStandardReportSubEx(ByVal rep As Reporter, ByVal pr As P11D_REPORTS, ByVal ReportDest As REPORT_TARGET, ByVal ReportOrient As REPORT_ORIENTATION, ByVal sReportName As String)
  Dim ee As Employee, InError As Boolean
  Dim j As Long, i As Long
  Dim eName As String
  Dim bSeparatePrintJobs As Boolean
  Dim sEmployeeLetterText As String
  Dim CurrentEmployee As Employee
  Dim PrintedEmployees As ObjectList
  Dim bIsEmail As Boolean
  Dim bPastCancel As Boolean
  Dim ben As IBenefitClass
  
  Dim RO As REPORT_ORIENTATION
  Dim rt As REPORT_TARGET
  
  On Error GoTo DoReportStandardSub_Err
  
  lbl.Caption = "Preparing report " & sReportName
  'cad who wrote this comment
  'nned to do as use global P11d32.ReportPrint.destination in sub functions
  bSeparatePrintJobs = p11d32.ReportPrint.SeparatePrintJobs
  If p11d32.ReportPrint.Destination = REPD_FILE_PRN Then p11d32.ReportPrint.SeparatePrintJobs = True
  
  bIsEmail = p11d32.ReportPrint.IsEmail(pr)
  
  rt = ReportDest
  If (p11d32.CurrentEmployer Is Nothing) Then Call Err.Raise(ERR_IS_NOTHING, "DoReportStandardSub", "The current employer is nothing.")
  
  Set CurrentEmployee = p11d32.CurrentEmployer.CurrentEmployee
  
  Call PrgAlignment(ALIGN_RIGHT)
  prg.TextAlignment = ALIGN_RIGHT
  
  If rep Is Nothing Then Set rep = ReporterNew
  
  RO = p11d32.ReportPrint.Orientation(pr)
  
  If bIsEmail Then
    'Set atms = New MailAttachments 'RK Email 19/03/03
  ElseIf Not p11d32.ReportPrint.Destination = REPD_FILE_HTML Then
    If (Not p11d32.ReportPrint.SeparatePrintJobs) Or (ReportDest = PREPARE_REPORT) Then
      If Not rep.InitReport("P11D report:" & sReportName, ReportDest, p11d32.ReportPrint.Orientation(pr), True) Then Call Err.Raise(ERR_DOREPORT, "DoSubReport", "Unable to initialise Reporter." & vbCrLf & "Unable to initialise report engine.")
    End If
  Else
    ReportDest = PREPARE_REPORT
  End If
    
  sEmployeeLetterText = p11d32.ReportPrint.StandardReportLetterFileText(pr)

  Set PrintedEmployees = New ObjectList
  
  If p11d32.ReportPrint.SelectedEmployees.Count > 0 Then
    Call PrgStartCaption(p11d32.ReportPrint.SelectedEmployees.Count, "Printing for employees", "Analysing employee", ValueOfMax)
    prg.Min = 0
    prg.value = 0
    prg.Max = p11d32.ReportPrint.SelectedEmployees.Count
    PrintedEmployees.Increment = p11d32.ReportPrint.SelectedEmployees.Count
  End If
  
  For i = 1 To p11d32.ReportPrint.SelectedEmployees.Count
     Set ee = p11d32.ReportPrint.SelectedEmployees(i)
     If ee Is Nothing Then GoTo NEXT_EMPLOYEE
     Set p11d32.CurrentEmployer.CurrentEmployee = p11d32.ReportPrint.SelectedEmployees(i)
     
     DoEvents
     If (m_Cancel) Then Call Err.Raise(ERR_PRINT_CANCEL, "DoReportStandardSub", "Printing cancelled")
     MDIMain.sts.Step
     Call prg.StepCaption(ee.FullName)
     If p11d32.ReportPrint.NonZeroTest(ee) Or p11d32.ReportPrint.ReportSettingsIgnoreZeroOnly(pr) Then  'km
       Set ben = ee
       If bIsEmail And (Len(ben.value(ee_Email_db)) = 0) Then GoTo NEXT_EMPLOYEE
       j = j + 1
       Call PrintedEmployees.Add(ee)
     Else
       GoTo NEXT_EMPLOYEE
     End If
     
     If pr = RPT_PRINTEDEMPLOYEES Then GoTo NEXT_EMPLOYEE_PRINTED
     
     If bIsEmail Then
       If Not p11d32.ReportPrint.ATCMail Is Nothing Then
         Call p11d32.ReportPrint.ATCMail.NewMessage 'RK Email 19/03/03
       Else
         Call SetCursor(vbArrow) 'RK reset pointer in case setup required
         Call p11d32.ReportPrint.InitATCMAIL
         Call ClearCursor
       End If
       'Call atms.RemoveAll 'RK Email 19/03/03
     Else
       If Not rep.InitReport(sReportName, ReportDest, RO, True) Then Call Err.Raise(ERR_DOREPORT, "DoSubReport", "Unable to initialise Reporter." & vbCrLf & "Unable to initialise report engine.")
     End If
     Call p11d32.ReportPrint.ReportTextToRep(rep, ee, pr, sEmployeeLetterText, RO, bIsEmail, sReportName) 'RK Email 19/03/03
     'Call ReportTextToRep(rep, ee, pr, sEmployeeLetterText, RO, bIsEmail, atms, sReportName) 'RK Email 19/03/03
         
     If p11d32.ReportPrint.Destination = REPD_FILE_HTML Then
      'export the file to a directory
      Call p11d32.ReportPrint.ExportReport(rep, ee, sReportName)
     End If
NEXT_EMPLOYEE_PRINTED:
     Call SetPanel2(IIf(bIsEmail, "Emailed for", "Printed for ") & j & " employee(s)")
NEXT_EMPLOYEE:
    Call ee.KillBenefitsEx(CurrentEmployee)
  Next i
  
  bPastCancel = True
  If p11d32.ReportPrint.PrintedEmployees And (pr <> RPT_PRINTEDEMPLOYEES) Then
    If Not rep.InitReport(p11d32.ReportPrint.Name(RPT_PRINTEDEMPLOYEES), rt, RO, True) Then Call Err.Raise(ERR_DOREPORT, "DoSubReport", "Unable to initialise Reporter." & vbCrLf & "Unable to initialise report engine.")
    Call Report_PrintedEmployees(rep, PrintedEmployees)
    Call rep.EndReport
    If (rt = PREPARE_REPORT) And (bIsEmail Or (p11d32.ReportPrint.Destination = REPD_FILE_HTML)) Then
      rep.PreviewReport
    End If
  ElseIf pr = RPT_PRINTEDEMPLOYEES Then
    Call Report_PrintedEmployees(rep, PrintedEmployees)
  End If
  
DoReportStandardSub_End:
  Call PrgStopCaption
  Call PrgAlignment(0)
  If Not rep Is Nothing Then
    If InError Then
      rep.AbortReport
      If Not bPastCancel Then
        Call rep.EndReport(True)
      End If
    Else
      If pr = RPT_PRINTEDEMPLOYEES Then
        Call rep.EndReport
        If (rt = PREPARE_REPORT) Then rep.PreviewReport
        GoTo DoReportStandardSubEx_End
      End If
      If (Not bIsEmail And (p11d32.ReportPrint.Destination <> REPD_FILE_HTML)) Then
        If pr = RPT_PRINTEDEMPLOYEES Or p11d32.ReportPrint.PrintedEmployees Then
          If (Not p11d32.ReportPrint.SeparatePrintJobs) Or (ReportDest = PREPARE_REPORT) Then Call rep.EndReport(InError)
          If (rt = PREPARE_REPORT) Then rep.PreviewReport
        Else
          If Not p11d32.ReportPrint.SeparatePrintJobs Or (rt = PREPARE_REPORT) Then rep.EndReport
          If (rt = PREPARE_REPORT) Then rep.PreviewReport
        End If
      End If
    End If
  End If
  
DoReportStandardSubEx_End:
  Me.Hide
  Unload Me
  Set rep = Nothing
  Set m_rep = Nothing
  If Not (p11d32.CurrentEmployer Is Nothing) Then Set p11d32.CurrentEmployer.CurrentEmployee = CurrentEmployee
  p11d32.ReportPrint.SeparatePrintJobs = bSeparatePrintJobs
  Exit Sub
DoReportStandardSub_Err:
  InError = True
  If ee Is Nothing Then
    eName = "(No current employee)"
  Else
    eName = ee.FullName
  End If
  Call ErrorMessage(ERR_ERROR, Err, "DoReportStandardSub", "Executing Report", "Error executing report " & sReportName & vbCrLf & "Employee = " & eName)
  Resume DoReportStandardSub_End
  Resume
End Sub
Private Sub cmdCancel_Click()
  m_Cancel = True
  
End Sub
Private Sub Form_Activate()
  Call DoStandardReportSubEx(m_rep, m_pr, m_ReportDest, m_ReportOrient, m_sReportName)
End Sub
