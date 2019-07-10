VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Object = "{A7CE771F-05B2-43CF-9650-ED841A9049FA}#1.0#0"; "ATC3FolderBrowser.ocx"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_MM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magnetic Media"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin atc3FolderBrowser.FolderBrowser fb 
      Height          =   510
      Left            =   90
      TabIndex        =   18
      Top             =   3735
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   900
   End
   Begin VB.CommandButton cmdMMViewOptions 
      Caption         =   "View &Options"
      Height          =   375
      Left            =   6930
      TabIndex        =   5
      Top             =   1665
      Width           =   1140
   End
   Begin VB.CommandButton cmdViewLastFile 
      Caption         =   "&View Last File"
      Height          =   375
      Left            =   6930
      TabIndex        =   4
      Top             =   1260
      Width           =   1140
   End
   Begin VB.CommandButton cmdPrintErrors 
      Caption         =   "&Print Errors"
      Height          =   375
      Left            =   6930
      TabIndex        =   7
      Top             =   2475
      Width           =   1140
   End
   Begin VB.CommandButton cmdViewErrors 
      Caption         =   "V&iew Errors"
      Height          =   375
      Left            =   6930
      TabIndex        =   6
      Top             =   2070
      Width           =   1140
   End
   Begin VB.CommandButton cmdViewMM 
      Caption         =   "&View File"
      Height          =   375
      Left            =   6930
      TabIndex        =   3
      Top             =   855
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6930
      TabIndex        =   2
      Top             =   450
      Width           =   1140
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   6930
      TabIndex        =   1
      Top             =   45
      Width           =   1140
   End
   Begin VB.Frame fraDataFormat 
      Caption         =   "Data format"
      Height          =   1140
      Left            =   6750
      TabIndex        =   15
      Top             =   4200
      Width           =   1335
      Begin VB.PictureBox pctFrame 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   45
         ScaleHeight     =   825
         ScaleWidth      =   1230
         TabIndex        =   17
         Top             =   225
         Width           =   1230
         Begin VB.OptionButton optDataFormat 
            Caption         =   "3.5"" floppy"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   11
            Top             =   420
            Width           =   1140
         End
         Begin VB.OptionButton optDataFormat 
            Caption         =   "Any"
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   10
            Top             =   45
            Value           =   -1  'True
            Width           =   1005
         End
      End
   End
   Begin TCSPROG.TCSProgressBar prgEmployee 
      Height          =   375
      Left            =   45
      TabIndex        =   14
      Top             =   4950
      Width           =   6645
      _cx             =   11721
      _cy             =   661
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
      Value           =   0
      BarBackColor    =   12632256
      BarForeColor    =   8388608
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   2
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   0   'False
      Increment       =   1
      TextAlignment   =   0
   End
   Begin atc2valtext.ValText vtSubReturnOf 
      Height          =   330
      Left            =   1710
      TabIndex        =   9
      Top             =   3390
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "1"
      Minimum         =   "1"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText vtSubReturn 
      Height          =   330
      Left            =   900
      TabIndex        =   8
      Top             =   3390
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "1"
      Minimum         =   "1"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin MSComctlLib.ListView lvMMEmployers 
      Height          =   3300
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   5821
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employer name"
         Object.Width           =   2222
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PAYE ref"
         Object.Width           =   1453
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Valid paye"
         Object.Width           =   1614
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "No. of employees"
         Object.Width           =   2485
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Submitter name"
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Submitter ref"
         Object.Width           =   1879
      EndProperty
   End
   Begin VB.Label lblWritingFile 
      Height          =   510
      Left            =   90
      TabIndex        =   16
      Top             =   4350
      Width           =   6555
   End
   Begin VB.Label Label1 
      Caption         =   "sub return"
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   3435
      Width           =   1095
   End
   Begin VB.Label lblOF 
      Caption         =   "of"
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   3435
      Width           =   240
   End
End
Attribute VB_Name = "F_MM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFrmGeneral
Implements IViewFile

Private m_InvalidVT As Control
Public OK As Boolean

Private Sub cmdCancel_Click()
  Me.Hide
  p11d32.MagneticMedia.UserDataSize = 0
End Sub

Private Sub cmdMMViewOptions_Click()
  F_Input.ValText.TypeOfData = VT_STRING
  F_Input.ValText.MaxLength = 2
  If F_Input.Start("Record Code", "Enter record code ie 2E", p11d32.MagneticMedia.RecordViewID) Then
    p11d32.MagneticMedia.RecordViewID = F_Input.ValText.Text
  End If
  Set F_Input = Nothing
End Sub


Private Sub cmdPrintErrors_Click()
  Call p11d32.MagneticMedia.Errors(PRINT_REPORT)
End Sub

Private Sub MMRun()
  On Error GoTo MMRun_ERR
  Call xSet("MMRun")
  
  Call SetCursor
  
  prgEmployee.Indicator = ValueOfMax
  If ValidateMMData Then
      
    Call p11d32.MagneticMedia.CreateMagneticMediaFiles(prgEmployee, lblWritingFile)
    If SetButtons(p11d32.MagneticMedia.ErrorCount > 0) Then
      If MsgBox("Warnings/Errors in magnetic media submission!" & vbCrLf & vbCrLf & "Do you wish to view the errors?", vbCritical Or vbYesNo, "Warnings and Errors") = vbYes Then
        p11d32.MagneticMedia.Errors (PREPARE_REPORT)
      End If
    End If
  End If
  
MMRun_END:
  prgEmployee.Indicator = None
  prgEmployee.value = 0
  Call ClearCursor
  Call xReturn("MMRun")
  Exit Sub
MMRun_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "MMRun", "MM Run", "Error running the magnetic media.")
  Resume MMRun_END
End Sub
Private Function ValidateMMData() As Boolean
  Dim ey As Employer
  Dim benEmployer As IBenefitClass, benEmployer2 As IBenefitClass
  Dim i As Long, j As Long
  Dim li As ListItem
  
  Call xSet("ValidateMMData")
  
  On Error GoTo ValidateMMData_ERR
  
  If InvalidFields(Me) Then Call Err.Raise(ERR_INVALID_FIELDS, "ValidateMMData", "Some of the data entry fields are invalid, please amend.")
  If Not FileExists(fb.Directory, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "ValidateMMData", "The directory " & fb.Directory & " does not exist.")
  
  For i = 1 To lvMMEmployers.listitems.Count
    Set li = lvMMEmployers.listitems(i)
    If Not li.Checked Then GoTo NEXT_ITEM
    Set ey = p11d32.Employers(li.Tag)
    If Not ey.MagneticMediaValid(True) Then GoTo ValidateMMData_END
    
    Set benEmployer = ey
    
    benEmployer.value(employer_MagneticMediaSelected) = li.Checked
    
    If Not benEmployer.value(employer_MagneticMediaSelected) Then GoTo NEXT_ITEM
    
    p11d32.MagneticMedia.SubmitterName = benEmployer.value(employer_SubmitterName_db)
    p11d32.MagneticMedia.SubmitterRef = benEmployer.value(employer_SubmitterRef_db)
        
    For j = 1 To lvMMEmployers.listitems.Count
      Set benEmployer2 = p11d32.Employers(li.Tag)
      If benEmployer.value(employer_MagneticMediaSelected) And Not benEmployer Is benEmployer2 Then
        If StrComp(benEmployer.value(employer_SubmitterRef_db), benEmployer2.value(employer_SubmitterRef_db), vbTextCompare) <> 0 Then
          Call Err.Raise(ERR_EMPLOYER_INVALID, "ValidateMMData", "Some employers selected have do not have the same submitter refs, please choose employers with the same references.")
        End If
      End If
    Next
NEXT_ITEM:
  Next
  
  'LK - Checking for a trailing \ when entering the file path
  
    
  p11d32.MagneticMedia.SubReturn = vtSubReturn.Text
  p11d32.MagneticMedia.SubReturnOf = vtSubReturnOf.Text
  
  For i = [MM_DF_FIRSTITEM] To [MM_DF_LASTITEM]
    If optDataFormat(i).value Then
      p11d32.MagneticMedia.DataFormat = i
      Exit For
    End If
  Next
  
  
  ValidateMMData = True
  
ValidateMMData_END:
  Call xSet("ValidateMMData")
  Exit Function
ValidateMMData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ValidateMMData", "Validate MM Data", "Error validating the data for Magnetic Media.")
  Resume ValidateMMData_END
  Resume
End Function

Private Sub cmdRun_Click()
  Call MMRun
End Sub

Private Sub cmdViewErrors_Click()
  Call p11d32.MagneticMedia.Errors(PREPARE_REPORT)
End Sub

Private Sub cmdViewLastFile_Click()
  Call F_ViewFile.ViewFile(p11d32.MagneticMedia.LastPathAndFileCreated, , p11d32.MagneticMedia.OutputDirectory, Me)
End Sub

Private Sub cmdViewMM_Click()
  Call F_ViewFile.ViewFile("", , p11d32.MagneticMedia.OutputDirectory, Me)
End Sub


Private Sub Form_Load()
  Call SetButtons(False)
  Set lvMMEmployers.SmallIcons = MDIMain.imlTree
  If Not FileExists(p11d32.MagneticMedia.OutputDirectory, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "F_MM", "The Magnetic Media directory does not exist = " & p11d32.MagneticMedia.OutputDirectory & ", unable to create export.")
  fb.Directory = p11d32.MagneticMedia.OutputDirectory
    
End Sub
Private Function SetButtons(bVisible As Boolean) As Boolean
  cmdPrintErrors.Visible = bVisible
  cmdViewErrors.Visible = bVisible
  SetButtons = bVisible
End Function

Private Sub Form_Unload(cancel As Integer)
  If Not p11d32 Is Nothing Then p11d32.MagneticMedia.UserDataSize = 0
End Sub

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  
End Property
Private Function HighLightRecord(rt As RichTextBox, ByVal sRecordID As String, Optional lCol As Long = vbBlue) As Long
  Dim l As Long, m As Long
  Dim lLens As String, lLenlf As Long

  On Error GoTo HighLightRecord_ERR
  
  Call xSet("HighLightRecord")
  

  If Len(sRecordID) = 0 Then GoTo HighLightRecord_END
  sRecordID = vbLf & sRecordID
  lLens = Len(sRecordID)
  lLenlf = Len(vbLf)
  
  l = rt.Find(sRecordID, 0)
  Do While l <> -1
    m = rt.Find(vbLf, l + lLens)
    If m <> -1 Then
      HighLightRecord = HighLightRecord + Abs(RTSelText(rt, l + lLenlf, (m - 1) - l, lCol))
    End If
    l = l + lLens
    l = rt.Find(sRecordID, l)
  Loop
  
HighLightRecord_END:
  Call xSet("HighLightRecord")
  Exit Function
HighLightRecord_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HighLightRecord", "High Light Record", "Error highlighting a record with ID " & sRecordID & ".")
  Resume HighLightRecord_END
  
End Function

Private Sub IViewFile_View(rt As RichTextLib.IRichText, sLabelInfo As String)
 On Error GoTo IViewFile_MagneticMedia_ERR
  
  Call xSet("IViewFile_MagneticMedia")
  
  sLabelInfo = HighLightRecord(rt, p11d32.MagneticMedia.RecordType(MM_REC_EMPLOYEE)) & " employees "
  If Len(p11d32.MagneticMedia.RecordViewID) Then
    sLabelInfo = sLabelInfo & ", " & HighLightRecord(rt, p11d32.MagneticMedia.RecordViewID, vbRed) & " type " & p11d32.MagneticMedia.RecordViewID & " records."
  End If
  
IViewFile_MagneticMedia_END:
  Call xReturn("IViewFile_MagneticMedia")
  Exit Sub
IViewFile_MagneticMedia_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IViewFile", "IView File Magnetic Media", "Error setting the text properties when loading a file form the magnetic media dialogue.")
  Resume IViewFile_MagneticMedia_END
End Sub

Public Sub lvMMEmployers_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  cmdRun.Enabled = ListViewAnyChecked(lvMMEmployers)
End Sub

Private Sub lvMMEmployers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call EditEmployerDetails(Button, Shift, x, y)
End Sub

Private Sub EditEmployerDetails(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim ibf As IBenefitForm2
  Dim li As ListItem, liMe As ListItem, liPrevious As ListItem
  
  On Error GoTo EditEmployerDetails_ERR
   
  Call xSet("EditEmployerDetails")
  
  If Not ((Button And vbRightButton) = vbRightButton) Then GoTo EditEmployerDetails_END
    
  Set ibf = CurrentForm
    
  If ibf Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, Err, "The benfit for is nothing") 'JN
  If Not ibf.benclass = BC_EMPLOYER Then Call Err.Raise(ERR_IS_NOTHING, Err, "Edit Employer Details") 'JN
  
  Set liMe = lvMMEmployers.HitTest(x, y)
  
  If liMe Is Nothing Then GoTo EditEmployerDetails_END
  
  For Each li In ibf.lv.listitems
    If li.Tag = liMe.Tag Then
      Set liPrevious = ibf.lv.SelectedItem
      Set ibf.lv.SelectedItem = li
      Call p11d32.EditEmployer(li.Tag)
      Call p11d32.MagneticMedia.UpdateListViewItem(liMe, p11d32.Employers(li.Tag))
      Set ibf.lv.SelectedItem = liPrevious
      Exit For
    End If
  Next
  
EditEmployerDetails_END: 'JN
  Call xReturn("EditEmployerDetails")
  Exit Sub
EditEmployerDetails_ERR: 'JN
  Call ErrorMessage(ERR_ERROR, Err, "EditEmployerDetails", "Edit Employer Details", "Error editing Employer Details from the magnetic media screen.")
  Resume EditEmployerDetails_END
  Resume
End Sub


Private Sub vtSubReturn_Change()
  If Not vtSubReturn.FieldInvalid Then vtSubReturnOf.Minimum = CLng(vtSubReturn.Text)
End Sub

Private Sub vtSubReturnOf_Change()
  If Not vtSubReturnOf.FieldInvalid Then vtSubReturn.Maximum = CLng(vtSubReturnOf.Text)
End Sub
Private Sub fb_Ended()
  p11d32.MagneticMedia.OutputDirectory = fb.Directory
End Sub

Private Sub fb_Started()
  fb.Directory = p11d32.MagneticMedia.OutputDirectory
End Sub

