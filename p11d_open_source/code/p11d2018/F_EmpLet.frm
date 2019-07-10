VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "atc2stat.ocx"
Object = "{D08C90A4-2337-4BE1-8137-EB1A093571A4}#1.0#0"; "atc2dmenu.ocx"
Begin VB.Form F_EmployeeLetter 
   Caption         =   "Employee Letter"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin atc2dmenu.DMenu dmenu 
      Left            =   1200
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin atc2stat.TCSStatus sts 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   8010
      Width           =   7275
      _ExtentX        =   12832
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
   Begin RichTextLib.RichTextBox rtEmpLet 
      Height          =   7980
      Left            =   0
      TabIndex        =   0
      Tag             =   "EQUALISE"
      Top             =   45
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   14076
      _Version        =   393217
      BackColor       =   12648447
      ScrollBars      =   3
      RightMargin     =   65535
      TextRTF         =   $"F_EmpLet.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As.."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "P&review"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "F_EmployeeLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Enum CURSOR_ACTION
  CA_SELECT_LEFT
  CA_SELECT_RIGHT
  CA_MOVE_LEFT
  CA_MOVE_RIGHT
  CA_NONE
End Enum

Private Type BLOCK_DEL
  LineTextCurrent As String
  StartSelPos As Long
  InShift As Boolean
  EraseKeyPressed As Boolean
  SelLength As Long
End Type

Private Enum LETTER_MODE
  LM_OPENFILE = 1
  LM_OPENTEXT
    End Enum

Private Enum SEL_COLOR_SET
  SCS_BLUE
  SCS_NORMAL
End Enum

Private BD As BLOCK_DEL
Private Panel As TCSPANEL
Private mCRS As clsFormResize
Private m_EmployeeLetterFileNumber As Long
Private m_dirty As Boolean
Private mb_ReadOnly As Boolean
Private mb_Untitled As Boolean
Private m_vbm As VBMenu
Private m_FileType As FILE_TYPES



Private Sub RecordKeyDown(KeyCode As Integer, Shift As Integer)

  On Error GoTo RecordKeyDown_ERR
  
  Call xSet("RecordKeyDown")
     
  BD.LineTextCurrent = rtEmpLet.Text
  BD.SelLength = rtEmpLet.SelLength
  
  If (Shift And vbShiftMask) Then
    If Not BD.InShift Then
      BD.InShift = True
      BD.StartSelPos = rtEmpLet.SelStart
    End If
  Else
    BD.InShift = False
    BD.StartSelPos = -1
  End If
  If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
    BD.EraseKeyPressed = True
  Else
    BD.EraseKeyPressed = False
  End If
  
RecordKeyDown_END:
  Call xReturn("RecordKeyDown")
  Exit Sub
RecordKeyDown_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RecordKeyDown", "Record Key Down", "Error recording the keydown in the employee letter.")
  Resume RecordKeyDown_END
  
End Sub
Private Function InsideBrace(lStartBrace As Long, lEndBrace As Long, sTextToSearch As String, lStartPos As Long) As Boolean
  Dim l As Long, m As Long, n As Long, o As Long
  
  On Error GoTo InsideBrace_ERR
  
  Call xSet("InsideBrace")
  
  lStartBrace = 0
  lEndBrace = 0
  
  If Len(sTextToSearch) Then
    l = InStr(lStartPos, sTextToSearch, "}", vbTextCompare)
    If l > 0 Then
      m = InStr(lStartPos, sTextToSearch, "{", vbTextCompare)
      If m = 0 Or m > l Then
        n = InStrRev(sTextToSearch, "{", lStartPos, vbTextCompare)
        If n > 0 Then
          o = InStrRev(sTextToSearch, "}", lStartPos - 1, vbTextCompare)
          If (o = 0) Or o > 0 And o < n Then
            lStartBrace = n
            lEndBrace = l
            InsideBrace = True
          End If
        End If
      End If
    Else
      lStartBrace = 0
      lEndBrace = 0
    End If
  End If

InsideBrace_END:
  Call xReturn("InsideBrace")
  Exit Function
InsideBrace_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "InsideBrace", "Inside Brace", "Error determining whether the caret is inside a set of braces.")
  Resume InsideBrace_END
      
      
End Function
Private Sub SelCodes(KeyCode As Integer, CA As CURSOR_ACTION)
  Dim lStartBrace As Long, lEndBrace As Long
  
  On Error GoTo SelCodes_ERR
  
  Call xSet("SelCodes")
  
  Select Case CA
    Case CA_NONE
    Case CA_MOVE_RIGHT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rtEmpLet.SelStart + 1 + Abs(BD.EraseKeyPressed)) Then
        If BD.EraseKeyPressed Then
          rtEmpLet.SelStart = lStartBrace - 1
          rtEmpLet.SelLength = lEndBrace - lStartBrace
          rtEmpLet.SelText = ""
        Else
          rtEmpLet.SelStart = lEndBrace
        End If
      End If
    Case CA_MOVE_LEFT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rtEmpLet.SelStart + 1) Then
          If BD.EraseKeyPressed Then
            rtEmpLet.SelStart = lStartBrace - 1
            rtEmpLet.SelLength = lEndBrace - lStartBrace
            rtEmpLet.SelText = ""
          Else
            rtEmpLet.SelStart = lStartBrace - 1
          End If
      End If
    Case CA_SELECT_RIGHT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rtEmpLet.SelStart + 1 + rtEmpLet.SelLength) Then
        Select Case KeyCode
          Case vbKeyLeft, vbKeyUp
            rtEmpLet.SelLength = (lStartBrace - 1) - rtEmpLet.SelStart
          Case vbKeyRight, vbKeyDown
            rtEmpLet.SelLength = lEndBrace - rtEmpLet.SelStart
        End Select
      End If
    Case CA_SELECT_LEFT
      If InsideBrace(lStartBrace, lEndBrace, BD.LineTextCurrent, rtEmpLet.SelStart + 1) Then
        Select Case KeyCode
          Case vbKeyLeft, vbKeyUp
              rtEmpLet.SelStart = lStartBrace - 1
              rtEmpLet.SelLength = BD.StartSelPos - (lStartBrace - 1)
              BD.InShift = False
          Case vbKeyRight, vbKeyDown
              rtEmpLet.SelStart = lEndBrace - 1
              rtEmpLet.SelLength = 0
        End Select
      End If
  End Select

SelCodes_END:
  Call xReturn("SelCodes")
  Exit Sub
SelCodes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SelCodes", "Sel Codes", "Error selecting an employee letter control code.")
  Resume SelCodes_END
  
End Sub

Private Sub dmenu_MenuClick(ByVal vbm As atc2DMenu.VBMenu, ByVal vbmi As atc2DMenu.VBMenuItem)
  Call ControlCodeClick(vbmi.Tag)
End Sub
Private Sub Form_Load()
  Set Panel = sts.AddPanel(100, , , "PanelEmpLet")
  rtEmpLet.Font.size = p11d32.ReportPrint.EmployeeLetterFontSize
  rtEmpLet.Font.Name = p11d32.ReportPrint.EmployeeLetterFontName
  Set mCRS = New clsFormResize
  Call mCRS.InitResize(Me, 9045, 7440)
  Call LoadControlCodes
  Call LoadLastLetter
End Sub

Public Function IsBackUpLetterFile(ByVal sPathAndFile As String) As Boolean
  On Error GoTo IsBackUpLetterFile_ERR
    
  Call xSet("IsBackUpLetterFile")
  'IsBackUpLetterFile = StrComp(sPathAndFile, p11d32.EmployeeLetterPath & p11d32.LetterFile & S_EMPLOYEE_LETTER_BACKUP_FILE_EXTENSION) = 0
  IsBackUpLetterFile = StrComp(sPathAndFile, p11d32.WorkingDirectory & S_USERDIR_ULETTERS & p11d32.LetterFile & S_EMPLOYEE_LETTER_BACKUP_FILE_EXTENSION) = 0
    
IsBackUpLetterFile_END:
  Call xReturn("IsBackUpLetterFile")
  Exit Function
IsBackUpLetterFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IsBackUpLetterFile", "Is Master File", "Error determining if the file " & sPathAndFile & " is a backup master letter file.")
  Resume IsBackUpLetterFile_END
End Function

Private Sub LoadControlCodes()
  
  Dim s As String
  Dim i As Long
  Dim vbm As VBMenu
  Dim vbmi As VBMenuItem
  
  
On Error GoTo LoadControlCodes_ERR
  
  Call xSet("LoadControlCodes")
  
  Set m_vbm = dmenu.Add("Menu")
  Set vbmi = m_vbm.Add(S_ELMC_MASTER, "&" & S_ELMC_MASTER, "")
  
  
  
  For i = [_ELMC_FIRST_ITEM] To [_ELMC_LAST_ITEM]
    s = EmployeeLetterMenuCaptions(i)
    Call m_vbm.Add(s, s, S_ELMC_MASTER)
  Next
  
  For i = EMPLOYEE_LETTER_CODE.ELC_FIRST_ITEM To EMPLOYEE_LETTER_CODE.ELC_LAST_ITEM
    s = EmployeeLetterCode(i, ELCT_MENU_CAPTION, False)
    Set vbmi = m_vbm.Add(s, s, EmployeeLetterCode(i, ELCT_MENU_PARENT, False))
    vbmi.Tag = i
  Next
  dmenu.hwnd = Me.hwnd
  
  
LoadControlCodes_END:
  Call xReturn("LoadControlCodes")
  Exit Sub
LoadControlCodes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LoadControlCodes", "Load Control Codes", "Error loading the control codes to the control code menu.")
  Resume LoadControlCodes_END
  Resume
End Sub
Public Function LoadLastLetter() As Long
  Dim sSource As String, sDestination As String
  Dim s As String
  On Error GoTo LoadLastLetter_ERR
  
  Call xSet("LoadLastLetter")
  
  If Not FileExists(p11d32.ReportPrint.EmployeeLetterPathAndFile) Then
    'check for letters directory
    If Not FileExists(p11d32.ReportPrint.SystemLettersPath, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "LoadLastLetter", "The directory " & p11d32.ReportPrint.SystemLettersPath & " does not exist, no employee letters to load.")
    If MsgBox("The last employee letter file does not exist." & vbCrLf & "File = " & p11d32.ReportPrint.EmployeeLetterPathAndFile & vbCrLf & "Do you wish to load the original file?", vbYesNo, "LoadLastLetter") = vbYes Then
      'check th original file
      'p11d32.EmployeeLetterPath = p11d32.LettersDirectoryMaster
      'If Not FileExists(p11d32.UserLettersDirectoryMaster, True) Then Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "LoadLastLetter", "The directory " & p11d32.UserEmployeeLetterPath & " does not exist, no user employee letters to load.")
      ' p11d32.UserEmployeeLetterPath = p11d32.UserLettersDirectoryMaster 'EK separation of user and application letters
      p11d32.ReportPrint.EmployeeLetterPath = p11d32.ReportPrint.SystemLettersPath
      p11d32.ReportPrint.EmployeeLetterFile = p11d32.LetterFile & S_EMPLOYEE_LETTER_FILE_EXTENSION
      If Not FileExists(p11d32.ReportPrint.EmployeeLetterPathAndFile) Then
        'recreate the original file
        If MsgBox("The original employee letter file does not exist." & vbCrLf & "File = " & p11d32.ReportPrint.EmployeeLetterPathAndFile & vbCrLf & "Do you wish to recreate the original file?", vbYesNo, "LoadLastLetter") = vbYes Then
          sSource = p11d32.ReportPrint.EmployeeLetterPath & p11d32.LetterFile & S_EMPLOYEE_LETTER_BACKUP_FILE_EXTENSION
          If FileExists(sSource) Then
            sDestination = p11d32.ReportPrint.EmployeeLetterPath & p11d32.LetterFile & S_EMPLOYEE_LETTER_FILE_EXTENSION
            If FileCopyEx(sSource, sDestination) Then
              p11d32.ReportPrint.EmployeeLetterFile = p11d32.LetterFile & S_EMPLOYEE_LETTER_FILE_EXTENSION
            Else
              Call Err.Raise(ERR_COPY_FAIL, "LoadLastLetter", "Unable to copy the file " & sSource & " to " & sDestination & ".")
            End If
          Else
            Call Err.Raise(ERR_FILE_NOT_EXIST, "LoadLastLetter", "The backup employee letter file " & sSource & " does not exist.")
          End If
        End If
      End If
    Else
      Call SetSave
      GoTo LoadLastLetter_END
    End If
  End If
  
  LoadLastLetter = OpenLetterFile(1, p11d32.ReportPrint.EmployeeLetterPathAndFile, , , LM_OPENTEXT)
  
LoadLastLetter_END:
  Call xReturn("LoadLastLetter")
  Exit Function
LoadLastLetter_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LoadLastLetter", "Load Last Letter", "Error loading an employee letter file ")
  Resume LoadLastLetter_END
  Resume
End Function
Private Sub SetSave()
  mnuFileSave.Enabled = CBoolean(m_FileType)
  mnuFileSaveAs.Enabled = True
  mb_ReadOnly = Not CBoolean(m_FileType)
End Sub
Private Sub CloseFile()
  On Error GoTo CloseFile_ERR
  
  Call xSet("CloseFile")
  
  If m_EmployeeLetterFileNumber > 0 Then
    Close m_EmployeeLetterFileNumber
    m_EmployeeLetterFileNumber = 0
  End If
  
CloseFile_END:
  Call xReturn("CloseFile")
  Exit Sub
CloseFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CloseFile", "Close File", "Error closing the current file, file handle = " & m_EmployeeLetterFileNumber & vbCrLf & vbCrLf & "File = " & p11d32.ReportPrint.EmployeeLetterPathAndFile)
  Resume CloseFile_END
End Sub
Private Function OpenLetterFile(bFileCaption As Boolean, ByVal sPathAndFile As String, Optional sText As String = "", Optional F_EmpLet As F_EmployeeLetter, Optional lm As LETTER_MODE) As Boolean
  Dim bMasterFile As Boolean
  Dim s As String
  Dim sPath As String, sFIle As String, sExt As String
  
  On Error GoTo OpenLetterFile_ERR
  
  Call xSet("OpenLetterFile")
     
  Call FileExistsAndNotOpenExclusive(sPathAndFile)
  
  If lm = LM_OPENTEXT Then
    rtEmpLet.Text = GetFileText(sPathAndFile)
    Call ColorCodes
    Call SetFileType
    Call SetSave   'IIf(Not CBoolean(m_FileType), True, False))
    'Call SetSave(True)
  End If
  
  If m_EmployeeLetterFileNumber > 0 Then Call Err.Raise(ERR_FILE_OPEN, "OpenLetterFile", "The file handle is non zero.")
  m_EmployeeLetterFileNumber = FreeFile

  'Open sPathAndFile For Input Lock Read Write As m_EmployeeLetterFileNumber
  If CBoolean(m_FileType) Then
    Open sPathAndFile For Input Lock Read Write As m_EmployeeLetterFileNumber
  Else
    Open sPathAndFile For Input Lock Read As m_EmployeeLetterFileNumber
  End If
  Call SplitPath(sPathAndFile, sPath, sFIle, sExt)
  p11d32.ReportPrint.EmployeeLetterFile = sFIle & sExt
  p11d32.ReportPrint.EmployeeLetterPath = sPath
  
  Panel.Caption = p11d32.ReportPrint.EmployeeLetterPathAndFile
  Me.Caption = "Employee Letter - "
  If bFileCaption Then
    Me.Caption = Me.Caption + p11d32.ReportPrint.EmployeeLetterFile
    mb_Untitled = False
    'If ReadOnly(sPathAndFile) Then
    If ReadOnly(sPathAndFile) Or (m_FileType <> FIT_USER_DEFINED) Then
      Me.Caption = Me.Caption & " [Read Only]"
      mb_ReadOnly = True
      Call CloseFile
    Else
      Call CloseFile
      mb_ReadOnly = False
    End If
  Else
    'JN has sorted
    Me.Caption = Me.Caption & S_UNTITLED
    mb_ReadOnly = False
    Call CloseFile
    mb_Untitled = True
  End If
  OpenLetterFile = True
  
OpenLetterFile_END:
  Call xReturn("OpenLetterFile")
  Exit Function
OpenLetterFile_ERR:
  Call ChangeFile(False, "")
  Call ErrorMessage(ERR_ERROR, Err, "OpenLetterFile", "Open Letter File", "Error opening the file " & sPathAndFile & ".")
  Resume OpenLetterFile_END
  Resume
End Function
Private Function ChangeFile(ByVal bFileCaption As Boolean, ByVal sNewPathAndFile As String) As Boolean
  Dim sMsg As String
    
  On Error GoTo ChangeFile_ERR
  
  Call xSet("ChangeFile")
  
  sMsg = "Are you sure you want to "
  
  If StrComp(sNewPathAndFile, p11d32.ReportPrint.EmployeeLetterPathAndFile, vbTextCompare) <> 0 Then 'JN
    sMsg = sMsg & "discard the changes you made to "
  Else
    sMsg = sMsg & "revert to the saved copy of "
  End If
  
  If mb_Untitled Then 'JN
      sMsg = sMsg & S_UNTITLED & "?"
  Else
    sMsg = sMsg & p11d32.ReportPrint.EmployeeLetterFile & "?"
  End If
   
  
  If m_dirty And Not p11d32.ReportPrint.IsMasterLetterFile(p11d32.ReportPrint.EmployeeLetterPathAndFile) Then 'JN
        If MsgBox(sMsg, vbQuestion Or vbOKCancel, "Change File") = vbOK Then
      Close m_EmployeeLetterFileNumber
      ChangeFile = True
      m_dirty = False
      Panel.Caption = ""
      rtEmpLet.Text = ""
    Else
      ChangeFile = False
      GoTo ChangeFile_END 'JN
    End If
  Else
    Close m_EmployeeLetterFileNumber
    rtEmpLet.Text = ""
    Panel.Caption = ""
    m_dirty = False
    ChangeFile = True
  End If
 
  If Len(sNewPathAndFile) > 0 Then Call OpenLetterFile(bFileCaption, sNewPathAndFile, , , LM_OPENTEXT)
  
ChangeFile_END:
  Call xReturn("ChangeFile")
  Exit Function
ChangeFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ChangeFile", "Close Current File", "Error closing the current employee letter file.")
  Resume ChangeFile_END
  Resume
  
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = Not ChangeFile(False, "")
End Sub

Private Sub Form_Resize()
  Call mCRS.Resize
End Sub
Private Sub FileOpen()
  Dim sPathAndFile As String
  Dim Flags As Long

  Dim sOldFileName As String
      
    Call xSet("FileOpen")
  
sOldFileName = p11d32.ReportPrint.EmployeeLetterPathAndFile

TRY_AGAIN:

  If HasFile Then Call CloseFile
  
  sPathAndFile = FileOpenDlg("Open letter file", "Employee letters (*" & S_EMPLOYEE_LETTER_FILE_EXTENSION & ")|*" & S_EMPLOYEE_LETTER_FILE_EXTENSION, p11d32.ReportPrint.EmployeeLetterPath)
    
  If Len(sPathAndFile) = 0 Then
    Call OpenLetterFile(Not mb_Untitled, sOldFileName, , , LM_OPENFILE)
    GoTo FileOpen_END
  End If
  
  If (InStr(1, sPathAndFile, S_EMPLOYEE_LETTER_FILE_EXTENSION, vbTextCompare) = 0) Then
    Call ErrorMessage(ERR_ERROR, Err, "FileOpen", "File Open", "The file you have have chosen does not have the file extension " & S_EMPLOYEE_LETTER_FILE_EXTENSION)
    GoTo TRY_AGAIN
  End If
  
  If Len(sPathAndFile) > 0 Then If Not ChangeFile(True, sPathAndFile) Then GoTo FileOpen_END
  
FileOpen_END:
  Call xReturn("FileOpen")
  End Sub
Private Sub ColorCodes()
  Dim i As Long
  
  On Error GoTo ColorCodes_ERR
  
  Call xSet("ColorCodes")
  
  For i = EMPLOYEE_LETTER_CODE.ELC_FIRST_ITEM To EMPLOYEE_LETTER_CODE.ELC_LAST_ITEM
    Call ColorCode(EmployeeLetterCode(i, ELCT_LETTER_FILE_CODES, False))
  Next
  rtEmpLet.SelStart = 0
  Call SetSelTextProperties(SCS_NORMAL)
  
ColorCodes_END:
  Call xReturn("ColorCodes")
  Exit Sub
ColorCodes_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ColorCodes", "Color Codes", "Error setting the color of the control codes in the current employee letter file.")
  Resume ColorCodes_END
End Sub
Private Sub ColorCode(sCode As String, Optional bFromMenu As Boolean = False)
  Dim l As Long
  
  On Error GoTo ColorCode_ERR
  
  Call xSet("ColorCode")
  
  
  If bFromMenu Then
    rtEmpLet.SelStart = rtEmpLet.SelStart - Len(sCode)
    rtEmpLet.SelLength = Len(sCode)
    Call SetSelTextProperties(SCS_BLUE)
    rtEmpLet.SelStart = rtEmpLet.SelStart + rtEmpLet.SelLength
  Else
    'bug with RT does not do ignore case for loop twice?
    Call ColorCodeEx(sCode)
    Call ColorCodeEx(LCase(sCode))
  End If
    
  
ColorCode_END:
  Call SetSelTextProperties(SCS_NORMAL)
  Call xReturn("ColorCode")
  Exit Sub
ColorCode_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ColorCode", "Color Code", "Error setting the color of the code " & sCode & " in the employee letter file.")
  Resume ColorCode_END
  
  
End Sub
Private Sub ColorCodeEx(sCode As String)
  Dim l As Long
  
  On Error GoTo ColorCodeEx_ERR
  
  Call xSet("ColorCodeEx")
  Do
      l = rtEmpLet.Find(sCode, l, , rtfWholeWord)
    If l <> -1 Then
      rtEmpLet.SelStart = l
      rtEmpLet.SelLength = Len(sCode)
      Call SetSelTextProperties(SCS_BLUE)
      rtEmpLet.SelStart = rtEmpLet.SelStart + Len(sCode)
      rtEmpLet.SelLength = 0
      Call SetSelTextProperties(SCS_NORMAL)
      l = l + 1
    Else
      Exit Do
    End If
    Loop While True
    
  
ColorCodeEx_END:
  Call SetSelTextProperties(SCS_NORMAL)
  Call xReturn("ColorCodeEx")
  Exit Sub
ColorCodeEx_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ColorCodeEx", "Color Code Ex", "Error setting the color of the code " & sCode & " in the employee letter file.")
  Resume ColorCodeEx_END
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mCRS = Nothing
  Set F_EmployeeLetter = Nothing
End Sub
Private Sub ControlCodeClick(Index As Long)
  Dim sCode As String
  
  sCode = EmployeeLetterCode(Index, ELCT_LETTER_FILE_CODES, False)
  rtEmpLet.SelText = sCode
  Call ColorCode(sCode, True)
  m_dirty = True
  
End Sub


Private Sub mnuFileExit_Click()
  Unload Me
End Sub
Private Sub FileNew()
  Dim sPathAndFile As String
  
  On Error GoTo FileNew_ERR
  
  Call xSet("FileNew")
  sPathAndFile = FullPath(p11d32.ReportPrint.SystemLettersPath) & p11d32.EmployeeLetterTemplateFile & S_EMPLOYEE_LETTER_FILE_EXTENSION
  If Not FileExists(sPathAndFile) Then Call Err.Raise(ERR_FILE_NOT_EXIST, "FileNew", "The file " & sPathAndFile & " does not exist.")
  
  Call CloseFile
  Call ChangeFile(False, sPathAndFile)
      
FileNew_END:
  Call xReturn("FileNew")
  Exit Sub
FileNew_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FileNew", "File New", "Error in file new letter.")
  Resume FileNew_END
End Sub
Private Sub mnuFileNew_Click()
  Call FileNew
End Sub

Private Sub mnuFileOpen_Click()
  Call FileOpen
End Sub

Private Sub mnuFilePreview_Click()
  Dim pr As P11D_REPORTS
  Dim es As EMPLOYEE_SELECTION
  On Error GoTo err_Err
  If (mb_Untitled) Then
    Call Err.Raise(ERR_PREVIEW, "Preview", "Please save the letter first")
  End If
  F_Print.FromEmployeeLetter = True
  pr = -1
  es = -1
  es = p11d32.ReportPrint.EmployeeSelection
  pr = p11d32.ReportPrint.DefaultReportIndex
  p11d32.ReportPrint.DefaultReportIndex = RPT_EMPLOYEE_LETTER
  'select only the current employee
  p11d32.ReportPrint.EmployeeSelection = ES_CURRENT
  Call F_Print.cmdPrintPreview_Click(1) 'cad todo need enum
err_End:
  If (pr <> -1) Then p11d32.ReportPrint.DefaultReportIndex = pr
  If (es <> -1) Then p11d32.ReportPrint.EmployeeSelection = es
  F_Print.FromEmployeeLetter = False
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Preview", "Preview", "Failed to preview the report")
  Resume err_End
End Sub

Private Sub mnuFileSave_Click()
  If mb_ReadOnly Or mb_Untitled Then
    Call FileSaveAs
  Else
    Call FileSave(p11d32.ReportPrint.EmployeeLetterPathAndFile, False)
  End If
End Sub
Private Function FileSave(ByVal sPathAndFile As String, bNewFile As Boolean) As Boolean
  Dim FS As FileSystemObject
  Dim ts As TextStream
  
  On Error GoTo FileSave_ERR
  
  Call xSet("FileSave")
  
  Set FS = New FileSystemObject
  
  If m_EmployeeLetterFileNumber <> 0 Then
    Close m_EmployeeLetterFileNumber
  End If
  
  If IsFileOpen(sPathAndFile, True) Then Call Err.Raise(ERR_FILE_OPEN_EXCLUSIVE, "FileSave", "The file " & sPathAndFile & " is opened exclusively.")
  
  Call TextFileSave(sPathAndFile, rtEmpLet.Text)
  
  m_dirty = False
  
  FileSave = True
  
FileSave_END:
  Call xReturn("FileSave")
  Exit Function
FileSave_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FileSave", "File Save", "Error saving the employee letter file " & sPathAndFile & ".")
  Resume FileSave_END
  Resume
End Function

Private Sub mnuFileSaveAs_Click()
  Call FileSaveAs
End Sub
Private Property Get HasFile() As Boolean
  HasFile = (m_EmployeeLetterFileNumber > 0)
End Property
Private Function FileSaveAs() As Boolean
  Dim sFIle As String
  Dim sExt As String
  Dim sPath As String
  Dim b_HasFileNumber As Boolean
  Dim s_OldFileName As String
  
  On Error GoTo FileSaveAs_ERR
  
  Call xSet("FileSaveAs")
  
  b_HasFileNumber = HasFile
  s_OldFileName = p11d32.ReportPrint.EmployeeLetterPathAndFile
  m_FileType = FIT_USER_DEFINED 'RK can only save user letters
    
TRY_AGAIN:
  If b_HasFileNumber Then Call CloseFile
    sFIle = FileSaveAsDlg("Save As Letter File", "Employee letters (*" & S_EMPLOYEE_LETTER_FILE_EXTENSION & ")|*" & S_EMPLOYEE_LETTER_FILE_EXTENSION, p11d32.ReportPrint.UserLettersPath)
    If Len(sFIle) = 0 Then
    Call OpenLetterFile(Not mb_Untitled, s_OldFileName, , , LM_OPENFILE)
    GoTo FileSaveAs_END
  End If
  
  
  Call SplitPath(sFIle, sPath, sFIle, sExt)
  If Len(sExt) > 0 And StrComp(sExt, S_EMPLOYEE_LETTER_FILE_EXTENSION, vbTextCompare) <> 0 Then
    Call ErrorMessage(ERR_ERROR, Err, "FileSave", "File Save", "The file you have chosen does not have the file extension " & S_EMPLOYEE_LETTER_FILE_EXTENSION)
    GoTo TRY_AGAIN
  End If
  If IsFormLoaded("F_PrintOptions") Then Call F_PrintOptions.AddNewLetterNode(sFIle, True)
  sFIle = FullPath(sPath) & sFIle & sExt
  
  If p11d32.ReportPrint.IsMasterLetterFile(sFIle) Then
    Call ErrorMessage(ERR_ERROR, Err, "FileSaveAs", "File Save As", "The file you have chosen is the same name as the master file, " & p11d32.LetterFile & S_EMPLOYEE_LETTER_FILE_EXTENSION)
    GoTo TRY_AGAIN
  End If
  
  If FileSave(sFIle, True) Then
    Call OpenLetterFile(True, sFIle, , , LM_OPENFILE)
    FileSaveAs = True
  End If
  
 p11d32.ReportPrint.EmployeeLetterPath = sPath
 Call SetFileType
FileSaveAs_END:
  Call xReturn("FileSaveAs")
  Exit Function
FileSaveAs_ERR:
  If Err.Number <> cdlCancel Then Call ErrorMessage(ERR_ERROR, Err, "FileSaveAs", "File Save As", "Error saving the employee letter file " & sFIle & ".")
  Resume FileSaveAs_END
  Resume
End Function
Private Function GetCursorAction(KeyCode As Integer) As CURSOR_ACTION
  On Error GoTo GetCursorAction_ERR
  
  Call xSet("GetCursorAction")

  If BD.EraseKeyPressed And BD.SelLength > 0 Then
    GetCursorAction = CA_NONE
    Exit Function
  End If
  
  If Not BD.InShift Then
    Select Case KeyCode
      Case vbKeyLeft, vbKeyUp, vbKeyBack
        GetCursorAction = CA_MOVE_LEFT
      Case vbKeyRight, vbKeyDown, vbKeyDelete
        GetCursorAction = CA_MOVE_RIGHT
    End Select
  Else
    Select Case rtEmpLet.SelStart
      Case BD.StartSelPos
        If rtEmpLet.SelLength > 0 Then
          GetCursorAction = CA_SELECT_RIGHT
        End If
      Case Is < BD.StartSelPos
        GetCursorAction = CA_SELECT_LEFT
    End Select
  End If
  
GetCursorAction_END:
  Call xReturn("GetCursorAction")
  Exit Function
GetCursorAction_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Get Cursor Action", "Get Cursor Action", "Error getting the cursor action")
  Resume GetCursorAction_END
End Function
Private Sub SetSelTextProperties(SCS As SEL_COLOR_SET)
  On Error GoTo SetSelTextProperties_ERR
  
  Call xSet("SetSelTextProperties")

  Select Case SCS
    Case SCS_BLUE
      rtEmpLet.SelItalic = True
      rtEmpLet.SelBold = True
      rtEmpLet.SelColor = vbBlue
    Case SCS_NORMAL
      rtEmpLet.SelItalic = False
      rtEmpLet.SelBold = False
      rtEmpLet.SelColor = vbBlack
  End Select
  
SetSelTextProperties_END:
  Call xReturn("SetSelTextProperties")
  Exit Sub
SetSelTextProperties_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetSelTextProperties", "Set SelText Properties", "Error setting the seltext font properties.")
  Resume SetSelTextProperties_END
End Sub

Private Sub rtEmpLet_KeyDown(KeyCode As Integer, Shift As Integer)
  m_dirty = True
  If KeyCode = 221 Or KeyCode = 219 Then '{}
    KeyCode = 0
  End If
  Call RecordKeyDown(KeyCode, Shift)
  If rtEmpLet.SelLength = 0 Then Call SetSelTextProperties(SCS_NORMAL)
End Sub

Private Sub rtEmpLet_KeyUp(KeyCode As Integer, Shift As Integer)
 
 Call SelCodes(KeyCode, GetCursorAction(KeyCode))
End Sub

Private Sub rtEmpLet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = vbRightButton Then
    Call m_vbm.Popup(S_ELMC_MASTER, X, Y)
  Else
    Call RecordKeyDown(-1, 0)
  End If
End Sub

Private Sub rtEmpLet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lStartBrace As Long
  
  If rtEmpLet.SelLength = 0 Then
    Call SelCodes(-1, CA_MOVE_LEFT)
  Else
    If InsideBrace(lStartBrace, 0, BD.LineTextCurrent, rtEmpLet.SelStart + 1) Then
      rtEmpLet.SelLength = 0
      rtEmpLet.SelStart = lStartBrace - 1
    ElseIf InsideBrace(lStartBrace, 0, BD.LineTextCurrent, (rtEmpLet.SelStart + rtEmpLet.SelLength + 1)) Then
      rtEmpLet.SelLength = 0
      rtEmpLet.SelStart = lStartBrace - 1
    End If
  End If
End Sub
Private Sub SetFileType()
  On Error GoTo SetFileType_Err
  Call xSet("SetFileType")
  If StrComp(p11d32.ReportPrint.EmployeeLetterPath, p11d32.ReportPrint.UserLettersPath, vbTextCompare) = 0 Then
    m_FileType = FIT_USER_DEFINED
  Else
    m_FileType = FIT_SYSTEM_DEFINED
  End If
  'Disable Save for system defined letters
  mnuFileSave.Enabled = CBoolean(m_FileType)
  
SetFileType_End:
  Call xReturn("SetFileType")
  Exit Sub

SetFileType_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetFileType", "Error in SetFileType", "Undefined error.")
  Resume SetFileType_End
End Sub

