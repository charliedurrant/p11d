VERSION 5.00
Object = "{770120E1-171A-436F-A3E0-4D51C1DCE486}#1.0#0"; "atc2stat.ocx"
Begin VB.Form F_EmployeeLetter 
   Caption         =   "Employee Letter"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin P11D2022.CodeEditor ce 
      Height          =   7980
      Left            =   0
      TabIndex        =   1
      Tag             =   "EQUALISE"
      Top             =   45
      Width           =   7260
      _extentx        =   4471
      _extenty        =   3201
   End
   Begin atc2stat.TCSStatus sts 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
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

Private Enum SEL_COLOR_SET
  SCS_BLUE
  SCS_NORMAL
End Enum

Private Enum LETTER_MODE
  LM_OPENFILE = 1
  LM_OPENTEXT
End Enum

Private Panel As TCSPANEL
Private mCRS As clsFormResize
Private m_EmployeeLetterFileNumber As Long
Private mb_ReadOnly As Boolean
Private mb_Untitled As Boolean
Private m_vbm As VBMenu
Private m_FileType As FILE_TYPES

Private Sub ce_MenuClick(ByVal vbm As atc2dmenu.VBMenu, ByVal vbmi As atc2dmenu.VBMenuItem)
  Call ControlCodeClick(vbmi.Tag)
End Sub

Private Sub dmenu_MenuClick(ByVal vbm As atc2dmenu.VBMenu, ByVal vbmi As atc2dmenu.VBMenuItem)
  Call ControlCodeClick(vbmi.Tag)
End Sub

Private Sub Form_Load()
  Set Panel = sts.AddPanel(100, , , "PanelEmpLet")
   
  ce.Font.size = p11d32.ReportPrint.EmployeeLetterFontSize
  ce.Font.Name = p11d32.ReportPrint.EmployeeLetterFontName
  Set mCRS = New clsFormResize
  Call mCRS.InitResize(Me, 9045, 7440)
  Call LoadControlCodes
  Call LoadLastLetter
  
End Sub

Public Function IsBackUpLetterFile(ByVal sPathAndFile As String) As Boolean
  On Error GoTo IsBackUpLetterFile_ERR
    
  Call xSet("IsBackUpLetterFile")
  'IsBackUpLetterFile = StrComp(sPathAndFile, p11d32.EmployeeLetterPath & p11d32.LetterFile & S_EMPLOYEE_LETTER_BACKUP_FILE_EXTENSION) = 0
  IsBackUpLetterFile = StrComp(sPathAndFile, p11d32.workingDirectory & S_USERDIR_ULETTERS & p11d32.LetterFile & S_EMPLOYEE_LETTER_BACKUP_FILE_EXTENSION) = 0
    
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
    
 
  Call EmployeeLetterAddCodes(Me.ce, False)
 
  ce.FormHWnd = Me.hwnd
  
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
    ce.Text = GetFileText(sPathAndFile)

    
    'Call ColorCodes
    Call SetFileType
    Call SetSave
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
  ce.Dirty = False
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
   
  
  If ce.Dirty And Not p11d32.ReportPrint.IsMasterLetterFile(p11d32.ReportPrint.EmployeeLetterPathAndFile) Then 'JN
      If MsgBox(sMsg, vbQuestion Or vbOKCancel, "Change File") = vbOK Then
      Close m_EmployeeLetterFileNumber
      ChangeFile = True
      Panel.Caption = ""
      ce.Text = ""
      ce.Dirty = False
    Else
      ChangeFile = False
      GoTo ChangeFile_END 'JN
    End If
  Else
    Close m_EmployeeLetterFileNumber
    ce.Text = ""
    Panel.Caption = ""
    ce.Dirty = False
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
Private Sub Form_Unload(Cancel As Integer)
  Set mCRS = Nothing
  Set F_EmployeeLetter = Nothing
End Sub
Private Function ControlCodeClick(Index As Long) As String
  Dim sCode As String
  
  sCode = EmployeeLetterCode(Index, ELCT_LETTER_FILE_CODES, False)
  ce.InsertCodeIntoText (sCode)
End Function
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
  On Error GoTo err_err
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
err_end:
  If (pr <> -1) Then p11d32.ReportPrint.DefaultReportIndex = pr
  If (es <> -1) Then p11d32.ReportPrint.EmployeeSelection = es
  F_Print.FromEmployeeLetter = False
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "Preview", "Preview", "Failed to preview the report")
  Resume err_end
End Sub

Private Sub mnuFileSave_Click()
  If mb_ReadOnly Or mb_Untitled Then
    Call FileSaveAs
  Else
    Call FileSave(p11d32.ReportPrint.EmployeeLetterPathAndFile, False)
  End If
End Sub
Private Function FileSave(ByVal sPathAndFile As String, bNewFile As Boolean) As Boolean
  Dim fs As FileSystemObject
  Dim ts As TextStream
  
  On Error GoTo FileSave_ERR
  
  Call xSet("FileSave")
  
  Set fs = New FileSystemObject
  
  If m_EmployeeLetterFileNumber <> 0 Then
    Close m_EmployeeLetterFileNumber
  End If
  
  If IsFileOpen(sPathAndFile, True) Then Call Err.Raise(ERR_FILE_OPEN_EXCLUSIVE, "FileSave", "The file " & sPathAndFile & " is opened exclusively.")
  
  Call TextFileSave(sPathAndFile, ce.Text)
  
  ce.Dirty = False
  
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
    sFIle = FileSaveAsDlg("Save As Letter File", "Employee letters (*" & S_EMPLOYEE_LETTER_FILE_EXTENSION & ")|*" & S_EMPLOYEE_LETTER_FILE_EXTENSION, p11d32.ReportPrint.UserLettersPathActual)
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
Private Sub SetFileType()
  On Error GoTo SetFileType_Err
  Call xSet("SetFileType")
  If StrComp(p11d32.ReportPrint.EmployeeLetterPath, p11d32.ReportPrint.UserLettersPathActual, vbTextCompare) = 0 Then
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
