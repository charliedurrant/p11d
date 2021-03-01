VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.1#0"; "atc2vtext.OCX"
Begin VB.Form F_FindDirectory 
   Caption         =   "Find Directory"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   525
      Width           =   1140
   End
   Begin VB.CommandButton cmdOutPutDir 
      Height          =   330
      Left            =   120
      Picture         =   "F_FindDirectory.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   375
      Width           =   375
   End
   Begin atc2valtext.ValText vtOutPutDir 
      Height          =   330
      Left            =   615
      TabIndex        =   1
      Top             =   360
      Width           =   3840
      _ExtentX        =   6773
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
      Text            =   ""
      TypeOfData      =   3
      AutoSelect      =   0
   End
End
Attribute VB_Name = "F_FindDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PS added form and code for TTP#184

Option Explicit

Private m_OK As Boolean
Private m_working_directory As String
Private TempDir As String
Public Property Get WorkingDirectory() As String
  WorkingDirectory = m_working_directory
End Property
Public Property Let WorkingDirectory(NewValue As String)
  m_working_directory = NewValue
End Property


Public Property Get OK() As Boolean
  OK = m_OK
End Property
Private Sub Form_Load()
    vtOutPutDir.Text = m_working_directory
    TempDir = m_working_directory
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    TempDir = vtOutPutDir.Text
    If validatedata Then
       m_working_directory = TempDir
        m_OK = True
        Me.Hide
    End If
End Sub

Private Sub cmdOutPutDir_Click()
    Dim cDir As String

    cDir = BrowseForFolderEx(Me.hwnd, TempDir, "Choose a Directory.")

    If Len(cDir) Then
        F_FindDirectory.vtOutPutDir.Text = FullPath(cDir)
        Call ChDriveUNC(cDir)
        Call ChDir(cDir)
    End If
End Sub

Private Function validatedata() As Boolean
    On Error GoTo ValidateData_ERR
  
    Call xSet("ValidateData")
  
    If FileExists(TempDir, True) Then
        TempDir = FullPath(TempDir)
        validatedata = True
    Else
      Call CreateWorkingDirectory(TempDir, "Working Directory", True)
      validatedata = False
    End If
  
ValidateData_END:
    Call xReturn("ValidateData")
    Exit Function
ValidateData_ERR:
    Call ErrorMessage(ERR_ERROR, Err, "ValidateData", "Validate Data", "Error validating directory")
    Resume ValidateData_END
End Function

Private Function CreateWorkingDirectory(sDefaultDirPath As String, sDefaultDirType As String, Optional bExportDir As Boolean = False) As Boolean
    On Error GoTo CreateWorkingDirectory_Err
    Call xSet("CreateWorkingDirectory")
    Dim sDirType_Name As String

    If MsgBox("The directory specified to be the " & sDefaultDirType & " does not exist." & vbCrLf & _
              "Do you want the working directory: " & sDefaultDirPath & " created?", vbOKCancel, "Missing working directory") = vbOK Then
        xMkdir (sDefaultDirPath)
      CreateWorkingDirectory = True
    End If
    
CreateWorkingDirectory_End:
    Call xReturn("CreateWorkingDirectory")
    Exit Function

CreateWorkingDirectory_Err:
    Call Err.Raise(ERR_DIRECTORY_CREATE, "CreateWorkingDirectory", "Could not create the directory " & sDefaultDirPath & " check rights to directory.")
    Resume CreateWorkingDirectory_End
End Function
