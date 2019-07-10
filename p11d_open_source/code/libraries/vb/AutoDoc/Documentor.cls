VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Documentor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum INI_PROCESS
  INI_READ = 1
  INI_WRITE
End Enum


Public LastPathAndFile As String
Public WorkingDir As String
Public VBPs As Collection

Public Function Initialise()
  Dim sPath As String
  
  Call IniFile(INI_READ)
  If FileExists(LastPathAndFile) Then
    Call SplitPath(LastPathAndFile, sPath)
    Call ShowMaximised(frmOutPut)
    ChDir (sPath)
  End If
  
End Function
Private Sub IniFile(IP As INI_PROCESS)
  Select Case IP
    Case INI_READ
      LastPathAndFile = GetIniEntry(S_INI_SECTION_FILE, "LastPathAndFile", "")
    Case INI_WRITE
      Call WriteIniEntry(S_INI_SECTION_FILE, "LastPathAndFile", LastPathAndFile)
  End Select
End Sub
Private Sub Class_Terminate()
  Call IniFile(INI_WRITE)
End Sub
Public Function LoadVBPs()
  Dim sPathAndFile As String
  Dim RF As IReadFile
  Dim sPath As String, s
  Dim VBP As VBP
  On Error GoTo ChoseFile_ERR
  
  
  Set
  Set VBP = New VBP
  Set RF = VBP
  Call SplitPath(sPathAndFile, sPath)
  WorkingDir = sPath
  LastPathAndFile = sPathAndFile
  Call RF.OpenFile(sPathAndFile)
  Call VBPToScreen
  
ChoseFile_END:
  Exit Function
ChoseFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ChoseFile", "Chose File", "Error chosing a file.")
  Resume ChoseFile_END
  Resume
End Function
Public Sub VBPToScreen()
  Dim s As String
  Dim Class As Class
  
  On Error GoTo FillScreen_ERR
  
  frmMain.txtVBPData = VBP.ScreenText
  frmMain.tvw.Nodes.Clear
  
  For Each Class In VBP.Classes
    Call frmMain.tvw.Nodes.Add(, , Class.Name, Class.Name, IMG_CROSS)
  Next
  If frmMain.tvw.Nodes.Count > 0 Then
    Call ClassToScreen(frmMain.tvw.Nodes(1).Text)
    frmMain.tvw.SelectedItem = frmMain.tvw.Nodes(1)
  End If
FillScreen_END:
  Exit Sub
FillScreen_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "FillScreen", "Fill Screen", "Error placing file to screen.")
  Resume FillScreen_END
End Sub
Public Sub ClassToScreen(sClassName As String)
  Dim s As String
  Dim Class As Class
  Dim FI As FunctionItem
  Dim param As Parameter
  
  Set Class = VBP.Classes(sClassName)
  
  frmOutPut.txtClassDetails = Class.ScreenText
  For Each FI In Class.Functions
    s = s & FI.ScreenText & vbCrLf
  Next
  
  frmOutPut.txtFunctions = s
End Sub
Private Function FunctionToScreen(FI As FunctionItem) As String
  
End Function
Private Function ParamToScreen()

End Function
