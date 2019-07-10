VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_AddRemoveFindFolders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add / Remove Folders"
   ClientHeight    =   4515
   ClientLeft      =   1335
   ClientTop       =   1455
   ClientWidth     =   3570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   926
      ButtonWidth     =   609
      ButtonHeight    =   767
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   972
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1440
      TabIndex        =   2
      Top             =   4080
      Width           =   972
   End
   Begin VB.TextBox txtStartAt 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   660
      Width           =   3075
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   345
      Left            =   3090
      Picture         =   "AddRemove.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   660
      Width           =   435
   End
   Begin VB.ListBox lst 
      Height          =   2595
      Left            =   30
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Folders to search:"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Enter folder:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   460
      Width           =   1095
   End
End
Attribute VB_Name = "F_AddRemoveFindFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ff As FindFiles
Private m_Cancel As Boolean
Public Function Start() As Boolean
  Dim vPaths As Variant
  Dim i As Long
  Dim s As String

  On Error GoTo Start_ERR

  Call lst.Clear

  For i = 1 To GetDelimitedValues(vPaths, p11d32.FindFilesDirList, True, True, ";")
    lst.AddItem (vPaths(i))
  Next

  tbar.Buttons(2).Enabled = lst.ListCount > 0
  
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  
  If Not m_Cancel Then
    For i = 0 To lst.ListCount - 1
      s = s & Trim$(lst.List(i))
      If i <> lst.ListCount - 1 Then
        s = s & ";"
      End If
    Next
    p11d32.FindFilesDirList = s
  End If

  Start = Not m_Cancel

Start_END:
  Exit Function

Start_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "StartFindFiles", "Error in Start", "Error in Start for Add/Remove folders to search.")
Resume Start_END
End Function
Private Sub cmdBrowse_Click()
  Dim sPath As String
  
  sPath = BrowseForFolderEx(Me.hwnd, p11d32.WorkingDirectory)
  txtStartAt.Text = sPath
  If Right$(txtStartAt.Text, 1) <> "\" Then txtStartAt.Text = txtStartAt.Text & "\"
  
  End Sub
Private Sub cmdCancel_Click()
  m_Cancel = True
  Me.Hide
End Sub
Private Sub cmdOK_Click()
  Dim i As Integer
  Dim s As String

  For i = 0 To lst.ListCount - 1
    lst.ListIndex = i
    s = s & lst.Text
    If i < lst.ListCount - 1 Then s = s & ";"
  Next
  p11d32.FindFilesDirList = s
  Call p11d32.DirList(Ini_Write)
  Me.Hide
        
End Sub
Private Sub Form_Load()
  Set ff = New FindFiles
  Call AddAddDelete(tbar)
  txtStartAt = p11d32.WorkingDirectory
  
End Sub

Private Sub tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Call AddDelClick(Button.Index)
End Sub
Private Function AddDelClick(ByVal lButtonIndex As Long) As Boolean
  Dim sFolder As String
  Dim i As Integer
  Dim l As Integer
  
  On Error GoTo AddDelClick_Err
  
  With lst
    Select Case lButtonIndex
      Case 1 'tick
        sFolder = FullPath(txtStartAt.Text)
        If Len(sFolder) = 0 Then
          txtStartAt.SetFocus
          Call Err.Raise(ERR_ZERO_LENGTH_STRING, "Add", "No folder entered")
        End If

        If Not FileExists(sFolder, True) Then
          txtStartAt.SetFocus
          txtStartAt.SelStart = 0
          txtStartAt.SelLength = Len(sFolder)
          Call Err.Raise(ERR_DIRECTORY_NOT_EXIST, "Add", "The folder " & sFolder & " does not exist.")
        End If
                  
        For i = 0 To lst.ListCount - 1
          If StrComp(lst.List(i), sFolder, vbTextCompare) = 0 Then
              txtStartAt.SetFocus
              txtStartAt.SelLength = Len(sFolder)
              Call Err.Raise(ERR_DIRECTORY_EXISTS, "Add", "The folder " & sFolder & " is already present in the list")
          End If
        Next
        
        txtStartAt = ""
        lst.AddItem sFolder
        txtStartAt.SetFocus
        
      Case 2 'cross
        l = .ListIndex
          If l < 0 Then
          Call Err.Raise(ERR_ZERO_LENGTH_STRING, "Delete", "No folder selected to delete")
        Else
          If .ListIndex = .ListCount - 1 Then
           .RemoveItem l
           .ListIndex = l - 1
          Else
           .RemoveItem l
           .ListIndex = l
          End If
        End If
      End Select
    
  End With

  tbar.Buttons(2).Enabled = lst.ListCount > 0
AddDelClick_End:
  Exit Function

AddDelClick_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddDelClick", "Add Delete Click", "Error interpreting a click for add/delete buttons.")
  Resume AddDelClick_End
End Function

Private Sub txtStartAt_Change()
  tbar.Buttons(1).Enabled = Len(txtStartAt) > 0
End Sub
