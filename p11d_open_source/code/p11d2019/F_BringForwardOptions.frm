VERSION 5.00
Begin VB.Form F_BringForwardOptions 
   Caption         =   "Bring Forward Sections"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   4650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "Select all / deselect all"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   2040
   End
   Begin VB.Frame fraOther 
      Caption         =   "Other"
      Height          =   1680
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   4560
      Begin VB.CheckBox chkComments 
         Caption         =   "Employee comments"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   1320
         Width           =   2760
      End
      Begin VB.CheckBox chkCDCs 
         Caption         =   "Company defined categories"
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   945
         Value           =   1  'Checked
         Width           =   3030
      End
      Begin VB.CheckBox chkSharedVans 
         Caption         =   "Shared vans"
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   630
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkCDBs 
         Caption         =   "Company defined benefits"
         Height          =   285
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Value           =   1  'Checked
         Width           =   3075
      End
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Options"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   225
      Value           =   1  'Checked
      Width           =   4470
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3420
      TabIndex        =   0
      Top             =   5040
      Width           =   1140
   End
End
Attribute VB_Name = "F_BringForwardOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub SettingsToScreen()
  On Error GoTo SettingsToScreen_Err
  
  Call xSet("SettingsToScreen")
  
  chkCDBs.value = BoolToChkBox(p11d32.BringForward.CDBs)
  chkCDCs.value = BoolToChkBox(p11d32.BringForward.CDCs)
  chkComments.value = BoolToChkBox(p11d32.BringForward.Comments)
  chkSharedVans.value = BoolToChkBox(p11d32.BringForward.SharedVans)
  
  
  Call HMITSectionsToScreen
  
SettingsToScreen_End:
  Call xReturn("SettingsToScreen")
  Exit Sub
SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error placing the setting to the screen for F_BringForwardOptions.")
  Resume SettingsToScreen_End
End Sub
Private Function CheckSelectAll() As Boolean
  Dim c As Control
  Dim b As Boolean
  
  For Each c In Me.Controls
    If TypeOf c Is CheckBox And (Not c Is chkSelectAll) Then
      b = (c.value = vbChecked)
      If Not b Then Exit For
    End If
  Next
  
  If b Then chkSelectAll.value = vbChecked
  CheckSelectAll = b
End Function
Private Function GetHMITSectionsToDisplay() As Long
  GetHMITSectionsToDisplay = -1
  Select Case p11d32.AppYear
    Case Else
      GetHMITSectionsToDisplay = -1 Xor (2 ^ HMIT_A)
    'Case Else
    '  Call ECASE("Invalid AppYear in get HMITSectionsToDisplay")
      'we take out the sections as appropriate from year to year
  End Select
  
End Function
Private Sub HMITSectionsToScreen()
  Dim lHMITToDisplay As Long
  Dim HS As HMIT_SECTIONS
  Dim bNotFirst As Boolean
  Dim i As Long, j As Long
  
  
  On Error GoTo HMITSectionsToScreen_ERR
  
  Call xSet("Start")
  
  
  lHMITToDisplay = GetHMITSectionsToDisplay
  
  For HS = HMIT_FIRST_ITEM To HMIT_LAST_ITEM
    j = 2 ^ HS
    If j And lHMITToDisplay Then
      If bNotFirst Then
        i = i + 1
        Load chkOptions(i)
        chkOptions(i).Visible = True
        chkOptions(i).Top = chkOptions(i - 1).Top + (chkOptions(i - 1).Height)
        chkOptions(i).Left = chkOptions(i - 1).Left
        chkOptions(i).Width = chkOptions(i - 1).Width
        chkOptions(i).Height = chkOptions(i - 1).Height
        chkOptions(i).Tag = HS
      Else
        chkOptions(i).Tag = HS
      End If
      If j And p11d32.BringForward.HMITSChosen Then
        chkOptions(i).value = vbChecked
      Else
        chkOptions(i).value = vbUnchecked
      End If
      
      chkOptions(i).Caption = p11d32.Rates.HMITSectionToHMITDescription(HS)
      bNotFirst = True
    End If
  Next
  fraOther.Top = chkOptions(i).Top + (2 * chkOptions(i).Height)
  cmdOK.Top = fraOther.Top + fraOther.Height + chkOptions(0).Height

  chkSelectAll.Top = cmdOK.Top
  
  Me.Height = cmdOK.Top + cmdOK.Width + chkOptions(0).Height - (Me.Height - Me.ScaleHeight) + 100
  
  Call CheckSelectAll

  
HMITSectionsToScreen_END:
  Call xReturn("HMITSectionsToScreen")
  Exit Sub
HMITSectionsToScreen_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HMITSectionsToScreen", "HMITSections To Screen", "Error placing the HMITsections to the screen.")
  Resume HMITSectionsToScreen_END
  Resume
End Sub

Private Sub chkSelectAll_Click()
  Dim c As Control
  
  For Each c In Me.Controls
    If TypeOf c Is CheckBox Then
      If Not c Is chkSelectAll Then
        c.value = chkSelectAll.value
      End If
    End If
  Next
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  Dim i As Long, j As Long
  
  Me.Hide
  For i = 0 To chkOptions.UBound
    If chkOptions(i).value = vbChecked Then
      j = j Or (2 ^ chkOptions(i).Tag)
    End If
  Next
  p11d32.BringForward.HMITSChosen = j
  p11d32.BringForward.CDBs = ChkBoxToBool(chkCDBs)
  p11d32.BringForward.CDCs = ChkBoxToBool(chkCDCs)
  p11d32.BringForward.Comments = ChkBoxToBool(chkComments)
  p11d32.BringForward.SharedVans = ChkBoxToBool(chkSharedVans)
  'AM To be removed
  
End Sub

Private Sub Form_Load()
  Call SettingsToScreen
End Sub

