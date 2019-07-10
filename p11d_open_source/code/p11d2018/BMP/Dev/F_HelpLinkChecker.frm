VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_HelpLinkChecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help links Checker"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save CSV"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   6015
      Begin VB.OptionButton optForms 
         Caption         =   "Forms"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optFormsControls 
         Caption         =   "Forms and Controls"
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdShowHelpLinks 
      Caption         =   "Show HelpLinks"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvHelpLinks 
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   6376
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblDir 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   6360
      Width           =   4335
   End
   Begin VB.Label lblDesc 
      Caption         =   "Visually check the help link is opening right page in Help viewer"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblSelect 
      Caption         =   "Select items with incorrect help link"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   4575
   End
End
Attribute VB_Name = "F_HelpLinkChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ms_FilenameAndPath As String

Private Sub Form_Load()
  On Error GoTo FormLoad_Err
  'ToDo - change refs
  
  ms_FilenameAndPath = App.Path & "\HelpLinksChecker.csv"
  lblDir.Caption = ms_FilenameAndPath
  optForms.value = True
  
  Call PopulateListView

FormLoad_End:
  Exit Sub
FormLoad_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Form_Load", "Check Help links Form load", "Error loading form")
  Resume FormLoad_End
End Sub

Private Sub cmdShowHelpLinks_Click()
  lvHelpLinks.listitems.Clear
  Call UpdateLvWithHelpLinks
End Sub

Private Sub PopulateListView()
  
  Const iCols As Integer = 4
  
  On Error GoTo PopulateListView_Err
  
  ' Add col headers for listview, set col widths
  'ObjName,ObjType,ObjHelpLink, Tooltip
  lvHelpLinks.ColumnHeaders.Add , , "ObjName", 3500  'lvHelpLinks.Width / iCols
  lvHelpLinks.ColumnHeaders.Add , , "ObjType", 800  'lvHelpLinks.Width / iCols
  lvHelpLinks.ColumnHeaders.Add , , "ObjHelpLink", 3250  'lvHelpLinks.Width / iCols
  lvHelpLinks.ColumnHeaders.Add , , "CtrlTooltip", 600
  
  lvHelpLinks.View = lvwReport  ' set view property to Report
  lvHelpLinks.FullRowSelect = True
  lvHelpLinks.LabelEdit = lvwManual  ' edit not enabled

PopulateListView_End:
  Exit Sub
PopulateListView_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PopulateListView", "PopulateListView", "Error setting up listview")
  Resume PopulateListView_End
End Sub

Private Sub UpdateLvWithHelpLinks()
  
  Dim li As ListItem
  Dim dctHFs As Dictionary
  Dim lFrmIndex As Long
  Dim lCtrlIndex As Long
  Dim cHF As HelpForm
  Dim cHC As HelpControl
  Dim lHelpFormsCount As Long
  Dim lHelpFormCtrlsCount As Long
  Dim sFrmName As String
  Dim sCtrlName As String
  
  On Error GoTo UpdateLvWithHelpLinks_Err
  
  'ObjName,ObjType,ObjHelpLink, Tooltip
  Set dctHFs = p11d32.Help.HelpForms
  lHelpFormsCount = dctHFs.Count - 1
  For lFrmIndex = 0 To lHelpFormsCount
    Set cHF = dctHFs.Items(lFrmIndex)
    sFrmName = dctHFs.Keys(lFrmIndex)
    Set li = lvHelpLinks.listitems.Add(, sFrmName, sFrmName)
    li.SubItems(1) = "frm"
    li.SubItems(2) = cHF.HelpLink
    li.SubItems(3) = ""
    If optFormsControls.value = True Then
      lHelpFormCtrlsCount = cHF.Controls.Count - 1
      For lCtrlIndex = 0 To lHelpFormCtrlsCount
        Set cHC = cHF.Controls.Items(lCtrlIndex)
        sCtrlName = cHF.Controls.Keys(lCtrlIndex)
        Set li = lvHelpLinks.listitems.Add(, sFrmName & "." & sCtrlName, sFrmName & "." & sCtrlName)
        li.SubItems(1) = "ctrl"
        li.SubItems(2) = cHC.HelpLink
        li.SubItems(3) = cHC.Tooltip
      Next lCtrlIndex
    End If
  Next lFrmIndex

UpdateLvWithHelpLinks_End:
  Exit Sub
UpdateLvWithHelpLinks_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateLvWithHelpLinks", "UpdateLvWithHelpLinks", "Error displaying data in listview")
  Resume UpdateLvWithHelpLinks_End
End Sub

Private Sub lvHelpLinks_ItemClick(ByVal Item As MSComctlLib.ListItem)
  DisplayListItemHelp (Item.SubItems(2))
  lvHelpLinks.SetFocus
End Sub

Private Sub DisplayListItemHelp(ByVal sHelpLink As String)
  Call p11d32.Help.ShowHelp(sDisplaySpecificHelpLink:=sHelpLink)
End Sub

Private Sub cmdSave_Click()
  Dim i As Long
  Dim li As ListItem
  Dim bDoSave As Boolean
  Dim s As String
  
  On Error GoTo err_Err
  
  bDoSave = False
  s = "Following helplinks are not displaying correct page: " & vbCrLf
  s = s & "ObjName,ObjType,ObjHelpLink, Tooltip" & vbCrLf

  For i = 1 To lvHelpLinks.listitems.Count
    Set li = lvHelpLinks.listitems(i)
    If li.Checked = True Then
      s = s & li.Text & "," & li.SubItems(1) & "," & li.SubItems(2) & "," & li.SubItems(3) & vbCrLf
      bDoSave = True
    End If
  Next i
  
  If bDoSave Then
    Call SaveTextFile(ms_FilenameAndPath, s)
    MsgBox "File saved."
  Else
    MsgBox "No items are selected to save."
  End If
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SaveCSV", "SaveCSV", "Error saving file:'" & ms_FilenameAndPath & "'")
  Resume err_End
End Sub

