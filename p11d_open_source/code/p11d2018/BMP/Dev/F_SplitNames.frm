VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_SplitNames 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analyse Names"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView tvwChoices 
      Height          =   1095
      Left            =   2835
      TabIndex        =   3
      Top             =   630
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwNameParts 
      Height          =   1095
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1931
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2655
      TabIndex        =   5
      Top             =   3240
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3690
      TabIndex        =   6
      Top             =   3240
      Width           =   960
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<<"
      Height          =   375
      Left            =   1935
      TabIndex        =   2
      Top             =   1125
      Width           =   780
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      Height          =   375
      Left            =   1935
      TabIndex        =   1
      Top             =   675
      Width           =   780
   End
   Begin VB.Label lblInstructions 
      Caption         =   "The following fix is designed to split names into individual parts. Please select the name parts and order them."
      Height          =   465
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   4470
   End
   Begin VB.Label lblCaption 
      Caption         =   "Examples"
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   1800
      Width           =   1770
   End
   Begin VB.Label lblExamples 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblExamples"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   45
      TabIndex        =   4
      Top             =   2070
      Width           =   4650
   End
End
Attribute VB_Name = "F_SplitNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const S_TITLE As String = "Title"
Private Const S_FIRSTNAME As String = "First name"
Private Const S_SURNAME As String = "Surname"
Private Const S_INITITALS As String = "Initials"

Private Const L_Title  As Long = 2
Private Const L_FIRSTNAME As Long = 4
Private Const L_SURNAME As Long = 8
Private Const L_INITIALS As Long = 16

Public Complete As Boolean
Private m_rsEmployeesToSplit As Recordset


Private Function AddExamples(rs As Recordset) As Boolean
  Dim r As String, s As String
  Dim i As Long
  On Error GoTo AddExamples_Err
  Call xSet("AddExamples")

  Do While Not rs.EOF
    r = "" & rs.Fields("Name").value
    r = ReplaceString(r, ",", " ")
    r = ReplaceString(r, ".", "")
    s = s & r & vbCrLf
    i = i + 1
    If i = 5 Then
      rs.MoveFirst
      Exit Do
    End If
    rs.MoveNext
  Loop
        
  lblExamples = s
  
AddExamples_End:
  Call xReturn("AddExamples")
  Exit Function

AddExamples_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddExamples", "Add Examples", "Error adding examples to the split names form.")
  Resume AddExamples_End
  Resume
End Function

Public Function InitSplit(db As Database, ByVal QueryID As QUERY_NAMES) As Boolean
  
  Dim n As Node
  
  On Error GoTo InitSplit_Err
  
  Call xSet("InitSplit")
  
  If QueryID <> SELECT_UNSPLIT_EMPLOYEES And QueryID <> SELECT_EMPLOYEES_NO_CDB Then Call Err.Raise(ERR_INVALID_QUERY, "InitSplit", "Invalid split query.")
  
  Set m_rsEmployeesToSplit = db.OpenRecordset(sql.Queries(QueryID), dbOpenDynaset)
      
  If Not (m_rsEmployeesToSplit.EOF And m_rsEmployeesToSplit.BOF) Then
    Call AddExamples(m_rsEmployeesToSplit)
    
    With tvwNameParts
        Set n = .Nodes.Add(, , S_TITLE, S_TITLE)
        n.Tag = L_Title
        Set n = .Nodes.Add(, , S_FIRSTNAME, S_FIRSTNAME)
        n.Tag = L_FIRSTNAME
        Set n = .Nodes.Add(, , S_INITITALS, S_INITITALS)
        n.Tag = L_INITIALS
        Set n = .Nodes.Add(, , S_SURNAME, S_SURNAME)
        n.Tag = L_SURNAME
    End With
'    Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  Else
    Complete = True
  End If
        
  
InitSplit_End:
  Set n = Nothing
  
  Call xReturn("InitSplit")
  Exit Function

InitSplit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitSplit", "Init Split", "Error initialising the split names form.")
  Resume InitSplit_End
  Resume
End Function

Private Sub cmdAdd_Click()
  Call PassNodes(tvwChoices, tvwNameParts)
End Sub

Private Sub cmdCancel_Click()
  Complete = False
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  'analyse selection
  Complete = AnalyseSelection
End Sub

Public Function AnalyseSelection() As Boolean
  Dim lRecord As Long, l As Long, lMax As Long, m As Long
  Dim n As Node
  Dim sName As String, sNamePart As String
  Dim bInitials As Boolean
  Dim bLookFromRight As Boolean
  
  On Error GoTo AnalyseSelection_Err
  Call xSet("AnalyseSelection")
  Call SetCursor
  If tvwChoices.Nodes.Count = 0 Then
    Call Err.Raise(ERR_NO_SPLIT_CHOICES, "AnalyseSelection", "Please select some name parts.")
    GoTo AnalyseSelection_End
  Else
      For Each n In tvwChoices.Nodes
        If n.Tag = L_INITIALS Then bInitials = True
      Next
      With m_rsEmployeesToSplit
        Call PrgStartCaption(Records(m_rsEmployeesToSplit), "Spliting employee names into their component parts.")
        Do While Not .EOF
          .Edit
          sName = .Fields("Name")
          sName = ReplaceString(sName, ",", " ")
          sName = ReplaceString(sName, ".", " ")
          lRecord = lRecord + 1
          
          m = 0
          
          For l = 1 To tvwChoices.Nodes.Count
            Set n = tvwChoices.Nodes(l)
            If n.Tag = L_INITIALS Then Exit For
            m = l
            Call GetWord(sNamePart, sName, False)
            Select Case n.Tag
              Case L_Title
                .Fields("Title") = sNamePart
              Case L_FIRSTNAME
                .Fields("FirstName") = sNamePart
              Case L_SURNAME
                .Fields("SurName") = sNamePart
            End Select
            sNamePart = ""
          Next l
          
          For l = tvwChoices.Nodes.Count To 1 Step -1
            If l = m Then Exit For
            Set n = tvwChoices.Nodes(l)
            If n.Tag = L_INITIALS Then Exit For
            Call GetWord(sNamePart, sName, True)
            Select Case n.Tag
              Case L_Title
                .Fields("Title") = sNamePart
              Case L_FIRSTNAME
                .Fields("FirstName") = sNamePart
              Case L_SURNAME
                .Fields("SurName") = sNamePart
            End Select
            sNamePart = ""
          Next l
          
          If bInitials Then .Fields("Initials") = Trim$(sName)
          
          If lRecord = 1 Then
            If MsgBox("Is the conversion correct ?" & vbCrLf & vbCrLf & .Fields("Name") & ": " & vbCrLf & vbCrLf & S_TITLE & " = " & .Fields("Title") & vbCrLf & S_FIRSTNAME & " = " & .Fields("FirstName") & vbCrLf & S_INITITALS & " = " & .Fields("Initials") & vbCrLf & S_SURNAME & " = " & .Fields("SurName"), vbOKCancel, "Check conversion") = vbCancel Then
              GoTo AnalyseSelection_End:
            End If
          End If
          
          .Update
          .MoveNext
          
          Call PrgStep
        Loop
      End With
      AnalyseSelection = True
      Me.Hide
  End If

AnalyseSelection_End:
  Call PrgStopCaption
  Call ClearCursor
  Call xReturn("AnalyseSelection")
  Exit Function

AnalyseSelection_Err:
  Call ClearEdit(m_rsEmployeesToSplit)
  Call ErrorMessage(ERR_ERROR, Err, "AnalyseSelection", "Analyse Selection", "Error analysing the name parts selected.")
  Resume AnalyseSelection_End
  Resume
End Function

Private Function GetWord(sDst As String, sSrc As String, bLookFromRight As Boolean) As Boolean
  Dim l As Long
  Dim lLen As Long
  
  On Error GoTo GetWord_Err
  Call xSet("GetWord")

  lLen = Len(sSrc)
  
  If bLookFromRight Then
    l = InStrBack(sSrc, " ", 1)
    If l > 1 Then
      sDst = Trim$(Right$(sSrc, lLen - l))
      sSrc = Trim$(Left$(sSrc, l - 1))
    Else
      sDst = sSrc
      sSrc = ""
    End If
  Else
    l = InStr(1, sSrc, " ")
    If l > 1 Then
      sDst = Trim$(Left$(sSrc, l - 1))
      sSrc = Trim$(Right$(sSrc, lLen - l))
    Else
      sDst = sSrc
      sSrc = ""
    End If
    
  End If
  

GetWord_End:
  Call xReturn("GetWord")
  Exit Function

GetWord_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetWord", "Get Word", "Error getting a name part.")
  Resume GetWord_End
End Function


Private Sub cmdRemove_Click()
  Call PassNodes(tvwNameParts, tvwChoices)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_rsEmployeesToSplit = Nothing
End Sub

Private Function PassNodes(tvwDst As TreeView, tvwSrc As TreeView) As Boolean
  Dim nNew As Node
  Dim nSelected As Node

  
  On Error GoTo PassNodes_Err
  Call xSet("PassNodes")
  
  If Not tvwSrc.SelectedItem Is Nothing Then
    Set nSelected = tvwSrc.SelectedItem
    Set nNew = tvwDst.Nodes.Add(, , nSelected.Key, nSelected.Text)
    nNew.Tag = nSelected.Tag
    tvwSrc.Nodes.Remove (nSelected.Index)
  End If
  
PassNodes_End:
  Set nNew = Nothing
  Set nSelected = Nothing
  Call xReturn("PassNodes")
  Exit Function

PassNodes_Err:
  If Err.Number <> 35600 And Err.Number <> 35602 And Err.Number <> 35601 Then
    Call ErrorMessage(ERR_ERROR, Err, "PassNodes", "Pass Nodes", "Error moving a node from tree view to tree view.")
    Resume PassNodes_End
  Else
    Resume Next
  End If
  Resume
End Function

Private Sub lvChoices_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvChoices_DragDrop(Source As Control, x As Single, y As Single)
      
End Sub

Private Sub TreeView2_BeforeLabelEdit(Cancel As Integer)

End Sub





