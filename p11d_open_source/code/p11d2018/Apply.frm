VERSION 5.00
Begin VB.Form F_ApplyCompanyDefined 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apply Company Defined Benefits"
   ClientHeight    =   6120
   ClientLeft      =   2010
   ClientTop       =   1380
   ClientWidth     =   8535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   8535
   Tag             =   "[CDB]"
   Begin VB.CommandButton cmdClearSelectionsSel 
      Caption         =   "Clear selections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7155
      TabIndex        =   11
      Top             =   4995
      Width           =   1320
   End
   Begin VB.CommandButton cmdClearSelectionsNotSel 
      Caption         =   "Clear selections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   10
      Top             =   5040
      Width           =   1320
   End
   Begin VB.CommandButton cmdUnselectAll 
      Caption         =   "<< Unse&lect All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3600
      TabIndex        =   4
      Top             =   2745
      Width           =   1350
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3600
      TabIndex        =   3
      Top             =   2295
      Width           =   1350
   End
   Begin VB.ListBox lstSel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   5040
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   585
      Width           =   3450
   End
   Begin VB.ListBox lstNotSel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   90
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   630
      Width           =   3450
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3600
      TabIndex        =   1
      Top             =   900
      Width           =   1350
   End
   Begin VB.CommandButton cmdUnselect 
      Caption         =   "< &Unselect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3600
      TabIndex        =   2
      Top             =   1350
      Width           =   1350
   End
   Begin VB.CommandButton B_Ok 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5400
      TabIndex        =   6
      Top             =   5670
      Width           =   1485
   End
   Begin VB.CommandButton B_Cancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7020
      TabIndex        =   7
      Top             =   5670
      Width           =   1485
   End
   Begin VB.Label Lab 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   252
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   840
      Width           =   3372
   End
   Begin VB.Label Lab 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To apply this CDB to an employee, select them on the table below and then 'tag' them by double-clicking or pressing the space bar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   585
      Index           =   1
      Left            =   45
      TabIndex        =   8
      Top             =   135
      Width           =   6915
   End
End
Attribute VB_Name = "F_ApplyCompanyDefined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public benefit As IBenefitClass

Private Type CDB_ASSINATION
  Assigned As Boolean
  WasAssigned As Boolean
  EmployeeIndex As Long
  BenefitIndex As Long
End Type

Private m_CDBAssinations() As CDB_ASSINATION
Private Sub B_Cancel_Click()
  Me.Hide
End Sub
Private Sub B_OK_Click()
  Call ApplyCompanyDefinedBenefits
End Sub
Private Sub cmdClearSelectionsNotSel_Click()
  Call ClearSelections(lstNotSel)
End Sub

Private Sub cmdClearSelectionsSel_Click()
  Call ClearSelections(lstSel)
End Sub
Private Sub ClearSelections(ByVal lst As ListBox)
  Dim i As Long
  Dim iCount As Long
  
  On Error GoTo err_Err
  
  iCount = lst.ListCount
  For i = 0 To iCount - 1
    lst.Selected(i) = False
  Next
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearSelections", "Clear Selections", Err.Description)
  Resume err_End:
  Resume
End Sub


Private Sub cmdSelect_Click()
  Call MoveItems(True, False)
End Sub

Private Sub cmdSelectAll_Click()
  Call MoveItems(True, True)
End Sub

Private Sub cmdUnselect_Click()
  Call MoveItems(False, False)
End Sub

Private Sub cmdUnselectAll_Click()
  Call MoveItems(False, True)
End Sub
'cad cdb
Public Function ViewAsignments(ByVal bDeleteBenefit As Boolean)
  Dim ee As Employee, oth As other
  Dim lCDBEmployeeBenefitIndex As Long
  Dim i As Long, lMax As Long
  Dim lEmployeeIndex As Long
  Dim rs As Recordset
  Dim sBenCode As String
  Dim ben As IBenefitClass
  
On Error GoTo ViewAsignments_ERR:
  
  Call xSet("ViewAsignments")
  
  Call SetCursor(vbHourglass)
  
  If benefit Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "ViewAsignments", "Benefit is nothing.")
  
  With F_Employees.LB
    lMax = .listitems.Count
    Call PrgStartCaption(lMax, "Analysing employees...")
    Set oth = benefit
    sBenCode = oth.PersonnelNumber
    ReDim m_CDBAssinations(1 To lMax)
    Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_CDB_LINKS)
    rs.Requery
    For i = 1 To lMax
      lEmployeeIndex = .listitems(i).Tag
      Set ee = p11d32.CurrentEmployer.employees(lEmployeeIndex)
      Set ben = ee
      If ee.BenefitsLoaded Then
        lCDBEmployeeBenefitIndex = ee.HasCDBBenefit(oth)
        
      Else
        rs.FindFirst ("P_Num='" & ee.PersonnelNumber & "' AND BenCode='" & sBenCode & "'")
        If Not rs.NoMatch Then lCDBEmployeeBenefitIndex = 1
      End If
      If lCDBEmployeeBenefitIndex > 0 Then
        m_CDBAssinations(i).Assigned = Not bDeleteBenefit
        m_CDBAssinations(i).WasAssigned = True
        m_CDBAssinations(i).EmployeeIndex = lEmployeeIndex
        m_CDBAssinations(i).BenefitIndex = lCDBEmployeeBenefitIndex
        'add the item to the added
        If Not bDeleteBenefit Then
          lstSel.AddItem (ben.value(ee_FullName) & " - " & ben.value(ee_PersonnelNumber_db))
          lstSel.ItemData(lstSel.NewIndex) = i
          lstSel.Selected(lstSel.NewIndex) = ben.value(ee_Selected)
        End If
        
        'set the item data to true to tell that it already had been linked
      Else
        'add the item to the not added
        If Not bDeleteBenefit Then
          lstNotSel.AddItem (ben.value(ee_FullName) & " - " & ben.value(ee_PersonnelNumber_db))
          lstNotSel.ItemData(lstNotSel.NewIndex) = i
          lstNotSel.Selected(lstNotSel.NewIndex) = ben.value(ee_Selected)
        End If
        m_CDBAssinations(i).Assigned = False
        m_CDBAssinations(i).WasAssigned = False
        m_CDBAssinations(i).EmployeeIndex = lEmployeeIndex
        
      End If
      
      lCDBEmployeeBenefitIndex = 0
      Call PrgStep
    Next
  End With
  
ViewAsignments_END:
  Call PrgStopCaption
  Call ClearCursor
  Set ee = Nothing
  Set rs = Nothing
  Call xReturn("ViewAsignments")
  Exit Function
ViewAsignments_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ViewAsignments", "View Asignments", "Error viewing the cdb assignments.")
  Resume ViewAsignments_END
  Resume
End Function

Private Function MoveItems(bSelected As Boolean, bMoveAll As Boolean) As Boolean
  Dim l As Long, lMax As Long, lEmployeeCount As Long
  
  On Error GoTo MoveItems_Err
  Call xSet("MoveItems")

  Call SetCursor(vbHourglass)
  Call LockWindowUpdate(Me.hwnd)
  
  If bSelected Then
    If lstNotSel.ListCount > 0 Then
      lMax = lstNotSel.ListCount
      Call PrgStartCaption(lMax, "Transferrring employees...")
      Do While l < lstNotSel.ListCount
        If lstNotSel.Selected(l) Or bMoveAll Then
          lstSel.AddItem (lstNotSel.List(l))
          lstSel.Selected(lstSel.NewIndex) = lstNotSel.Selected(l)
          lstSel.ItemData(lstSel.NewIndex) = lstNotSel.ItemData(l)
          m_CDBAssinations(lstNotSel.ItemData(l)).Assigned = True
          lstNotSel.RemoveItem (l)
          
        Else
          l = l + 1
        End If
        lEmployeeCount = lEmployeeCount + 1
        Call PrgStep
      Loop
    End If
  Else
    If lstSel.ListCount > 0 Then
      lMax = lstNotSel.ListCount
      Call PrgStartCaption(lMax, "Transferrring employees...")
      Do While l < lstSel.ListCount
        If lstSel.Selected(l) Or bMoveAll Then
          lstNotSel.AddItem (lstSel.List(l))
          lstNotSel.ItemData(lstNotSel.NewIndex) = lstSel.ItemData(l)
          m_CDBAssinations(lstSel.ItemData(l)).Assigned = False
          lstSel.RemoveItem (l)
        Else
          l = l + 1
        End If
        lEmployeeCount = lEmployeeCount + 1
        Call PrgStep
      Loop
    End If
  End If

MoveItems_End:
  Call LockWindowUpdate(0)
  Call PrgStopCaption
  ClearCursor
  Call xReturn("MoveItems")
  Exit Function

MoveItems_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MoveItems", "Move Items", "Error transferring employees for company defined benefits.")
  Resume MoveItems_End
  Resume
End Function

'cad cdb stuff
Public Function ApplyCompanyDefinedBenefits() As Boolean
  Dim l As Long, lMax As Long
  Dim ee As Employee
  Dim sBenCode As String
  Dim oth As other
  Dim benEE As IBenefitClass
  Dim rs As Recordset
  Dim iBenIndex As Long
  
  On Error GoTo ApplyCompanyDefinedBenefits_Err
  
  Call xSet("ApplyCompanyDefinedBenefits")
  
  If benefit Is Nothing Then Call Err.Raise(ERR_BEN_IS_NOTHING, "ViewAsignments", "Benefit is nothing.")
  'delete links in DB
  SetCursor (vbArrowHourglass)
  Set oth = benefit
  sBenCode = oth.PersonnelNumber
  Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_CDB_LINKS)
  lMax = UBound(m_CDBAssinations)
  Call PrgStartCaption(lMax, "Applying benefits...", , Percentage)
  
  For l = 1 To lMax
    If m_CDBAssinations(l).WasAssigned And Not m_CDBAssinations(l).Assigned Then
      'delete
      
      Set ee = p11d32.CurrentEmployer.employees(m_CDBAssinations(l).EmployeeIndex)
      Set benEE = ee
      If ee.BenefitsLoaded Then
      
        iBenIndex = ee.HasCDBBenefit(benefit)
        If iBenIndex > 0 Then
          Call ee.benefits.Remove(iBenIndex)
        End If
        benEE.NeedToCalculate = True
      End If
      Call RemoveCDBAssignment(rs, ee, sBenCode)
    ElseIf (Not m_CDBAssinations(l).WasAssigned) And m_CDBAssinations(l).Assigned Then
      'new assinations
      Set ee = p11d32.CurrentEmployer.employees(m_CDBAssinations(l).EmployeeIndex)
      Set benEE = ee
      If ee.BenefitsLoaded Then
        Call ee.CDBBenefitAdd(benefit)
        benEE.NeedToCalculate = True
      End If
      rs.AddNew
      rs.Fields("P_Num").value = ee.PersonnelNumber
      rs.Fields("BenCode").value = sBenCode
      rs.Update
    End If
    Call PrgStep
  Next
  
ApplyCompanyDefinedBenefits_End:
  Unload Me
  Call PrgStopCaption
  ClearCursor
  Call xReturn("ApplyCompanyDefinedBenefits")
  Exit Function

ApplyCompanyDefinedBenefits_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ApplyCompanyDefinedBenefits", "Apply Company Defined Benefits", "Error applying company defined benefits.")
  Resume ApplyCompanyDefinedBenefits_End
  Resume
End Function

Private Sub Form_Load()
  Call ViewAsignments(False)
End Sub

