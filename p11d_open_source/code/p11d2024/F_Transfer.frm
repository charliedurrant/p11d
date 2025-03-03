VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Begin VB.Form F_Transfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Employees"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin TCSPROG.TCSProgressBar prg 
      Height          =   330
      Left            =   135
      TabIndex        =   4
      Top             =   4005
      Width           =   5460
      _cx             =   4203935
      _cy             =   4194886
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
      Value           =   50
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
      Percentage      =   0
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5670
      TabIndex        =   2
      Top             =   810
      Width           =   1095
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "&Transfer"
      Default         =   -1  'True
      Height          =   375
      Left            =   5670
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3570
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   6297
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fix level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Select the employer to transfer to...."
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   2715
   End
End
Attribute VB_Name = "F_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LastListItem As ListItem

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Function Validate() As Boolean
  Dim j As Long, i As Long
  Dim li As ListItem
  Dim ben As IBenefitClass
  
  On Error GoTo Validate_err
  
  Call xSet("Validate")
  
  For i = 1 To lv.listitems.Count
    Set li = lv.listitems(i)
    If li.SmallIcon = IMG_SELECTED Then
      j = i
      Exit For
    End If
  Next
    
  If j > 0 Then
    Set ben = p11d32.Employers(lv.listitems(j).Tag)
    If ben.value(employer_FixLevel_db) <> p11d32.TemplateFixlevel Then Call Err.Raise(ERR_FIX_LEVEL_LOW, "Validate", "The fixel level is below " & p11d32.TemplateFixlevel & " please load the file to update the fix level.")
    Set ben = p11d32.CurrentEmployer
    If ben.value(employer_NoOfSelectedEmployees) = 0 Then Call Err.Raise(ERR_NO_EMPLOYEES_SELECTED, "Validate", "The current employer has no employees selected.")
    Set lv.SelectedItem = lv.listitems(j)
    Validate = True
  End If
  
Validate_end:
  Call xReturn("Validate")
  Exit Function
Validate_err:
  Call ErrorMessage(ERR_ERROR, Err, "Validate", "Validate", "Error trying to validate the employer selected.")
  Resume Validate_end
  Resume
End Function
Private Function Transfer() As Long
  Dim ben As IBenefitClass, benDstEmployer As IBenefitClass, benSrcEmployer As IBenefitClass
  Dim i As Long, j As Long
  Dim sqlt As SQLQUERIES_TRANSFER
  Dim dstEmployer As Employer
  Dim bOtherDone As Boolean
  Dim ibf As IBenefitForm2
  Dim li As ListItem
  Dim employees As ObjectList
  
  On Error GoTo Transfer_ERR
  Call xSet("Transfer")
  
  Set sqlt = New SQLQUERIES_TRANSFER
  
  'open the db and perform the queries per employee
  Set ibf = F_Employees
  Set dstEmployer = p11d32.Employers(lv.SelectedItem.Tag)
  Set benDstEmployer = dstEmployer
  If Not dstEmployer.ValidateEx(benDstEmployer.value(employer_PathAndFile), False, True, True) Then Call Err.Raise(ERR_EMPLOYER_DB, "Transfer", "The employer " & benDstEmployer.value(employer_Name_db) & " can not be opened exclusively.")
  Set benSrcEmployer = p11d32.CurrentEmployer
  prg.Max = benSrcEmployer.value(employer_NoOfSelectedEmployees)
  prg.Indicator = ValueOfMax
  Call SetCursor
  Set employees = p11d32.CurrentEmployer.employees
  For i = 1 To employees.Count
    Set ben = employees(i)
    If Not ben Is Nothing Then
      If ben.value(ee_Selected) Then
        'transfer employee
        prg.StepCaption ("Transferring employee " & ben.value(ee_PersonnelNumber_db))
        Call p11d32.CurrentEmployer.db.Execute(sqlt.Queries(GENERAL_TRANSFER, p11d32.Rates.BenClassTo(BC_EMPLOYEE, BCT_TABLE), benDstEmployer.value(employer_PathAndFile), ben.value(ee_PersonnelNumber_db)))
        'transfer address
        Call p11d32.CurrentEmployer.db.Execute(sqlt.Queries(GENERAL_TRANSFER, "T_Addresses", benDstEmployer.value(employer_PathAndFile), ben.value(ee_PersonnelNumber_db)))
        'Transfer Benefits
        For j = [BC_FIRST_ITEM] To BC_UDM_BENEFITS_LAST_ITEM
          Select Case j
            Case BC_LOAN_OTHER_H
              Call TransferSpecialBen(sqlt, p11d32.CurrentEmployer, dstEmployer, ben, OTHER_LOAN_TRANSFER_DATA, OTHER_LOAN_TRANSFER_KEYS)
            Case BC_QUALIFYING_RELOCATION_J
              Call TransferSpecialBen(sqlt, p11d32.CurrentEmployer, dstEmployer, ben, RELOC_TRANSFER_DATA, RELOC_TRANSFER_KEYS)
            Case BC_EMPLOYEE_CAR_E
              Call TransferSpecialBen(sqlt, p11d32.CurrentEmployer, dstEmployer, ben, EECAR_TRANSFER_DATA, EECAR_TRASNSFER_KEYS)
            Case BC_FUEL_F, BC_NON_QUALIFYING_RELOCATION_N, BC_SERVICES_PROVIDED_K
              'do nothing as done in car and relocation,assets at disposal,home phone
            Case Else
              If IsBenOtherClass(j) Then
                If Not bOtherDone Then
                  Call p11d32.CurrentEmployer.db.Execute(sqlt.Queries(OTHER_TRANSFER, benDstEmployer.value(employer_PathAndFile), ben.value(ee_PersonnelNumber_db)))
                  bOtherDone = True
                End If
              Else
                Call p11d32.CurrentEmployer.db.Execute(sqlt.Queries(GENERAL_TRANSFER, p11d32.Rates.BenClassTo(j, BCT_TABLE), benDstEmployer.value(employer_PathAndFile), ben.value(ee_PersonnelNumber_db)))
              End If
          End Select
        Next
        bOtherDone = False
        Set ibf.lv.SelectedItem = GetBenListItem(ibf, i)
        If ibf.lv.SelectedItem Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "Transfer", "The list item to remove is nothing.")
        ibf.RemoveBenefit (i)
        Set ben = Nothing
        Transfer = Transfer + 1
      End If
    End If
  Next
  
  If Not dstEmployer Is Nothing Then
    dstEmployer.db.Close
    Set dstEmployer.db = Nothing
  End If
  
  If Transfer > 0 Then
    'Call BenScreenSwitchEnd(F_Employees)
    prg.Caption = "Finished transfers...."
  End If
  
Transfer_END:
  Call ClearCursor
  Call xReturn("Transfer")
  Exit Function
Transfer_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Transfer", "Transfer", "Error transferring employees.")
  Resume Transfer_END
  Resume
End Function
Private Sub TransferSpecialBen(sqlt As SQLQUERIES_TRANSFER, benSrc As IBenefitClass, benDst As IBenefitClass, benEmployee As IBenefitClass, lSqlTransferData As QUERY_NAMES_TRANSFER, lSQLTransferkeys As QUERY_NAMES_TRANSFER)
' All keys changed to GUIDS so no need to update
  Dim rs As Recordset
  Dim eySrc As Employer, eyDst As Employer
  Dim sql As String
  
  Set eySrc = benSrc
  Set eyDst = benDst
  
  On Error GoTo TransferSpecialBen_ERR
  
  Call xSet("TransferSpecialBen")
  
  
  sql = sqlt.Queries(lSqlTransferData, benDst.value(employer_PathAndFile), benEmployee.value(ee_PersonnelNumber_db))
  eySrc.db.Execute (sql)
  
  sql = sqlt.Queries(lSQLTransferkeys, benDst.value(employer_PathAndFile), benEmployee.value(ee_PersonnelNumber_db))
  eySrc.db.Execute (sql)
  
  
TransferSpecialBen_END:
  Call xReturn("TransferSpecialBen")
  Exit Sub
TransferSpecialBen_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "TransferSpecialBen", "Transfer Special Ben", "Error transferring a benefit.")
  
  Resume TransferSpecialBen_END
  Resume
End Sub
Private Sub cmdTransfer_Click()
  If Validate Then
    Call Transfer
  End If
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not LastListItem Is Nothing Then LastListItem.SmallIcon = IMG_UNSELECTED
  Item.SmallIcon = IMG_SELECTED
  Set LastListItem = Item
End Sub
