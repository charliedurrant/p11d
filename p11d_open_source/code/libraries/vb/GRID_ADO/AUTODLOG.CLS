VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ParentAC As AutoClass
Private m_CurFrame As Frame
Private m_DialogForm As Form
Public key As String
Private WithEvents dtaDial As Data
Attribute dtaDial.VB_VarHelpID = -1

Private Const lxStart = 60
Private Const lyStart = 60
Private Const lySep = 180
Private Const lColWidth = 2600

Public Property Set DialogForm(NewValue As Object)
  Set m_DialogForm = NewValue
End Property

Public Property Set DataControl(NewValue As Object)
  Set dtaDial = NewValue
End Property

Private Sub dtaDial_Reposition()
  Call ResetDrop
End Sub

Private Sub dtaDial_Resize()
  Call Redraw
End Sub

Public Sub InitDialog()
  Dim ACol As AutoCol
  Dim ColSet As Column
  Dim l As Long, x As Long, y As Long
  Dim lHeight As Long, i As Integer, j As Integer
  
  On Error GoTo InitDialog_Err
  If dtaDial.Recordset.BOF And dtaDial.Recordset.EOF Then
    dtaDial.Recordset.AddNew
  End If
  If ParentAC.lTabs > 0 Then
    Call CreateTabs
    m_DialogForm.tabDialog.Visible = True
  Else
    m_DialogForm.tabDialog.Visible = False
  End If
  l = 1
  Set m_CurFrame = m_DialogForm.fmeDialog(l)
  x = lxStart
  y = lyStart
  For Each ACol In ParentAC
    If Len(ACol.TabTitle) > 0 Then
      Set m_CurFrame = m_DialogForm.fmeDialog(l)
      l = l + 1
      m_CurFrame.ZOrder
      x = lxStart
      y = lyStart
    End If
    If ACol.NewColumn Then
      x = x + (lColWidth * 2)
      y = lyStart
    End If
    If Not ACol.Hide Then
      lHeight = AddCtrl(ACol, x, y)
      y = y + lHeight + lySep
    End If
    j = j + 1
  Next ACol
  NoItems = 6
  
  Exit Sub
  
InitDialog_Err:
  Call Err.Raise(ERR_INITDIALOG, "InitDialog", "Unable to initialise AutoDialog" & vbCrLf & Err.Description)
End Sub

Private Sub CreateTabs()
  Dim ACol As AutoCol
  Dim i As Long
  
  On Error Resume Next
  Call xSet("CreateTabs")
  i = 1
  For Each ACol In ParentAC
    If ACol.TabTitle <> "" Then
      If m_DialogForm.tabDialog.Tabs.Count = 1 And i = 1 Then
        m_DialogForm.tabDialog.Tabs(1).key = ACol.TabTitle
        m_DialogForm.tabDialog.Tabs(1).Caption = ACol.TabTitle
        i = i + 1
      Else
        m_DialogForm.tabDialog.Tabs.Add i, ACol.TabTitle, ACol.TabTitle
        Load m_DialogForm.fmeDialog(i)
        i = i + 1
      End If
    End If
  Next ACol
  Call xReturn("CreateTabs")
End Sub

Private Function AddCtrl(ACol As AutoCol, ByVal x As Long, ByVal y As Long) As Long

  Call xSet("AddCtrl")
  If AddLabel(ACol, x, y) Then
    If Len(ACol.QueryName) > 0 Then
      AddCtrl = AddDrop(ACol, x + lColWidth, y)
      If AddCtrl = -1 Then
        AddCtrl = AddText(ACol, x + lColWidth, y)
      End If
    ElseIf ACol.dbDataType = TYPE_BOOL Then
      AddCtrl = AddCheck(ACol, x + lColWidth, y)
      If AddCtrl = -1 Then
        AddCtrl = AddText(ACol, x + lColWidth, y)
      End If
    Else
      AddCtrl = AddText(ACol, x + lColWidth, y)
    End If
    If Not dtaDial.Recordset.Updatable Then Me.EOFAction = vbMoveLast
  Else
    AddCtrl = -1
  End If
  Call xReturn("AddCtrl")
End Function

Private Function AddLabel(ACol As AutoCol, ByVal x As Long, ByVal y As Long) As Boolean
    
  On Error GoTo AddLabel_err
  With m_DialogForm
    Load .lblDialog(ACol.lID)
    .lblDialog(ACol.lID).Caption = ACol.DialogCaption
    .lblDialog(ACol.lID).Left = x
    .lblDialog(ACol.lID).Top = y
    Set .lblDialog(ACol.lID).Container = m_CurFrame
    .lblDialog(ACol.lID).Visible = True
  End With
  AddLabel = True
  
AddLabel_end:
  Exit Function
  
AddLabel_err:
  AddLabel = False
  Resume AddLabel_end
End Function

Private Function AddText(ACol As AutoCol, ByVal x As Long, ByVal y As Long) As Long
  Dim txt As TextBox
  
  On Error GoTo AddText_Err
  Select Case ACol.Alignment
    Case ALIGN_LEFT
      Load m_DialogForm.txtLdialog(ACol.lID)
      Set txt = m_DialogForm.txtLdialog(ACol.lID)
    Case ALIGN_CENTER
      Load m_DialogForm.txtCdialog(ACol.lID)
      Set txt = m_DialogForm.txtCdialog(ACol.lID)
    Case ALIGN_RIGHT
      Load m_DialogForm.txtRdialog(ACol.lID)
      Set txt = m_DialogForm.txtRdialog(ACol.lID)
  End Select
    
  txt.DataField = ACol.DataField
  If dtaDial.Recordset.Updatable Then
    txt.Enabled = Not ACol.NoEdit
  Else
    txt.Enabled = False
  End If
  txt.Left = x
  txt.Top = y
  Set txt.Container = m_CurFrame
  txt.Visible = True
  AddText = txt.height
  
AddText_End:
  Exit Function
  
AddText_Err:
  AddText = -1
  Resume AddText_End
End Function

Private Function AddDrop(ACol As AutoCol, ByVal x As Long, ByVal y As Long) As Long
  Dim db As Database
  Dim qry As QueryDef
  Dim rs As Recordset
  Dim str As String
  Dim i As Long
    
  On Error GoTo AddDrop_Err
  
  With m_DialogForm
    Load .cmbdialog(ACol.lID)
    Set db = gwsMain.Databases(ParentAC.DatabaseName)
    Set qry = db.QueryDefs(ACol.QueryName)
    For i = 1 To ACol.NoParams
      qry.Parameters(i - 1) = dtaDial.Recordset.Fields(ACol.GetParam(i)).Value
    Next i
    If ACol.NoParams > 0 Then .cmbdialog(ACol.lID).Tag = "TcsCombo"
    .cmbdialog(ACol.lID).DBName = db.Name
    Set .cmbdialog(ACol.lID).rsData = .dtadialog.Recordset
    .cmbdialog(ACol.lID).DataField = ACol.DataField
    Set .cmbdialog(ACol.lID).rsList = qry.OpenRecordset(dbOpenSnapshot, dbFailOnError)
    .cmbdialog(ACol.lID).BoundColumn = .cmbdialog(ACol.lID).rsList.Fields(0).Name
    .cmbdialog(ACol.lID).ListField = .cmbdialog(ACol.lID).rsList.Fields(.cmbdialog(ACol.lID).rsList.Fields.Count - 1).Name
    .cmbdialog(ACol.lID).UpdateControls
    .cmbdialog(ACol.lID).Left = x
    .cmbdialog(ACol.lID).Top = y
    .cmbdialog(ACol.lID).Style = ACol.QueryStyle
    .cmbdialog(ACol.lID).NoItems = ACol.Items
    Set .cmbdialog(ACol.lID).Container = m_CurFrame
    If Not .dtadialog.Recordset.Updatable Then
      .cmbdialog(ACol.lID).Enabled = False
    End If
    .cmbdialog(ACol.lID).Visible = True
    AddDrop = .cmbdialog(ACol.lID).height
  End With
  
AddDrop_End:
  Set rs = Nothing
  Set qry = Nothing
  Set db = Nothing
  Exit Function
  
AddDrop_Err:
  AddDrop = m_DialogForm.cmbdialog(0).height
  Resume AddDrop_End
End Function

Private Sub ShowDrop(pseCol As AutoCol, x As Double, y As Double, Width As Double, height As Double)
  Dim db As Database
  Dim qry As QueryDef
  Dim rs As Recordset
  Dim str As String
  Dim i As Long
  
  On Error GoTo ShowDrop_Err
  Call xSet("ShowDrop")
  
  Set db = gwsMain.Databases(ParentAC.DatabaseName)
  Set qry = db.QueryDefs(pseCol.QueryName)
  For i = 1 To pseCol.NoParams
    qry.Parameters(i - 1) = dtaDial.Recordset.Fields(pseCol.GetParam(i)).Value
  Next i
  m_DialogForm.dtaList.DatabaseName = db.Name
  m_DialogForm.dblGrid.DataField = pseCol.DataField
  Set m_DialogForm.dtaList.Recordset = qry.OpenRecordset(dbOpenSnapshot, dbFailOnError)
  m_DialogForm.dblGrid.ListField = m_DialogForm.dtaList.Recordset.Fields(m_DialogForm.dtaList.Recordset.Fields.Count - 1).Name
  m_DialogForm.dblGrid.BoundColumn = m_DialogForm.dtaList.Recordset.Fields(0).Name
  m_DialogForm.dblGrid.Left = x
  m_DialogForm.dblGrid.Top = y + height
  m_DialogForm.dblGrid.height = 0 * 6 'apf Grid.RowHeight
  m_DialogForm.dblGrid.Visible = True
  m_DialogForm.dblGrid.SetFocus
  
ShowDrop_End:
  Set rs = Nothing
  Set qry = Nothing
  Set db = Nothing
  Call xReturn("ShowDrop")
  Exit Sub

ShowDrop_Err:
  Call ErrorMessage(ERR_ERRORSILENT, Err, "ShowDrop", "Adding control", "Cannot add dropdown for " & pseCol.DataField)
  Resume ShowDrop_End
End Sub

Private Function AddCheck(ACol As AutoCol, ByVal x As Long, ByVal y As Long) As Long
    
  On Error GoTo AddCheck_err
  With m_DialogForm
    Load .chkdialog(ACol.lID)
    .chkdialog(ACol.lID).DataField = ACol.DataField
    .chkdialog(ACol.lID).Left = x
    .chkdialog(ACol.lID).Top = y
    Set .chkdialog(ACol.lID).Container = m_CurFrame
    If Not .dtadialog.Recordset.Updatable Then
      .chkdialog(ACol.lID).Enabled = False
    Else
      .chkdialog(ACol.lID).Enabled = Not ACol.NoEdit
    End If
    .chkdialog(ACol.lID).Visible = True
    AddCheck = .chkdialog(ACol.lID).height
  End With
  
AddCheck_end:
  Exit Function
  
AddCheck_err:
  AddCheck = -1
  Resume AddCheck_end
End Function

Property Let Font(ByVal NewValue As String)
  Dim ctrl As Control
  
  On Error Resume Next
  For Each ctrl In m_DialogForm.Controls
    If ctrl.Tag <> "NonAuto" Then
      ctrl.Font.Name = NewValue
      ctrl.FontName = NewValue
    End If
  Next ctrl
End Property

Property Let FontSize(ByVal NewValue As Long)
  Dim ctrl As Control
  
  On Error Resume Next
  
  For Each ctrl In m_DialogForm.Controls
    If ctrl.Tag <> "NonAuto" Then
      ctrl.Font.Size = NewValue
      ctrl.FontSize = NewValue
    End If
  Next ctrl
  dtaDial.height = 285
End Property

Property Let BOFAction(ByVal NewValue As VBRUN.DataBOFconstants)
  dtaDial.BOFAction = NewValue
End Property

Property Let EOFAction(ByVal NewValue As VBRUN.DataEOFConstants)
  dtaDial.EOFAction = NewValue
End Property

Public Sub Refresh()
  Dim x As Control
  
  On Error GoTo DialogRefresh_Err
  For Each x In m_DialogForm.Controls
    If x.Index > 0 Then Unload x
  Next x
  Call Me.InitDialog
  
DialogRefresh_End:
  Exit Sub
  
DialogRefresh_Err:
  '343 No Array, 361 Can't unload data controls, 362 Can't unload frames
  Select Case Err.Number
    Case 343, 361, 362
          Resume Next
  End Select
  Call ErrorMessage(ERR_ERROR, Err, "Refresh", "Refresh dialog", "Failed to Refresh Dialog")
  Resume DialogRefresh_End
End Sub

Private Sub ResetDrop()
  Dim ctrl As Control
  Dim ACol As AutoCol
  Dim qry As QueryDef
  Dim db As Database
  Dim i As Long
  
  On Error GoTo ResetDrop_Err
  Call xSet("ResetDrop")
  If Not (dtaDial.Recordset.BOF And dtaDial.Recordset.BOF) Then
    For Each ctrl In m_DialogForm.Controls
      If ctrl.Tag = "TcsCombo" Then
        If ctrl.Index > 0 Then
          Set ACol = ParentAC.Item(ctrl.DataField)
          Set db = gwsMain.Databases(ParentAC.DatabaseName)
          Set qry = db.QueryDefs(ACol.QueryName)
          For i = 1 To ACol.NoParams
            qry.Parameters(i - 1) = ctrl.rsData.Fields(ACol.GetParam(i)).Value
          Next i
          Set ctrl.rsList = qry.OpenRecordset(dbOpenSnapshot, dbFailOnError)
        End If
      End If
    Next ctrl
  End If

ResetDrop_End:
  Set ACol = Nothing
  Set qry = Nothing
  Set db = Nothing
  Call xReturn("ResetDrop")
  Exit Sub
  
ResetDrop_Err:
  Call ErrorMessage(ERR_ERRORSILENT, Err, "ResetDrop", "ERR_DIALOG", "Can't reset dropdowns")
  Resume ResetDrop_End
End Sub

Private Sub Redraw()
  Dim lbl As Label
  Dim ctrl As Control
  Dim i As Long, x As Long
  
  On Error GoTo Redraw_Err
  For i = 1 To m_DialogForm.Controls.Count - 1
    If TypeOf m_DialogForm.Controls(i) Is Label Then
      Set lbl = m_DialogForm.Controls(i)
      If lbl.Name = "lblDialog" Then
        If lbl.Index > 0 Then
          lbl.height = dtaDial.height
          If ObjectExists(m_DialogForm.txtRdialog(lbl.Index)) Then
            Set ctrl = m_DialogForm.txtRdialog(lbl.Index)
          ElseIf ObjectExists(m_DialogForm.txtCdialog(lbl.Index)) Then
            Set ctrl = m_DialogForm.txtCdialog(lbl.Index)
          ElseIf ObjectExists(m_DialogForm.txtLdialog(lbl.Index)) Then
            Set ctrl = m_DialogForm.txtLdialog(lbl.Index)
          ElseIf ObjectExists(m_DialogForm.cmbdialog(lbl.Index)) Then
            Set ctrl = m_DialogForm.cmbdialog(lbl.Index)
          ElseIf ObjectExists(m_DialogForm.chkdialog(lbl.Index)) Then
            Set ctrl = m_DialogForm.chkdialog(lbl.Index)
          End If
          If lbl.Top = lxStart Then
            x = lbl.Top
          Else
            lbl.Top = x
            ctrl.Top = x
          End If
          x = x + lbl.height + lySep
          ctrl.height = lbl.height
        End If
      Else
        lbl.Top = m_DialogForm.dtadialog.Top
        lbl.height = m_DialogForm.dtadialog.height
      End If
    End If
  Next i

Redraw_End:
  Exit Sub
  
Redraw_Err:
  Call ErrorMessage(ERR_ERRORSILENT, Err, "Redraw", "ERR_DIALOG", "Can't redraw Dialog")
  Resume Redraw_End:
End Sub

Private Function ObjectExists(ctrl As Control) As Boolean
  Dim x As Long
  
  On Error GoTo ObjectExists_err
  x = ctrl.height
  ObjectExists = True
  
ObjectExists_end:
  Exit Function
  
ObjectExists_err:
  ObjectExists = False
  Resume ObjectExists_end
End Function

Public Property Let NoItems(ByVal iNew_Value As Integer)
  Dim ctrl As Control
  
  For Each ctrl In m_DialogForm.Controls
    If ctrl.Name = "cmbDialog" Then
      ctrl.NoItems = iNew_Value
    End If
  Next ctrl
End Property

Public Sub FindInBox()
  Dim ctrl As Control
  
  On Error GoTo FindInBox_Err
  Set ctrl = Screen.ActiveForm.ActiveControl
  Call ParentAC.Find(ctrl.DataField, DAOtoDatatype(dtaDial.Recordset.Fields(ctrl.DataField).Type))
  Exit Sub
  
FindInBox_Err:
  Call ECASE("FindInBox")
End Sub
