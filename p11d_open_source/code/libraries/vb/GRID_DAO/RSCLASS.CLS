VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rsClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function AddGrid(frm As Object, rs As Object, sDb As String, rsCodes As Object) As String
  Dim grd As AutoGrid
    
  On Error GoTo AddGrid_Err
  Call xSet("AddGrid")
  Set grd = New AutoGrid
  frm.dtaAutogrid.DatabaseName = sDb
  Call grd.InitGrid(rs.Clone, frm.dtaAutogrid, frm.dbgAutoGrid, rsCodes)
  grd.dtaAutogrid.Caption = CStr(frm.dbgAutoGrid.Row + 1) & " of " & CStr(frm.dtaAutogrid.Recordset.RecordCount)
  AutoGrids.Add grd, grd.rsAutogrid.Name
  AddGrid = grd.key
  Set grd = Nothing
  Call xReturn("AddGrid")
  Exit Function
  
AddGrid_Err:
  Set grd = Nothing
  Call xReturn("AddGrid")
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
  
End Function

Public Sub RemoveGrid(sKey)
  Dim grd As AutoGrid
  
  On Error GoTo RemoveGrid_err
  Set grd = AutoGrids(sKey)
  grd.SaveCol
  AutoGrids.Remove sKey
  Set grd = Nothing
  
  Exit Sub
  
RemoveGrid_err:
  Resume Next ' zzzz gs change can't get excited about the clean up error so ignoring it
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

Public Sub ClearGrids()
  Dim grd As AutoGrid
  Call xSet("ClearGrids")
  
  On Error Resume Next
  
  For Each grd In AutoGrids
    grd.SaveCol
    AutoGrids.Remove grd.key
  Next grd
  Call xReturn("ClearGrids")
  Set grd = Nothing
End Sub

Public Function GridCount(Optional sKeys As Variant) As Long
  
  On Error GoTo GridCount_err
  If IsArray(sKeys) Then
    For i = 1 To AutoGrids.Count
      sKeys(i) = AutoGrids.Item(i)
    Next i
  End If
  GridCount = AutoGrids.Count
  Exit Function
  
GridCount_err:
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Public Function Grid(key As Variant) As Object
  
  On Error GoTo Grid_err
  
  Set Grid = AutoGrids(key)
  Exit Function
Grid_err:
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Public Function AddDialog(frm As Object, rs As Object, sDb As String, rsCodes As Object) As String
  Dim dlg As New AutoDialog
  
  On Error GoTo AddDialog_Err
  Call xSet("AddDialog")
  frm.dtadialog.DatabaseName = sDb
  Set frm.dtadialog.Recordset = rs
  Call dlg.InitDialog(frm, rs, rsCodes)
  AutoDialogs.Add dlg, dlg.key
  AddDialog = dlg.key
  Call xReturn("AddDialog")
  Exit Function

AddDialog_Err:
  Set dlg = Nothing
  Call xReturn("AddDialog")
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Public Sub RemoveDialog(sKey As String)
  On Error Resume Next
  AutoDialogs.Remove sKey
End Sub

Public Sub ClearDialogs()
  Dim dlg As AutoDialog
  
  On Error Resume Next
  Call xSet("ClearDialogs")
  
  For Each dlg In AutoDialogs
    AutoDialogs.Remove dlg.key
  Next dlg
  
  Set dlg = Nothing
  Call xReturn("ClearDialogs")
End Sub

Public Function DialogCount(Optional sKeys As Variant) As Long
  
  On Error GoTo DialogCount_err
  If IsArray(sKeys) Then
    For i = 1 To AutoDialogs.Count
      sKeys(i) = AutoDialogs.Item(i)
    Next i
  End If
  DialogCount = AutoDialogs.Count
  Exit Function
  
DialogCount_err:
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Public Function Dialog(key As Variant) As AutoDialog
  
  On Error GoTo Dialog_err
  
  Set Dialog = AutoDialogs(key)
  Exit Function
Dialog_err:
  Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function


Public Function SyncGrid(Dialog As Object, Grid As Object) As Boolean
  Dim dlg As New AutoDialog
  Dim grd As AutoGrid
  
  On Error GoTo SyncGrid_Err
  Call xSet("SyncGrid")
  
  SyncGrid = True
  Set dlg = Dialog
  Set grd = Grid
  Call grd.SetPos(dlg.GetPos)
  grd.Refresh
  
SyncGrid_End:
  Call xReturn("SyncGrid")
  Exit Function

SyncGrid_Err:
  SyncGrid = False
  Resume SyncGrid_End
End Function

Public Function SyncDialog(Grid As Object, Dialog As Object) As Boolean
  Dim dlg As New AutoDialog
  Dim grd As New AutoGrid
  
  On Error GoTo SyncDialog_Err
  Call xSet("SyncDialog")
  
  SyncDialog = True
  Set dlg = Dialog
  Set grd = Grid
  Call dlg.SetPos(grd.GetPos)
  dlg.Refresh
  
SyncDialog_End:
  Set grd = Nothing
  Set dlg = Nothing
  Call xReturn("SyncDialog")
  Exit Function

SyncDialog_Err:
  SyncDialog = False
  Resume SyncDialog_End
End Function

