Attribute VB_Name = "TreeViewCode"
Option Explicit


Public Function FillTreeView(TRVIEW As TreeView, RepDataSets As DataSetCollection, Optional ByVal Key As String) As Boolean
  Dim dSet As ReportDataSet
  Dim Fld As ReportField
  Dim nod As node, uDSetName As String
  
  For Each dSet In RepDataSets
    uDSetName = UCase$(dSet.Name)
    If dSet.Parent Is Nothing Then
      Set nod = TRVIEW.Nodes.Add(, , uDSetName, dSet.Name, NOTSELECTED_PARENT)
      nod.Bold = True
    Else
      Set nod = TRVIEW.Nodes.Add(Key, tvwChild, Key & DATASET_SEP & uDSetName, dSet.Name, NOTSELECTED_PARENT)
      nod.Bold = True
    End If
    For Each Fld In dSet.cFields
      If Len(Key) > 0 Then
        Set nod = TRVIEW.Nodes.Add(Key & DATASET_SEP & uDSetName, tvwChild, Fld.KeyString, Fld.Description, NOTSELECTED_NODE)
      Else
        Set nod = TRVIEW.Nodes.Add(uDSetName, tvwChild, Fld.KeyString, Fld.Description, NOTSELECTED_NODE)
      End If
    Next Fld
    If Len(Key) > 0 Then
      Call FillTreeView(TRVIEW, dSet.Children, Key & DATASET_SEP & uDSetName)
    Else
      Call FillTreeView(TRVIEW, dSet.Children, uDSetName)
    End If
  Next dSet
End Function

Public Function ModifyTreeView(FrmRepWiz As Frm_RepWiz, RepFields As Collection) As Boolean
  Dim rFld As ReportField, nod As node
  
  On Error GoTo ModifyTreeView_Err
  Call xSet("ModifyTreeView")

  For Each rFld In RepFields
    Set nod = FrmRepWiz.TrV_Fields.Nodes(rFld.KeyString)
    Call FrmRepWiz.TrV_Fields_NodeClick(nod)
    'Call FrmRepWiz.NodeClicked(nod, False, False, -1)
  Next rFld

ModifyTreeView_End:
  Call xReturn("ModifyTreeView")
  Exit Function

ModifyTreeView_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ModifyTreeView", "ERR_UNDEFINED", "Error filling Tree View after load.")
  Resume ModifyTreeView_End
End Function

