Attribute VB_Name = "ADORecordSet"
Option Explicit

Public Function NewRS() As Recordset
  Dim rs As adodb.Recordset, flds As adodb.Fields
  
  Set rs = New adodb.Recordset
  Set flds = rs.Fields
  Call flds.Append("x", adWChar, 64)
  Call flds.Append("y", adWChar, 64)
  Call rs.Open(
  rs.AddNew
    rs.Fields("x").Value = "ooomn"
    rs.Fields("y").Value = "ooomnd"
  rs.Update
  
End Function
