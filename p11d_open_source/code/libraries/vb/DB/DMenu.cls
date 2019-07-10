VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DebugMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public m_DB As Database
Public m_DBName As String
Public m_DBClass As DBClass
Implements IDebugMenu

Private Function IDebugMenu_DatabaseMenuItem(ByVal MenuName As String) As Boolean
  Dim FileName As String
  
  On Error GoTo DatabaseMenuItem_err
  If StrComp(MenuName, "REPAIR_COMPACT") = 0 Then
    If (m_DB Is Nothing) And (Len(m_DBName) > 0) Then
      FileName = m_DBName
    Else
      FileName = FileOpenDlg("Choose Database to Repair and Compact", "Access Databases|*.MDB|Abacus Database Files|*.ABD|All Files|*.*||", CurDir$)
    End If
    If Len(FileName) > 0 Then
      Call m_DBClass.RepairCompactDB(FileName, MODE_COMPACT + MODE_REPAIR)
    End If
    IDebugMenu_DatabaseMenuItem = True
  ElseIf StrComp(MenuName, "SQL") = 0 Then
    If Not m_SQLExplorer Is Nothing Then
      If m_SQLExplorer.OpenExclusive And Not (m_DB Is Nothing) Then Err.Raise ERR_EXPLORESQL, "DatabaseMenuItem", m_SQLExplorer.OpenExclusiveMessage
      If Not m_DB Is Nothing Then
        Call m_SQLExplorer.Show(m_DB)
      ElseIf Len(m_DBName) > 0 Then
        Call m_SQLExplorer.Show(m_DBName)
      Else
        Call m_SQLExplorer.Show
      End If
      IDebugMenu_DatabaseMenuItem = True
    End If
  Else
    IDebugMenu_DatabaseMenuItem = False
  End If
  
DatabaseMenuItem_end:
  Exit Function
  
DatabaseMenuItem_err:
  IDebugMenu_DatabaseMenuItem = False
  Call ErrorMessage(ERR_ERROR, Err, "DatabaseMenuItem", "Error executing menu " & MenuName, "Unable to execute menuaction.")
  Resume DatabaseMenuItem_end
  Resume
End Function

Private Function IDebugMenu_RepairCompactAvailable() As Boolean
  IDebugMenu_RepairCompactAvailable = True
End Function

Private Function IDebugMenu_SQLAvailable() As Boolean
  IDebugMenu_SQLAvailable = (Not m_SQLExplorer Is Nothing)
End Function
