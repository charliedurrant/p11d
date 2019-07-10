VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParsedQueryRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_rsType As RS_TYPE
Private m_SnapShot As Boolean
Private m_rs As recordset

Public Property Get rsType() As RS_TYPE
  rsType = m_rsType
End Property

Public Property Get SnapShot() As Boolean
  SnapShot = m_SnapShot
End Property

Public Property Get rs() As recordset
  Set rs = m_rs
End Property

Public Sub Kill()
  Set m_rs = Nothing
End Sub

Public Sub Setup(ByVal db As database, ByVal rsType As RS_TYPE, ByVal SnapShot As Boolean)
  Dim dOpt As Long
  
  On Error GoTo Setup_ERR
  If db Is Nothing Then Exit Sub
  Call xSet("Setup")
  If SnapShot Then
    dOpt = dbOpenSnapshot
  Else
    dOpt = dbOpenDynaset
  End If
  m_rsType = rsType
  m_SnapShot = SnapShot
  
  Select Case rsType
    Case QUERY_HEADER_RS
      Set m_rs = db.OpenRecordset("select * from sys_Queries_Header", dOpt)
    ' redundant
'    Case QUERY_DETAIL_RS
'      Set m_rs = db.OpenRecordset("select * from sys_Queries_Detail", dOpt)
    Case QUERY_RELATIONS_RS
      Set m_rs = db.OpenRecordset("select * from sys_Queries_relations", dOpt)
'    Case QUERY_SYS_OBJ
'     Set m_rs = db.OpenRecordset("select name, dateupdate from Msysobjects", dOpt)
    Case Else
      ECASE "Invalid rs_type - " & CStr(rsType)
      m_rsType = 0
  End Select
  
Setup_END:
  Call xReturn("Setup")
  Exit Sub
Setup_ERR:
  Resume Setup_END
End Sub
