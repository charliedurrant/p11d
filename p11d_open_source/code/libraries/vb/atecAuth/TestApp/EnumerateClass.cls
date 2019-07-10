VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumerateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ILDAPNotify
Public mLDAPHelp As LDAPHelper2
Private m_usercount As Long

Private Function ILDAPNotify_LDAPNotify(ByVal ldapObject As ActiveDs.IADs) As Boolean
  Dim c As Collection
  Dim lp As LDAPProperty, email As String
  
  On Error GoTo notify_err
  m_usercount = m_usercount + 1
  Set c = mLDAPHelp.GetAllProperties2(ldapObject, "department", "adsPath")
  If InCollection(c, "department") Then
    Set lp = c.Item("department")
    email = lp.Values
  Else
    email = "(Unknown)"
  End If
  frmMain.lblCount.Caption = "User: " & m_usercount
  frmMain.lblCount.Refresh
  frmMain.txtResults.Text = frmMain.txtResults.Text & ldapObject.Name & " dept:" & email & vbCrLf
  frmMain.txtResults.Refresh
  ILDAPNotify_LDAPNotify = True
  
  Exit Function
notify_err:
  frmMain.txtResults.Text = frmMain.txtResults.Text & "ERROR: (" & Err.Number & ")" & Err.Description & " Source: " & Err.Source
  ILDAPNotify_LDAPNotify = False
End Function

