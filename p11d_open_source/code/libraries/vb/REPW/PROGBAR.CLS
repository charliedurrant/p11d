VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UDMNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'cad what is this thing, can't remove as it puclic
Option Explicit
Implements IBaseNotify

Public SPanel As TCSPANEL

Public Sub IBaseNotify_Notify(ByVal Current As Long, ByVal Max As Long, ByVal Message As String)
  If Not SPanel Is Nothing Then
    SPanel.Caption = CStr(Current) & "," & CStr(Max) & "," & Message
  End If
  Debug.Print Current, Max, Message
End Sub
