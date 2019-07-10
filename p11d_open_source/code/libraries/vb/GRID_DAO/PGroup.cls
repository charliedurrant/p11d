VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseGroup"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = -1
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "GROUP"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim fmt() As Byte, tlen As Long
  Dim aCol As AutoCol, p As Long, p0 As Long, i As Long
  
  If (paramcount <> 0) And (paramcount <> 1) Then Call Err.Raise(ERR_AUTOPARSE, "ParseGroup", "Expected either 0 or 1 parameters found " & CStr(paramcount))
  Set aCol = objCurItem
  If FormatRemove Then
    aCol.bGroup = False
    aCol.bGroupPageBreak = False
    aCol.bGroupHeader = False
    aCol.bGroupCount = False
    aCol.bGroupTotal = False
    aCol.SumLevel = 0
  Else
    aCol.bGroup = True
    If paramcount = 1 Then
      p = 1
      fmt = UCase$(params(1))
      tlen = LenB(params(1)) - 1
      For i = 0 To tlen Step 2
        Select Case fmt(i)
          Case vbKeyP
            aCol.bGroupPageBreak = True
          Case vbKeyH
            aCol.bGroupHeader = True
          Case vbKeyC
            aCol.bGroupCount = True
          Case vbKeyT
           aCol.bGroupTotal = True
           aCol.SumLevel = Val(Mid$(params(1), p + 1))
          Case Else
            If Not ((aCol.SumLevel > 0) And (InStr("0123456789", Chr$(fmt(i))) <> 0)) Then
              Err.Raise ERR_INVALIDGROUP, "ParseGroup", "Invalid Group format flags: " & UCase$(params(1))
            End If
        End Select
        p = p + 1
      Next i
    End If
  End If
  IParseItem_OnAction = PARSESTATUS_OK
End Function


