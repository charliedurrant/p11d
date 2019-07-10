VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTFTextHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RTF_MARKUP
  Start As Long
  Length As Long
  Colour As ColorConstants
  Bold As Boolean
End Type

Private mMarkups() As RTF_MARKUP
Private mMarkupCur As Long
Private mMarkupCount As Long

Private mRTOffset As Long
Private mRTString As String

Public Sub BeginRTFMarkups(rt As RichTextBox)
  mRTOffset = Len(rt.Text)
  mRTString = ""
  mMarkupCur = 0
End Sub


Public Sub AddRTFText(ByVal NewString As String, ByVal Colour As ColorConstants, ByVal Bold As Boolean)
  Dim p0 As Long, p1 As Long
  
  If Len(Trim$(ReplaceString(NewString, vbCrLf, ""))) = 0 Then Exit Sub
  mMarkupCur = mMarkupCur + 1
  If mMarkupCur > mMarkupCount Then
    mMarkupCount = mMarkupCount + 10
    ReDim Preserve mMarkups(1 To mMarkupCount)
  End If
  mMarkups(mMarkupCur).Start = Len(mRTString) + mRTOffset
  mMarkups(mMarkupCur).Length = Len(NewString)
  mMarkups(mMarkupCur).Colour = Colour
  mMarkups(mMarkupCur).Bold = Bold
  mRTString = mRTString & NewString
End Sub

Public Sub ApplyMarkups(rt As RichTextBox)
  Dim i As Long
  
  Call LockWindowUpdate(rt.hWnd)
  rt.Text = rt.Text & mRTString
  For i = 1 To mMarkupCur
    rt.SelStart = mMarkups(i).Start
    rt.SelLength = mMarkups(i).Length
    rt.SelColor = mMarkups(i).Colour
    rt.SelBold = mMarkups(i).Bold
    rt.SelLength = 0
  Next i
  Call LockWindowUpdate(0)
End Sub

