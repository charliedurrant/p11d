VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Name As String
Public Size As Long
Public ForeColor As Long
Public Bold As Boolean
Public Italic As Boolean
Public Strikethru As Boolean
Public Underline As Boolean
Public Parent As ReportField

Public Function FontString() As String
  Dim FS As String, Opt As String
  FS = "{FONT=" & Chr$(34) & Name & Chr$(34) & ","
  FS = FS & CStr(Size) & ","
  Opt = ""
  If Bold Then Opt = Opt & "B"
  If Underline Then Opt = Opt & "U"
  If Italic Then Opt = Opt & "I"
  If Strikethru Then Opt = Opt & "S"
  If Not Parent Is Nothing Then
    Select Case Parent.Alignment
      Case ALIGN_LEFT
        Opt = Opt & "L"
      Case TCSBASE.ALIGN_CENTER
        Opt = Opt & "C"
      Case ALIGN_RIGHT
        Opt = Opt & "R"
    End Select
  End If
  If Opt = "" Then Opt = "N"
  FS = FS & Opt & "}"
  FontString = FS
End Function

Public Function FontStringRestricted(FontName As String, FontSize As Long) As String
  Dim FS As String, Opt As String
  FS = "{FONT=" & Chr$(34) & FontName & Chr$(34) & ","
  FS = FS & CStr(FontSize) & ","
  Opt = ""
  If Bold Then Opt = Opt & "B"
  If Underline Then Opt = Opt & "U"
  If Italic Then Opt = Opt & "I"
  If Strikethru Then Opt = Opt & "S"
  If Not Parent Is Nothing Then
    Select Case Parent.Alignment
      Case ALIGN_LEFT
        Opt = Opt & "L"
      Case TCSBASE.ALIGN_CENTER
        Opt = Opt & "C"
      Case ALIGN_RIGHT
        Opt = Opt & "R"
    End Select
  End If
  If Opt = "" Then Opt = "N"
  FS = FS & Opt & "}"
  FontStringRestricted = FS
End Function

Public Function SaveString(Token As String) As String
  Dim s As String
  s = "(Font="
  s = s & Token
  s = s & "," & Name
  s = s & "," & CStr(Size)
  s = s & "," & CStr(ForeColor)
  s = s & "," & CStr(Bold)
  s = s & "," & CStr(Italic)
  s = s & "," & CStr(Strikethru)
  s = s & "," & CStr(Underline)
  s = s & ")"
  SaveString = s
End Function

Public Sub LoadFontFromString(ByVal Param As String, ByVal Arg As String)
  Param = UCase$(Param)
  Select Case Param
    Case UCase$("Name")
      Name = Arg
    Case UCase$("Size")
      Size = CLng(Val(Arg))
    Case UCase$("ForeColor")
      ForeColor = CLng(Val(Arg))
    Case UCase$("Bold")
      Bold = CBool(Arg)
    Case UCase$("Italic")
      Italic = CBool(Arg)
    Case UCase$("Strikethru")
      Strikethru = CBool(Arg)
    Case UCase$("Underline")
      Underline = CBool(Arg)
    Case Else
  End Select
End Sub

Public Function CopyTo(ToFontDetails As FontDetails) As Boolean
  With ToFontDetails
    .Name = Name
    .Size = Size
    .Bold = Bold
    .Italic = Italic
    .Underline = Underline
    .Strikethru = Strikethru
    .ForeColor = ForeColor
    Set .Parent = Parent
  End With
End Function

Private Sub Class_Initialize()
  Call Reinitialise
End Sub

Public Sub Reinitialise()
  Bold = False
  Italic = False
  Name = "Arial"
  Size = 12
  Strikethru = False
  Underline = False
End Sub

Public Function FontDescription() As String
  Dim s As String
  s = Name & ", Size " & Trim$(CStr(Size))
  If Bold Then s = s & ", Bold"
  If Italic Then s = s & ", Italics"
  If Underline Then s = s & ", Underline"
  If Strikethru Then s = s & ", Strikethrough"
  FontDescription = s
End Function

Public Function FontDescriptionRestricted(FontName As String, FontSize As Long) As String
  Dim s As String
  s = FontName & ", Size " & Trim$(CStr(FontSize))
  If Bold Then s = s & ", Bold"
  If Italic Then s = s & ", Italics"
  If Underline Then s = s & ", Underline"
  If Strikethru Then s = s & ", Strikethrough"
  FontDescriptionRestricted = s
End Function

