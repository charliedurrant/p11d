VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunctionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String
Public ReturnValue As String
Public Description As String
Public IsFunction As Boolean
Public PropertyType As PROPERTY_TYPE
Public AddedToCollection As Boolean
Private mDescriptionLong As String
Public ReturnValueDescription As String
Public Params As Collection
Public Parent As Class
Public Key As String
Public Selected As Boolean
Public FileFunctionStartPos As Long

Private mCategory As String
Private m_SearchText As String
Private mRTHelp As RTFTextHelp

Public Function ParametersString(Optional ByVal PropType As PROPERTY_TYPE = PROPERTY_NONE) As String
  Dim s As String
  Dim param As Parameter
  Dim i As Long, MaxParam As Long
  
  MaxParam = Params.Count
  If PropType = PROPERTY_GET Then MaxParam = MaxParam - 1
  For i = 1 To MaxParam
    Set param = Params.Item(i)
    s = s & param.Text
    If i < MaxParam Then s = s & ", "
  Next
  ParametersString = s
End Function
Public Property Let Category(ByVal NewValue As String)
  Dim i As Long
  
  If Not Parent.GlobalNameSpace Then Err.Raise 380, "Category", "Cannot have categories for functions in classes which are not global." & vbCrLf & "Class: " & Parent.Name & " Function: " & Me.Name
  i = FindString(CategoryMaps, NewValue)
  If i <> -1 Then
    Category = CategoryValues(i)
  Else
    mCategory = NewValue
  End If
End Property

Public Property Get Category() As String
  If Len(mCategory) = 0 Then
    Category = "Uncategorised"
  Else
    Category = mCategory
  End If
End Property

Public Sub kill()
  Set Parent = Nothing
End Sub

Private Sub Class_Initialize()
  Set Params = New Collection
  Set mRTHelp = New RTFTextHelp
End Sub
 
Public Sub InitialScreenOutput(rt As RichTextBox)
  Dim s As String
  Dim sParam As String
  Dim param As Parameter, pcount As Long
  Dim i As Long
  
  Call mRTHelp.BeginRTFMarkups(rt)
  Call mRTHelp.AddRTFText(Name & vbCrLf, &H800000, True)
  If IsFunction Or (PropertyType = PROPERTY_NONE) Then
    If IsFunction Then
      s = "Function "
    Else
      s = "Sub "
    End If
    s = s & Name & "("
    s = s & ParametersString & ")" & IIf(IsFunction, " as " & ReturnValue, "") & vbCrLf
  Else
    If (PropertyType And PROPERTY_GET) = PROPERTY_GET Then
      s = s & "Property Get " & Name & "(" & ParametersString(PROPERTY_GET) & ")" & " as " & ReturnValue & vbCrLf
    End If
    If (PropertyType And PROPERTY_LET) = PROPERTY_LET Then
      s = s & "Property Let " & Name & "(" & ParametersString(PROPERTY_LET) & ")" & vbCrLf
    End If
    If (PropertyType And PROPERTY_SET) = PROPERTY_SET Then
      s = s & "Property Set " & Name & "(" & ParametersString(PROPERTY_SET) & ")" & vbCrLf
    End If
  End If
  Call mRTHelp.AddRTFText(s & vbCrLf, vbBlack, False)
  
  Call mRTHelp.AddRTFText(Description & vbCrLf, &H800000, False)
  Call mRTHelp.AddRTFText(DescriptionLong & vbCrLf, vbBlue, False)
  
  s = "": pcount = 0
  For Each param In Params
    If Len(param.Description) > 0 Then
      s = s & "  " & param.Name & " " & param.Description & vbCrLf
      pcount = pcount + 1
    End If
  Next param
  If pcount > 0 Then
    If pcount = 1 Then
      s = "Variable:" & vbCrLf & s
    Else
      s = "Variables:" & vbCrLf & s
    End If
    Call mRTHelp.AddRTFText(s, vbRed, False)
  End If
  
  If IsFunction Then
    If Len(Me.ReturnValueDescription) Then
      Call mRTHelp.AddRTFText("Return value:" & vbCrLf & "  " & Me.ReturnValueDescription & vbCrLf, vbRed, False)
    End If
  End If
  
End Sub

Public Sub AddScreenOutput(ByVal ScreenText As String, ByVal Colour As ColorConstants, ByVal Bold As Boolean)
  Call mRTHelp.AddRTFText(ScreenText & vbCrLf, Colour, Bold)
End Sub

Public Sub OutputToScreen(rt As RichTextBox)
  Call mRTHelp.ApplyMarkups(rt)
  rt.SelStart = 1
End Sub


Public Property Get SearchText() As String
  Dim sParam As String
  Dim param As Parameter
  Dim i As Long
  
  If Len(m_SearchText) > 0 Then
    SearchText = m_SearchText
    Exit Property
  End If
  
  m_SearchText = m_SearchText & Name & " "
  m_SearchText = m_SearchText & Description & " "
  For Each param In Params
    If Len(param.Description) > 0 Then m_SearchText = m_SearchText & param.Description & " "
  Next
  m_SearchText = m_SearchText & Name & " "
  m_SearchText = m_SearchText & ParametersString & IIf(IsFunction, ReturnValue, "") & " "
  SearchText = m_SearchText
End Property


Public Property Get Name() As String
  Name = mName
End Property

Public Property Let Name(ByVal NewValue As String)
  mName = Trim$(NewValue)
End Property

Public Property Get DescriptionLong() As String
  DescriptionLong = mDescriptionLong
End Property

Public Property Let DescriptionLong(ByVal NewValue As String)
  mDescriptionLong = Trim$(mDescriptionLong)
  If Len(mDescriptionLong) > 0 Then mDescriptionLong = mDescriptionLong & " "
  mDescriptionLong = mDescriptionLong & Trim$(NewValue)
End Property
