VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mValue As Variant
Private mDatatype As Long
Private mDTLength As Long

Public Property Get Value() As Variant
  Value = mValue
End Property

Public Property Get Datatype() As String
  Datatype = mDatatype
End Property

Public Property Let DTLength(ByVal lNewVal As Long)
  mDTLength = lNewVal
End Property

Public Property Get DTLength() As Long
  DTLength = mDTLength
End Property

Friend Sub SetDatatype(ByVal sXMLType As String)
  Select Case UCase(sXMLType)
    Case "INT"
      mDatatype = adInteger
    Case "STRING"
      mDatatype = adVarChar
    Case "DATE"
      mDatatype = adDate
    Case Else
      Err.Raise ERR_INVALID_SPEC, , "Unrecognised parameter data type [ " & sXMLType & " ]"
  End Select
End Sub
  

Friend Sub SetValue(ByVal sValue As String)
  
  sValue = TrimChar(Trim(sValue), """")
  Select Case mDatatype
    Case adInteger
      mValue = GetTypedValueFromString(sValue, TYPE_LONG)
    Case adVarChar
      mValue = GetTypedValueFromString(sValue, TYPE_STR)
    Case adDate
      mValue = GetTypedValueFromString(sValue, TYPE_DATE)
  End Select
End Sub
