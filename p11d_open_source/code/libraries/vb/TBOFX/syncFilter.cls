VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "syncFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FILTER_TYPE
  INCLUDE_FILTER
  EXCLUDE_FILTER
End Enum

Private msFilterString As String
Private mFilterType As FILTER_TYPE
Private mAction As Filter_Action
Private mExceptions() As String
Private mbExceptions As Boolean

Public Function SetupFilter(sFilterString As String, lFilterType As FILTER_TYPE, lFilterAction As Filter_Action, Optional sExceptions As String = "") As Boolean

  On Error GoTo SetupFilter_Err
  
  If Len(sFilterString) = 0 Then Err.Raise ERR_NO_FILTER_STRING, "SetupFilter", "Filter string cannot be a zero length string."
  msFilterString = sFilterString
  mFilterType = lFilterType
  mAction = lFilterAction
  If Len(sExceptions) Then
    mExceptions = Split(sExceptions, ";")
    mbExceptions = True
  End If
  SetupFilter = True
  
SetupFilter_End:
  Exit Function

SetupFilter_Err:
  SetupFilter = False
  Call ErrorMessage(ERR_ERROR, Err, "SetupFilter", "Setup Filter", "Error setting database synchronisation filter.")
  Resume SetupFilter_End
End Function

Public Function FilterApplies(sCheckValue As String) As Boolean
  Dim i As Long
  
  If mbExceptions Then
    For i = LBound(mExceptions) To UBound(mExceptions)
      If StrComp(mExceptions(i), sCheckValue, vbTextCompare) = 0 Then
        FilterApplies = False
        Exit Function
      End If
    Next
  End If
  FilterApplies = sCheckValue Like msFilterString
  If mFilterType = EXCLUDE_FILTER Then
    FilterApplies = Not FilterApplies
  End If
End Function

Public Property Get StructureFilter() As Boolean
  StructureFilter = (mAction And STRUCTURE_FILTER) > 0
End Property
Public Property Get DataFilter() As Boolean
  DataFilter = (mAction And DATA_FILTER) > 0
End Property
Public Property Get FilterString() As String
  FilterString = msFilterString
End Property
Public Property Get FilterType() As FILTER_TYPE
  FilterType = mFilterType
End Property
Public Property Get FilterActions() As Filter_Action
  FilterActions = mAction
End Property

