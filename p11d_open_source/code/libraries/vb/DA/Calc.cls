VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum REL_TYPE
  LIST_REL
  TREE_TOP_REL
  TREE_BID_REL
  TREE_KID_REL
  TREE_KIDBID_REL
End Enum
Public Property Get HasParents() As Boolean

End Property
Public Property Let HasParents(b As Boolean)

End Property
Public Property Get CircCount() As Integer

End Property
Public Property Let CircCount(newval As Integer)

End Property

Public Sub addparent(pname As String)
  ' adds to the class another it depends on
End Sub

Public Sub addchild(c As Calc, assoc As Boolean)
  ' adds to the class another that depends on it
End Sub

Public Sub removechild(c As Calc)
  ' removes a child
End Sub

Public Sub removeparent(c As Calc)
  'removoves a parent
End Sub

Public Sub addassocparent(pname As String)
  ' adds an associated (circular) parent
End Sub

Public Sub addassocparenttoCollection(pCalc As Calc)
  ' adds an associated (circular) parent both to the array and the collection
  ' used for iteratively added calcs
  
End Sub
Public Sub addchildtoCollection(pCalc As Calc)
  ' adds a child both to the array and the collection
  ' used for iteratively added calcs
  
End Sub
Public Sub addparenttoCollection(pCalc As Calc)
  ' adds an parent both to the array and the collection
  ' used for iteratively added calcs
  
End Sub
Public Sub LogRelations(obj As Object, i As REL_TYPE, Optional nn As Node = Nothing)
  ' logs all relations to log file
  
  
End Sub

Public Property Get name() As String
  ' returns object name
End Property

Public Sub Calc()
  ' asks an object to recalc if its calced flag is false
End Sub

Public Property Let calced(parent As Calc, val As Boolean)
  ' sets the recalc status
End Property
Public Property Get boocalced() As Boolean

  ' gets the recalc status
End Property

Public Sub touchparents()
  ' used by transients when edited
End Sub
