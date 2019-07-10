VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtrlAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 
Public sName As String
Public sngXFactor As Single
Public sngYFactor As Single
Public sngWidthFactor As Single
Public sngHeightFactor As Single
Public sKey As String
Public ResizeAttributes As Long
Public sngTopDistance As Single
Public sngBottomDistance As Single
Public sngRightDistance As Single
Public sngLeftDistance As Single
Public sngDesignWidth As Single
Public sngDesignHeight As Single
Public sngBuddyDesignDistance As Single
Public sngFontHeightFactor As Single
Public sngFontWidthFactor As Single
Public Resized As Boolean
Public sngCntrlOldWidth As Single
Public bContainer As Boolean
Public sBuddy$
Public lBuddy&
Public cContainer As Object
Public cMe As Control
Public sContainer As String
Public bCanMove As Boolean
