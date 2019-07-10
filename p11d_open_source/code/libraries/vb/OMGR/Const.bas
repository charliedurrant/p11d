Attribute VB_Name = "Const"
Option Explicit

Public Enum TCSOMGR_ERRORS
  ERR_CREATEPOOL = TCSOMGR_ERROR
End Enum

Public Type TPoolObject
  obj As IPoolObject
  rcPtr As Long
  next As Long  ' Index of the next object
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal Source As Long, ByVal cbCopy As Long)

Public Function GetRefCounterAddr(iUkn As IUnknown) As Long
  GetRefCounterAddr = ObjPtr(iUkn) + 4
End Function

