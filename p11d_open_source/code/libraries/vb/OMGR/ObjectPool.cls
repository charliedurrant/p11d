VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectPoolManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mObjectList() As TPoolObject
Private mFreeList As Long
Private mAllocList As Long
Private mTemplateObject As IPoolObject

Private mObjName As String
Private mPoolSize As Long
Private mPoolIncrement As Long
Private mPoolForceFree As Boolean

Private Sub Class_Initialize()
  mPoolForceFree = False
End Sub

Private Sub Class_Terminate()
  Call RemovalAll
End Sub

Private Sub RemovalAll()
  Dim i As Long
  On Error Resume Next
  For i = mPoolSize To 1 Step -1
    Call mObjectList(i).obj.KillObject
    Set mObjectList(i).obj = Nothing
  Next i
  mPoolSize = 0
End Sub

Public Function CreatePool(ObjectToPool As IPoolObject, ByVal InitialPoolSize As Long, Optional ByVal PoolIncrement As Long = 256) As Boolean
  Dim i As Long
  
  On Error GoTo CreatePool_Err
  If InitialPoolSize < 1 Then Err.Raise ERR_CREATEPOOL, "CreatePool", "Unable to create an Object Pool with initial size of " & InitialPoolSize
  Set mTemplateObject = ObjectToPool
  mObjName = TypeName(ObjectToPool)
  mPoolSize = InitialPoolSize
  mPoolIncrement = PoolIncrement
  ReDim mObjectList(1 To mPoolSize) As TPoolObject
  mAllocList = 0: mFreeList = 1
  For i = 1 To mPoolSize
    Set mObjectList(i).obj = mTemplateObject.CreateNewObject
    mObjectList(i).next = i + 1
    mObjectList(i).rcPtr = GetRefCounterAddr(mObjectList(i).obj)
  Next i
  mObjectList(mPoolSize).next = 0
  CreatePool = True
  
CreatePool_End:
  Exit Function
  
CreatePool_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CreatePool", "Create Pool", "Unable to Create Pool")
  Call RemovalAll
  Resume CreatePool_End
End Function

Public Function NewObject() As IPoolObject
  'Dim tmpPool() As TPoolObject
  Dim NewObjIndex As Long, i As Long
  
  On Error GoTo NewObject_err
  If mFreeList = 0 Then
    If mPoolForceFree Then Call Me.Free
    If mFreeList = 0 Then
      mFreeList = mPoolSize + 1
      
'      ReDim tmpPool(1 To mPoolSize) As TPoolObject
'      For i = 1 To mPoolSize
'        Set tmpPool(i) = mObjectList(i)
'        Set mObjectList(i) = no
'      Next i
      mPoolSize = mPoolSize + mPoolIncrement
      ReDim Preserve mObjectList(1 To mPoolSize) As TPoolObject
      
      For i = mFreeList To mPoolSize
        Set mObjectList(i).obj = mTemplateObject.CreateNewObject
        mObjectList(i).next = i + 1
        mObjectList(i).rcPtr = GetRefCounterAddr(mObjectList(i).obj)
      Next i
      mObjectList(mPoolSize).next = 0
    End If
  End If
  
  ' remove from free list
  NewObjIndex = mFreeList
  mFreeList = mObjectList(mFreeList).next
  
  ' add to alloc free list
  mObjectList(NewObjIndex).next = mAllocList
  mAllocList = NewObjIndex
  
  Set NewObject = mObjectList(NewObjIndex).obj
NewObject_end:
  Exit Function
  
NewObject_err:
  Call ErrorMessage(ERR_ERROR, Err, "NewObject", "Allocate New Object", "Unable to allocate a new object of type " & mObjName)
  Set NewObject = Nothing
  Resume NewObject_end
End Function

Public Function Free() As Long
  Dim FreeCount As Long
  Dim refCount As Long
  Dim LastAllocIndex As Long, CurAllocIndex As Long, NextAllocIndex As Long
    
  On Error GoTo Free_err
  FreeCount = 0
  LastAllocIndex = 0
  CurAllocIndex = mAllocList
  Do While CurAllocIndex > 0
    refCount = -1
    Call CopyMemory(VarPtr(refCount), mObjectList(CurAllocIndex).rcPtr, 4)
    If refCount = 1 Then
      FreeCount = FreeCount + 1
      ' remove from alloc list
      NextAllocIndex = mObjectList(CurAllocIndex).next
      If LastAllocIndex = 0 Then
        mAllocList = NextAllocIndex
      Else
        mObjectList(LastAllocIndex).next = NextAllocIndex
      End If
      
      ' clear out and place on free list
      Call mObjectList(CurAllocIndex).obj.ClearObject
      mObjectList(CurAllocIndex).next = mFreeList
      mFreeList = CurAllocIndex
      
      CurAllocIndex = NextAllocIndex
    Else
      LastAllocIndex = CurAllocIndex
      CurAllocIndex = mObjectList(CurAllocIndex).next
    End If
  Loop
  Free = FreeCount
  
Free_end:
  Exit Function
  
Free_err:
  Call ErrorMessage(ERR_ERROR, Err, "Free", "Free Objects", "Unable to free object of type " & mObjName)
  Resume Free_end
End Function

Public Sub PoolInfo(Name As String, nFreeObjects As Long, nAllocObjects As Long)
  Dim nCount As Long, nNext As Long
  
  On Error GoTo PoolInfo_err
  Name = mObjName
  nCount = 0
  nNext = mFreeList
  Do While nNext <> 0
    nCount = nCount + 1
    nNext = mObjectList(nNext).next
  Loop
  nFreeObjects = nCount
  nAllocObjects = (mPoolSize - nCount)
  
PoolInfo_end:
  Exit Sub
  
PoolInfo_err:
  Call ErrorMessage(ERR_ERROR, Err, "PoolInfo", "Poll Information", "Unable to retrieve pool information for Object Pool: " & mObjName)
  Resume PoolInfo_end
End Sub

Public Property Get FreeOnNewObject() As Boolean
  FreeOnNewObject = mPoolForceFree
End Property

Public Property Let FreeOnNewObject(ByVal NewValue As Boolean)
  mPoolForceFree = NewValue
End Property
