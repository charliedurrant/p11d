Attribute VB_Name = "FPCS"
Option Explicit

Public ErFPCSCounter As Long
Public ErFPCSGrid As New Collection

'Private Const BALSTEP As Long = 5
Private m_parent As IBenefitClass
Private m_readfromdb As Boolean
Private m_sbookmark As String
Private m_dirty  As Boolean
Private m_InvalidFields As Long
Private m_FPCSCounter As Long
Private lFPCSSize As Long
Private m_FPCSGrid() As clsFPCSBand


Public Function FPCSReadDb() As Long
Dim db As Database, q As New SQLQUERIES, rs As Recordset
Dim rsFPCS As Recordset
Dim m_FPCSItems As clsFPCSBand

Dim i As Long

  On Error GoTo clsErFPCS_FPCSReadDb_Err
  xSet "clsErFPCS_FPCSReadDb"
  If m_readfromdb Then GoTo clsErFPCS_FPCSReadDb_end
  Set rsFPCS = CurrentEmployer.db.OpenRecordset(q.Queries(SELECT_ER_FPCS), dbOpenForwardOnly, dbFailOnError)
  Do While Not rsFPCS.EOF
    Set m_FPCSItems = New clsFPCSBand
    m_FPCSItems.Scheme = "" & rsFPCS.Fields("FPCS")
    m_FPCSItems.BandName = "" & rsFPCS.Fields("CCBand")
    m_FPCSItems.AboveMileage = IIf(IsNull(rsFPCS.Fields("MilesAbove")), 0, rsFPCS.Fields("MilesAbove"))
    m_FPCSItems.AboveEngineSize = IIf(IsNull(rsFPCS.Fields("EngineAbove")), 0, rsFPCS.Fields("EngineAbove"))
    m_FPCSItems.rate = IIf(IsNull(rsFPCS.Fields("RateMiles")), 0, rsFPCS.Fields("RateMiles"))
    Call AddBand(m_FPCSItems)
    Set m_FPCSItems = Nothing
    rsFPCS.MoveNext
  Loop
  m_readfromdb = True
  FPCSReadDb = True
clsErFPCS_FPCSReadDb_end:
  Set rsFPCS = Nothing
  xReturn "clsErFPCS_FPCSReadDb"
  Exit Function
clsErFPCS_FPCSReadDb_Err:
  Resume clsErFPCS_FPCSReadDb_end
End Function


Public Property Get FPCSGrid(Index As Long)
  Set FPCSGrid = m_FPCSGrid(Index)
End Property

Public Function AddBand(FPCSItem As clsFPCSBand) As Long

  On Error GoTo AddBand_Err
  Call xSet("AddBand")
  ErFPCSCounter = ErFPCSCounter + 1
  
  'If ErFPCSCounter > lFPCSSize Then
  '  lFPCSSize = lFPCSSize + 5
  '  ReDim Preserve ErFPCSGrid(lFPCSSize)
  'End If
  ErFPCSGrid.Add Item:=FPCSItem

AddBand_End:
  Call xReturn("AddBand")
  Exit Function

AddBand_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBand", "ERR_UNDEFINED", "Undefined error.")
  Resume AddBand_End
End Function

Public Property Get FPCSCounter() As Long
  FPCSCounter = m_FPCSCounter
End Property

Public Property Let FPCSCounter(NewCount As Long)
  m_FPCSCounter = NewCount
End Property

Public Function WriteBands() As Boolean
  Dim db As Database
  Dim q As New SQLQUERIES
  Dim i As Long
  Dim rs As Recordset
  Dim stype As String
  On Error GoTo WriteBands_Err
  Call xSet("WriteBands")
  
  Set db = m_parent.Parent.Parent.db
  
  'Delete old entries
  'If Not (m_FPCSItems(FPCS_FPCSkey) = Empty) Then
  Call db.Execute(q.Queries(DELETE_ER_FPCS, , ""))
 
  Set rs = db.OpenRecordset(q.Queries(SELECT_ER_FPCS), dbOpenDynaset, dbFailOnError)
  If m_FPCSCounter > 0 Then
    For i = 0 To m_FPCSCounter - 1
      rs.AddNew
      rs.Fields("FPCS") = m_FPCSGrid(i).Scheme
      rs.Fields("CCBand") = m_FPCSGrid(i).BandName
      rs.Fields("EngineAbove") = m_FPCSGrid(i).AboveEngineSize
      rs.Fields("MilesAbove") = m_FPCSGrid(i).AboveMileage
      rs.Fields("RateMiles") = m_FPCSGrid(i).rate
      rs.Update
    Next i
End If

WriteBands_End:
  Set db = Nothing
  Call xReturn("WriteBands")
  Exit Function

WriteBands_Err:
  Call ErrorMessage(ERR_ERROR, Err, "WriteBands", "ERR_UNDEFINED", "Undefined error.")
  Resume WriteBands_End
End Function

