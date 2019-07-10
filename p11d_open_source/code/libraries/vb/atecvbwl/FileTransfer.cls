VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiPartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cadxx completely replaced
Option Explicit
Private m_MaxKiloBytes As Long

Public Property Get MaxKiloBytes() As Long
  MaxKiloBytes = m_MaxKiloBytes
End Property

Public Property Let MaxKiloBytes(ByVal NewValue As Long)
  If NewValue <= 0 Then Call Err.Raise(ERR_MULTI_PART_FORM, "MaxKiloBytes", "MaxKiloBytes can not be less than 1")
  m_MaxKiloBytes = NewValue
End Property

Private Function BstrToChar(ByVal s As String) As String
  Dim i As Long, v As Variant
  
  On Error GoTo BstrToChar_ERR
  For i = 1 To Len(s)
    v = v & ChrB(AscB(Mid(s, i, 1)))
  Next
  BstrToChar = v
  Exit Function

BstrToChar_ERR:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "BstrToChar", COMPONENT_NAME), Err.Description
  Resume
End Function

Private Function CharToBstr(ByVal byteStr As String) As String
  Dim i As Long
  Dim qs As QString
  
  On Error GoTo CharToBstr_Err
  Set qs = New QString
  For i = 1 To LenB(byteStr)
    qs.Append Chr(AscB(MidB(byteStr, i, 1)))
  Next
  CharToBstr = qs
  
CharToBstr_End:
  Exit Function
CharToBstr_Err:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "CharToBstr", COMPONENT_NAME), Err.Description
  Resume
End Function

'todo error numbers
Private Function MultiPartFormData(ByRef endpos As Long, ByVal bdata As Variant, StartPos As Long, ByVal sname As String) As String
  Dim p0 As Long, p1 As Long
  Dim s As String, v As Variant, i As Long
  On Error GoTo MultiPartFormData_ERR
  
  p0 = InStrB(StartPos, bdata, BstrToChar(sname))
  If p0 = 0 Then Exit Function
  p0 = p0 + Len(sname)
  p1 = InStrB(p0, bdata, BstrToChar(Chr(10)))
  s = CharToBstr(MidB(bdata, p0, (p1 - p0) - 1))
  i = GetDelimitedValues(v, s, True, True, ";")
  If i > 1 Then
    endpos = StartPos
  Else
    'endpos = p1 - 1
  End If
  
  MultiPartFormData = v(1)
  endpos = p1
  
MultiPartFormData_END:
  Exit Function
MultiPartFormData_ERR:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "MultiPartFormData", COMPONENT_NAME), Err.Description
  Resume
End Function
Private Function GetLine(ByRef NewPos As Long, ByVal StartPos As Long, ByRef BinaryData As Variant) As String
  Dim i As Long
  i = InStrB(StartPos, BinaryData, BstrToChar(Chr(10)))
  GetLine = CharToBstr(MidB(BinaryData, StartPos, (i - 1) - StartPos))
  If i > 0 Then
    NewPos = i + 1
  End If
End Function
Public Function Process(ByVal Request As Request, Response As Response) As Collection
  Dim files As Collection
  Dim BinaryData As Variant
  Dim p0 As Long, p1 As Long, pEndOfFile As Long
  Dim v As Variant
  Dim Boundary As String, EndOfFile As String
  Dim sLine As String
  Dim f As MultiPartFormElement, sFile As String, sPath As String, sExt As String, sFileAndPath As String
  Dim b() As Byte
  Dim ivalues As Long
  Dim i As Long
  
  On Error GoTo Process_ERR
  
  p0 = 1
  Set files = New Collection
  Boundary = Request.ServerVariables("HTTP_CONTENT_TYPE")
  If InStr(1, Boundary, "multipart/form-data", vbTextCompare) = 0 Then GoTo Process_END
  Boundary = "--" & Trim$(Mid$(Boundary, InStr(1, Boundary, "boundary=", vbBinaryCompare) + Len("boundary=")))
  EndOfFile = Boundary & "--"
  Boundary = BstrToChar(Boundary)
  BinaryData = Request.BinaryRead(Request.TotalBytes)
  
  pEndOfFile = InStrB(1, BinaryData, BstrToChar(EndOfFile))
  p1 = 1
  p0 = InStrB(p0, BinaryData, Boundary)
  Do While ((p0 <> 0))
    p0 = p0 + LenB(Boundary)
    p1 = p0 + 2
    Set f = New MultiPartFormElement
    
    
    sLine = GetLine(p1, p1, BinaryData)
    sLine = Replace(sLine, "Content-Disposition: form-data;", "", , , vbTextCompare)
    ivalues = GetDelimitedValues(v, sLine, True, True, ";")
    f.Name = Replace(v(1), "name=", "")
    If ivalues = 2 Then
      f.FileName = Replace(v(2), "filename=", "")
      Call SplitPath(f.FileName, sPath, sFile, sExt)
      f.FileName = sFile & sExt
      f.Path = Fullpath(sPath)
      sLine = GetLine(p1, p1, BinaryData)
      f.ContentType = Replace(sLine, "Content-Type: ", "")
      p0 = InStrB(p0 + 1, BinaryData, Boundary)
      p1 = p1 + 2
      f.Data = MidB(BinaryData, p1, ((p0) - p1))
      
      
      
      f.DataType = MPFDT_BINARY
    Else
      f.DataType = MPFDT_TEXT
      f.Data = Replace(GetLine(p1, p1 + 2, BinaryData), Chr$(160), " ")
      p0 = InStrB(p0, BinaryData, Boundary)
    End If
    
    Call files.Add(f, f.Name)
    
    If pEndOfFile <= p0 Then Exit Do
  Loop
Process_END:
  Set Process = files
  Exit Function
Process_ERR:
  Call Err.Raise(Err.Number, ErrorSourceComponentEx(Err, "Process", COMPONENT_NAME), Err.Description)
  Resume
End Function
Private Sub Class_Initialize()
  m_MaxKiloBytes = 2000000
End Sub

