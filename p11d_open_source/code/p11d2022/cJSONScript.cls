VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJSONScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim dictVars As New Dictionary
Dim plNestCount As Long


Public Function Eval(sJSON As String) As String
   Dim sb As New QString
   
   Dim o As Object
   Dim c As Object
   Dim i As Long
   
   Set o = JSON.parse(sJSON)
   If (JSON.GetParserErrors = "") And Not (o Is Nothing) Then
      For i = 1 To o.Count
         Select Case VarType(o.Item(i))
         Case vbNull
            sb.Append "null"
         Case vbDate
            sb.Append CStr(o.Item(i))
         Case vbString
            sb.Append CStr(o.Item(i))
         Case Else
            Set c = o.Item(i)
            sb.Append ExecCommand(c)
         End Select
      Next
   Else
      MsgBox JSON.GetParserErrors, vbExclamation, "Parser Error"
   End If
   Eval = sb.bstr
End Function

Public Function ExecCommand(ByRef obj As Variant) As String
   Dim sb As New QString
   
   If plNestCount > 40 Then
      ExecCommand = "ERROR: Nesting level exceeded."
   Else
      plNestCount = plNestCount + 1
      
      Select Case VarType(obj)
         Case vbNull
            sb.Append "null"
         Case vbDate
            sb.Append CStr(obj)
         Case vbString
            sb.Append CStr(obj)
         Case vbObject
            
            Dim i As Long
            Dim j As Long
            Dim this As Object
            Dim Key
            Dim paramKeys
            
            If TypeName(obj) = "Dictionary" Then
               Dim sOut As String
               Dim sRet As String
   
               Dim keys
               keys = obj.keys
               For i = 0 To obj.Count - 1
                  sRet = ""
             
                  Key = keys(i)
                  If VarType(obj.Item(Key)) = vbString Then
                     sRet = obj.Item(Key)
                  Else
                     Set this = obj.Item(Key)
                  End If
                  
                  ' command implementation
                  Select Case LCase(Key)
                  Case "alert":
                     MsgBox ExecCommand(this.Item("message")), vbInformation, ExecCommand(this.Item("title"))
                     
                  Case "input":
                     sb.Append InputBox(ExecCommand(this.Item("prompt")), ExecCommand(this.Item("title")), ExecCommand(this.Item("default")))
                     
                  Case "switch"
                     sOut = ExecCommand(this.Item("default"))
                     sRet = LCase(ExecCommand(this.Item("case")))
                     For j = 0 To this.Item("items").Count - 1
                        If LCase(this.Item("items").Item(j + 1).Item("case")) = sRet Then
                           sOut = ExecCommand(this.Item("items").Item(j + 1).Item("return"))
                           Exit For
                        End If
                     Next
                     sb.Append sOut
                  
                  Case "set":
                     If dictVars.Exists(this.Item("name")) Then
                        dictVars.Item(this.Item("name")) = ExecCommand(this.Item("value"))
                     Else
                        dictVars.Add this.Item("name"), ExecCommand(this.Item("value"))
                     End If
                     
                  Case "get":
                     sRet = ExecCommand(dictVars(CStr(this.Item("name"))))
                     If sRet = "" Then
                        sRet = ExecCommand(this.Item("default"))
                     End If
                     
                     sb.Append sRet
                     
                  Case "if"
                     Dim val1 As String
                     Dim val2 As String
                     Dim bRes As Boolean
                     val1 = ExecCommand(this.Item("value1"))
                     val2 = ExecCommand(this.Item("value2"))
                     
                     bRes = False
                     Select Case LCase(this.Item("type"))
                     Case "eq" ' =
                        If LCase(val1) = LCase(val2) Then
                           bRes = True
                        End If
                        
                     Case "gt" ' >
                        If val1 > val2 Then
                           bRes = True
                        End If
                     
                     Case "lt" ' <
                        If val1 < val2 Then
                           bRes = True
                        End If
                     
                     Case "gte" ' >=
                        If val1 >= val2 Then
                           bRes = True
                        End If
                     
                     Case "lte" ' <=
                        If val1 <= val2 Then
                           bRes = True
                        End If
                     
                     End Select
                     
                     If bRes Then
                        sb.Append ExecCommand(this.Item("true"))
                     Else
                        sb.Append ExecCommand(this.Item("false"))
                     End If
                     
                  Case "return"
                     sb.Append obj.Item(Key)
                  
                     
                  Case Else
                     If TypeName(this) = "Dictionary" Then
                        paramKeys = this.keys
                        For j = 0 To this.Count - 1
                           If j > 0 Then
                              sRet = sRet & ","
                           End If
                           sRet = sRet & CStr(this.Item(paramKeys(j)))
                        Next
                     End If
                     
                     
                     sb.Append "<%" & UCASE(Key) & "(" & sRet & ")%>"
                     
                  End Select
               Next i
               
            ElseIf TypeName(obj) = "Collection" Then
   
               Dim value
               For Each value In obj
                  sb.Append ExecCommand(value)
               Next value
               
            End If
            Set this = Nothing
   
         Case vbBoolean
            If obj Then sb.Append "true" Else sb.Append "false"
         
         Case vbVariant, vbArray, vbArray + vbVariant
         
         Case Else
            sb.Append Replace(obj, ",", ".")
      End Select
      plNestCount = plNestCount - 1
   End If
   
   ExecCommand = sb.bstr
   Set sb = Nothing
   
End Function




