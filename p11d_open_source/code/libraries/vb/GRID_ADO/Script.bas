Attribute VB_Name = "Scripting"
Option Explicit
Private mScriptInitialised As Boolean
'Scripting additions
Private mScriptCtrl As ScriptControl
Public mScriptFunctions As AutoScriptFunctions


Public Sub InitScriptControl()
  On Error GoTo InitScriptControl_err:
  
  If Not mScriptInitialised Then
    mScriptInitialised = True
    Set mScriptFunctions = New AutoScriptFunctions
    Set mScriptCtrl = New ScriptControl
    mScriptCtrl.Language = "VBScript"
    mScriptCtrl.AllowUI = False
    mScriptCtrl.UseSafeSubset = True
    Call mScriptCtrl.AddObject("AutoHelp", mScriptFunctions, True)
  End If
  
InitScriptControl_end:
  Exit Sub
  
InitScriptControl_err:
  Call ErrorMessage(ERR_ERROR, Err, "InitScriptControl", "Initialise Scripting", "Unable to Initialise Script Control." & vbCrLf & "Grid Updates/Additions may not work correctly.")
  Resume InitScriptControl_end
End Sub

Public Sub ResetScriptControl()
  On Error Resume Next
  If mScriptInitialised Then
    Call mScriptCtrl.Reset
    Set mScriptCtrl = Nothing
    Set mScriptFunctions = Nothing
    mScriptInitialised = False
  End If
End Sub
  
Public Function AutoEvaluate(ByVal Expression As String) As Variant
  If Not mScriptCtrl Is Nothing Then AutoEvaluate = mScriptCtrl.Eval(Expression)
End Function

