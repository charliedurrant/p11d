VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFPCSVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements tcseg.IValidate

Private Function IValidate_RowValidation(CurrentRow As Long) As Boolean
  'MsgBox ("ValidateRow")
  IValidate_RowValidation = True
End Function
