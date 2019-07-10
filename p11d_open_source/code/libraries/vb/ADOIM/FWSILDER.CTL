VERSION 5.00
Begin VB.UserControl FWSlider 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   ScaleHeight     =   3960
   ScaleWidth      =   285
   Begin VB.Line Line10 
      X1              =   72
      X2              =   120
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Line Line9 
      X1              =   168
      X2              =   108
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   3840
   End
End
Attribute VB_Name = "FWSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

RaiseEvent MouseUp(Button, Shift, x, y)

End Sub
