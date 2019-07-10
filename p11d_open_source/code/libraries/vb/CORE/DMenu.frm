VERSION 5.00
Begin VB.Form frmDebugMenu 
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuBreak 
         Caption         =   "&Break"
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&System"
         Index           =   0
         Begin VB.Menu mnuSystemItem 
            Caption         =   "&Environment"
            Index           =   0
         End
         Begin VB.Menu mnuSystemItem 
            Caption         =   "&Components"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSystemItem 
            Caption         =   "&Application"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Database"
         Index           =   1
         Begin VB.Menu mnuDatabaseItem 
            Caption         =   "&Repair and Compact"
            Index           =   0
         End
         Begin VB.Menu mnuDatabaseItem 
            Caption         =   "&SQL View"
            Index           =   1
         End
      End
      Begin VB.Menu mnuUserDebugItem 
         Caption         =   "&Application"
         Begin VB.Menu mnuOtherItem 
            Caption         =   "AppItem"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmDebugMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Call gTCSEventClass.FillDebugMenu(Me)
End Sub

Private Sub mnuBreak_Click()
  Call gTCSEventClass.DebugMenu(0, MNU_BREAK, "BREAK")
End Sub

Private Sub mnuDatabaseItem_Click(Index As Integer)
  Call gTCSEventClass.DebugMenu(CLng(Index), MNU_DATABASE, mnuDatabaseItem(Index).Tag)
End Sub

Private Sub mnuOtherItem_Click(Index As Integer)
  Call gTCSEventClass.DebugMenu(CLng(Index), MNU_APPLICATION, mnuOtherItem(Index).Tag)
End Sub

Private Sub mnuSystemItem_Click(Index As Integer)
  Call gTCSEventClass.DebugMenu(CLng(Index), MNU_SYSTEM, mnuSystemItem(Index).Tag)
End Sub
