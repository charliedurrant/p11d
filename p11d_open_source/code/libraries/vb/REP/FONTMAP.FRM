VERSION 5.00
Begin VB.Form frmFontMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font Mappings"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Fontmap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMap 
      Caption         =   "&Map"
      Default         =   -1  'True
      Height          =   375
      Left            =   1950
      TabIndex        =   2
      Top             =   2700
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3300
      TabIndex        =   3
      Top             =   2700
      Width           =   1260
   End
   Begin VB.ComboBox cmbMapTo 
      Height          =   315
      Left            =   75
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1950
      Width           =   4515
   End
   Begin VB.ListBox lstRequired 
      Height          =   1815
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4515
   End
End
Attribute VB_Name = "frmFontMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdMap_Click()
  If (CStr(Me.lstRequired.List(Me.lstRequired.ListIndex)) <> "") And (CStr(Me.cmbMapTo.Text) <> "") Then
    Call WriteIniEntry("PRINTER", CStr(Me.lstRequired.List(Me.lstRequired.ListIndex)), CStr(Me.cmbMapTo.Text))
  End If
End Sub
