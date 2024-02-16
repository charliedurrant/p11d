VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form F_TestsResults 
   Caption         =   "Tests results"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton buttonOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   9360
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox richTextBoxResults 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   15901
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"F_TestsREsults.frx":0000
   End
End
Attribute VB_Name = "F_TestsResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Let results(value As String)
  richTextBoxResults.Text = value
  
End Property
