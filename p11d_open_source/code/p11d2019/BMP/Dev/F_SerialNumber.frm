VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_SerialNumber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Number"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin atc2valtext.ValText ValSerialNumber 
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   1005
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   23
      Text            =   ""
      TypeOfData      =   3
      AutoSelect      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2115
      TabIndex        =   2
      Top             =   1410
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   1410
      Width           =   1005
   End
   Begin VB.Label lblInstruction 
      Caption         =   "Enter Serial Number:"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   600
      Width           =   3030
   End
   Begin VB.Label lblLicenceType 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblLicencetype"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   3030
   End
End
Attribute VB_Name = "F_SerialNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_OK As Boolean

Private Sub cmdCancel_Click()
  Unload Me
End Sub
Public Function Start() As Boolean
  Call PopulateFields
  m_OK = False
  Call p11d32.Help.ShowForm(Me, vbModal)
  If m_OK Then
    Start = Not ValSerialNumber.FieldInvalid
  End If
End Function
Private Sub cmdOK_Click()
  p11d32.SerialNumber = Replace(ValSerialNumber.Text, "-", "")
  If CBoolean(p11d32.LicenceType) Then
     m_OK = True
     Me.Hide
  Else
     MsgBox "The Serial Number is invalid", vbExclamation, "Invalid Serial number"
  End If
End Sub
Public Sub PopulateFields()

  On Error GoTo PopulateFields_Err
  Call xSet("PopulateFields")
  ValSerialNumber.Text = p11d32.SerialNumber
  If p11d32.LicenceType = LT_STANDARD Then
        Me.Caption = "P11D Serial Number"
        lblLicenceType = "Registered version: P11D Standard Edition"
        lblInstruction = "Enter Serial Number:"
  ElseIf p11d32.LicenceType = LT_INTRANET Then
        Me.Caption = "P11D Serial Number"
        lblLicenceType = "Registered version: P11D Intranet Edition"
        lblInstruction = "Enter Serial Number:"
  ElseIf p11d32.LicenceType = LT_UDM Then
        Me.Caption = "P11D Serial Number"
        lblLicenceType = "Registered version: P11D UDM Edition"
        lblInstruction = "Enter Serial Number:"
  ElseIf p11d32.LicenceType = LT_SHORT Then
        Me.Caption = "P11D Serial Number"
        lblLicenceType = "Registered version: P11D Short Edition"
        lblInstruction = "Enter Serial Number:"
  ElseIf p11d32.LicenceType = LT_DEMO Then
        Me.Caption = "P11D Serial Number"
        lblLicenceType = "Registered version: P11D Demo"
        lblInstruction = "Enter Serial Number:"
  Else 'LT_UNLICENSED
        Me.Caption = "P11D Serial Number"
        lblLicenceType = "Unregistered version"
        lblInstruction = "Please enter a valid Serial Number:"
  End If

PopulateFields_End:
  Call xReturn("PopulateFields")
  Exit Sub

PopulateFields_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PopulateFields", "Error in PopulateFields", "Undefined error.")
  Resume PopulateFields_End
End Sub

