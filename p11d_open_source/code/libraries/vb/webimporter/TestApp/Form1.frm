VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   6465
   ClientTop       =   6240
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton aplus 
      Caption         =   "Aplus"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aplus_Click()
Dim x As Import
Dim sImportID As String
Dim zz As WebErrors
Dim ConnTest As ADODB.Connection


Set x = New Import
Set zz = New WebErrors
Set ConnTest = New ADODB.Connection

x.ImporterDatabaseTarget = DB_TARGET_ORACLE
sImportID = x.ImportFile("c:\APLUSIMPORT\TESTSpec.xml", XML_FILE, "c:\APLUSIMPORT\aplustest.csv", "bbalraj")
If Len(sImportID) = 0 Then
    MsgBox "Import Failed"
  Else
    MsgBox (sImportID)
   ' Call x.UndoImport(sImportID, "BALI", "c:\bin\ImportSpec.xml")   'If Not x.UndoImport(sImportID, "BALI", "c:\ing\ImportSpec.xml", XML_FILE) Then MsgBox "Undo failed"
  End If
Set zz = Nothing
  Set x = Nothing
  Set ConnTest = Nothing

End Sub

Private Sub cmdTest_Click()
Dim x As Import
Dim sImportID As String
Dim zz As WebErrors
Dim ConnTest As ADODB.Connection

On Error GoTo err

Set x = New Import
Set zz = New WebErrors
Set ConnTest = New ADODB.Connection

  ConnTest.ConnectionString = "PROVIDER=SQLOLEDB;Data Source=BB-ADSSVR;Initial Catalog=WebImporter;User ID=sa;Password=ukcentral"
  x.ImporterDatabaseTarget = DB_TARGET_SQLSERVER
  'sImportID = x.ImportFile("c:\bin\importSpec.xml", XML_FILE, "c:\bin\imp4.csv", "bbalraj")
  sImportID = x.ImportFile("c:\ing\importSpec.xml", XML_FILE, "c:\ing\test3.csv", "bbalraj")
  'sImportID = x.ImportFile("c:\malc\xmlSpec.xml", XML_FILE, "c:\malc\test.csv", "bbalraj")
  'sImportID = x.ImportFile("c:\mc\spec\utn_static_spec.xml", XML_FILE, "c:\mc\utnstatic.csv", "bbalraj")
  'sImportID = x.ImportFile("c:\mc\spec\utn_values_spec.xml", XML_FILE, "c:\mc\utnvalues.csv", "bbalraj")
  'sImportID = x.ImportFile("c:\mc\spec\utn_sectors_spec.xml", XML_FILE, "c:\mc\utnsectors.csv", "bbalraj")
  
  If Len(sImportID) = 0 Then
    MsgBox "Import Failed"
  Else
    MsgBox (sImportID)
   'If Not x.UndoImport(sImportID, "BALI", "c:\bin\ImportSpec.xml", XML_FILE) Then MsgBox "Undo failed"
  
  End If
  Set zz = Nothing
  Set x = Nothing
  Set ConnTest = Nothing
  Exit Sub

err:
  Call ErrorMessage(ERR_ERROR, err, "Import Test App", "Import err", err.Description)
End Sub



