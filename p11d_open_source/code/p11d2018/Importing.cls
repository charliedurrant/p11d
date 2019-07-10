VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Importing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPostProcessImport



'// import strings the user will see from the import wizard
Private m_db As Database
Private Const S_SRCFIELD_DEF As String = "P_NUM"
Private Const S_DSTFIELD_DEF As String = S_SRCFIELD_DEF


Private Function ImportName(ImpID As QUERY_NAMES) As String
  
On Error GoTo ImportName_ERR

  Call xSet("ImportName")

  Select Case ImpID
    Case IMP_EMPLOYEES
      ImportName = "Employees"
    Case IMP_HOMEPHONES
      ImportName = "Home phones"
    Case IMP_MOBILEPHONES
      ImportName = "Mobile phones"
    Case IMP_ENTERTAINMENT
      ImportName = "Entertainment"
    Case IMP_CHAUFFEUR
      ImportName = "Chauffeur"
    Case IMP_CREDITCARDS_AND_VOUCHERS
      ImportName = "Credit cards and vouchers"
    Case IMP_EDUCATION
      ImportName = "Education"
    Case IMP_GENERAL
      ImportName = "General"
    Case IMP_MEDICAL
      ImportName = "Medical"
    Case IMP_NOTIONAL_PAYMENTS
      ImportName = "Notional payments"
    Case IMP_NURSERY
      ImportName = "Nursery"
    Case IMP_SUBSCRIPTIONS
      ImportName = "Subscriptions"
    Case IMP_TAXPAID
      ImportName = "Tax paid"
    Case IMP_TRAVEL_AND_SUBSISTENCE
      ImportName = "Travel and subsistence"
      
    Case Else
      ECASE ("ImportName")
  End Select
  
ImportName_END:
  Call xReturn("ImportName")
  Exit Function
ImportName_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ImportName", "Import Name", "Unable to get the import name for ID " & "ImpID" & ".")
  Resume ImportName_END
End Function


Public Function InitImport() As Boolean
  Dim ic As ImportClass
  Dim iw As ImportWizard
  Dim q As SQLQUERIES
  Dim db As Database
  Dim Employer As Employer
    
  On Error GoTo InitImport_Err
  Call xSet("InitImport")
  
  Set q = New SQLQUERIES
    
  If Not F_ErScreen.LB.SelectedItem Is Nothing Then
    'close employer
    Call LoadEmployer(Nothing, False)
    'get ref
    Set Employer = Employers.Item(F_ErScreen.LB.SelectedItem.Tag)
    'open
    If Not Employer Is Nothing Then
      Call Employer.OpenEmployer
        Set m_db = Employer.db
        
        Set ic = New ImportClass
        Set ic.Workspace = P11DWS
        Set iw = ic.ImportWizard
        Call AddImportQueries(iw)
         
        Call iw.AddFileExtension("Comma delimeted files (*.csv)", "*.csv", True, True)
        Call iw.AddFileExtension("Text files (*.txt)", "*.txt")
        Call iw.AddFileExtension("All files (*.*)", "*.*")
        
        iw.AutoOpenSourceFile = True
        iw.TitlePrefix = "P11D98 Importer  -  "
        iw.AllowChangeUpdateType = True
        
        Do
          iw.StartWizard
        Loop While iw.ImportAnother
                
        InitImport = True
      End If
    End If
  End If
  


InitImport_End:
  If Not ic Is Nothing Then ic.KillImporter
  Set ic = Nothing
  Set iw = Nothing
  Call xReturn("InitImport")
  Exit Function

InitImport_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitImport", "ERR_UNDEFINED", "Undefined error.")
  Resume InitImport_End
End Function

Public Function CheckParentNotExists(sDeleteSQL As String, sSourceTable As String, sParentTable As String, sSrcField As String, sParentField As String) As Long
  Dim sSelectSQL As String, sSQL
  Dim rs As Recordset
  
  On Error GoTo CheckParentNotExists_Err
  Call xSet("CheckParentNotExists")

  If Len(sSourceTable) > 0 And Len(sParentTable) > 0 And Len(sSrcField) And Len(sParentField) Then
  
    sSQL = sSourceTable & "." & sSrcField & ", " & sParentTable & "." & sParentField & "FROM " & sSourceTable & "LEFT JOIN " & sParentTable & " ON " & sSourceTable & "." & sSrcField & _
    " = " & sParentTable & "." & sParentField & " Where " & sParentTable & "." & sParentField & " is null"
    
    sSelectSQL = "SELECT DISTINCTROW " & sSQL
    sDeleteSQL "DELETE DISTINCTROW " & sSQL
    
    Set rs = m_db.OpenRecordset(sSelectSQL, dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
      rs.MoveLast
      CheckParentNotExists = rs.RecordCount
    End If
  Else
    'err zzzz
  End If


CheckParentNotExists_End:
  Call xReturn("CheckParentNotExists")
  Exit Function

CheckParentNotExists_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckParentNotExists", "ERR_UNDEFINED", "Undefined error.")
  Resume CheckParentNotExists_End
End Function

Private Function AddImportQueries(iw As ImportWizard) As Boolean
  Dim q As SQLQUERIES
  Dim sKeyFields() As String
  
  On Error GoTo AddImportQueries_Err
  Call xSet("AddImportQueries")

  Set q = New SQLQUERIES
  
  'format is
  '1 add query
  '2 Add keys
  '3 add defaults
  
  With iw
    ReDim sKeyFields(0 To 0)
    sKeyFields(0) = "P_NUM"
    Call .AddSQL(q.Queries(IMP_EMPLOYEES), m_db, ImportName(IMP_EMPLOYEES), , , UPDATE_FIRST, sKeyFields, , Me)
    
    ReDim sKeyFields(0 To 1)
    sKeyFields(0) = "P_NUM"
    sKeyFields(1) = "Item"
    Call .AddSQL(q.Queries(IMP_HOMEPHONES), m_db, ImportName(IMP_HOMEPHONES), , , UPDATE_FIRST, sKeyFields, , Me)
    
    'iw.AddDefaultStatic(
  
    'Call iw.AddSQL(q.Queries(IMP_HOMEPHONES), Employer.db, ImportName(IMP_HOMEPHONES), , , UPDATE_FIRST, sKeyFields)
    'Call iw.AddSQL(q.Queries(IMP_MOBILEPHONES), Employer.db, ImportName(IMP_MOBILEPHONES), , , UPDATE_FIRST, sKeyFields)
    'Call iw.AddSQL(q.Queries(IMP_CHAUFFEUR), Employer.db, ImportName(IMP_CHAUFFER), , , UPDATE_FIRST, sKeyFields)
    
'    Call iw.AddSQL(q.Queries(IMP_CREDITCARDS_AND_VOUCHERS), Employer.db, ImportName(IMP_CREDITCARDS_AND_VOUCHERS), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_EDUCATION), Employer.db, ImportName(IMP_EDUCATION), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_ENTERTAINMENT), Employer.db, ImportName(IMP_ENTERTAINMENT), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_GENERAL), Employer.db, ImportName(IMP_GENERAL), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_MEDICAL), Employer.db, ImportName(IMP_MEDICAL), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_NOTIONAL_PAYMENTS), Employer.db, ImportName(IMP_NOTIONAL_PAYMENTS), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_NURSERY), Employer.db, ImportName(IMP_NURSERY), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_SUBSCRIPTIONS), Employer.db, ImportName(IMP_SUBSCRIPTIONS), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_TAXPAID), Employer.db, ImportName(IMP_TAXPAID), , , UPDATE_FIRST, sKeyFields, , IPPI)
'    Call iw.AddSQL(q.Queries(IMP_TRAVEL_AND_SUBSISTENCE), Employer.db, ImportName(IMP_TRAVEL_AND_SUBSISTENCE), , , UPDATE_FIRST, sKeyFields, , IPPI)
    
  End With


AddImportQueries_End:
  Call xReturn("AddImportQueries")
  Exit Function

AddImportQueries_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddImportQueries", "ERR_UNDEFINED", "Undefined error.")
  Resume AddImportQueries_End
  Resume
End Function



Private Function IPostProcessImport_Process(ic As TCSIMP.ImportClass, RollBack As Boolean) As Long
  Dim sSrcTable As String, sParentTable As String
  Dim rs As Recordset
  Dim sErrorMessage As String, sErrorDescription1 As String, sErrorDescription2 As String
  Dim sButtonDescription As String
  Dim sImport As String
  Dim sDeleteSQL As String
  
  Dim lNumberNotExisting As Long
  
  On Error GoTo IPostProcessImport_ERR
  
  sParentTable = "T_EMPLOYEES"
  
  If Not ic Is Nothing Then
    If Not ic.ImportWizard Is Nothing Then
      If Not ic.ImportWizard.CurrentDest Is Nothing Then
        sImport = ic.ImportWizard.CurrentDest.Name
        Select Case sImport
          Case ImportName(IMP_HOMEPHONES)
            sSrcTable = "T_BENPHONE"
            sErrorDescriptionPart1 = "There were "
            sErrorDescriptionPart2 = " employee records for the phones imported"
        End Select
        
        If Len(sSrcTable) Then
          lNumberNotExisting = CheckParentNotExists(sDeleteSQL, sSrcTable, sParentTable, S_SRCFIELD_DEF, S_DSTFIELD_DEF)
          If lNumberNotExisting > 0 Then
            sErrorMessage = "There were errors encountered during the " & iImport & " import." & vbCrLf
            
            If lNumberNotExisting = ic.ImportedLineCount Then
              'no records were valid
              sErrorMessage = sErrorMessage & sErrorDescriptionPart1 & "no" & sErrorDescription2 & vbCrLf & vbCrLf
              sButtonDescription = "Press close to continue."
              sErrorMessage = sErrorMessage & sButtonDescription
               
              l = MultiDialog("Errors in import", sErrorMessage, "&Close")
              RollBack = True
            Else
              'some were valid
              sErrorMessage = sErrorMessage & sErrorDescriptionPart1 & "some" & sErrorDescription2
              sButtonDescription = "Press discard invalid to save the valid records or discard all to lose all imported records."
              sErrorMessage = sErrorMessage & sButtonDescription
              
              l = MultiDialog("Errors in import", sErrorMessage, "Discard &Invalid", "Discard &All")
              If l = 1 Then
                RollBack = False
                m_db.Execute (sDeleteSQL)
              Else
                RollBack = True
              End If
            End If
            
            IPostProcessImport_Process = lNumberNotExisting
          End If
        End If
          
      End If
    End If
  End If
  
  
  
IPostProcessImport_END:
  Call xReturn("IPostProcessImport")
  Exit Function
IPostProcessImport_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IPostProcessImport", "Post Process Import", "Error in post processing the import for " & ic.ImportWizard.CurrentDest.Name & ".")
  Resume IPostProcessImport_END

End Function
