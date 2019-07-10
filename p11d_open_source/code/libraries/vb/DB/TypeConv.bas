Attribute VB_Name = "TypeConv"
Option Explicit

Public Function DAOtoDatatypeEx(ByVal daoType As Long) As DATABASE_FIELD_TYPES
  Select Case daoType
    Case dbBoolean
      DAOtoDatatypeEx = TYPE_BOOL
    Case dbBinary, dbByte, dbBigInt, dbInteger, dbLong
      DAOtoDatatypeEx = TYPE_LONG
    Case dbChar
      DAOtoDatatypeEx = TYPE_STR
    Case dbCurrency, dbDecimal, dbFloat, dbDouble, dbNumeric, dbSingle
      DAOtoDatatypeEx = TYPE_DOUBLE
    Case dbDate, dbTime, dbTimeStamp
      DAOtoDatatypeEx = TYPE_DATE
    Case dbGUID, dbText, dbMemo
      DAOtoDatatypeEx = TYPE_STR
    Case dbLongBinary
      DAOtoDatatypeEx = TYPE_BLOB
    Case Else ' dbLongBinary, dbLongBinary
      Call ECASE("Unrecognised data Type: " & CStr(daoType))
  End Select
End Function

Public Function RDOtoDatatypeEx(ByVal rdoType As Long) As DATABASE_FIELD_TYPES
  RDOtoDatatypeEx = DAOtoDatatypeEx(RDOtoDAO_Datatype(rdoType))
End Function

Public Function VarTypetoDatatypeEx(ByVal vbType As VbVarType) As DATABASE_FIELD_TYPES
  Select Case vbType
    Case vbCurrency, vbDecimal, vbDouble, vbSingle
      VarTypetoDatatypeEx = TYPE_DOUBLE
    Case vbInteger, vbLong, vbByte
      VarTypetoDatatypeEx = TYPE_LONG
    Case vbDate
      VarTypetoDatatypeEx = TYPE_DATE
    Case vbString
      VarTypetoDatatypeEx = TYPE_STR
    Case vbBoolean
      VarTypetoDatatypeEx = TYPE_BOOL
    Case Else
      Call ECASE("Unrecognised data Type: " & CStr(vbType))
  End Select
End Function

Private Function RDOtoDAO_Datatype(ByVal rdoType As Long) As Long
  Select Case rdoType
    Case rdTypeCHAR
      RDOtoDAO_Datatype = dbChar
    Case rdTypeNUMERIC
      RDOtoDAO_Datatype = dbNumeric
    Case rdTypeDECIMAL
      RDOtoDAO_Datatype = dbDecimal
    Case rdTypeINTEGER
      RDOtoDAO_Datatype = dbInteger
    Case rdTypeSMALLINT
      RDOtoDAO_Datatype = dbInteger
    Case rdTypeFLOAT
      RDOtoDAO_Datatype = dbFloat
    Case rdTypeREAL
      RDOtoDAO_Datatype = dbDouble
    Case rdTypeDOUBLE
      RDOtoDAO_Datatype = dbDouble
    Case rdTypeDATE
      RDOtoDAO_Datatype = dbDate
    Case rdTypeTIME
      RDOtoDAO_Datatype = dbTime
    Case rdTypeTIMESTAMP
      RDOtoDAO_Datatype = dbTimeStamp
    Case rdTypeVARCHAR
      RDOtoDAO_Datatype = dbText
    Case rdTypeLONGVARCHAR
      RDOtoDAO_Datatype = dbMemo
    Case rdTypeBINARY
      RDOtoDAO_Datatype = dbBinary
    Case rdTypeLONGVARBINARY
      RDOtoDAO_Datatype = dbLongBinary
    Case rdTypeBIGINT
      RDOtoDAO_Datatype = dbBigInt
    Case rdTypeTINYINT
      RDOtoDAO_Datatype = dbByte
    Case rdTypeBIT
      RDOtoDAO_Datatype = dbByte
  End Select
End Function

