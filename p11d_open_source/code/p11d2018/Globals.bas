Attribute VB_Name = "Globals"
Option Explicit

Public gbForceExit As Boolean
Public gbAllowAppExit As Boolean
' app globals

Public p11d32 As p11d32
Public sql As SQLQUERIES

Public gPreAlloc As PreAllocate
Public li As listitems

Public CurrentForm As Form

Public Type REPORT_SETTINGS
  Name As String
  Orientation As REPORT_ORIENTATION
  GroupHeader As Boolean 'if the report is just a group category in the rports treeview
  ParentName As String
  IgnoreZeroOnly As Boolean   'km
End Type

Public Type BEN_DATA_STATIC
  DataType As DATABASE_FIELD_TYPES
  MMFieldSize As Long
  MMRequired As Boolean
  RequiresCalculate As Boolean
End Type

Public Type BEN_DATA_LINK
  Initialised As Boolean
  BenefitTable As BENEFIT_TABLES
  UDMFields() As Variant
  UDMDescriptions() As Variant 'descriptions of UDM fields has same UBound as UDMfields if AllowZeroLwngth
  UDMFieldIDs() As Long 'fields id relate to enum values for benefits
  UDMUBound As Long 'lbound always one based related to above arrays
  StaticData() As BEN_DATA_STATIC
End Type

Public Type BEN_DATA
  DataLinks([BC_FIRST_ITEM] To [BC_REAL_BENEFITS_LAST_ITEM]) As BEN_DATA_LINK
  UDMTableOffsets() As BEN_CLASS
  UDMTableOffsetUBound As Long
End Type

Public Type COMPANY_CAR_CHECK
  Employee_db As String
  PersonnelNumber_db As String
  Amended As Boolean
  AvailableFrom_db As Date
  OldAvailableTo_db As Date
  NewAvailableTo_db As Date
  AvailableToAmended As Boolean
  Registration_db As String
  RegistrationReplaced_db As String
  Replaced_db As Boolean
  Replacement_db As Boolean
  SecondCar_db As Boolean       ' MP DB in T_BenCar
  MakeAndModel_db As String     ' MP DB in T_BenCar
  MakeModelReplaced_db As String
  DateCarReplaced_db As Date    ' MP DB in T_BenCar
  OldDateRegistered_db As Date
  NewDateRegistered_db As Date
  DateRegisteredAmended As Boolean
  Comments_db As String
  EmployeeStartDate_db As Date
  EmployeeLeaveDate_db As Date
  'IK 17/06/2003
  DaysUnavailable_db As Long    'MP DB ToDo Confirm not in use & del (fld read from T_BenCar)
  FuelAvailableFrom_db As Date
  FuelOldAvailableTo_db As Date
'MP DB (not used)  FuelNewAvailableTo As Date
'MP DB (not used)    FuelAvailableToAmended As Boolean
  OldNumberOfUsers_db As Long
  NewNumberOfUsers_db As Long
End Type

