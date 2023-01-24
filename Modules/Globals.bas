' --------------------------------------------
' MODULE FOR GLOBAL VARIABLES
' Here are all the variables that are used
' throughout the Excel File
' --------------------------------------------

' Global variables
Public Const GAdminPassword As String = "franci2021"
Public Const GMaxHistoryEntries As Integer = 25

' Table names
Public Const GWorkerTableName As String = "MitarbeiterList"
Public Const GVacationsTableName As String = "FeiertageList"
Public Const GHistoryTableName As String = "Historie"
Public Const GAbsencesTableName As String = "UrlaubsplanDB"
Public Const GShiftsTableName As String = "SchichtplanDB"
Public Const GAbsencesModifyTableName As String = "Urlaubsplan"

' Enums
Public Enum WeekStatus
  ' The week is not yet initialized AKA no data contained
  Emtpy
  ' The workdays are normal (all the same plan) and there can
  ' be shifts planned in the weekend
  Normal
  ' The week is not normal, every day can have different value
  Custom
End Enum

Public Enum WeekDays
  ' No weekDay specified
  Unknown
  Monday
  Tuesday
  Wednesday
  Thursday
  Friday
  'Sathurday
  'Sunday
End Enum
