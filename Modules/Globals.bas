Attribute VB_Name = "Globals"
' --------------------------------------------
' MODULE FOR GLOBAL VARIABLES
' Here are all the variables that are used
' throughout the Excel File
' --------------------------------------------

' Global variables
Public Const GAdminPassword As String = "franci2021"
Public Const GMaxHistoryEntries As Integer = 15

' Table names
Public Const GWorkerTableName As String = "MitarbeiterList"
Public Const GVacationsTableName As String = "FeiertageList"
Public Const GHistoryTableName As String = "Historie"
Public Const GAbsencesTableName As String = "UrlaubsplanDB"
Public Const GShiftsTableName As String = "SchichtplanDB"

' Sheet names
Public Const GUserSheetName As String = "Personal"
Public Const GVacationsSheetName As String = "Feiertage"
Public Const GHistorySheetName As String = "Historie"
Public Const GAbsencesSheetName As String = "AH_DB"
Public Const GShiftsSheetName As String = "SP_DB"
Public Const GWeekModifySheetName As String = "Bearbeiten"

' Enums
Public Enum WeekStatus
 Emtpy ' The week is not yet initialize AKA no data contained
 Normal ' The week is normal, every weekday has the same plan (except Weekend)
 Custom ' The is not normal, every day has different values
End Enum
