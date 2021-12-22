VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --------------------------------------------
' DAY VIEW FUNCTIONS
' Manage methods for getting all the data for
' Sheet6 (ÜB Tag)
' --------------------------------------------
Public sheet As Worksheet
Private weekData As Collection

Private dayToLoadCell, activeWeekCell, activeDayCell As Range

Public Sub Worksheet_Activate()
  Set sheet = Sheet6
  
  Set dayToLoadCell = Range("Sh6_DayToLoadCell")
  Set activeWeekCell = Range("Sh6_ActiveWeek")
  Set activeDayCell = Range("Sh6_ActiveDay")
  'Set actualYearCell = Sheet3.Range("Global_ActualYear")
  
  DayDataLoad ' Needed to avoid errors when returning to the sheet
End Sub

Public Sub DayDataLoad()
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Get the day data
  ShiftDataLoad
  
  ' POPOLARE LA TABELLA QUI <----------

End Sub

Private Sub ShiftDataLoad()
  Dim shiftList As New SchichtList
  
  ' Get the day data
  shiftList.DayRead dayToLoadCell.Value
  Set weekData = shiftList.weekData
End Sub
