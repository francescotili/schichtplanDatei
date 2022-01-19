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
  activeDayCell = dayToLoadCell.Value
  ShiftDataLoad
  Populate
  
  ' Get the absence data
  'AbsenceDataLoad

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  'sheet.Protect Password:=GAdminPassword
End Sub

Private Sub ShiftDataLoad()
  Dim shiftList As New SchichtList
  
  ' Get the day data
  shiftList.DayRead dayToLoadCell.Value
  Set weekData = shiftList.weekData 'Contains only 4 shiftCodes
End Sub

Private Sub AbsenceDataLoad()
  
End Sub

Private Sub Populate()
  ' This function populate the various sections of the Worksheet.
  ' Warning: before calling this function, please load and pass the right data!
  
  Dim startRow, startCol As Integer
  Dim numOfWorkplaces As Integer
  Dim i, j As Long
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  EmptyPlan
  
  ' Set variable for the view
  startCol = 9 ' Column I
  startRow = 6
  numOfWorkplaces = Sheet1.ListObjects(GShiftsTableName).DataBodyRange.Columns.Count - 3
  For i = 1 To numOfWorkplaces
    For j = 1 To 4
      If Len(CStr(weekData.Item(j).workPlace(i - 1))) <> 0 Then
        sheet.Cells((startRow + 2 * (i - 1)), (startCol + (j - 1))) = CStr(weekData.Item(j).workPlace(i - 1))
      End If
    Next j
  Next i
End Sub

Private Sub EmptyPlan()
  Dim startCell, endCell As String
  
  startCell = "I6"
  endCell = "L71"
  Range(startCell & ":" & endCell).ClearContents
End Sub

Private Sub EmptyAbsences()
  Dim startCell, endCell As String
  
  startCell = "N6"
  endCell = "P71"
  Range(startCell & ":" & endCell).ClearContents
End Sub
