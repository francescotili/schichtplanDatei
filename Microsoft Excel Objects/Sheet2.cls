VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --------------------------------------------
' WEEK BEARBEITEN PUBLIC FUNCTIONS
' Manage methods for getting all the data for
' Sheet2 (Bearbeiten)
' --------------------------------------------
Public sheet As Worksheet

Private nextDayBtn, prevDayBtn As Shape
Private weekStatusCell, activeDayCell As Range

Public Sub Worksheet_Activate()
  Set sheet = ActiveWorkbook.Worksheets(GWeekModifySheetName)
  
  Set nextDayBtn = Me.sheet.Shapes("Sh2_NextDayBtn")
  Set prevDayBtn = Me.sheet.Shapes("Sh2_PreviousDayBtn")
  
  Set weekStatusCell = Range("Sh2_WeekStatusCell")
  Set activeDayCell = Range("Sh2_ActiveDay")
End Sub

Public Sub WeekDataLoad()
  Dim shiftList As New SchichtList
  Dim i, j As Long
  Dim weekNumber As Integer
  
  ' Get the week number
  weekNumber = CInt(Range("Sh2_WeekToLoadCell").Value)
  
  ' Get the week data
  shiftList.WeekRead (weekNumber)
  
  ' Worflow based on week status (read at monday)
  Select Case shiftList.weekData.Item(1).shiftStatus
    Case WeekStatus.Emtpy
      SetEmptyMode
    Case WeekStatus.Normal
      SetNormalMode
    Case WeekStatus.Custom
      SetCustomMode
  End Select
  Populate shiftList.weekData, 1 ' Monday
End Sub

Private Sub Populate(weekData As Collection, weekDay As Integer)
  ' This function take the day of the week and populate the various
  ' sections of the Worksheet.
  ' Warning: before calling this function, please load and pass the right data!
  
  Dim stanzStartRow, qualiStartRow, lagerStartRow, bueroStartRow, galvaStartRow, instStartRow As Integer
  Dim shiftsStartCol, startWorkPlace As Integer
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
'  sheet.Unprotect Password:=GAdminPassword
  
  ' Set variable for the view
  shiftStartCol = 7 'G
  
  ' Stanzerei
  stanzStartRow = 7
  startWorkPlace = 1
  endWorkPlace = 11
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((stanzStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (weekDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Qualität
  qualiStartRow = 41
  startWorkPlace = 12
  endWorkPlace = 13
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((qualiStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (weekDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Lager+Trommlerei
  lagerStartRow = 48
  startWorkPlace = 14
  endWorkPlace = 15
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((lagerStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (weekDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Büro
  bueroStartRow = 55
  startWorkPlace = 16
  endWorkPlace = 20
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((bueroStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (weekDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Galvanik
  galvaStartRow = 71
  startWorkPlace = 21
  endWorkPlace = 26
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((galvaStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (weekDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Instandsetzung
  instaStartRow = 90
  startWorkPlace = 27
  endWorkPlace = 33
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((instaStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (weekDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Enable screen updating and protection
  Application.ScreenUpdating = True
'  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub ToggleMode()
  Select Case weekStatusCell.Value
    Case WeekStatus.Emtpy
      SetNormalMode
    Case WeekStatus.Normal
      SetCustomMode
    Case WeekStatus.Custom
      SetNormalMode
  End Select
End Sub

Private Sub SetNormalMode()
  weekStatusCell.Value = WeekStatus.Normal
  activeDayCell.Value = 0
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
  
End Sub

Private Sub SetCustomMode()
  weekStatusCell.Value = WeekStatus.Custom
  activeDayCell.Value = 1
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoTrue
End Sub

Private Sub SetEmptyMode()
  weekStatusCell.Value = WeekStatus.Emtpy
  activeDayCell.Value = 0
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
End Sub

Public Sub GoToNextDay()
  Select Case activeDayCell.Value
    Case 1 'Montag > Dienstag
      activeDayCell.Value = 2
      prevDayBtn.Visible = msoTrue
    Case 2 'Dienstag > Mittwoch
      activeDayCell.Value = 3
    Case 3 'Mittwoch > Donnerstag
      activeDayCell.Value = 4
    Case 4 'Donnerstag > Freitag
      activeDayCell.Value = 5
      nextDayBtn.Visible = msoFalse
  End Select
End Sub

Public Sub GoToPreviousDay()
  Select Case activeDayCell.Value
    Case 5 'Freitag > Donnerstag
      activeDayCell.Value = 4
      nextDayBtn.Visible = msoTrue
    Case 4 'Donnerstag > Mittwoch
      activeDayCell.Value = 3
    Case 3 'Mittwoch > Dienstag
      activeDayCell.Value = 2
    Case 2 'Dienstag > Montag
      activeDayCell.Value = 1
      prevDayBtn.Visible = msoFalse
  End Select
End Sub
