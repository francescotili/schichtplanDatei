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

Private isMoChecked, isDiChecked, isMiChecked, isDoChecked, isFrChecked As Range
Private nextDayBtn, prevDayBtn As Shape
Private weekStatusCell, actualTagCell As Range

Private Sub Worksheet_Activate()
  Set sheet = ActiveWorkbook.Worksheets(GWeekModifySheetName)
  
  Set isMoChecked = Range("Sh2_IsMoChecked")
  Set isDiChecked = Range("Sh2_IsDiChecked")
  Set isMiChecked = Range("Sh2_IsMiChecked")
  Set isDoChecked = Range("Sh2_IsDoChecked")
  Set isFrChecked = Range("Sh2_IsFrChecked")
  
  Set nextDayBtn = Me.sheet.Shapes("Sh2_NextDayBtn")
  Set prevDayBtn = Me.sheet.Shapes("Sh2_PreviousDayBtn")
  
  Set weekStatusCell = Range("Sh2_WeekStatusCell")
  Set actualTagCell = Range("Sh2_ModifyingTagLabel")
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
  actualTagCell.Value = "MONTAG bis FREITAG"
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
  isMoChecked.Value = True
  isDiChecked.Value = True
  isMiChecked.Value = True
  isDoChecked.Value = True
  isFrChecked.Value = True
End Sub

Private Sub SetCustomMode()
  weekStatusCell.Value = WeekStatus.Custom
  actualTagCell.Value = "MONTAG"
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoTrue
  isMoChecked.Value = True
  isDiChecked.Value = False
  isMiChecked.Value = False
  isDoChecked.Value = False
  isFrChecked.Value = False
End Sub

Private Sub SetEmptyMode()
  weekStatusCell.Value = WeekStatus.Emtpy
  actualTagCell.Value = "MONTAG bis FREITAG"
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
  isMoChecked.Value = True
  isDiChecked.Value = True
  isMiChecked.Value = True
  isDoChecked.Value = True
  isFrChecked.Value = True
End Sub

Public Sub GoToNextDay()
  Select Case actualTagCell.Value
    Case "MONTAG"
      actualTagCell.Value = "DIENSTAG"
      isMoChecked.Value = False
      isDiChecked.Value = True
      prevDayBtn.Visible = msoTrue
    Case "DIENSTAG"
      actualTagCell.Value = "MITTWOCH"
      isDiChecked.Value = False
      isMiChecked.Value = True
    Case "MITTWOCH"
      actualTagCell.Value = "DONNERSTAG"
      isMiChecked.Value = False
      isDoChecked.Value = True
    Case "DONNERSTAG"
      actualTagCell.Value = "FREITAG"
      isDoChecked.Value = False
      isFrChecked.Value = True
      nextDayBtn.Visible = msoFalse
  End Select
End Sub

Public Sub GoToPreviousDay()
  Select Case actualTagCell.Value
    Case "FREITAG"
      actualTagCell.Value = "DONNERSTAG"
      isFrChecked.Value = False
      isDoChecked.Value = True
      nextDayBtn.Visible = msoTrue
    Case "DONNERSTAG"
      actualTagCell.Value = "MITTWOCH"
      isDoChecked.Value = False
      isMiChecked.Value = True
    Case "MITTWOCH"
      actualTagCell.Value = "DIENSTAG"
      isMiChecked.Value = False
      isDiChecked.Value = True
    Case "DIENSTAG"
      actualTagCell.Value = "MONTAG"
      isDiChecked.Value = False
      isMoChecked.Value = True
      prevDayBtn.Visible = msoFalse
  End Select
End Sub
