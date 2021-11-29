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

Private Sub Worksheet_Activate()
  Set sheet = ActiveWorkbook.Worksheets(GWeekModifySheetName)

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
  Select Case Range("Sh2_WeekStatusCell").Value
    Case WeekStatus.Emtpy
      SetNormalMode
    Case WeekStatus.Normal
      SetCustomMode
    Case WeekStatus.Custom
      SetNormalMode
  End Select
End Sub

Private Sub SetNormalMode()
  Range("Sh2_WeekStatusCell").Value = WeekStatus.Normal
  Range("Sh2_ModifyingTagLabel").Value = "MONTAG bis FREITAG"
  Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoFalse
  Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoFalse
  Me.sheet.Shapes("Sh2_MoCheck").OLEFormat.Object.Value = 1 ' -4146 is unchecked
  Me.sheet.Shapes("Sh2_DiCheck").OLEFormat.Object.Value = 1
  Me.sheet.Shapes("Sh2_MiCheck").OLEFormat.Object.Value = 1
  Me.sheet.Shapes("Sh2_DoCheck").OLEFormat.Object.Value = 1
  Me.sheet.Shapes("Sh2_FrCheck").OLEFormat.Object.Value = 1
End Sub

Private Sub SetCustomMode()
  Range("Sh2_WeekStatusCell").Value = WeekStatus.Custom
  Range("Sh2_ModifyingTagLabel").Value = "MONTAG"
  Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoFalse
  Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoTrue
  Range("Sh2_IsMoChecked").Value = True
  Range("Sh2_IsDiChecked").Value = False
  Range("Sh2_IsMiChecked").Value = False
  Range("Sh2_IsDoChecked").Value = False
  Range("Sh2_IsFrChecked").Value = False
End Sub

Private Sub SetEmptyMode()
  Range("Sh2_WeekStatusCell").Value = WeekStatus.Emtpy
  Range("Sh2_ModifyingTagLabel").Value = "MONTAG bis FREITAG"
  Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoFalse
  Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoFalse
  Range("Sh2_IsMoChecked").Value = True
  Range("Sh2_IsDiChecked").Value = True
  Range("Sh2_IsMiChecked").Value = True
  Range("Sh2_IsDoChecked").Value = True
  Range("Sh2_IsFrChecked").Value = True
End Sub

Public Sub GoToNextDay()
  Select Case Range("Sh2_ModifyingTagLabel").Value
    Case "MONTAG"
      Range("Sh2_ModifyingTagLabel").Value = "DIENSTAG"
      Range("Sh2_IsMoChecked").Value = False
      Range("Sh2_IsDiChecked").Value = True
      Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoTrue
    Case "DIENSTAG"
      Range("Sh2_ModifyingTagLabel").Value = "MITTWOCH"
      Range("Sh2_IsDiChecked").Value = False
      Range("Sh2_IsMiChecked").Value = True
    Case "MITTWOCH"
      Range("Sh2_ModifyingTagLabel").Value = "DONNERSTAG"
      Range("Sh2_IsMiChecked").Value = False
      Range("Sh2_IsDoChecked").Value = True
    Case "DONNERSTAG"
      Range("Sh2_ModifyingTagLabel").Value = "FREITAG"
      Range("Sh2_IsDoChecked").Value = False
      Range("Sh2_IsFrChecked").Value = True
      Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoFalse
  End Select
End Sub

Public Sub GoToPreviousDay()
  Select Case Range("Sh2_ModifyingTagLabel").Value
    Case "FREITAG"
      Range("Sh2_ModifyingTagLabel").Value = "DONNERSTAG"
      Range("Sh2_IsFrChecked").Value = False
      Range("Sh2_IsDoChecked").Value = True
      Me.sheet.Shapes("Sh2_NextDayBtn").Visible = msoTrue
    Case "DONNERSTAG"
      Range("Sh2_ModifyingTagLabel").Value = "MITTWOCH"
      Range("Sh2_IsDoChecked").Value = False
      Range("Sh2_IsMiChecked").Value = True
    Case "MITTWOCH"
      Range("Sh2_ModifyingTagLabel").Value = "DIENSTAG"
      Range("Sh2_IsMiChecked").Value = False
      Range("Sh2_IsDiChecked").Value = True
    Case "DIENSTAG"
      Range("Sh2_ModifyingTagLabel").Value = "MONTAG"
      Range("Sh2_IsDiChecked").Value = False
      Range("Sh2_IsMoChecked").Value = True
      Me.sheet.Shapes("Sh2_PreviousDayBtn").Visible = msoFalse
  End Select
End Sub
