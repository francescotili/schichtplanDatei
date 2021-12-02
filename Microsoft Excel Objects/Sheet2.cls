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
Private weekData As Collection
Private weekNumber As Double

Private nextDayBtn, prevDayBtn As Shape
Private weekToLoadCell, weekStatusCell, activeDayCell As Range

Public Sub Worksheet_Activate()
  Set sheet = ActiveWorkbook.Worksheets(GWeekModifySheetName)
  
  Set nextDayBtn = Me.sheet.Shapes("Sh2_NextDayBtn")
  Set prevDayBtn = Me.sheet.Shapes("Sh2_PreviousDayBtn")
  
  Set weekToLoadCell = Range("Sh2_WeekToLoadCell")
  Set weekStatusCell = Range("Sh2_WeekStatusCell")
  Set activeDayCell = Range("Sh2_ActiveDay")
  
  WeekDataLoad ' Needed to avoid errors when returning to the sheet
End Sub

Public Sub WeekDataLoad()
  Dim shiftList As New SchichtList
  
  ' Get the week data
  weekNumber = CInt(weekToLoadCell.Value)
  Load
  
  ' Worflow based on week status (read at monday)
  Select Case weekData.Item(WeekDay.Monday).shiftStatus
    Case WeekStatus.Emtpy
      'SetEmptyMode
      SetNormalMode
    Case WeekStatus.Normal
      SetNormalMode
    Case WeekStatus.Custom
      SetCustomMode
  End Select
  Populate Monday
End Sub

Public Sub WeekDataSave()
  If weekStatusCell.Value = WeekStatus.Custom Then
    Dim window As New Sheet2_CustomSave
    With window
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
    End With
  Else
    Dim msgAnswer As Integer
    Dim msgText, msgTitle As String
    
    msgText = "Möchten Sie den Schichtplan wirklich speichern?" & vbCrLf & "Vorhandene Daten der ganzen Kalenderwoche werden überschrieben!"
    msgTitle = "Fortfahren?"
    msgAnswer = MsgBox(msgText, vbQuestion + vbYesNo, msgTitle)
  End If
  
  ' Read every schicht for the day or selected days
  ' Build SchichtDay objects and pass them to the Schichtplan class for saving
  
End Sub

Private Sub Load()
  Dim shiftList As New SchichtList
  
  ' Get the week data
  shiftList.WeekRead weekNumber
  Set weekData = shiftList.weekData
End Sub

Private Sub Populate(day As WeekDay)
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
      sheet.Cells((stanzStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (day - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Qualität
  qualiStartRow = 41
  startWorkPlace = 12
  endWorkPlace = 13
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((qualiStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (day - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Lager+Trommlerei
  lagerStartRow = 48
  startWorkPlace = 14
  endWorkPlace = 15
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((lagerStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (day - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Büro
  bueroStartRow = 55
  startWorkPlace = 16
  endWorkPlace = 20
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((bueroStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (day - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Galvanik
  galvaStartRow = 71
  startWorkPlace = 21
  endWorkPlace = 26
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((galvaStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (day - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Instandsetzung
  instaStartRow = 90
  startWorkPlace = 27
  endWorkPlace = 33
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((instaStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (day - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Enable screen updating and protection
  Application.ScreenUpdating = True
'  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub ToggleMode()
  Select Case weekStatusCell.Value
'    Case WeekStatus.Emtpy
'      SetNormalMode
    Case WeekStatus.Normal
      SetCustomMode
    Case WeekStatus.Custom
      SetNormalMode
  End Select
End Sub

Private Sub SetNormalMode()
  weekStatusCell.Value = WeekStatus.Normal
  activeDayCell.Value = WeekDay.Unknown
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
End Sub

Private Sub SetCustomMode()
  weekStatusCell.Value = WeekStatus.Custom
  activeDayCell.Value = WeekDay.Monday
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoTrue
End Sub

Private Sub SetEmptyMode()
  weekStatusCell.Value = WeekStatus.Emtpy
  activeDayCell.Value = WeekDay.Unknown
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
End Sub

Public Sub GoToNextDay()
  activeDayCell.Value = activeDayCell.Value + 1
  If activeDayCell.Value > WeekDay.Monday Then prevDayBtn.Visible = msoTrue
  If activeDayCell.Value = WeekDay.Friday Then nextDayBtn.Visible = msoFalse
  Populate activeDayCell.Value
End Sub

Public Sub GoToPreviousDay()
  activeDayCell.Value = activeDayCell.Value - 1
  If activeDayCell.Value = WeekDay.Monday Then prevDayBtn.Visible = msoFalse
  If activeDayCell.Value < WeekDay.Friday Then nextDayBtn.Visible = msoTrue
  Populate activeDayCell.Value
End Sub

