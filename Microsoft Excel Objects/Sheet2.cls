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
Private weekToLoadCell, weekStatusCell, activeDayCell, activeWeekCell, changesDetectedCell As Range

' Strings
Private Const losignDataWarningText = "Im Schichtplan wurden nicht gespeicherte änderungen erkannt!" & vbCrLf & "Wenn Sie fortfahren, gehen die verloren. Fortfahren?"
Private Const savingDataWarningText = "Möchten Sie den Schichtplan wirklich speichern?" & vbCrLf & "Vorhandene Daten der ganzen Kalenderwoche werden überschrieben!"

Private Sub Sh2_WeekSaveBtn_Click()
  WeekDataSave
End Sub

Public Sub Worksheet_Activate()
  Set sheet = ActiveWorkbook.Worksheets(GWeekModifySheetName)
  
  Set nextDayBtn = Me.sheet.Shapes("Sh2_NextDayBtn")
  Set prevDayBtn = Me.sheet.Shapes("Sh2_PreviousDayBtn")
  
  Set weekToLoadCell = Range("Sh2_WeekToLoadCell")
  Set weekStatusCell = Range("Sh2_WeekStatusCell")
  Set activeDayCell = Range("Sh2_ActiveDay")
  Set activeWeekCell = Range("Sh2_ActiveWeek")
  Set changesDetectedCell = Range("Sh2_ChangesDetected")
  
  WeekDataLoad ' Needed to avoid errors when returning to the sheet
End Sub

Public Sub WeekDataLoad()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(losignDataWarningText, vbExclamation + vbYesNo, "ACHTUNG")
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If

Proceed:
  ' Get the week data
  weekNumber = CInt(weekToLoadCell.Value)
  activeWeekCell.Value = weekNumber
  Load
  
  ' Worflow based on week status (read at monday)
  Select Case weekData.Item(WeekDay.Monday).shiftStatus
    Case WeekStatus.Emtpy
      SetEmptyMode
    Case WeekStatus.Normal
      SetNormalMode
      Populate Monday
    Case WeekStatus.Custom
      SetCustomMode
      Populate Monday
  End Select
Abort:
End Sub

Private Sub WeekDataSave()
  Select Case weekStatusCell.Value
  Case WeekStatus.Custom
    Dim window As New Sheet2_CustomSave
    With window
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
    End With
  Case WeekStatus.Emtpy, WeekStatus.Normal
    Dim msgAnswer As Integer
    msgAnswer = MsgBox(savingDataWarningText, vbQuestion + vbYesNo, "Fortfahren?")
    If msgAnswer = vbYes Then Save WeekDay.Unknown
  End Select
End Sub

Public Sub Save(passedDay As WeekDay)
  ' Read every schicht for the day or selected days
  ' Define two cases for empty/normal and for custom
  '   If normal
  '     Build SchichtDay object varying the date component and save them every step
  '   If custom
  '     Build SchichtDay object based on the days selected and save them every step
  
  Dim startRow As Integer
  Dim shiftsStartCol, startWorkPlace, endWorkPlace As Integer
  Dim workPlaces(3, 32) As String
  Dim i, j As Long
  Dim shiftList As New SchichtList
  Dim shiftData As New SchichtDay
  
  ' Set variable for the view
  shiftStartCol = 7 ' Columnn G
  
  ' Read data for every shift and save them in a collection of workPlaces
  For j = 1 To 4
    ' Stanzerei
    startRow = 7
    startWorkPlace = 1
    endWorkPlace = 11
    For i = 1 To (endWorkPlace - startWorkPlace + 1)
      workPlaces(j - 1, startWorkPlace + i - 2) = CStr(sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))).Value)
    Next i
    
    ' Qualitaet
    startRow = 41
    startWorkPlace = 12
    endWorkPlace = 13
    For i = 1 To (endWorkPlace - startWorkPlace + 1)
      workPlaces(j - 1, startWorkPlace + i - 2) = CStr(sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))).Value)
    Next i
    
    ' Lager+Trommlerei
    startRow = 48
    startWorkPlace = 14
    endWorkPlace = 15
    For i = 1 To (endWorkPlace - startWorkPlace + 1)
      workPlaces(j - 1, startWorkPlace + i - 2) = CStr(sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))).Value)
    Next i
    
    ' Buero
    startRow = 55
    startWorkPlace = 16
    endWorkPlace = 20
    For i = 1 To (endWorkPlace - startWorkPlace + 1)
      workPlaces(j - 1, startWorkPlace + i - 2) = CStr(sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))).Value)
    Next i
    
    ' Galvanik
    startRow = 71
    startWorkPlace = 21
    endWorkPlace = 26
    For i = 1 To (endWorkPlace - startWorkPlace + 1)
      workPlaces(j - 1, startWorkPlace + i - 2) = CStr(sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))).Value)
    Next i
    
    ' Instandsetzung
    startRow = 90
    startWorkPlace = 27
    endWorkPlace = 33
    For i = 1 To (endWorkPlace - startWorkPlace + 1)
      workPlaces(j - 1, startWorkPlace + i - 2) = CStr(sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))).Value)
    Next i
  Next j
  
  ' Now save the different days based on the WeekModus
  Dim d, s As Long
  Dim tempWorkPlaces(32) As String
  Select Case activeDayCell.Value
    Case WeekDay.Unknown ' Normal Modus
      saveHistory ("Der Schichtplan für KW" & activeWeekCell.Value & " wurde geändert")
      
      For d = 1 To 5 ' Days of the week
        For s = 1 To 4 ' Shifts
          shiftData.Init weekData.Item(4 * (d - 1) + s).shiftDate, Normal, Right(weekData.Item(4 * (d - 1) + s).shiftCode, 1)
          ' Populate shiftData with a loop
          For i = 0 To 32
            tempWorkPlaces(i) = workPlaces(s - 1, i)
          Next i
          
          shiftData.Populate tempWorkPlaces
          shiftData.shiftStatus = WeekStatus.Normal
          
          ' Now save the SchichtDay (shiftData) object
          shiftList.WeekSave shiftData
        Next s
      Next d
    Case Else
      saveHistory ("Der Schichtplan für den " & weekData.Item(4 * (passedDay - 1) + 1).shiftDate & " (KW" & activeWeekCell.Value & ") wurde geändert")
      For d = 1 To 5 ' Days of the week
        For s = 1 To 4 ' Shifts
          shiftData.Init weekData.Item(4 * (d - 1) + s).shiftDate, Normal, Right(weekData.Item(4 * (d - 1) + s).shiftCode, 1)
          ' Populate shiftData with a loop
          If d = passedDay Then
            For i = 0 To 32
              tempWorkPlaces(i) = workPlaces(s - 1, i)
            Next i
          Else
            For i = 0 To 32
              tempWorkPlaces(i) = CStr(weekData.Item(CLng(s + (d - 1) * 4)).workPlace(CLng(i)))
            Next i
          End If
          
          shiftData.Populate tempWorkPlaces
          shiftData.shiftStatus = WeekStatus.Custom
          
          ' Now save the SchichtDay (shiftData) object
          shiftList.WeekSave shiftData
        Next s
      Next d
  End Select
  
  ' Reset changes detected
  changesDetectedCell.Value = False
End Sub

Private Sub Load()
  Dim shiftList As New SchichtList
  
  ' Get the week data
  shiftList.WeekRead weekNumber
  Set weekData = shiftList.weekData
End Sub

Private Sub Populate(passedDay As WeekDay)
  ' This function take the day of the week and populate the various
  ' sections of the Worksheet.
  ' Warning: before calling this function, please load and pass the right data!
  
  Dim stanzStartRow, qualiStartRow, lagerStartRow, bueroStartRow, galvaStartRow, instStartRow As Integer
  Dim shiftsStartCol, startWorkPlace, endWorkPlace As Integer
  Dim i, j As Long
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
'  sheet.Unprotect Password:=GAdminPassword
  
  ' Set variable for the view
  shiftStartCol = 7 ' Columnn G
  
  ' Stanzerei
  stanzStartRow = 7
  startWorkPlace = 1
  endWorkPlace = 11
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((stanzStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Qualität
  qualiStartRow = 41
  startWorkPlace = 12
  endWorkPlace = 13
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((qualiStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Lager+Trommlerei
  lagerStartRow = 48
  startWorkPlace = 14
  endWorkPlace = 15
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((lagerStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Büro
  bueroStartRow = 55
  startWorkPlace = 16
  endWorkPlace = 20
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((bueroStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Galvanik
  galvaStartRow = 71
  startWorkPlace = 21
  endWorkPlace = 26
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((galvaStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Instandsetzung
  instaStartRow = 90
  startWorkPlace = 27
  endWorkPlace = 33
  For i = 1 To (endWorkPlace - startWorkPlace + 1)
    For j = 1 To 4
      sheet.Cells((instaStartRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1)).workPlace(i + startWorkPlace - 2))
    Next j
  Next i
  
  ' Reset changes detected
  changesDetectedCell.Value = False
  
  ' Enable screen updating and protection
  Application.ScreenUpdating = True
'  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub ToggleMode()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(losignDataWarningText, vbExclamation + vbYesNo, "ACHTUNG")
    If answer = vbYes Then GoTo ToggleMode Else GoTo NoToggle
  Else
    GoTo ToggleMode
  End If
ToggleMode:
  Select Case weekStatusCell.Value
    Case WeekStatus.Emtpy
      SetNormalMode
    Case WeekStatus.Normal
      SetCustomMode
    Case WeekStatus.Custom
      SetNormalMode
  End Select
NoToggle:
End Sub

Private Sub SetNormalMode()
  weekStatusCell.Value = WeekStatus.Normal
  activeDayCell.Value = WeekDay.Unknown
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
  Populate Monday
  changesDetectedCell.Value = False
End Sub

Private Sub SetCustomMode()
  weekStatusCell.Value = WeekStatus.Custom
  activeDayCell.Value = WeekDay.Monday
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoTrue
  Populate Monday
  changesDetectedCell.Value = False
End Sub

Private Sub SetEmptyMode()
  weekStatusCell.Value = WeekStatus.Emtpy
  activeDayCell.Value = WeekDay.Unknown
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
  EmptyPlan
  changesDetectedCell.Value = False
End Sub

Public Sub GoToNextDay()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(losignDataWarningText, vbExclamation + vbYesNo, "ACHTUNG")
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If
Proceed:
  activeDayCell.Value = activeDayCell.Value + 1
  If activeDayCell.Value > WeekDay.Monday Then prevDayBtn.Visible = msoTrue
  If activeDayCell.Value = WeekDay.Friday Then nextDayBtn.Visible = msoFalse
  Populate activeDayCell.Value
Abort:
End Sub

Public Sub GoToPreviousDay()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(losignDataWarningText, vbExclamation + vbYesNo, "ACHTUNG")
    Debug.Print answer
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If
Proceed:
  activeDayCell.Value = activeDayCell.Value - 1
  If activeDayCell.Value = WeekDay.Monday Then prevDayBtn.Visible = msoFalse
  If activeDayCell.Value < WeekDay.Friday Then nextDayBtn.Visible = msoTrue
  Populate activeDayCell.Value
Abort:
End Sub

Private Sub EmptyPlan()
  Dim shiftStartCell, shiftEndCell As String
  
  shiftStartCell = "G7"
  shiftEndCell = "J110"
  Range(shiftStartCell & ":" & shiftEndCell).ClearContents
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  Sh2_WeekSaveBtn.Enabled = False
  If Not Intersect(Target, Range("G7:J110")) Is Nothing Then ' Monitor changes on Schichtplan
    changesDetectedCell.Value = True
    Sh2_WeekSaveBtn.Enabled = True
  End If
End Sub
