' --------------------------------------------
' WEEK BEARBEITEN PUBLIC FUNCTIONS
' Manage methods for getting all the data for
' Sheet2 (Bearbeiten)
' --------------------------------------------
Public sheet As Worksheet
Private weekData As Collection
Private weekNumber As Double

Private nextDayBtn, prevDayBtn As Shape
Private weekToLoadCell, weekStatusCell, activeDayCell, actualYearCell, activeWeekCell, changesDetectedCell, lastUsedCell, maxKW As Range

Private Sub Sh2_WeekSaveBtn_Click()
  WeekDataSave
  Range(lastUsedCell.Value).Select
End Sub

Public Sub Worksheet_Activate()
  Set sheet = Sheet2
  
  Set nextDayBtn = Me.sheet.Shapes("Sh2_NextDayBtn")
  Set prevDayBtn = Me.sheet.Shapes("Sh2_PreviousDayBtn")
  
  Set weekToLoadCell = Range("Sh2_WeekToLoadCell")
  Set weekStatusCell = Range("Sh2_WeekStatusCell")
  Set activeDayCell = Range("Sh2_ActiveDay")
  Set activeWeekCell = Range("Sh2_ActiveWeek")
  Set actualYearCell = Sheet3.Range("Global_ActualYear")
  Set changesDetectedCell = Range("Sh2_ChangesDetected")
  Set lastUsedCell = Range("Sh2_LastCell")
  Set maxKW = Sheet3.Range("Global_MaxKW")
  
  WeekDataLoad ' Needed to avoid errors when returning to the sheet
End Sub

Public Sub WeekDataLoad()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(str_sh2_losingDataWarn, vbExclamation + vbYesNo, str_warnBoxTitle)
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If

Proceed:
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Get the week data
  weekNumber = CInt(weekToLoadCell.Value)
  activeWeekCell.Value = weekNumber
  Load
  
  ' Worflow based on week status (read at monday)
  Select Case weekData.Item(WeekDays.Monday).shiftStatus
    Case WeekStatus.Emtpy
      SetEmptyMode
    Case WeekStatus.Normal
      SetNormalMode
    Case WeekStatus.Custom
      SetCustomMode
  End Select
  
  ' Get the absence data
  AbsenceDataLoad

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
Abort:
End Sub

Private Sub AbsenceDataLoad()
  Dim absenceList As New AbwesenheitsList
  Dim workerList As New MitarbeiterList
  Dim weekDates(4) As Date
  Dim dayVacation() As String
  Dim workersNames() As String
  Dim startCol, startRow As Integer
  Dim i, j As Long

  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Set coordinates
  startCol = 12
  startRow = 7
  
  ' Get the days
  For i = 1 To 5
    weekDates(i - 1) = 7 * (activeWeekCell.Value - 1) + DateSerial(actualYearCell.Value, 1, 4) - WeekDay(DateSerial(actualYearCell.Value, 1, 4), 2) + i
  Next i
  
  ' Get data and populate
  i = 0
  For Each WeekDate In weekDates
    dayVacation = absenceList.DayAbsences(DateSerial(Year(WeekDate), Month(WeekDate), day(WeekDate)))
    
    Select Case UBound(dayVacation)
    Case Is > 98
      moreVacation = UBound(dayVacation) - 97
      ReDim Preserve dayVacation(97)
      ReDim workersNames(98)

      For j = 0 To UBound(dayVacation)
        workerList.Search dayVacation(j)
        workersNames(j) = workerList.worker.visName
      Next j
      workersNames(98) = StringFormat(str_sh2_moreAbsence, moreVacation)

      For j = 0 To UBound(workersNames)
        sheet.Cells(startRow + j, startCol + i) = workersNames(j)
      Next j
    Case Is > 0
      ReDim workersNames(UBound(dayVacation))

      For j = 0 To UBound(dayVacation)
        workerList.Search dayVacation(j)
        workersNames(j) = workerList.worker.visName
      Next j
      
      For j = 0 To UBound(workersNames)
        sheet.Cells(startRow + j, startCol + i) = workersNames(j)
      Next j
    Case Is = 0
      If dayVacation(0) <> "0" Then
        workerList.Search dayVacation(0)
        sheet.Cells(startRow, startCol + i) = workerList.worker.visName
      End If
    End Select
    
    i = i + 1
  Next

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
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
    msgAnswer = MsgBox(str_sh2_savingWarn, vbQuestion + vbYesNo, "Fortfahren?")
    If msgAnswer = vbYes Then
      Sheet1.Unprotect Password:=GAdminPassword
      Save WeekDays.Unknown
      Sheet1.Protect Password:=GAdminPassword
    End If
  End Select
End Sub

Public Sub Save(passedDay As WeekDays)
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
    startRow = 7
    For i = 1 To 33
      workPlaces(j - 1, i - 1) = CStr(sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))).Value)
    Next i
  Next j
  
  ' Now save the different days based on the WeekModus
  Dim d, s As Long
  Dim tempWorkPlaces(32) As String
  Select Case activeDayCell.Value
    Case WeekDays.Unknown ' Normal Modus
      saveHistory StringFormat(str_sh2_historySPWU, activeWeekCell.Value)
      
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
      saveHistory StringFormat(str_sh2_historySPWK, weekData.Item(4 * (passedDay - 1) + 1).shiftDate, activeWeekCell.Value)
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

Private Sub Populate(passedDay As WeekDays)
  ' This function take the day of the week and populate the various
  ' sections of the Worksheet.
  ' Warning: before calling this function, please load and pass the right data!
  
  Dim stanzStartRow, qualiStartRow, lagerStartRow, bueroStartRow, galvaStartRow, instStartRow As Integer
  Dim shiftsStartCol, startWorkPlace, endWorkPlace As Integer
  Dim i, j As Long
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  EmptyPlan
  EmptyAbsences
  
  ' Set variable for the view
  shiftStartCol = 7 ' Columnn G
  
  ' Stanzerei
  startRow = 7
  For i = 1 To 33
    For j = 1 To 4
      If Len(CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i - 1))) <> 0 Then
        sheet.Cells((startRow + 3 * (i - 1)), (shiftStartCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i - 1))
      End If
    Next j
  Next i
  
  ' Reset changes detected
  changesDetectedCell.Value = False
  
  ' Enable screen updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub ToggleMode()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(str_sh2_losingDataWarn, vbExclamation + vbYesNo, str_warnBoxTitle)
    If answer = vbYes Then GoTo ToggleMode Else GoTo NoToggle
  Else
    GoTo ToggleMode
  End If
ToggleMode:
  Select Case weekStatusCell.Value
    Case WeekStatus.Emtpy
      SetCustomMode
    Case WeekStatus.Normal
      SetCustomMode
    Case WeekStatus.Custom
      SetNormalMode
  End Select
NoToggle:
End Sub

Private Sub SetNormalMode()
  weekStatusCell.Value = WeekStatus.Normal
  activeDayCell.Value = WeekDays.Unknown
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
  Populate Monday
  changesDetectedCell.Value = False
End Sub

Private Sub SetCustomMode()
  weekStatusCell.Value = WeekStatus.Custom
  activeDayCell.Value = WeekDays.Monday
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoTrue
  Populate Monday
  changesDetectedCell.Value = False
End Sub

Private Sub SetEmptyMode()
  weekStatusCell.Value = WeekStatus.Emtpy
  activeDayCell.Value = WeekDays.Unknown
  prevDayBtn.Visible = msoFalse
  nextDayBtn.Visible = msoFalse
  EmptyPlan
  EmptyAbsences
  changesDetectedCell.Value = False
End Sub

Public Sub GoToNextDay()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(str_sh2_losingDataWarn, vbExclamation + vbYesNo, str_warnBoxTitle)
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If
Proceed:
  activeDayCell.Value = activeDayCell.Value + 1
  If activeDayCell.Value > WeekDays.Monday Then prevDayBtn.Visible = msoTrue
  If activeDayCell.Value = WeekDays.Friday Then nextDayBtn.Visible = msoFalse
  Populate activeDayCell.Value
Abort:
End Sub

Public Sub GoToPreviousDay()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(str_sh2_losingDataWarn, vbExclamation + vbYesNo, str_warnBoxTitle)
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If
Proceed:
  activeDayCell.Value = activeDayCell.Value - 1
  If activeDayCell.Value = WeekDays.Monday Then prevDayBtn.Visible = msoFalse
  If activeDayCell.Value < WeekDays.Friday Then nextDayBtn.Visible = msoTrue
  Populate activeDayCell.Value
Abort:
End Sub

Private Sub EmptyPlan()
  Dim startCell, endCell As String
  
  startCell = "G7"
  endCell = "J105"
  Range(startCell & ":" & endCell).ClearContents
End Sub

Private Sub EmptyAbsences()
  Dim startCell, endCell As String
  
  startCell = "L7"
  endCell = "P105"
  Range(startCell & ":" & endCell).ClearContents
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  'Sh2_WeekSaveBtn.Enabled = False
  If Not Intersect(Target, Range("G7:J110")) Is Nothing Then ' Monitor changes on Schichtplan
    changesDetectedCell.Value = True
    'Sh2_WeekSaveBtn.Enabled = True
  End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  Dim ComboObj As OLEObject
  Dim ComboListArr As Variant
  
  Select Case Application.CutCopyMode
    Case Is = False
      ' Initialize Combo Object for Autofill
      Set ComboObj = sheet.OLEObjects("TempCombo")
      With ComboObj
        .Visible = False
        .ListFillRange = "=Personal!" & Sheet9.ListObjects(GWorkerTableName).ListColumns(4).DataBodyRange.Address
        .LinkedCell = ""
      End With
      
      ' Move and visualize ComboBox if needed: only in worker range and if cell is empty
      If Not Application.Intersect(Target, Range("G7:J110")) Is Nothing Then
        If Target.Areas.Count = 1 _
        And Target.Columns.Count = 1 _
        And Target.Rows.Count = 3 _
        And Target.Cells(1).Value = "" Then
          lastUsedCell.Value = Target.Address
          sheet.Unprotect Password:=GAdminPassword
          
          With ComboObj
            .Visible = True
            .Left = Target.Left + 1
            .Top = Target.Top + 1
            .Width = Target.Width - 1
            .Height = Target.Height - 1
            .LinkedCell = Target.Address
          End With
    
          ComboObj.Activate
          Me.TempCombo.DropDown
          
          sheet.Protect Password:=GAdminPassword
        End If
      End If
    Case Is = xlCopy
      ' do nothing
    Case Is = xlCut
      MsgBox str_sh2_errCutMode
      Application.CutCopyMode = False ' Cancel clipboard and block cut operation
  End Select
End Sub

Private Sub TempCombo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Select Case KeyCode
    Case 9 ' Tab
      changesDetectedCell.Value = True
      Application.ActiveCell.Offset(0, 1).Activate
    Case 13 ' Enter
      changesDetectedCell.Value = True
      Application.ActiveCell.Offset(1, 0).Activate
    Case 27 ' ESC
      With sheet.OLEObjects("TempCombo")
        .Object.Value = ""
        .Visible = False
      End With
      Application.ActiveCell.Activate
  End Select
End Sub

Private Sub TempCombo_Click()
  changesDetectedCell.Value = True
End Sub

Public Sub MasterWeekImport()
  Dim startCol, startRow As Integer
  
  If weekStatusCell.Value <> 0 Then
    answer = MsgBox(str_sh2_losingDataImport, vbExclamation + vbYesNo, str_warnBoxTitle)
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If
  
Proceed:
  startRowMW = 7
  startRow = 7
  startCol = 7
  
  Application.ScreenUpdating = False
  EmptyPlan
  
  If activeWeekCell.Value Mod 2 = 0 Then
    ' Even (gerade)
    startColMW = 7
  Else
    ' Odd (ungerade)
    startColMW = 14
  End If
    
  For j = 1 To 4
    For i = 1 To 33
      sheet.Cells(startRow + 3 * (i - 1), startCol + j - 1).Value = CStr(Sheet11.Cells(startRowMW + 3 * (i - 1), startColMW + j - 1).Value)
    Next i
  Next j
  
  Application.ScreenUpdating = True

Abort:
End Sub

Public Sub NextWeekLoad()
  If (weekToLoadCell.Value + 1) <= maxKW Then
    weekToLoadCell.Value = weekToLoadCell.Value + 1
    WeekDataLoad
  Else
    MsgBox str_sh2_errNextYear
  End If
End Sub

Public Sub PreviousWeekLoad()
  If (weekToLoadCell.Value - 1) <> 0 Then
    weekToLoadCell.Value = weekToLoadCell.Value - 1
    WeekDataLoad
  Else
    MsgBox str_sh2_errPrevYear
  End If
End Sub
