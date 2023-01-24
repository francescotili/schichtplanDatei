VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --------------------------------------------
' WEEK VIEW FUNCTIONS
' Manage methods for getting all the data for
' Sheet7 (UB Woche)
' --------------------------------------------
Public sheet As Worksheet
Private weekData As Collection
Private weekNumber As Double

Private weekToLoadCell, activeWeekCell, actualYearCell, weekStatusCell, weekStatusLabel, maxKW As Range

Public Sub Worksheet_Activate()
  Set sheet = Sheet7
  
  Set weekToLoadCell = Range("Sh7_WeekToLoadCell")
  Set activeWeekCell = Range("Sh7_ActiveWeek")
  Set weekStatusCell = Range("Sh7_WeekStatusCell")
  Set actualYearCell = Sheet3.Range("Global_ActualYear")
  Set weekStatusLabel = Range("Sh7_WeekStatusLabel")
  Set maxKW = Sheet3.Range("Global_MaxKW")
  
  WeekDataLoad ' Needed to avoid errors when returning to the sheet
End Sub

Public Sub WeekDataLoad()
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Get the week data
  weekNumber = CInt(weekToLoadCell.Value)
  activeWeekCell.Value = weekNumber
  ShiftDataLoad
  
  ' Workflow based on week status (read at monday)
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
  startCol = 14
  startRow = 6
  
  ' Get the days
  For i = 1 To 5
    weekDates(i - 1) = 7 * (activeWeekCell.Value - 1) + DateSerial(actualYearCell.Value, 1, 4) - WeekDay(DateSerial(actualYearCell.Value, 1, 4), 2) + i
  Next i
  
  ' Get data and populate
  i = 0
  For Each WeekDate In weekDates
    dayVacation = absenceList.DayAbsences(DateSerial(Year(WeekDate), Month(WeekDate), day(WeekDate)))
    
    Select Case UBound(dayVacation)
    Case Is > 65
      moreVacation = UBound(dayVacation) - 64
      ReDim Preserve dayVacation(64)
      ReDim workersNames(65)

      For j = 0 To UBound(dayVacation)
        workerList.Search dayVacation(j)
        workersNames(j) = workerList.worker.visName
      Next j
      workersNames(64) = StringFormat(str_sh7_moreAbsence, moreVacation)

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

Public Sub ShiftDataLoad()
  Dim shiftList As New SchichtList
  
  ' Get the week data
  shiftList.WeekRead weekNumber
  Set weekData = shiftList.weekData
End Sub

Private Sub Populate(passedDay As WeekDays)
  ' This function take the day of the week and populate the various
  ' sections of the Worksheet.
  ' Warning: before calling this function, please load and pass the right data!
  
  Dim startRow, startCol As Integer
  Dim numOfWorkplaces As Integer
  Dim i, j As Long
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  EmptyPlan
  EmptyAbsences
  
  ' Set variable for the view
  startCol = 9 ' Columnn I
  startRow = 6
  numOfWorkplaces = Sheet1.ListObjects(GShiftsTableName).DataBodyRange.Columns.Count - 3
  For i = 1 To numOfWorkplaces
    For j = 1 To 4
      If Len(CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i - 1))) <> 0 Then
        sheet.Cells((startRow + 2 * (i - 1)), (startCol + (j - 1))) = CStr(weekData.Item(j + (passedDay - 1) * 4).workPlace(i - 1))
      End If
    Next j
  Next i
  
  ' Enable screen updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
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
  endCell = "R71"
  Range(startCell & ":" & endCell).ClearContents
End Sub

Private Sub SetNormalMode()
  weekStatusCell.Value = WeekStatus.Normal
  weekStatusLabel.Value = str_sh7_weekStatusLabelNM
  Populate Monday
End Sub

Private Sub SetCustomMode()
  weekStatusCell.Value = WeekStatus.Custom
  weekStatusLabel.Value = str_sh7_weekStatusLabelCM
  EmptyPlan
  EmptyAbsences
End Sub

Private Sub SetEmptyMode()
  weekStatusCell.Value = WeekStatus.Emtpy
  weekStatusLabel.Value = str_sh7_weekStatusLabelEM
  EmptyPlan
  EmptyAbsences
End Sub

Public Sub NextWeekLoad()
  If (weekToLoadCell.Value + 1) <= maxKW Then
    weekToLoadCell.Value = weekToLoadCell.Value + 1
    WeekDataLoad
  Else
    MsgBox str_sh7_errNextYear
  End If
End Sub

Public Sub PreviousWeekLoad()
  If (weekToLoadCell.Value - 1) <> 0 Then
    weekToLoadCell.Value = weekToLoadCell.Value - 1
    WeekDataLoad
  Else
    MsgBox str_sh7_errPrevYear
  End If
End Sub

Public Sub GoToActualWeek()
  If actualYearCell = Year(Date) Then
    weekToLoadCell.Value = WorksheetFunction.WeekNum(Now, 21)
    WeekDataLoad
  Else
    MsgBox str_sh7_errTodayYear
  End If
End Sub

Public Sub ModifyWeek()
  Sheet2.Range("Sh2_WeekToLoadCell").Value = activeWeekCell.Value
  Sheet2.Activate
End Sub
