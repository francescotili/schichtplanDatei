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
' Sheet6 (UB Tag)
' --------------------------------------------
Public sheet As Worksheet
Private weekData As Collection

Private dayToLoadCell, activeWeekCell, activeDayCell As Range

Public Sub Worksheet_Activate()
  Set sheet = Sheet6
  
  Set dayToLoadCell = Range("Sh6_DayToLoadCell")
  Set activeWeekCell = Range("Sh6_ActiveWeek")
  Set activeDayCell = Range("Sh6_ActiveDay")
  
  DayDataLoad ' Needed to avoid errors when returning to the sheet
End Sub

Public Sub DayDataLoad()
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  activeDayCell = dayToLoadCell.Value
  
  EmptyPlan
  EmptyAbsences
  
  ' Get the day data
  ShiftDataLoad
  
  ' Get the absence data
  AbsenceDataLoad
  
  Populate

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Private Sub ShiftDataLoad()
  Dim shiftList As New SchichtList
  
  ' Get the day data
  shiftList.DayRead dayToLoadCell.Value
  Set weekData = shiftList.weekData 'Contains only 4 shiftCodes
End Sub

Private Sub AbsenceDataLoad()
  Dim absenceList As New AbwesenheitsList
  Dim workerList As New MitarbeiterList
  Dim weekDates(2) As Date
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
  For i = 1 To 3
    weekDates(i - 1) = dayToLoadCell.Value + i - 1
  Next i
  
  ' Get data and populate
  i = 0
  For Each weekDate In weekDates
    dayVacation = absenceList.DayAbsences(DateSerial(year(weekDate), Month(weekDate), day(weekDate)))
    
    Select Case UBound(dayVacation)
    Case Is > 65
      moreVacation = UBound(dayVacation) - 64
      ReDim Preserve dayVacation(64)
      ReDim workersNames(65)
      
      For j = 0 To UBound(dayVacation)
        workerList.Search dayVacation(j)
        workersNames(j) = workerList.worker.visName
      Next j
      workersNames(64) = StringFormat(str_sh6_moreAbsence, moreVacation)
      
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

Private Sub Populate()
  ' This function populate the various sections of the Worksheet.
  ' Warning: before calling this function, please load and pass the right data!
  
  Dim startRow, startCol As Integer
  Dim numOfWorkplaces As Integer
  Dim i, j As Long
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Set variable for the view
  startCol = 9 ' Column I
  startRow = 6
  numOfWorkplaces = Sheet1.ListObjects(GShiftsTableName).DataBodyRange.Columns.Count - 3
  For i = 1 To numOfWorkplaces
    For j = 1 To 4
      If Len(CStr(weekData.Item(j).workPlace(i - 1))) <> 0 Then
        If Not IsWorkerAbsent(CStr(weekData.Item(j).workPlace(i - 1))) Then
          sheet.Cells((startRow + 2 * (i - 1)), (startCol + (j - 1))) = CStr(weekData.Item(j).workPlace(i - 1))
        End If
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

Private Function IsWorkerAbsent(passedName As String) As Boolean
  Dim absentWorkers() As Variant
  Dim result As Boolean
  result = False
  
  ' Get all the absences
  absentWorkers = Range("N6:N71")
  
  ' Check if passedName is found
  For i = 1 To UBound(absentWorkers)
    If absentWorkers(i, 1) = passedName Then
      result = True
      Exit For
    End If
  Next i
  
  ' Return result
  IsWorkerAbsent = result
End Function

Public Sub NextDayLoad()
  If year(dayToLoadCell.Value) = year(dayToLoadCell.Value + 1) Then
    dayToLoadCell.Value = dayToLoadCell.Value + 1
    DayDataLoad
  Else
    MsgBox str_sh6_errNextYear
  End If
End Sub

Public Sub PreviousDayLoad()
  If year(dayToLoadCell.Value) = year(dayToLoadCell.Value - 1) Then
    dayToLoadCell.Value = dayToLoadCell.Value - 1
    DayDataLoad
  Else
    MsgBox str_sh6_errPrevYear
  End If
End Sub

Public Sub GoToToday()
  If year(dayToLoadCell.Value) = year(Date) Then
    dayToLoadCell.Value = Date
    DayDataLoad
  Else
    MsgBox str_sh6_errTodayYear
  End If
End Sub
