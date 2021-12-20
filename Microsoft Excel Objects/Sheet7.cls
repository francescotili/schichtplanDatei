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
' Sheet7 (ÜB Woche)
' --------------------------------------------
Public sheet As Worksheet
Private weekNumber As Double

Private weekToLoadCell, activeWeekCell, actualYearCell As Range

Public Sub Worksheet_Activate()
  Set sheet = Sheet7
  
  Set weekToLoadCell = Range("Sh7_WeekToLoadCell")
  Set activeWeekCell = Range("Sh7_ActiveWeek")
  Set weekStatusCell = Range("Sh7_WeekStatusCell")
  Set actualYearCell = Sheet3.Range("Global_ActualYear")
  
  WeekDataLoad ' Needed to avoid errors when returning to the sheet
End Sub

Public Sub WeekDataLoad()
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Get the week data
  weekNumber = CInt(weekToLoadCell.Value)
'  activeWeekCell.Value = weekNumber
'  Load
  
  ' Workflow based on week status (read at monday)
'  TODO!

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
    weekDates(i - 1) = DateAdd("ww", activeWeekCell.Value - 1, DateSerial(actualYearCell.Value, 1, 2 + i))
  Next i
  
  ' Get data and populate
  i = 0
  For Each weekDate In weekDates
    dayVacation = absenceList.DayAbsences(DateSerial(year(weekDate), Month(weekDate), day(weekDate)))
    
    Select Case UBound(dayVacation)
    Case Is > 32
      moreVacation = UBound(dayVacation) - 31
      ReDim Preserve dayVacation(31)
      ReDim workersNames(32)

      For j = 0 To UBound(dayVacation)
        workerList.Search dayVacation(j)
        workersNames(j) = workerList.worker.visName
      Next j
      workersNames(32) = "und " & moreVacation & " weitere"

      For j = 0 To UBound(workersNames)
        sheet.Cells(startRow + j, startCol + i) = workersNames(j)
      Next j
    Case Is > 0
      ReDim workersNames(UBound(dayVacation))

      For j = 0 To UBound(dayVacation)
        workerList.Search dayVacation(j)
        workersNames(j) = workerList.worker.visName
      Next j
      
      Debug.Print UBound(workersNames)
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
