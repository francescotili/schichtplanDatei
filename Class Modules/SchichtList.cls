' --------------------------------------------
' SCHICHT LIST CLASS
' Manage methods for getting all the shifts
' or save/search/delete week of shifts
' --------------------------------------------

Public data As Collection
Public weekData As Collection
Private sheet As Worksheet
Private table As ListObject

Private Const customFont As String = "Franklin Gothic Medium Cond"

Private Sub Class_Initialize()
  Set sheet = Sheet1
  Set table = sheet.ListObjects.Item(GShiftsTableName)
End Sub

Public Sub Load()
  ' Get all the shifts for all the days and build a
  ' Collection of shifts objects

  Dim shiftsData As Variant
  Set data = New Collection
  Dim i, j As Long
  Dim tempShift As SchichtDay
  Dim workPlace(32) As String

  If Not table.DataBodyRange Is Nothing Then
    shiftsData = table.DataBodyRange

    ' Parse data in Shifts objects
    For i = LBound(shiftsData) To UBound(shiftsData)
      ' Initialize SchichtDay object with read data
      Set tempShift = New SchichtDay
      tempShift.Init DateValue(shiftsData(i, 1)), CLng(shiftsData(i, 2)), Right(CStr(shiftsData(i, 3)), 1)
      
      ' Create array of workplaces for that day
      For j = 4 To 36 ' Column for every workplace
        workPlace(j - 4) = CStr(shiftsData(i, j))
      Next j
      tempShift.Populate workPlace
      
      data.Add tempShift
    Next i
  End If
End Sub

Public Sub WeekSave(passedData As SchichtDay)
  Dim i, j, rowToSave As Long
  
  ' Load the data
  If Not table.DataBodyRange Is Nothing Then
    shiftsData = table.DataBodyRange
    
    ' Search for the shiftCode to save
    For i = LBound(shiftsData) To UBound(shiftsData)
      If StrComp(shiftsData(i, 3), passedData.shiftCode, vbTextCompare) = 0 Then
        rowToSave = i
        Exit For
      End If
    Next i
    
    ' Save the row
    table.DataBodyRange(rowToSave, 2).Value = passedData.shiftStatus
    For j = 4 To 36 ' Columns for every workplace
      If Len(passedData.workPlace(j - 4)) <> 0 Then
        table.DataBodyRange(rowToSave, j).Value = passedData.workPlace(j - 4)
      Else
        If Len(table.DataBodyRange(rowToSave, j).Value) <> 0 Then
          table.DataBodyRange(rowToSave, j).ClearContents
        End If
      End If
    Next j
  End If
End Sub

Public Sub WeekRead(weekNumber As Double)
  Dim startDate, endDate As Date
  Dim numOfDay As Long
  Dim tempShiftCodes() As String
  Set shiftCodes = CreateObject("System.Collections.ArrayList")
  Dim shiftYear As Integer
  Dim i, j As Double
  Set weekData = New Collection
  
  ' Load shifts data from table
  Me.Load
  
  ' Define the year, reading the first date of shifts database
  shiftYear = Year(data.Item(1).shiftDate)
  
  ' Generate start and end dates from passed week number
  ' From https://www.snb-vba.eu/VBA_ISO_weeknummer_en.html
  startDate = 7 * (weekNumber - 1) + DateSerial(shiftYear, 1, 4) - WeekDay(DateSerial(shiftYear, 1, 4), 2) + 1
  endDate = 7 * (weekNumber - 1) + DateSerial(shiftYear, 1, 4) - WeekDay(DateSerial(shiftYear, 1, 4), 2) + 5
  
  ' Generate shiftCodes to search and read
  numOfDay = DateDiff("d", startDate, endDate)
  For i = 0 To numOfDay
    tempShiftCodes = generateShiftsCode(startDate + i)
    For j = LBound(tempShiftCodes) To UBound(tempShiftCodes)
      shiftCodes.Add tempShiftCodes(j)
    Next j
  Next i
  
  ' Filter the data with the requested shifts and populate weekData collection
  For i = 1 To data.Count
    For j = 1 To shiftCodes.Count
      If shiftCodes(j - 1) = data.Item(i).shiftCode Then
        weekData.Add data.Item(i)
      End If
    Next j
  Next i
End Sub

Public Sub DayRead(dayDate As Date)
  Dim tempShiftCodes() As String
  Set shiftCodes = CreateObject("System.Collections.ArrayList")
  Dim shiftYear As Integer
  Dim i, j As Double
  Set weekData = New Collection
  
  ' Load shifts data from table
  Me.Load
  
  ' Define the year, reading the first date of shifts database
  shiftYear = Year(dayDate)
  
  ' Generate shiftCodes to search and read
  tempShiftCodes = generateShiftsCode(dayDate)
  For i = LBound(tempShiftCodes) To UBound(tempShiftCodes)
    shiftCodes.Add tempShiftCodes(i)
  Next i
  
  ' Filter the data with the requested shifts and populate weekData collection
  For i = 1 To data.Count
    For j = 1 To shiftCodes.Count
      If shiftCodes(j - 1) = data.Item(i).shiftCode Then
        weekData.Add data.Item(i)
      End If
    Next j
  Next i
End Sub

Public Sub GenerateYear(passedYear As Integer)
  Dim startDate, endDate As Date
  Dim numOfDays As Long
  Dim shiftCodes() As String
  Dim i, j As Double
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Generate start und end date from form
  startDate = DateSerial(passedYear, 1, 1)
  endDate = DateSerial(passedYear, 12, 31)
  
  ' Clear shifts table data
  With table
    If Not .DataBodyRange Is Nothing Then
      .DataBodyRange.Delete
    End If
  End With
  
  ' Populate the rows with all the day of the year
  numOfDays = DateDiff("d", startDate, endDate)
  For i = 0 To numOfDays
    shiftCodes = generateShiftsCode(startDate + i)
    For j = 1 To 4 ' For the 4 shifts of the day
      table.ListRows.Add
      table.DataBodyRange((4 * i + j), 1).Value = startDate + i
      table.DataBodyRange((4 * i + j), 2) = WeekStatus.Emtpy
      table.DataBodyRange((4 * i + j), 3) = shiftCodes(j - 1)
    Next j
  Next i
  
  ' Save history entry
  saveHistory (str_resetShiftplan)
  
  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Private Function generateShiftsCode(inputDate As Date) As String()
  Dim shiftNames(3) As String
  Dim shiftCodes(3) As String
  Dim i As Double

  ' Define names of shifts
  shiftNames(0) = "X"
  shiftNames(1) = "F"
  shiftNames(2) = "S"
  shiftNames(3) = "N"
  
  For i = 0 To 3
    shiftCodes(i) = _
      Year(inputDate) & _
      Format(Month(inputDate), "00") & _
      Format(day(inputDate), "00") & _
      shiftNames(i)
  Next i
  
  generateShiftsCode = shiftCodes
End Function


