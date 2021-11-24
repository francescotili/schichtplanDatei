VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SchichtList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------------------
' SCHICHT LIST CLASS
' Manage methods for getting all the shifts
' or save/search/delete week of shifts
' --------------------------------------------

Public data As Collection
Private sheet As Worksheet
Private table As ListObject

Private Const customFont As String = "Franklin Gothic Medium Cond"

Private Sub Class_Initialize()
  Set sheet = ActiveWorkbook.Worksheets(GShiftsSheetName)
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

  Set tempShift = New SchichtDay
  tempShift.Init DateSerial(2021, 11, 23), "F"

  If Not table.DataBodyRange Is Nothing Then
    shiftsData = table.DataBodyRange

    ' Parse data in Shifts objects
    For i = LBound(shiftsData) To UBound(shiftsData)
      ' Initialize SchichtDay object with read data
      Set tempShift = New SchichtDay
      tempShift.Init DateValue(shiftsData(i, 1)), CStr(shiftsData(i, 2))
      
      ' Create array of workplaces for that day
      For j = 4 To 36 ' Column for every workplace
        workPlace(j - 4) = CStr(shiftsData(i, j))
      Next j
      tempShift.Populate workPlace
      
      data.Add tempShift
    Next i
    
    'Debug.Print data.Item(1).workPlace(2)
  End If
End Sub

'Public Sub WeekSave()
'  ' Save or modify the passed week
'
'End Sub

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
      table.DataBodyRange((4 * i + j), 2) = Right(shiftCodes(j - 1), 1)
      table.DataBodyRange((4 * i + j), 3) = shiftCodes(j - 1)
    Next j
  Next i
  
  ' Save history entry
  saveHistory ("DER SCHICHTPLAN WURDE ZURÜCKGESETZT")
  
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
      year(inputDate) & _
      Format(Month(inputDate), "00") & _
      Format(Day(inputDate), "00") & _
      shiftNames(i)
  Next i
  
  generateShiftsCode = shiftCodes
End Function
