VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbwesenheitsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sheet As Worksheet
Private table As ListObject

Private Const workerColumnWidth As Double = 4
Private Const startWorkerColumn As Integer = 2

Private Sub Class_Initialize()
  Set sheet = Sheet8
  Set table = sheet.ListObjects.Item(GAbsencesTableName)
End Sub

Public Sub WorkerAdd(worker As Mitarbeiter)
  Dim i As Integer
  Dim workerColumn As Double
  Dim workerRow As Double
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Add new column at last position
  table.ListColumns.Add
  ' Change column header
  With table.HeaderRowRange(1, table.ListColumns.Count)
    .Value = worker.personalCode
    .ColumnWidth = workerColumnWidth
    .Orientation = xlUpward
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub GenerateYear(Year As Integer)
  Dim startDate, endDate As Date
  Dim numOfDays As Long
  Dim i As Integer
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Generate start und end date from form
  startDate = DateSerial(Year, 1, 1)
  endDate = DateSerial(Year, 12, 31)
  
  ' Clear vacation table data
  With table
    If Not .DataBodyRange Is Nothing Then
      .DataBodyRange.Delete
    End If
  End With
  
  ' Populate the rows with all the day of the year
  numOfDays = DateDiff("d", startDate, endDate)
  For i = 0 To numOfDays
    table.ListRows.Add
    table.DataBodyRange((i + 1), 1).Value = startDate + i
  Next i
  
  ' Save history entry
  saveHistory (str_resetAbsenceStatus)

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Public Sub WorkerDelete(Optional personalCode As String = "null")
  Dim i As Double
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  If personalCode = "null" Then ' Delete all workers column
    If table.ListColumns.Count < startWorkerColumn Then
      For i = startWorkerColumn To table.ListColumns.Count
        table.ListColumns(startWorkerColumn).Delete
      Next i
    End If
  Else ' Search for user and delete
    For i = startWorkerColumn To table.DataBodyRange.Columns.Count
      If table.HeaderRowRange(1, i).Value = personalCode Then
        table.ListColumns(i).Delete
      End If
    Next i
  End If

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Public Property Get Mitarbeiter(personalCode As String) As String()
  Dim response() As String
  Dim numOfDays As Integer
  Dim workerCol As Integer
  Dim i As Long
  
  ' Search column of worker
  For i = 2 To table.HeaderRowRange.Columns.Count
    If table.HeaderRowRange(1, i) = personalCode Then
      workerCol = i
    End If
  Next i
  
  ' Get worker absence days
  ReDim response(table.DataBodyRange.Rows.Count - 1)
  For i = 1 To table.DataBodyRange.Rows.Count
    response(i - 1) = table.DataBodyRange(i, workerCol)
  Next i
  
  Mitarbeiter = response
End Property

Public Property Get DayAbsences(passedDate As Date) As String()
  Dim response() As String
  Dim numOfAbsence As Integer
  Dim i, j, dayRow As Long
  
  ' Find the row
  For i = 1 To table.DataBodyRange.Rows.Count
    If table.DataBodyRange(i, 1).Value = passedDate Then
      dayRow = i
      Exit For
    End If
  Next i
  
  ' Retrieve data and build the response
  numOfAbsence = 0
  For i = 2 To table.DataBodyRange.Columns.Count
    If table.DataBodyRange(dayRow, i).Value <> "" Then
      numOfAbsence = numOfAbsence + 1
      ReDim Preserve response(numOfAbsence - 1)
      response(numOfAbsence - 1) = CStr(table.HeaderRowRange(1, i).Value)
    End If
  Next i
  
  If numOfAbsence = 0 Then
    ReDim response(0)
    response(0) = "0"
  End If
  
  DayAbsences = response
End Property

Public Property Get AllMitarbeiter() As String()
  Dim response() As String
  Dim numOfWorkers As Integer
  Dim i As Long
  
  numOfWorkers = table.DataBodyRange.Columns.Count - 1
  ReDim response(numOfWorkers - 1)
  For i = 2 To table.DataBodyRange.Columns.Count
    response(i - 2) = table.HeaderRowRange(1, i)
  Next i
  
  AllMitarbeiter = response
End Property

Public Sub Save(passedCode As String, absenceData() As String)
  Dim workerCol As Integer
  Dim i, j As Long
  
  ' Search column of worker
  For i = 2 To table.HeaderRowRange.Columns.Count
    If table.HeaderRowRange(1, i) = passedCode Then
      workerCol = i
      
      ' Save data
      For j = LBound(absenceData) To UBound(absenceData)
        If Len(absenceData(j)) <> 0 Then
          table.DataBodyRange(j + 1, workerCol) = absenceData(j)
        Else
          If Len(table.DataBodyRange(j + 1, workerCol).Value) <> 0 Then
            table.DataBodyRange(j + 1, workerCol).ClearContents
          End If
        End If
      Next j
    End If
  Next i
End Sub

Public Sub ProtectionOFF()
  sheet.Unprotect Password:=GAdminPassword
End Sub

Public Sub ProtectionON()
  sheet.Protect Password:=GAdminPassword
End Sub
