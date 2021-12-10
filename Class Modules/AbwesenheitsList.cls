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

Public Sub GenerateYear(year As Integer)
  Dim startDate, endDate As Date
  Dim numOfDays As Long
  Dim i As Integer
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Generate start und end date from form
  startDate = DateSerial(year, 1, 1)
  endDate = DateSerial(year, 12, 31)
  
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
  saveHistory ("DER ABWESENHEITPLAN WURDE ZURÜCKGESETZT")

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

Property Get GetMitarbeiter(personalCode As String) As String()
  
End Property

Public Property Get AllMitarbeiter() As String()
  Dim response() As String
  Dim numOfWorkers As Integer
  Dim i As Long
  
  numOfWorkers = table.DataBodyRange.Columns.Count - 2
  ReDim response(numOfWorkers)
  For i = 2 To table.DataBodyRange.Columns.Count
    response(i - 2) = table.HeaderRowRange(1, i)
  Next i
  
  AllMitarbeiter = response
End Property
