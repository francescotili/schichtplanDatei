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

Private Const workerColumnWidth As Double = 4.5
Private Const abteilungFontSize As Double = 8
Private Const customFont As String = "Franklin Gothic Medium Cond"

Private Sub Class_Initialize()
  Set sheet = ActiveWorkbook.Worksheets(GUrlaubsSheetName)
  Set table = sheet.ListObjects.Item(GUrlaubsTableName)
End Sub

Public Sub WorkerAdd(worker As Mitarbeiter)
  Dim i As Integer
  Dim workerColumn As Double
  Dim workerRow As Double
  
  ' Add new column at last position
  table.ListColumns.Add
  ' Change column header
  With table.HeaderRowRange(1, table.ListColumns.Count)
    .Value = worker.name & " " & worker.surname
    .ColumnWidth = workerColumnWidth
    .Orientation = xlUpward
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With
End Sub

Public Function WorkerSearch(workerPCode As String) As String

End Function

Public Sub GenerateYear()
  Dim startDate, endDate As Date
  Dim numOfDays As Long
  Dim i As Integer
  
  ' To be dynamicized - Format is M/d/yyyy
  startDate = DateSerial(2021, 1, 1)
  endDate = DateSerial(2021, 12, 31)
  
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
End Sub
