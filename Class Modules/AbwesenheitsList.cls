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
    .Font.name = customFont
  End With
  
  ' Add also PCode, Vacation, etc
  workerColumn = table.HeaderRowRange(1, table.ListColumns.Count).Column
  workerRow = table.HeaderRowRange(1, table.ListColumns.Count).Row
  With Cells(workerRow - 1, workerColumn)
    .Value = worker.department
    .Font.Size = abteilungFontSize
  End With
  With Cells(workerRow - 2, workerColumn)
    .Value = worker.personalCode
    .Orientation = xlUpward
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .Font.name = customFont
  End With
  With Cells(workerRow - 8, workerColumn)
    .Value = worker.vacationTotal
    .HorizontalAlignment = xlHAlignCenter
  End With
  
End Sub

Public Function WorkerSearch(workerPCode As String) As String

End Function
