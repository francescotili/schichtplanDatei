VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' --------------------------------------------
' ABSENCE MANAGER PUBLIC FUNCTIONS
' Manage methods for getting all the data for
' Sheet5 (Abwesenheitsplan)
' --------------------------------------------
Public sheet As Worksheet
Private table As ListObject

Private changesDetectedCell As Range

' Table styles
Private Const workerColumnWidth As Double = 5
Private Const startWorkerColumn As Integer = 6
Private Const holidayRow As Integer = 3
Private Const personalCodeRow As Integer = 9
Private Const departmentRow As Integer = 10

' Strings
Private Const losignDataWarningText = "Im Abwesenheitplan wurden nicht gespeicherte änderungen erkannt!" & vbCrLf & "Wenn Sie fortfahren, gehen die verloren. Fortfahren?"

Public Sub Worksheet_Activate()
  Set sheet = Sheet5
  Set table = sheet.ListObjects.Item(GAbsencesModifyTableName)
  
  Set changesDetectedCell = Range("Sh5_ChangesDetected")
  
  'AbsenceDataLoad
  TableInitialize
End Sub

Private Sub AbsenceDataLoad()
  If changesDetectedCell.Value = True Then
    answer = MsgBox(losignDataWarningText, vbExclamation + vbYesNo, "ACHTUNG")
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If

Proceed:
  Dim absenceList As New AbwesenheitsList
  Dim workersCodes() As String
  
  ' Get the mitarbeiter and check columns
  workersCodes = absenceList.AllMitarbeiter
  
  ' Add or remove columns if needed, mantein the custom order!
  ' Display messages if workers where modified
  ' Get absence list for every worker and then populate the respective rows

Abort:
End Sub

Public Sub TableInitialize()
  ' Delete every worker column
  If table.ListColumns.Count > (startWorkerColumn - 1) Then
    maxColumn = table.ListColumns.Count
    For i = startWorkerColumn To maxColumn
      table.ListColumns(startWorkerColumn).Delete
    Next i
  End If
  
  ' Initialize MitarbeiterList class
  Dim workersList As New MitarbeiterList
  
  ' Get the worker and generate columns
  Dim absenceList As New AbwesenheitsList
  Dim workersCodes() As String
  workersCodes = absenceList.AllMitarbeiter
  Dim currentColumn As Integer
  
  ' Generate columns for every worker
  For i = LBound(workersCodes) To UBound(workersCodes)
    table.ListColumns.Add
    
    workersList.Search workersCodes(i)
    With table.HeaderRowRange(1, i + startWorkerColumn)
      .Value = workersList.worker.visName
      .ColumnWidth = workerColumnWidth
      .Orientation = xlUpward
      .HorizontalAlignment = xlHAlignCenter
      .VerticalAlignment = xlVAlignCenter
      
      currentColumn = .Column
    End With
    
    ' Add maximal holiday days
    With Cells(holidayRow, currentColumn)
      .Value = workersList.worker.vacationTotal
      .Font.Bold = True
      .HorizontalAlignment = xlHAlignCenter
      .VerticalAlignment = xlVAlignCenter
    End With
    
    ' Add all the remaining formulas
    ' with Cells.Formula = "=FORMULA()"
    
    ' Add personal code
    With Cells(personalCodeRow, currentColumn)
      .Value = workersList.worker.personalCode
      .Orientation = xlUpward
      .Font.Size = 10
      .HorizontalAlignment = xlHAlignCenter
      .VerticalAlignment = xlVAlignCenter
    End With
    
    ' Add department
    With Cells(departmentRow, currentColumn)
      .Value = workersList.worker.department
      .Font.Size = 8
      .HorizontalAlignment = xlHAlignCenter
      .VerticalAlignment = xlVAlignCenter
    End With
  Next i
End Sub
