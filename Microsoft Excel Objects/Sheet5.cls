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
' ABSENCE MANAGER FUNCTIONS
' Manage methods for getting all the data for
' Sheet5 (Abwesenheitsplan)
' --------------------------------------------
' TODO LIST
' 1. Custom ordering of workers
' 2. Collapsible detail of holidays (rows 5-8)
' --------------------------------------------
Public sheet As Worksheet
Private table As ListObject

Private changesDetectedCell As Range

' Table styles
Private Const workerColumnWidth As Double = 5
Private Const startWorkerColumn As Integer = 6
Private Const statisticsBeginRow As Integer = 3

' Strings
Private Const losignDataWarningText = "Im Abwesenheitplan wurden nicht gespeicherte änderungen erkannt!" & vbCrLf & "Wenn Sie fortfahren, gehen die verloren. Fortfahren?"
Private Const newWorkerAddedText = "Ein neuer Mitarbeiter wurde dem Plan hinzugefugt"
Private Const workerDeletedText = "Ein Mitarbeiter wurde aus dem Plan gelöscht"

Public Sub Worksheet_Activate()
  Set sheet = Sheet5
  Set table = sheet.ListObjects.Item(GAbsencesModifyTableName)
  
  Set changesDetectedCell = Range("Sh5_ChangesDetected")
  
  AbsenceDataLoad
End Sub

Private Sub AbsenceDataLoad()
  Dim i, j As Long
  
  If changesDetectedCell.Value = True Then
    answer = MsgBox(losignDataWarningText, vbExclamation + vbYesNo, "ACHTUNG")
    If answer = vbYes Then GoTo Proceed Else GoTo Abort
  Else
    GoTo Proceed
  End If

Proceed:
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword

  Dim absenceList As New AbwesenheitsList
  Dim workersCodesDB() As String
  Dim workersCodes() As String
  Dim workerAbsence() As String
  Dim workerCol As Integer
  
  ' Get the workers from database
  workersCodesDB = absenceList.AllMitarbeiter
  
  ' Get the workers from the table
  Dim startCol, endCol As Integer
  startCol = table.HeaderRowRange(1, startWorkerColumn).Column
  endCol = table.HeaderRowRange(1, table.DataBodyRange.Columns.Count).Column
  ReDim workersCodes(endCol - startCol)
  For i = startCol To endCol
    workersCodes(i - startCol) = Cells(statisticsBeginRow + 6, i)
  Next i
  
  ' Check if all the workers in DB are in the table, if not then add it
  For Each workerDB In workersCodesDB
    If Not (UBound(Filter(workersCodes, workerDB)) > -1) Then
      AddWorker CStr(workerDB)
      MsgBox newWorkerAddedText & vbCrLf & "Mitarbeiter-Nr. " & workerDB
    End If
  Next
  
  ' Check if all the workers in the table are in the DB, if not then delete it
  For Each worker In workersCodes
    If Not (UBound(Filter(workersCodesDB, worker)) > -1) Then
      MsgBox workerDeletedText & vbCrLf & "Mitarbeiter-Nr. " & worker
      DeleteWorker CStr(worker)
    End If
  Next
  
  ' Populate the data with new content
  For Each workerDB In workersCodesDB
    ' Get content for the worker
    workerAbsence = absenceList.Mitarbeiter(CStr(workerDB))
    
    ' Find the worker column
    For i = 0 To table.DataBodyRange.Columns.Count - 6
      If CStr(workerDB) = CStr(Cells(statisticsBeginRow + 6, startWorkerColumn + 4 + i).Value) Then
        workerCol = startWorkerColumn + i
        Exit For
      End If
    Next i
    
    ' Populate the respective rows
    table.ListColumns(workerCol).DataBodyRange.ClearContents
    For j = 1 To table.DataBodyRange.Rows.Count
      If Len(workerAbsence(j - 1)) <> 0 Then
        table.DataBodyRange(j, workerCol).Value = workerAbsence(j - 1)
      Else
        If Len(table.DataBodyRange(j, workerCol).Value) <> 0 Then
          table.DataBodyRange(j, workerCol).ClearContent
        End If
      End If
    Next j
  Next
  
  ' Reset change detected Cell
  changesDetectedCell.Value = False

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword

Abort:
End Sub

Private Sub AddWorker(workerPCode As String)
  Dim currentColumn As Long
  Dim currentColumnAa As String
  table.ListColumns.Add
  
  ' Initialize MitarbeiterList class
  Dim workersList As New MitarbeiterList
  workersList.Search workerPCode
  
  ' Write header
  With table.HeaderRowRange(1, table.HeaderRowRange.Columns.Count)
    .Value = workersList.worker.visName
    .ColumnWidth = workerColumnWidth
    .Orientation = xlUpward
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    currentColumn = .Column
  End With
  
  currentColumnAa = columnToLetter(currentColumn)
  
  ' File header border
  With Cells(1, currentColumn).Borders(xlEdgeBottom)
    .LineStyle = XlLineStyle.xlContinuous
    .Color = RGB(0, 154, 155)
    .Weight = xlThick
  End With
  
  ' Urlaubsanpruch Cell
  With Cells(statisticsBeginRow, currentColumn)
    .Value = workersList.worker.vacationTotal
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Resturlaub Cell (FORMULA)
  With Cells(statisticsBeginRow + 1, currentColumn)
    .Formula = "=" & currentColumnAa & "3-SUM(" & currentColumnAa & "5;" & currentColumnAa & "6;" & currentColumnAa & "8)"
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Nur Urlaub Cell (FORMULA)
  With Cells(statisticsBeginRow + 2, currentColumn)
    .Formula = "=COUNTIF(Urlaubsplan[" & workersList.worker.visName & "],""=U"")"
    .Font.Bold = False
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Nur Freizeit Cell (FORMULA)
  With Cells(statisticsBeginRow + 3, currentColumn)
    .Formula = "=COUNTIF(Urlaubsplan[" & workersList.worker.visName & "],""=F"")"
    .Font.Bold = False
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Nur Krank Cell (FORMULA)
  With Cells(statisticsBeginRow + 4, currentColumn)
    .Formula = "=COUNTIF(Urlaubsplan[" & workersList.worker.visName & "],""=K"")"
    .Font.Bold = False
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Nur Sonstiges Cell (FORMULA)
  With Cells(statisticsBeginRow + 5, currentColumn)
    .Formula = "=COUNTIF(Urlaubsplan[" & workersList.worker.visName & "],""=S"")"
    .Font.Bold = False
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Add personal code
  With Cells(statisticsBeginRow + 6, currentColumn)
    .Value = workersList.worker.personalCode
    .Orientation = xlUpward
    .Font.Size = 10
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Add department
  With Cells(statisticsBeginRow + 7, currentColumn)
    .Value = workersList.worker.department
    .Font.Size = 8
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With
  
  ' Reset changes detected
  changesDetectedCell.Value = False
End Sub

Private Sub DeleteWorker(workerPCode As String)
  ' Get the workers from the table
  Dim startCol, endCol As Integer
  startCol = table.HeaderRowRange(1, startWorkerColumn).Column
  endCol = table.HeaderRowRange(1, table.DataBodyRange.Columns.Count).Column
  For i = startCol To endCol
    If CStr(Cells(statisticsBeginRow + 6, i).Value) = workerPCode Then
      sheet.Columns(i).EntireColumn.Delete
    End If
  Next i
  
  ' Reset changes detected
  changesDetectedCell.Value = False
End Sub

Public Sub TableInitialize(year As Integer)
  ' These declaration are needed when the Sub is called from other Sheets
  Set sheet = Sheet5
  Set table = sheet.ListObjects.Item(GAbsencesModifyTableName)
  Set changesDetectedCell = Range("Sh5_ChangesDetected")
  Dim startDate, endDate As Date
  Dim numOfDays, i As Long
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Delete every worker column
  If table.ListColumns.Count > (startWorkerColumn - 1) Then
    maxColumn = table.ListColumns.Count
    For i = startWorkerColumn To maxColumn
      sheet.Columns(startWorkerColumn + 4).EntireColumn.Delete
    Next i
  End If
  
  ' Delete all the rows
  With table
    If Not .DataBodyRange Is Nothing Then
      .DataBodyRange.Delete
    End If
  End With
  
  ' Get the workers Data
  Dim absenceList As New AbwesenheitsList
  Dim workersCodes() As String
  workersCodes = absenceList.AllMitarbeiter
  
  ' Generate columns for every worker
  For i = LBound(workersCodes) To UBound(workersCodes)
    AddWorker CStr(workersCodes(i))
  Next i
  
  ' Generate rows with date
  startDate = DateSerial(year, 1, 1)
  endDate = DateSerial(year, 12, 31)
  numOfDays = DateDiff("d", startDate, endDate)
  For i = 0 To numOfDays
    table.ListRows.Add
    table.DataBodyRange((i + 1), 1).Value = startDate + i
  Next i
  
  ' Reset change detected Cell
  changesDetectedCell.Value = False

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  Sh5_PlanSaveBtn.Enabled = False
  If Target.row > 11 Then ' Everything under Row 11 will trigger
    changesDetectedCell.Value = True
    Sh5_PlanSaveBtn.Enabled = True
  End If
End Sub
