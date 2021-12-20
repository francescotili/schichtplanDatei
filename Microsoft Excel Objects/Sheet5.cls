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

Public sheet As Worksheet
Private table As ListObject

Private changesDetectedCell, detailsHiddenCell As Range
Private moveColRBtn, moveColLBtn, toggleDetails As Shape

' Table styles
Private Const workerColumnWidth As Double = 5
Private Const startWorkerColumn As Integer = 6
Private Const statisticsBeginRow As Integer = 3

' Strings
Private Const losignDataWarningText = "Im Abwesenheitplan wurden nicht gespeicherte änderungen erkannt!" & vbCrLf & "Wenn Sie fortfahren, gehen die verloren. Fortfahren?"
Private Const newWorkerAddedText = "Ein neuer Mitarbeiter wurde dem Plan hinzugefugt"
Private Const workerDeletedText = "Ein Mitarbeiter wurde aus dem Plan gelöscht"
Private Const savingDataWarningText = "Möchten Sie den Abwesenheitplan wirklich speichern?" & vbCrLf & "Vorhandene Daten werden überschrieben!"

Public Sub Worksheet_Activate()
  Set sheet = Sheet5
  Set table = sheet.ListObjects.Item(GAbsencesModifyTableName)
  
  Set moveColRBtn = Me.sheet.Shapes("Sh5_MoveColRightBtn")
  Set moveColLBtn = Me.sheet.Shapes("Sh5_MoveColLeftBtn")
  Set toggleDetails = Me.sheet.Shapes("Sh5_ToggleDtailsVisBtn")
  
  Set changesDetectedCell = Range("Sh5_ChangesDetected")
  Set detailsHiddenCell = Range("Sh5_DetailHidden")
  
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
          table.DataBodyRange(j, workerCol).ClearContents
        End If
      End If
    Next j
  Next
  
  ' Reset change detected Cell
  changesDetectedCell.Value = False

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword, AllowFormattingColumns:=True

Abort:
End Sub

Private Sub Sh5_PlanSaveBtn_Click()
  AbsenceDataSave
End Sub

Private Sub AbsenceDataSave()
  Dim msgAnswer As Integer
  msgAnswer = MsgBox(savingDataWarningText, vbQuestion + vbYesNo, "Fortfahren?")
  If msgAnswer = vbYes Then GoTo Proceed Else GoTo Abort

Proceed:
  Dim i, j As Long
  Dim absenceList As New AbwesenheitsList
  Dim workersCodesDB() As String
  Dim workerAbsence() As String
  ReDim workerAbsence(table.DataBodyRange.Rows.Count - 1)
  Dim modCounter As Long
  
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  absenceList.ProtectionOFF
  
  ' Get the workers from database
  workersCodesDB = absenceList.AllMitarbeiter
  
  ' Scan every worker and save data
  For Each workerDB In workersCodesDB
    For i = startWorkerColumn To table.DataBodyRange.Columns.Count
      If Cells(statisticsBeginRow + 6, table.HeaderRowRange(1, i).Column) = CStr(workerDB) Then
        For j = 1 To table.DataBodyRange.Rows.Count
          workerAbsence(j - 1) = table.DataBodyRange(j, i)
        Next j
        absenceList.Save CStr(workerDB), workerAbsence
      End If
    Next i
  Next
  
  ' Reset change detected Cell
  changesDetectedCell.Value = False
  
  ' Log event in History
  saveHistory ("Der Abwesenheitsplan wurde geändert")

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword, AllowFormattingColumns:=True
  absenceList.ProtectionON
  
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
    .Formula = "=" & currentColumnAa & "3-SUM(" & currentColumnAa & "5," & currentColumnAa & "6," & currentColumnAa & "8)"
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
  ProtectShapes
  
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
  workersCodes = absenceList.AllMitarbeiter ' PROBLEMA QUA! NON CI SONO PIU´ GLI OPERAI NO!? DEVO PRENDERLI DA MITARBEITER LIST!
  
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
  sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  Sh5_PlanSaveBtn.Enabled = False
  If Target.row > 11 Then ' Everything under Row 11 will trigger
    changesDetectedCell.Value = True
    Sh5_PlanSaveBtn.Enabled = True
  End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  Select Case Target.Column
  Case Is < 10
    moveColRBtn.Visible = msoFalse
    moveColLBtn.Visible = msoFalse
  Case Is = 10
    moveColRBtn.Visible = msoTrue
    moveColLBtn.Visible = msoFalse
  Case Is < (table.DataBodyRange.Columns.Count + 4)
    moveColRBtn.Visible = msoTrue
    moveColLBtn.Visible = msoTrue
  Case Is = (table.DataBodyRange.Columns.Count + 4)
    moveColRBtn.Visible = msoFalse
    moveColLBtn.Visible = msoTrue
  Case Is > (table.DataBodyRange.Columns.Count + 4)
    moveColRBtn.Visible = msoFalse
    moveColLBtn.Visible = msoFalse
  End Select
  
  moveColRBtn.Left = Target.Left + 20
  moveColLBtn.Left = Target.Left - 5
End Sub

Public Sub MoveColToRight()
  Dim activeCol As Long
  
  sheet.Unprotect Password:=GAdminPassword
  activeCol = CLng(ActiveCell.Column)
  sheet.Columns(activeCol).Cut
  sheet.Columns(activeCol + 2).Insert
  sheet.Columns(activeCol).ColumnWidth = workerColumnWidth
  sheet.Columns(activeCol + 1).Select
  
  RestoreButtons
End Sub

Public Sub MoveColToLeft()
  Dim activeCol As Long
  
  sheet.Unprotect Password:=GAdminPassword
  activeCol = CLng(ActiveCell.Column)
  sheet.Columns(activeCol).Cut
  sheet.Columns(activeCol - 1).Insert
  sheet.Columns(activeCol + 1).ColumnWidth = workerColumnWidth
  sheet.Columns(activeCol - 1).Select
  
  RestoreButtons
End Sub

Private Sub RestoreButtons()
  With sheet.Shapes("Sh5_MoveColRightBtn")
    .Width = 15
    .Left = ActiveCell.Left + 20
  End With
  With sheet.Shapes("Sh5_MoveColLeftBtn")
    .Width = 15
    .Left = ActiveCell.Left - 5
  End With
  sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True
End Sub

Private Sub ProtectShapes()
  sheet.Shapes("Sh5_MoveColRightBtn").Left = 0
  sheet.Shapes("Sh5_MoveColLeftBtn").Left = 0
End Sub

Public Sub ToggleDetailsVisibility()
  If detailsHiddenCell.Value = True Then
    sheet.Unprotect Password:=GAdminPassword
    Rows("5:8").EntireRow.Hidden = True
    toggleDetails.TextFrame.Characters.Text = "+"
    detailsHiddenCell.Value = False
    sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True
  Else
    sheet.Unprotect Password:=GAdminPassword
    Rows("5:8").EntireRow.Hidden = False
    toggleDetails.TextFrame.Characters.Text = "–"
    detailsHiddenCell.Value = True
    sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True
  End If
End Sub

