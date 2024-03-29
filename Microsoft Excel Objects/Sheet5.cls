' --------------------------------------------
' ABSENCE MANAGER FUNCTIONS
' Manage methods for getting all the data for
' Sheet5 (Abwesenheitsplan)
' --------------------------------------------

Public sheet As Worksheet
Private table As ListObject

Private changesDetectedCell As Range
Private moveColRBtn, moveColLBtn As Shape

' Table styles
Private Const workerColumnWidth As Double = 5
Private Const startWorkerColumn As Integer = 6
Private Const personalCodeRow As Integer = 3

Public Sub Worksheet_Activate()
  Set sheet = Sheet5
  Set table = sheet.ListObjects.Item(GAbsencesModifyTableName)
  
  Set moveColRBtn = Me.sheet.Shapes("Sh5_MoveColRightBtn")
  Set moveColLBtn = Me.sheet.Shapes("Sh5_MoveColLeftBtn")
  
  Set changesDetectedCell = Range("Sh5_ChangesDetected")
  
  AbsenceDataLoad
End Sub

Private Sub AbsenceDataLoad()
  Dim i, j As Long
  
  ' Code commented to avoid data corruption when the user decide to not load data
  'If changesDetectedCell.Value = True Then
  '  answer = MsgBox(str_sh5_losingDataWarn, vbExclamation + vbYesNo, "ACHTUNG")
  '  If answer = vbYes Then GoTo Proceed Else GoTo Abort
  'Else
  '  GoTo Proceed
  'End If

'Proceed:
  ' Disable screen updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword

  Dim absenceList As New AbwesenheitsList
  Dim workersCodesDB() As String
  Dim workersCodes() As String
  Dim workerAbsence() As String
  Dim workerTableCol, workerCol As Integer
  
  ' Get the workers from database
  workersCodesDB = absenceList.AllMitarbeiter
  
  ' Get the workers from the table
  Dim startCol, endCol As Integer
  startCol = table.HeaderRowRange(1, startWorkerColumn).Column
  endCol = table.HeaderRowRange(1, table.DataBodyRange.Columns.Count).Column
  ReDim workersCodes(endCol - startCol)
  For i = startCol To endCol
    workersCodes(i - startCol) = Cells(personalCodeRow, i)
  Next i
  
  ' Check if all the workers in DB are in the table, if not then add it
  For Each workerDB In workersCodesDB
    If Not IsInArray(CStr(workerDB), workersCodes) Then
      AddWorker CStr(workerDB)
      MsgBox StringFormat(str_sh5_newWorkerAdded, workerDB)
    End If
  Next
  
  ' Check if all the workers in the table are in the DB, if not then delete it
  For Each worker In workersCodes
    If Not IsInArray(CStr(worker), workersCodesDB) Then
      MsgBox StringFormat(str_sh5_workerDeleted, worker)
      DeleteWorker CStr(worker)
    End If
  Next
  
  ' Populate the data with new content
  Dim workersList As New MitarbeiterList
  For Each workerDB In workersCodesDB
    ' Get content for the worker
    workerAbsence = absenceList.Mitarbeiter(CStr(workerDB))
    
    ' Find the worker column
    For i = 0 To table.DataBodyRange.Columns.Count - 6
      If CStr(workerDB) = CStr(Cells(personalCodeRow, startWorkerColumn + 4 + i).Value) Then
        workerTableCol = startWorkerColumn + i
        workerCol = table.DataBodyRange(1, workerTableCol).Column
        Exit For
      End If
    Next i
    
    ' Populate the respective columns
    table.ListColumns(workerTableCol).DataBodyRange.ClearContents
    For j = 1 To table.DataBodyRange.Rows.Count
      If Len(workerAbsence(j - 1)) <> 0 Then
        table.DataBodyRange(j, workerTableCol).Value = workerAbsence(j - 1)
      Else
        If Len(table.DataBodyRange(j, workerTableCol).Value) <> 0 Then
          table.DataBodyRange(j, workerTableCol).ClearContents
        End If
      End If
    Next j
  
    ' Check if all the information are up to date (only department and name)
    workersList.Search CStr(workerDB)
    With table.HeaderRowRange(1, workerTableCol)
      If .Value <> workersList.worker.visName Then
        .Value = CStr(workersList.worker.visName)
      End If
    End With
    With Cells(4, workerCol)
      If .Value <> workersList.worker.department Then
        If Len(workersList.worker.department) > 5 Then
          .Value = Left(workersList.worker.department, 5) & "."
        Else
          .Value = CStr(workersList.worker.department)
        End If
      End If
    End With
  Next
  
  ' Reset change detected Cell
  changesDetectedCell.Value = False

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True

'Abort:
End Sub

Private Sub Sh5_PlanSaveBtn_Click()
  AbsenceDataSave
End Sub

Private Sub AbsenceDataSave()
  Dim msgAnswer As Integer
  msgAnswer = MsgBox(str_sh5_savingWarn, vbQuestion + vbYesNo, "Fortfahren?")
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
      If Cells(personalCodeRow, table.HeaderRowRange(1, i).Column) = CStr(workerDB) Then
        For j = 1 To table.DataBodyRange.Rows.Count
          ' Convert "a" to "A"
          If CStr(table.DataBodyRange(j, i)) = "a" Then
            table.DataBodyRange(j, i) = "A"
          End If
          ' Convert "s" to "S"
          If CStr(table.DataBodyRange(j, i)) = "s" Then
            table.DataBodyRange(j, i) = "S"
          End If
          ' Convert "h" to "H"
          If CStr(table.DataBodyRange(j, i)) = "h" Then
            table.DataBodyRange(j, i) = "H"
          End If
          ' Convert "e" to "E"
          If CStr(table.DataBodyRange(j, i)) = "e" Then
            table.DataBodyRange(j, i) = "E"
          End If
          
          workerAbsence(j - 1) = table.DataBodyRange(j, i)
        Next j
        absenceList.Save CStr(workerDB), workerAbsence
      End If
    Next i
  Next
  
  ' Reset change detected Cell
  changesDetectedCell.Value = False
  
  ' Log event in History
  saveHistory (str_sh2_history)

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True
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

  ' Add personal code
  With Cells(personalCodeRow, currentColumn)
    .Value = workersList.worker.personalCode
    .Orientation = xlUpward
    .Font.Size = 10
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With

  ' Add department
  With Cells(personalCodeRow + 1, currentColumn)
    If Len(workersList.worker.department) > 5 Then
      .Value = Left(workersList.worker.department, 5) & "."
    Else
      .Value = workersList.worker.department
    End If
    .Font.Size = 8
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
  End With
  
  ' Restore conditional formatting
  ResetConditionalFormatting
  
  ' Reset changes detected
  changesDetectedCell.Value = False
End Sub

Private Sub DeleteWorker(workerPCode As String)
  Dim startCol, endCol As Integer
  startCol = table.HeaderRowRange(1, startWorkerColumn).Column
  endCol = table.HeaderRowRange(1, table.DataBodyRange.Columns.Count).Column
  
  ' Protect buttons for moving columns
  ProtectShapes
  
  ' Scan and delete worker
  For i = startCol To endCol
    If CStr(Cells(personalCodeRow, i).Value) = workerPCode Then
      sheet.Columns(i).EntireColumn.Delete
    End If
  Next i
  
  ' Restore conditional formatting
  ResetConditionalFormatting
  
  ' Reset changes detected
  changesDetectedCell.Value = False
End Sub

Public Sub TableInitialize(Year As Integer)
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
  'If table.ListColumns.Count > (startWorkerColumn - 1) Then
  '  maxColumn = table.ListColumns.Count
  '  For i = startWorkerColumn To maxColumn
  '    sheet.Columns(startWorkerColumn + 4).EntireColumn.Delete
  '  Next i
  'End If
  
  ' Delete all the rows
  With table
    If Not .DataBodyRange Is Nothing Then
      .DataBodyRange.Delete
    End If
  End With
  
  ' Get the workers Data
  'Dim absenceList As New AbwesenheitsList
  'Dim workersCodes() As String
  'workersCodes = absenceList.AllMitarbeiter
  
  ' Generate columns for every worker
  'For i = LBound(workersCodes) To UBound(workersCodes)
  '  AddWorker CStr(workersCodes(i))
  'Next i
  
  ' Generate rows with date
  startDate = DateSerial(Year, 1, 1)
  endDate = DateSerial(Year, 12, 31)
  numOfDays = DateDiff("d", startDate, endDate)
  For i = 0 To numOfDays
    table.ListRows.Add
    table.DataBodyRange((i + 1), 1).Value = startDate + i
  Next i
  
  ' Reset conditional formatting
  table.DataBodyRange.Interior.Color = xlNone
  ResetConditionalFormatting
  
  ' Reset change detected Cell
  changesDetectedCell.Value = False

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  Sh5_PlanSaveBtn.Enabled = False
  If Target.row > 5 Then ' Everything under Row 5 will trigger
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
  
  moveColRBtn.Left = Target.Left + 25
  moveColLBtn.Left = Target.Left - 10
  moveColRBtn.Width = 20
  moveColLBtn.Width = 20
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
    .Width = 20
    .Left = ActiveCell.Left + 25
  End With
  With sheet.Shapes("Sh5_MoveColLeftBtn")
    .Width = 20
    .Left = ActiveCell.Left - 10
  End With

  ' Enable screen Updating and protection
  Application.ScreenUpdating = True
  sheet.Protect Password:=GAdminPassword, AllowFormattingCells:=True
End Sub

Private Sub ProtectShapes()
  sheet.Shapes("Sh5_MoveColRightBtn").Left = 0
  sheet.Shapes("Sh5_MoveColLeftBtn").Left = 0
End Sub

Private Sub ResetConditionalFormatting()
  Dim tempRange As Range
  
  ' Delete preexistent conditional formatting
  table.DataBodyRange.FormatConditions.Delete
  
  ' Add conditional formatting for months
  Set tempRange = Range(table.DataBodyRange(1, 1).Address & ":" & table.DataBodyRange(table.DataBodyRange.Rows.Count, 3).Address)
  With tempRange
    .FormatConditions.Add Type:=xlExpression, Formula1:="=WEEKDAY($E6;11)>5"
    With .FormatConditions(1)
      .StopIfTrue = False
      .Font.Color = RGB(0, 154, 155)
    End With
  End With
  
  ' Add conditional formatting for months dividers
  Set tempRange = Range(table.DataBodyRange(1, 1).Address & ":" & table.DataBodyRange(table.DataBodyRange.Rows.Count, table.DataBodyRange.Columns.Count).Address)
  With tempRange
    .FormatConditions.Add Type:=xlExpression, Formula1:="=$G6<>$G7"
    With .FormatConditions(2)
      .StopIfTrue = False
      With .Borders.Item(4)
        .LineStyle = xlContinuous
        .Color = RGB(0, 100, 112)
      End With
    End With
  End With
  
  ' Add conditional formatting for week divider
  Set tempRange = Range(table.DataBodyRange(1, 1).Address & ":" & table.DataBodyRange(table.DataBodyRange.Rows.Count, table.DataBodyRange.Columns.Count).Address)
  With tempRange
    .FormatConditions.Add Type:=xlExpression, Formula1:="=$F6<>$F7"
    With .FormatConditions(3)
      .StopIfTrue = False
      With .Borders.Item(4)
        .LineStyle = xlDashDot
        .Color = RGB(0, 154, 155)
      End With
    End With
  End With
  
  ' Add conditional formatting for vacation type, if workers exist
  If table.DataBodyRange.Columns.Count > 5 Then
    Set tempRange = Range(table.DataBodyRange(1, 6).Address & ":" & table.DataBodyRange(table.DataBodyRange.Rows.Count, table.DataBodyRange.Columns.Count).Address)
    
    With tempRange
      .Orientation = xlHorizontal ' Reset text orientation
      .Locked = False ' Reset Locked status
      
      ' Add conditional formatting for vacation type A (Abwesend)
      .FormatConditions.Add Type:=xlExpression, Formula1:="=J6=""A"""
      .FormatConditions.Add Type:=xlExpression, Formula1:="=J6=""S"""
      .FormatConditions.Add Type:=xlExpression, Formula1:="=J6=""H"""
      .FormatConditions.Add Type:=xlExpression, Formula1:="=J6=""E"""
    End With
    
    ' Modify conditional formatting - do not modify the order, Excel has a lot of bugs
    ' with conditional formatting and VBA
    With table.DataBodyRange
      With .FormatConditions(4)
        .StopIfTrue = False
        .Borders.Color = RGB(32, 55, 100)
        .Font.Color = RGB(32, 55, 100)
        .Font.Bold = True
      End With
    End With
    With table.DataBodyRange
      With .FormatConditions(5)
        .StopIfTrue = False
        .Borders.Color = RGB(55, 86, 35)
        .Font.Color = RGB(55, 86, 35)
        .Font.Bold = True
      End With
    End With
    With table.DataBodyRange
      With .FormatConditions(6)
        .StopIfTrue = False
        .Borders.Color = RGB(0, 154, 155)
        .Font.Color = RGB(0, 154, 155)
        .Font.Bold = True
      End With
    End With
    With table.DataBodyRange
      With .FormatConditions(7)
        .StopIfTrue = False
        .Borders.Color = RGB(131, 60, 12)
        .Font.Color = RGB(131, 60, 12)
        .Font.Bold = True
      End With
    End With
    
    With table.DataBodyRange
      With .FormatConditions(4)
        .Interior.Color = RGB(180, 198, 231)
      End With
    End With
    With table.DataBodyRange
      With .FormatConditions(5)
        .Interior.Color = RGB(198, 224, 180)
      End With
    End With
    With table.DataBodyRange
      With .FormatConditions(6)
        .Interior.Color = RGB(217, 236, 239)
      End With
    End With
    With table.DataBodyRange
      With .FormatConditions(7)
        .Interior.Color = RGB(252, 228, 214)
      End With
    End With
  End If
End Sub

