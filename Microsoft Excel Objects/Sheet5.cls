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
Private Const workerColumnWidth As Double = 4
Private Const abteilungFontSize As Double = 8
Private Const customFont As String = "Franklin Gothic Medium Cond"
Private Const startWorkerColumn As Integer = 6

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
  
  ' Get the worker and generate columns
  Dim absenceList As New AbwesenheitsList
  Dim workersCodes() As String
  workersCodes = absenceList.AllMitarbeiter
  
  ' Generate columns for every worker
  For i = LBound(workersCodes) To UBound(workersCodes)
    table.ListColumns.Add
    With table.HeaderRowRange(1, i + startWorkerColumn)
      .Value = workersCodes(i)
      .ColumnWidth = workerColumnWidth
      .Orientation = xlUpward
      .HorizontalAlignment = xlHAlignCenter
      .VerticalAlignment = xlVAlignCenter
    End With
  Next i
End Sub
