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

Private changesDetectedCell As Range

' Strings
Private Const losignDataWarningText = "Im Abwesenheitplan wurden nicht gespeicherte änderungen erkannt!" & vbCrLf & "Wenn Sie fortfahren, gehen die verloren. Fortfahren?"

Public Sub Worksheet_Activate()
  Set sheet = Sheet5
  
  Set changesDetectedCell = Range("Sh5_ChangesDetected")
  
  AbsenceDataLoad
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
  
  ' Get the mitarbeiter and generate columns
  workersCodes = absenceList.AllMitarbeiter
  
  For i = LBound(workersCodes) To UBound(workersCodes)
    Debug.Print workersCodes(i)
  Next i
  
  ' Generate columns for every worker
  ' Display messages if workers where modified (?)
  ' Get absence list for every worker and then populate rows

Abort:
End Sub
