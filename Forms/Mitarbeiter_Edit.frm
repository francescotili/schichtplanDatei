Public Sub Initialize(workerPCode As String)
  ' Load user from table
  Dim workersList As New MitarbeiterList
  workersList.Search workerPCode

  ' Populate form fields
  textboxName.Value = workersList.worker.name
  textboxSurname.Value = workersList.worker.surname
  textboxPersonalCode.Value = workerPCode
  textboxDepartment.Value = workersList.worker.department
  'textboxVacationTotal.Value = workersList.worker.vacationTotal
  textboxVisNameOverride.Value = workersList.worker.visName
End Sub

Private Sub textboxName_Change()
  ' When the name changes, it changes also the visualisation name
  textboxVisNameOverride.Value = textboxSurname.Value & " " & Left(textboxName.Value, 1) & "."
End Sub

Private Sub textboxSurname_Change()
  ' When the surname changes, it changes also the visualisation name
  textboxVisNameOverride.Value = textboxSurname.Value & " " & Left(textboxName.Value, 1) & "."
End Sub

Private Sub Button_Save_Click()
  Dim worker As New Mitarbeiter
  Dim workersList As New MitarbeiterList
    
  ' Create entry
  worker.Init _
    textboxPersonalCode.Value, _
    textboxName.Value, _
    textboxSurname.Value, _
    textboxVisNameOverride.Value, _
    textboxDepartment.Value, _
    0 ' Empty field for future use cases
    
  ' Save or edit the worker
  workersList.Save worker
    
  Unload Me
End Sub

Private Sub Button_Cancel_Click()
  Unload Me
End Sub

