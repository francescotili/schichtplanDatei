Private Sub textboxName_Change()
  ' When the name changes, it changes also the visualisation name
  textboxVisNameOverride.Value = textboxSurname.Value & " " & Left(textboxName.Value, 1) & "."
  CheckValidity
End Sub

Private Sub textboxSurname_Change()
  ' When the surname changes, it changes also the visualisation name
  textboxVisNameOverride.Value = textboxSurname.Value & " " & Left(textboxName.Value, 1) & "."
  CheckValidity
End Sub

Private Sub textboxPersonalCode_Change()
  CheckValidity
End Sub

Private Sub UserForm_Initialize()
  CheckValidity
  textboxDepartment.Value = "Stanzerei"
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

Private Sub CheckValidity()
  Button_Save.Enabled = _
    CBool(Len(textboxName.Value)) And _
    CBool(Len(textboxSurname.Value)) And _
    CBool(Len(textboxPersonalCode.Value))
End Sub
