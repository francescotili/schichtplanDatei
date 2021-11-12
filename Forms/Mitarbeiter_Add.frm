VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mitarbeiter_Add 
   Caption         =   "Neu Mitarbeiter"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   OleObjectBlob   =   "Mitarbeiter_Add.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mitarbeiter_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub textboxName_Change()
  ' When the name changes, it changes also the visualisation name
  textboxVisNameOverride.Value = Left(textboxName.Value, 1) & ". " & textboxSurname.Value
  CheckValidity
End Sub

Private Sub textboxSurname_Change()
  ' When the surname changes, it changes also the visualisation name
  textboxVisNameOverride.Value = Left(textboxName.Value, 1) & ". " & textboxSurname.Value
  CheckValidity
End Sub

Private Sub textboxPersonalCode_Change()
  CheckValidity
End Sub

Private Sub UserForm_Initialize()
  CheckValidity
  textboxDepartment.Value = "OPT6"
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
    textboxVacationTotal.Value
    
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
