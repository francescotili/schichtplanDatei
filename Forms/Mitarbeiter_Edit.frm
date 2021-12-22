VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mitarbeiter_Edit 
   Caption         =   "UserForm1"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4515
   OleObjectBlob   =   "Mitarbeiter_Edit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mitarbeiter_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
  textboxVisNameOverride.Value = Left(textboxName.Value, 1) & ". " & textboxSurname.Value
End Sub

Private Sub textboxSurname_Change()
  ' When the surname changes, it changes also the visualisation name
  textboxVisNameOverride.Value = Left(textboxName.Value, 1) & ". " & textboxSurname.Value
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
    0 ' TODO: Use this field for "fremdarbeiter"
    
  ' Save or edit the worker
  workersList.Save worker
    
  Unload Me
End Sub

Private Sub Button_Cancel_Click()
  Unload Me
End Sub
