VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResetFile 
   Caption         =   "Datei zurücksetzen"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4050
   OleObjectBlob   =   "ResetFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ResetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub JahrTextbox_Change()
  If (JahrTextbox.Value >= 1900) And (JahrTextbox.Value < 2300) Then
    Proceed.Enabled = True
  Else
    Proceed.Enabled = False
  End If
End Sub

Private Sub Proceed_Click()
  Application.StatusBar = "Datei wurde zurückgesetzt. Bitte warten!"

  ' Generate new year on shifts list
  If Shifts_Check = True Then
    Dim shiftList As New SchichtList
    shiftList.GenerateYear JahrTextbox.Value ' Sanitized in JahrTextbox_Change()
  End If
  
  ' Generate new year on vacation list
  If Vacation_Check = True Then
    Dim vacationList As New AbwesenheitsList
    vacationList.GenerateYear JahrTextbox.Value ' Sanitized in JahrTextbox_Change()
  End If
  
  ' Delete workers database and update also vacation list
  If Workers_Check = True Then
    Dim workersList As New MitarbeiterList
    workersList.Reset
  End If
  
  Application.StatusBar = False
  Unload Me
End Sub

Private Sub UserForm_Initialize()
  Vacation_Check.Value = True
  Shifts_Check.Value = True
End Sub
