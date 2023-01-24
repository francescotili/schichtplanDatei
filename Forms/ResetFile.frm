Private Sub JahrTextbox_Change()
  If (JahrTextbox.Value >= 1900) And (JahrTextbox.Value < 2300) Then
    Proceed.Enabled = True
  Else
    Proceed.Enabled = False
  End If
End Sub

Private Sub Proceed_Click()
  ' Generate new year on shifts list
'  If Shifts_Check = True Then
    Application.StatusBar = str_fRF_shiftPlanReset
    Dim shiftList As New SchichtList
    shiftList.GenerateYear JahrTextbox.Value ' Sanitized in JahrTextbox_Change()
'  End If
  
  ' Generate new year on vacation list
'  If Vacation_Check = True Then
    Application.StatusBar = str_fRF_absencePlanReset
    Dim vacationList As New AbwesenheitsList
    vacationList.GenerateYear JahrTextbox.Value ' Sanitized in JahrTextbox_Change()
    Application.StatusBar = str_fRF_absenceTableReset
    Sheet5.TableInitialize JahrTextbox.Value
'  End If
  
  ' Delete workers database and update also vacation list
  ' To be reimplemented, because it doesn't work well with Sheet5.TableInitialize
'  If Workers_Check = True Then
'    Application.StatusBar = "Personallist wurde zurückgesetzt. Bitte warten!"
'    Dim workersList As New MitarbeiterList
'    workersList.Reset
'  End If
  
  ' Update global variable
  Range("Global_ActualYear").Value = JahrTextbox.Value
  
  ' Reset all the views
  Dim firstDay As Date
  firstDay = DateSerial(JahrTextbox.Value, 1, 1)
  Range("Sh6_DayToLoadCell").Value = firstDay
  Range("Sh7_WeekToLoadCell").Value = 1
  Range("Sh2_WeekToLoadCell").Value = 1
  
  Application.StatusBar = False
  Unload Me
  MsgBox "Erfolgreich zurückgesetzt!"
End Sub

Private Sub UserForm_Initialize()
'  Vacation_Check.Value = True
'  Shifts_Check.Value = True
End Sub

