Attribute VB_Name = "Strings"
' --------------------------------------------
' MODULE FOR STRINGS LOCALIZATION
' Here are all the messages used throught the
' project.
' --------------------------------------------

' Sheet 2
Public Const str_sh2_losingDataWarn = "Im Schichtplan wurden nicht gespeicherte änderungen erkannt!" & vbCrLf & "Wenn Sie fortfahren, gehen die verloren. Fortfahren?"
Public Const str_sh2_savingWarn = "Möchten Sie den Schichtplan wirklich speichern?" & vbCrLf & "Vorhandene Daten der ganzen Kalenderwoche werden überschrieben!"
Public Const str_sh2_moreAbsence = "Und {0} weitere"
Public Const str_sh2_historySPWU = "Der Schichtplan für KW{0} wurde geändert"
Public Const str_sh2_historySPWK = "Der Schichtplan für den {0} (KW {1}) wurde geändert"

' Sheet 5
Public Const str_sh5_losingDataWarn = "Im Abwesenheitplan wurden nicht gespeicherte änderungen erkannt!" & vbCrLf & "Wenn Sie fortfahren, gehen die verloren. Fortfahren?"
Public Const str_sh5_newWorkerAdded = "Ein neuer Mitarbeiter wurde dem Plan hinzugefügt"
Public Const str_sh5_workerDeleted = "Ein Mitarbeiter wurde aus dem Plan gelöscht"
Public Const str_sh5_savingWarn = "Möchten Sie den Abwesenheitplan wirklich speichern?" & vbCrLf & "Vorhandene Daten werden überschrieben!"
Public Const str_sh2_history = "Der Abwesenheitsplan wurde geändert"

' Sheet 6
Public Const str_sh6_moreAbsence = "Und {0} weitere"

' Sheet 7
Public Const str_sh7_weekStatusLabelNM = "Regelmäßige Woche"
Public Const str_sh7_weekStatusLabelCM = "Unregelmäßige Woche"
Public Const str_sh7_weekStatusLabelEM = "Leere Woche"
Public Const str_sh7_moreAbsence = "Und {0} weitere"

' Sheet 10
Public Const str_sh10_history = "Im Registerkarte ""Feiertage"" wurden änderungen festgestellt"

' Form - Mitarbeiter_Manage
Public Const str_fMM_noWorkerSelected = "Bitte einen Mitarbeiter aus der Liste auswählen!"
Public Const str_fMM_deleteConfirmation = "ACHTUNG!" & vbCrLf & "Mitarbeiter-Nr. {0} wird endgultig gelöscht." & vbCrLf & "Wirklich fortfahren ??"
Public Const str_fMM_historyDeletion = "Mitarbeiter-Nr. {0} gelöscht"

' Form - ResetFile
Public Const str_fRF_shiftPlanReset = "Schichtplandatei wurde zurückgesetzt. Bitte warten!"
Public Const str_fRF_absencePlanReset = "Abwesenheitsplan wurde zurückgesetzt. Bitte warten!"
Public Const str_fRF_absenceTableReset = "Abwesenheitstabelle wurde zurückgesetzt. Bitte warten!"

' Global
Public Const str_warnBoxTitle = "ACHTUNG"

' Reset functions history
Public Const str_resetAbsenceStatus = "DER ABWESENHEITPLAN WURDE ZURÜCKGESETZT"
Public Const str_resetWorkerList = "DIE MITARBEITERLISTE WURDE GELÖSCHT"
Public Const str_resetShiftplan = "DER SCHICHTPLAN WURDE ZURÜCKGESETZT"
