VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sheet2_CustomSave 
   Caption         =   "Speichern"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   OleObjectBlob   =   "Sheet2_CustomSave.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sheet2_CustomSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
  Dim activeDayCell As Range
  Set activeDayCell = Range("Sh2_ActiveDay")
  
  Select Case activeDayCell.Value
  Case WeekDay.Monday
    Me.MoCheckbox.Value = True
  Case WeekDay.Tuesday
    Me.DiCheckbox.Value = True
  Case WeekDay.Wednesday
    Me.MiCheckbox.Value = True
  Case WeekDay.Thursday
    Me.DoCheckbox.Value = True
  Case WeekDay.Friday
    Me.FrCheckbox.Value = True
  End Select
End Sub
