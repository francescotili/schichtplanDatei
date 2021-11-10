VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mitarbeiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declare properties of Mitarbeiter class
Public name, surname, personalCode, department, vacationTotal, visName As String

Public Sub Init(personalCode, name, surname, visName, department, vacationTotal As String)
  Me.personalCode = personalCode
  Me.name = name
  Me.surname = surname
  Me.visName = visName
  Me.department = department
  Me.vacationTotal = vacationTotal
End Sub
