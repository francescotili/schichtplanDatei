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
