VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Test()
  Dim vacationList As New AbwesenheitsList
  Dim worker As New Mitarbeiter
  
  worker.Init _
    "20130", _
    "Max", _
    "Mustermann", _
    "M. Mustermann", _
    "OPT6E", _
    "30"
  
  vacationList.WorkerAdd worker
End Sub
