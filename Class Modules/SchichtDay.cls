VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SchichtDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declare properties of SchichtDay class
Public shiftDate As Date
Public shiftStatus As WeekStatus
Public shiftCode As String
Private workPlaces(32) As String ' You cannot have a Public array in a Class Module, sadly

Public Sub Init(shiftDate As Date, shiftStatus As WeekStatus, shiftName As String)
  Me.shiftDate = shiftDate
  Me.shiftStatus = shiftStatus
  Me.shiftCode = Year(shiftDate) & Format(Month(shiftDate), "00") & Format(day(shiftDate), "00") & shiftName
End Sub

Public Sub Populate(shiftData() As String)
  Dim i As Long
  For i = 0 To 32 ' Number of work places
    workPlaces(i) = shiftData(i)
  Next i
End Sub

Public Property Get shiftPlan() As String()
  shiftPlan = workPlaces
End Property

Public Property Get workPlace(index As Long) As String
  workPlace = workPlaces(index)
End Property
