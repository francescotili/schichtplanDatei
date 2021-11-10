VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declare the variables of the User object
Public name As String
Public location As String
Public computer As String
Public ID As String

Private Sub Class_Initialize()
    name = Application.userName
    location = Environ("MQADSITE")
    computer = Environ("COMPUTERNAME") & "/" & Environ("USERNAME")
    ID = Environ("USERNAME")
End Sub
