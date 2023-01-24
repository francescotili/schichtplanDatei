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
