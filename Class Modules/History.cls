VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declare the variables of the History object
Public eventName As String
Private currentUser As New User
Private sheet As Worksheet
Private table As ListObject

Private Sub Class_Initialize()
  Me.eventName = ""
  Set sheet = Sheet4
  Set table = sheet.ListObjects.Item(GHistoryTableName)
End Sub

Public Sub Save()
  Dim entry As Variant
  Dim i As Integer
    
  ' Create array Value for History
  entry = Array( _
    CDbl(Now()), _
    currentUser.name, _
    eventName, _
    currentUser.location, _
    currentUser.computer _
  )
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
    
  ' Add new row at first position
  table.ListRows.Add Position:=1
  ' Populate the row with data
  For i = LBound(entry) To UBound(entry)
    table.DataBodyRange(1, (i + 1)).Value = entry(i)
  Next i
  
  ' Enable screen Updating and protection
  sheet.Protect Password:=GAdminPassword
  Application.ScreenUpdating = True
End Sub

Public Sub Clean()
  Dim i As Long
  
  ' Disable screen Updating and protection
  Application.ScreenUpdating = False
  sheet.Unprotect Password:=GAdminPassword
  
  ' Clean all the old entries
  If table.ListRows.Count > GMaxHistoryEntries Then
    For i = (GMaxHistoryEntries + 1) To table.ListRows.Count
      table.ListRows(GMaxHistoryEntries + 1).Delete
    Next i
  End If
End Sub
