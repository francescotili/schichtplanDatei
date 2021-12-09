VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mitarbeiter_Manage 
   Caption         =   "Mitarbeiter verwalten"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "Mitarbeiter_Manage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mitarbeiter_Manage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public userList As MitarbeiterList

Private Sub Btn_MitarbeiterAdd_Click()
  ' Close window
  Unload Me
  
  ' Open the creation form
  Dim form As New Mitarbeiter_Add
  With form
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
  End With
End Sub

Private Sub Btn_MitarbeiterDelete_Click()
  ' Get selected user
  Dim selectedPCode As String
  selectedPCode = GetSelectedWorker
  
  If selectedPCode = "" Then ' No worker selected
    MsgBox "Bitte einen Mitarbeiter aus der Liste auswählen!"
  Else
    ' Ask for confirmation
    Dim Msg, Style, Title
    Msg = "ACHTUNG!" & vbCrLf & "Mitarbeiter-Nr. " & selectedPCode & " wird endgultig gelöscht." & vbCrLf & "Wirklich fortfahren ??"
    Style = vbYesNo + vbExclamation + vbDefaultButton2
    response = MsgBox(Msg, Style, Title)
    If response = vbYes Then ' User wants to continue
      Dim workersData As New MitarbeiterList
      workersData.Delete (selectedPCode)
      UpdateData
      saveHistory ("Mitarbeiter-Nr. " & selectedPCode & " gelöscht")
    End If
  End If
End Sub

Private Sub Btn_MitarbeiterEdit_Click()
  ' Get selected user
  Dim selectedPCode As String
  selectedPCode = GetSelectedWorker
  
  If selectedPCode = "" Then ' No worker selected
    MsgBox "Bitte einen Mitarbeiter aus der Liste auswählen!"
  Else
    ' Close window
    Unload Me
    ' Open the editing form
    Dim form As New Mitarbeiter_Edit
    With form
      .Initialize selectedPCode
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
    End With
  End If
End Sub

Private Function GetSelectedWorker() As String
  Dim x As Long
  Dim selectedPCode As String
  selectedPCode = ""
  
  ' Get selected element
  For x = 0 To MitarbeiterList.ListCount - 1
    If MitarbeiterList.Selected(x) = True Then
      selectedPCode = MitarbeiterList.list(x, 0)
    End If
  Next x
  
  GetSelectedWorker = selectedPCode
End Function

Private Sub UserForm_Activate()
  UpdateData
End Sub

Public Sub UpdateData()
  Dim User As Mitarbeiter
  Dim x, i As Long
  Dim entryList() As String
  
  Set userList = New MitarbeiterList
  userList.Load
  
  ' User list contains 3 column: PCode, Nachname, Vorname
  MitarbeiterList.ColumnCount = 3
  If userList.data.Count > 0 Then
    ReDim entryList(1 To userList.data.Count, 1 To 3)
    For i = 1 To userList.data.Count
      entryList(i, 1) = userList.data.Item(i).personalCode
      entryList(i, 2) = userList.data.Item(i).surname
      entryList(i, 3) = userList.data.Item(i).name
    Next
    MitarbeiterList.list() = entryList
  Else
    Btn_MitarbeiterEdit.Enabled = False
    Btn_MitarbeiterDelete.Enabled = False
  End If
End Sub
