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
    MsgBox str_fMM_noWorkerSelected
  Else
    ' Ask for confirmation
    Dim Msg, Style, Title
    Msg = StringFormat(str_fMM_deleteConfirmation, selectedPCode)
    Style = vbYesNo + vbExclamation + vbDefaultButton2
    response = MsgBox(Msg, Style, Title)
    If response = vbYes Then ' User wants to continue
      Dim workersData As New MitarbeiterList
      workersData.Delete (selectedPCode)
      UpdateData
      saveHistory StringFormat(str_fMM_historyDeletion, selectedPCode)
    End If
  End If
End Sub

Private Sub Btn_MitarbeiterEdit_Click()
  ' Get selected user
  Dim selectedPCode As String
  selectedPCode = GetSelectedWorker
  
  If selectedPCode = "" Then ' No worker selected
    MsgBox str_fMM_noWorkerSelected
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
  Dim X As Long
  Dim selectedPCode As String
  selectedPCode = ""
  
  ' Get selected element
  For X = 0 To MitarbeiterList.ListCount - 1
    If MitarbeiterList.Selected(X) = True Then
      selectedPCode = MitarbeiterList.list(X, 0)
    End If
  Next X
  
  GetSelectedWorker = selectedPCode
End Function

Private Sub UserForm_Activate()
  UpdateData
End Sub

Public Sub UpdateData()
  Dim User As Mitarbeiter
  Dim X, i As Long
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
