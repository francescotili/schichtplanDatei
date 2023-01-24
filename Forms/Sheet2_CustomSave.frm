Private Sub SaveButton_Click()
  Application.ScreenUpdating = False
  Sheet1.Unprotect Password:=GAdminPassword
  If Me.MoCheckbox.Value Then
    Sheet2.Save Monday
    Sheet2.WeekDataLoad
  End If
  If Me.DiCheckbox.Value Then
    Sheet2.Save Tuesday
    Sheet2.WeekDataLoad
  End If
  If Me.MiCheckbox.Value Then
    Sheet2.Save Wednesday
    Sheet2.WeekDataLoad
  End If
  If Me.DoCheckbox.Value Then
    Sheet2.Save Thursday
    Sheet2.WeekDataLoad
  End If
  If Me.FrCheckbox.Value Then
    Sheet2.Save Friday
    Sheet2.WeekDataLoad
  End If
  Sheet2.Unprotect Password:=GAdminPassword
  Application.ScreenUpdating = True
    
  Unload Me
End Sub

Private Sub UserForm_Activate()
  Dim activeDayCell As Range
  Set activeDayCell = Range("Sh2_ActiveDay")
  
  Select Case activeDayCell.Value
  Case WeekDays.Monday
    Me.MoCheckbox.Value = True
    Me.MoCheckbox.Enabled = False
  Case WeekDays.Tuesday
    Me.DiCheckbox.Value = True
    Me.DiCheckbox.Enabled = False
  Case WeekDays.Wednesday
    Me.MiCheckbox.Value = True
    Me.MiCheckbox.Enabled = False
  Case WeekDays.Thursday
    Me.DoCheckbox.Value = True
    Me.DoCheckbox.Enabled = False
  Case WeekDays.Friday
    Me.FrCheckbox.Value = True
    Me.FrCheckbox.Enabled = False
  End Select
End Sub

