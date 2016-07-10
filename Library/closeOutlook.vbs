On Error Resume Next
Set Outlook = GetObject(, "Outlook.Application")
If Err = 0 Then
Outlook.Quit()
End If