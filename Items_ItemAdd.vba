Private Sub Items_ItemAdd(ByVal item As Object)

  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem

  If TypeName(item) = "MailItem" Then
    revealURLdomainInNewMail item

    ' *************************
    ' Can do other things here.
    ' *************************

  End If

  
ProgramExit:
  Exit Sub

ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit

End Sub

