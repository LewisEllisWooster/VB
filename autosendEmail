Private Sub AutoSend_Click()
'Email Userform 3
''''''''''''''''''''''''''''
''''Textbox parameters based on userform input


Dim Mail_Object, Mail_Single As Variant
 
        Email_Subject = "Complaint" & " " & "Section:" & " " & UserForm2.ComboBox1.Value
 
        'nameList = Sheets("Sheet1").Range("A1").Value
        Email_Send_To = Me.TextBox1.Value
 
 
        Email_Cc = Me.TextBox2.Value
        Email_Bcc = Me.TextBox3.Value
        Email_Body = Me.TextBox4.Value
 
        Set Mail_Object = CreateObject("Outlook.Application")
        Set Mail_Single = Mail_Object.CreateItem(o)
If Application.Version < 12 Then
With Mail_Single
            .Subject = Email_Subject
            .To = Email_Send_To
 
            .CC = Email_Cc
            .BCC = Email_Bcc
            .Body = Email_Body
            .send
End With
ElseIf Application.Version >= 12 Then

Dim selected As String
With Application.FileDialog(msoFileDialogOpen)
    .InitialFileName = newpath10
    .AllowMultiSelect = True
    .Show
     If .SelectedItems.Count = 1 Then
        selected = .SelectedItems(1)
    End If
End With
'If selected <> "" Then
   'Open selected For Output As #n
'End If


With Mail_Single
            .Subject = Email_Subject
            .To = Email_Send_To
 
            .CC = Email_Cc
            .BCC = Email_Bcc
            .Body = Email_Body
            If selected <> "" Then
            .Attachments.Add (selected)
            Else
            End If
            '.Display
            .send
            
            End With
End If
        MsgBox "E-mail successfully sent"
        Application.DisplayAlerts = False

Me.Hide
MsgBox ("Please Press Submit Complaint To Add It To The Database")
UserForm2.Show
End Sub
