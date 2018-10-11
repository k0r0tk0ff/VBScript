Sub getMaterialIDFromErrorMsg()
    Dim Inbox As Outlook.MAPIFolder						 ' Inbox Folder
    Dim ErrorMessageFolder As Outlook.MAPIFolder 		 ' Folder with error messages
    Set Inbox = Session.GetDefaultFolder(olFolderInbox)
    Set ErrorMessageFolder = Inbox.Folders("send errors").Folders("Error registerMsg")
    MsgBox ("Finds: " & ErrorMessageFolder.Items.Count)
    Dim Item As Outlook.MailItem 						' Mail message
    Set Items = ErrorMessageFolder.Items 				' All messages in folder
													'RegExp for match <MaterialID>ID in digits</MaterialID>

    Dim Reg1 As RegExp
    Dim M1 As MatchCollection
    Dim M As Match
    Set Reg1 = New RegExp
    With Reg1
        .Pattern = "<MaterialID>\d*</MaterialID>"
        .Global = True
    End With

    Dim OutputFile As Integer
    OutputFile = FreeFile() 							'Create File
    Open "E:\materialId.txt" For Output As #OutputFile 	'Open File
														'Find by RegExp Matches in all messages in folder
    For Each Item In Items
       If TypeOf Item Is Outlook.MailItem Then
            Dim oMail As Outlook.MailItem: Set oMail = Item
            If Reg1.Test(oMail.Body) Then
                Set M1 = Reg1.Execute(oMail.Body)
                For Each M In M1
                    Print #OutputFile, M.Value 
														'Write to File
														'If match counf > 1 break loop
                    If M1.Count > 1 Then Exit For
                Next
            End If
       End If
    Next

    Close #OutputFile 'Close File
End Sub
