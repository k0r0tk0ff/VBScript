
' Скрипт выполняет поиск в заданной папке ("PROBA") писем,
'  в которых имеется строка "Первичный ключ: {INTEGER_VALUE}"
'  и сохраняет найденные строки в файле "E:\materialId.txt"

Sub GetValueUsingRegEx()
  ' Need to add
  ' Microsoft VBScript Regular Expressions 5.5
 
    Dim olMail As Outlook.MailItem
    Dim Reg1 As RegExp
    Dim M1 As MatchCollection
    Dim M As Match
    
    Dim FolderName As String
    Dim ErrorMessageFolder As Folder
    
    FolderName = "PROBA"
    Set ErrorMessageFolder = FindInFolders(Application.Session.Folders, FolderName)
    
    Dim Item As Outlook.MailItem                       'Mail message
    Set Items = ErrorMessageFolder.Items               'All messages
    MsgBox ("Finds: " & ErrorMessageFolder.Items.Count)
	
	
	' \s* = скрытые пробелы
    ' \d* = цифры
    ' \w* = цифро-буквенные выражения
    
    Set Reg1 = New RegExp
    With Reg1
        .Pattern = "Первичный ключ: \d*"
        .Global = True
    End With
    
    Dim OutputFile As Integer
    OutputFile = FreeFile()                             'Create File
    Open "E:\materialId.txt" For Output As #OutputFile  'Open File
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

Function FindInFolders(TheFolders As Outlook.Folders, Name As String)
  Dim SubFolder As Outlook.MAPIFolder
   
  On Error Resume Next
   
  Set FindInFolders = Nothing
   
  For Each SubFolder In TheFolders
    If LCase(SubFolder.Name) Like LCase(Name) Then
      Set FindInFolders = SubFolder
      Exit For
    Else
      Set FindInFolders = FindInFolders(SubFolder.Folders, Name)
      If Not FindInFolders Is Nothing Then Exit For
    End If
  Next
End Function
