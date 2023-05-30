Private Sub Document_Close()
    
    'WritePassword:=strPWD
    'obscure the filename and enable password. But make sure you'll be able to read it
    'save backups, potentially to a hidden folder somewhere obscure
    
    'the date is not really important in the file name, as it's saved in the parameters anyway
    
    'add ".doc", to the saved files, prior to opening them. They're saved without ".doc" on purpose.
    
    'no file extension as an extra protection, you'll add it yourself when opening the files
    'date obfuscated on purpose
    
    'TODO: have that file saved with attribute hidden

    Application.Dialogs(wdDialogFileSaveAs).Show
    
    ActiveDocument.SaveAs2 FileName:="C:\Users\User\Documents\" _
                & "processing saving " _
                & Format(Time, "s.m.h") _
                & "." & Format(Date, "m.d"), _
    FileFormat:=wdFormatDocumentDefault, _
    AddToRecentFiles:=False

End Sub
