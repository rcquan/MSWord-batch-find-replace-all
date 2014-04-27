Attribute VB_Name = "BatchReplaceAll"
 Option Explicit

Public Sub BatchReplaceAll()

Dim FirstLoop As Boolean
Dim myFile As String
Dim FilePath As String
Dim myDir As String
Dim myDoc As Document
Dim Response As Long


'Allows user to select directory and
'assigns input to string variable

With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    If .Show <> -1 Then MsgBox "No folder selected! Exiting sub...":
    myDir = .SelectedItems(1)
End With

'Pauses run to make sure that user does
'not execute VBA on wrong folder

If MsgBox("Are you sure this is the correct working " & _
    "directory?", vbYesNo) = vbNo Then Exit Sub


FilePath = "" & myDir & Chr(92) & ""

'Error handler for whenever
'the FindReplace dialog is closed

On Error Resume Next

'Boolean expression to test whether first loop
'Used so that the FindReplace dialog is
'only displayed for the first Word document

FirstLoop = True

'Set the directory and type of file to batch process
'NOTE .doc extension picks up .docx files as well

myFile = Dir$(FilePath & "*.doc")

While myFile <> ""

    'Open document
    Set myDoc = Documents.Open(FilePath & myFile)

    If FirstLoop Then

        'Display dialog on first loop only

        Dialogs(wdDialogEditReplace).Show

        FirstLoop = False

        Response = MsgBox("Do you want to process " & _
        "the rest of the files in this folder", vbYesNo)
        If Response = vbNo Then Exit Sub

    Else

        'On subsequent loops (files), original
        'ReplaceAll is executed without
        'displaying the dialog box again

        With Dialogs(wdDialogEditReplace)
            .ReplaceAll = 1
            .Execute
        End With

    End If

    'Close the modified document after saving changes

    myDoc.Close SaveChanges:=wdSaveChanges

    'Move to next file in folder

    myFile = Dir$()

Wend

End Sub


