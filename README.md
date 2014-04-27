MSWord-batch-find-replace-all
=============================

A Visual Basic macro that allows the user to find and replace a string of text for all MS Word documents in a specified directory and saves the modified files in the same directory.

System Requirements
===================
* MS Word 2007 or higher
* Windows Operating system

Installation
============
There are several options to deploy this package:

1. **.docm** - Open the .docm file (Macro-enabled MS Word Template) and follow the instructions
2. **.bas** - Under the Developer tab in MS Word, navigate to Visual Basic -> File -> Import File. Choose the .bas file from the directory.
3. **.vbs** - Open the .vbs in a text editor such as Sublime Text or Notepad. Under the Developer tab in MS word, navigate to Visual Basic -> Insert -> Module. Copy and paste the contents of the .vbs file into the blank module. Exit Visual Basic. You can now call the VBA macro.

How to Use
=================
1. Save a copy of your original files in case you want to refer back to them. This macro will overwrite current files.
2. Prepare a folder of .doc or .docx files you wish to execute a batch find and replace command on. For example, you might have 200 Word documents with the same subject title and you wish to update the subject title to reflect new changes. 
3. Execute the MSWord-batch-find-replace-all macro.
4. Specify the directory you wish to execute macro on (where all your files are contained).
5. Confirm selection. The first document will appear.
6. Find and replace desired text. When finished, click "close" in the dialog box.
7. Choose "Yes" to apply find and replace all to the rest of the documents in the folder. Choose "No" to exit.
8. Done! Check documents to verify changes.
