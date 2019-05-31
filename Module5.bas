Attribute VB_Name = "Module5"
Sub SelectSourceFolder()
    Dim diaFolder As FileDialog

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    Cells(1, 1).Value = diaFolder.SelectedItems(1) & "\"
    
    Set diaFolder = Nothing
End Sub


Sub SelectOutputFile()
    Dim diaFolder As FileDialog

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    Cells(1, 14).Value = diaFolder.SelectedItems(1) & "\Output.xlsx"
    MsgBox diaFolder.SelectedItems(1)

    Set diaFolder = Nothing
End Sub
