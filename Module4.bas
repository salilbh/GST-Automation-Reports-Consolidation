Attribute VB_Name = "Module4"
Sub Clear_FilesInfo()
Attribute Clear_FilesInfo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("J5").Select
    Selection.ClearContents
End Sub
