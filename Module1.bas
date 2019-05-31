Attribute VB_Name = "Module1"

Sub ConslidateWorkbooks()


Dim MyTimer As Double
Dim StartTime As Double
Dim MinutesElapsed As String
Dim EstimatedTotalTime As String

'Remember time when macro starts
StartTime = Timer


Application.ScreenUpdating = False
     
    Dim xRow As Long
    Dim xDirect$, xFname$, InitialFoldr$
     
    InitialFoldr$ = Cells(1, 2).Value '<<< Startup folder to begin searching from
     
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Please select a folder to list Files from"
        .InitialFileName = InitialFoldr$
        .Show
        
        If .SelectedItems.Count <> 0 Then
            xDirect$ = .SelectedItems(1) & "\"
            xFname$ = Dir(xDirect$, 7)
            Do While xFname$ <> ""
                Cells(2, 1).Select
                ActiveCell.Offset(xRow) = xFname$
                xRow = xRow + 1
                xFname$ = Dir
            Loop
        End If
    End With




'Created by Sumit Bansal from https://trumpexcel.com
Dim folderPath As String
Dim FileName As String
Dim Sheet As Worksheet
Dim AA As Long
Dim AB As Long
Dim Wkb As Long






AA = Application.WorksheetFunction.CountA(Range("A1:A1000000"))


Workbooks.Open FileName:=Cells(1, 14).Value






AB = 1


For Wkb = 2 To AA

AB = AB + 1

Windows("Combine Excel Files 4.xlsm").Activate


Cells(5, 10).Value = AA

folderPath = Cells(1, 1).Value

FileName = Cells(AB, 1).Value

Cells(AB, 2).Value = folderPath & FileName
Cells(AB, 3).Value = Wkb

Workbooks.Open FileName:=folderPath & FileName, ReadOnly:=False
 

 
 
 
' copy sheet b2b
    Windows(FileName).Activate
    Sheets("B2B").Select
    Rows("1:6").Select
    Selection.delete Shift:=xlUp
    
    
        
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    Dim T As Long
    T = Application.WorksheetFunction.CountA(Range("O1:O1048575"))
    Cells(1, 16).Value = FileName
    If Cells(1, 1) <> "" Then
        Range(Cells(1, 16), Cells(T, 16)).Select
    Selection.FillDown
    End If
    
    
    
    If Cells(1, 1).Value <> "" Then
    
    Windows("Output.xlsx").Activate
    Sheets("B2B").Select
    Dim A As Long
    A = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 7
    Cells(A, 1).Select
    
        Windows(FileName).Activate
    Range("A1:P1").Select
    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    

    End If
    

' copy sheet B2BA

    Windows(FileName).Activate

    Sheets("B2BA").Select
    Rows("1:7").Select
    Selection.delete Shift:=xlUp
    
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    

    T = Application.WorksheetFunction.CountA(Range("C1:C1048575"))
    Cells(1, 18).Value = FileName
    If Cells(1, 1) <> "" Then
        Range(Cells(1, 18), Cells(T, 18)).Select
    Selection.FillDown
    End If
    
    
    If Cells(1, 1).Value <> "" Then
            
    Windows("Output.xlsx").Activate
    Sheets("B2BA").Select
    Dim B As Long
    B = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 8
    Cells(B, 1).Select
    
    
    Windows(FileName).Activate
        Range("A1:R1").Select
    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
     
     
     ' copy sheet CDNR

    Windows(FileName).Activate
    Sheets("CDNR").Select
    Rows("1:6").Select
    Selection.delete Shift:=xlUp
    
    
        
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    T = Application.WorksheetFunction.CountA(Range("A1:A1048575"))
    Cells(1, 15).Value = FileName
    If Cells(1, 1) <> "" Then
        Range(Cells(1, 15), Cells(T, 15)).Select
        Selection.FillDown
    End If
    
    
    
    If Cells(1, 1).Value <> "" Then
    
    Windows("Output.xlsx").Activate
    Sheets("CDNR").Select
    Dim C As Long
    C = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 7
    Cells(C, 1).Select
    
        Windows(FileName).Activate
    Range("A1:O1").Select
    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
    
         ' copy sheet CDNRA

    Windows(FileName).Activate

    Sheets("CDNRA").Select
    Rows("1:7").Select
    Selection.delete Shift:=xlUp
    
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    T = Application.WorksheetFunction.CountA(Range("D1:D1048575"))
    Cells(1, 18).Value = FileName
    If Cells(1, 1).Value <> "" Then
        Range(Cells(1, 18), Cells(T, 18)).Select
        Selection.FillDown
    End If
    
    
    
    If Cells(1, 1).Value <> "" Then
            
    Windows("Output.xlsx").Activate
    Sheets("CDNRA").Select
    Dim D As Long
    D = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 8
    Cells(D, 1).Select
    
    
    Windows(FileName).Activate
        Range("A1:R1").Select
    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
     
         ' copy sheet ISD

    Windows(FileName).Activate
    Sheets("ISD").Select
    Rows("1:6").Select
    Selection.delete Shift:=xlUp
    
    
        
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    T = Application.WorksheetFunction.CountA(Range("B1:B1048575"))
    Cells(1, 16).Value = FileName
    If Cells(1, 1) <> "" Then
        Range(Cells(1, 16), Cells(T, 16)).Select
        Selection.FillDown
    End If
    
    
    If Cells(1, 1).Value <> "" Then
    
    Windows("Output.xlsx").Activate
    Sheets("ISD").Select
    Dim E As Long
    E = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 7
    Cells(E, 1).Select
    
        Windows(FileName).Activate
    Range("A1:P1").Select
    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
             ' copy sheet ISDA

    Windows(FileName).Activate

    Sheets("ISDA").Select
    Rows("1:7").Select
    Selection.delete Shift:=xlUp
    
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    T = Application.WorksheetFunction.CountA(Range("E1:E1048575"))
    Cells(1, 19).Value = FileName
    If Cells(1, 1) <> "" Then
        Range(Cells(1, 19), Cells(T, 19)).Select
        Selection.FillDown
    End If
    
    
    If Cells(1, 1).Value <> "" Then
            
    Windows("Output.xlsx").Activate
    Sheets("ISDA").Select
    Dim F As Long
    F = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 8
    Cells(F, 1).Select
    
    
    Windows(FileName).Activate
        Range("A1:S1").Select
    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
     
     
     
             ' copy sheet TDS

    Windows(FileName).Activate
    Sheets("TDS").Select
    Rows("1:6").Select
    Selection.delete Shift:=xlUp
    
    
        
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    T = Application.WorksheetFunction.CountA(Range("A1:A1048575"))
    
    Cells(1, 8).Value = FileName
    If T = 1 And Cells(1, 1) <> "" Then
'        Range(Cells(1, 8), Cells(1, 8)).Select
        Range("A1:H1").Select
        
    ElseIf T > 1 And Cells(1, 1) <> "" Then
            Range(Cells(1, 8), Cells(T, 8)).Select
            Selection.FillDown
    End If
    
    
    If Cells(1, 1).Value <> "" Then
    
    Windows("Output.xlsx").Activate
    Sheets("TDS").Select
    Dim G As Long
    G = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 7
    Cells(G, 1).Select
    
        Windows(FileName).Activate
'    Range("A1:H1").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
    
             ' copy sheet TDSA

    Windows(FileName).Activate
    Sheets("TDSA").Select
    Rows("1:6").Select
    Selection.delete Shift:=xlUp
    
    
        
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    T = Application.WorksheetFunction.CountA(Range("A1:A1048575"))
    Cells(1, 9).Value = FileName
    If Cells(1, 1) <> "" Then
        Range(Cells(1, 9), Cells(T, 9)).Select
        Selection.FillDown
    End If
    
    
    If Cells(1, 1).Value <> "" Then
    
    Windows("Output.xlsx").Activate
    Sheets("TDSA").Select
    Dim H As Long
    H = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 7
    Cells(H, 1).Select
    
        Windows(FileName).Activate
    Range("A1:I1").Select
    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
    
    
             ' copy sheet TCS

    Windows(FileName).Activate
    Sheets("TCS").Select
    Rows("1:6").Select
    Selection.delete Shift:=xlUp
    
    
        
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.delete
    
    T = Application.WorksheetFunction.CountA(Range("A1:A1048575"))
    Cells(1, 10).Value = FileName
    If Cells(1, 1) <> "" And Application.WorksheetFunction.CountA(Range("A1:A1000000")) > 1 Then
        Range(Cells(1, 10), Cells(T, 10)).Select
        Selection.FillDown
    ElseIf Cells(1, 1) <> "" And Application.WorksheetFunction.CountA(Range("A1:A1000000")) = 1 Then
        Range(Cells(1, 10), Cells(T, 10)).Select
    
    End If
    
    
    If Cells(1, 1).Value <> "" Then
    
    Windows("Output.xlsx").Activate
    Sheets("TCS").Select
    Dim I As Long
    I = Application.WorksheetFunction.CountA(Range("A7:A1000000")) + 7
    Cells(I, 1).Select
    
        Windows(FileName).Activate
    
    If Application.WorksheetFunction.CountA(Range("A7:A1000000")) > 1 Then
    
    Range("A1:J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    ElseIf Application.WorksheetFunction.CountA(Range("A7:A1000000")) = 1 Then
    
    Range("A1:J1").Select
    Selection.Copy
    
    
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End If
    End If
    
     
 'For Each Sheet In ActiveWorkbook.Sheets
 
 
 'Sheet.Copy After:=ThisWorkbook.Sheets(1)
 
 
 'Next Sheet
 
 
 Windows(FileName).Activate
 ActiveWindow.Close False
 
 
 
 Application.ScreenUpdating = True
 
    MyTimer = Timer
         
    MinutesElapsed = Format((MyTimer - StartTime) / 86400, "hh:mm:ss")
    EstimatedTotalTime = Format(((MyTimer - StartTime) / Wkb * AA) / 86400, "hh:mm:ss")
    
    Application.StatusBar = "Progress: " & Wkb & " of " & AA & ": " & (Wkb / AA) * 100 & "% in elapsed time " & MinutesElapsed & " out of Total Estimated Time of : " & EstimatedTotalTime
    DoEvents


Application.ScreenUpdating = False


Next Wkb


Windows("Output.xlsx").Activate
ActiveWorkbook.Save
Windows("Combine Excel Files 4.xlsm").Activate
ActiveWorkbook.Save

Application.ScreenUpdating = True
End Sub




