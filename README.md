# GST-Automation-Reports-Consolidation
Consolidation of Returns GSTR 2A to match input tax credits
_______________________________________
File "Steps to depict working of macros.docx" contains all the steps and information to help describe the steps in using the project

The project was created for BFSI Sector to consolidate downloaded GSTR 2A from GST website through Excel VBA accuracy and speed as the volume of such files is huge and accuracy is of utmost important to match Input Tax Credit. 




#EXCEL VBA CODE



Sub ConslidateWorkbooks()


Dim MyTimer As Double
Dim StartTime As Double
Dim MinutesElapsed As String
Dim EstimatedTotalTime As String

'Remember time when macro starts'
StartTime = Timer


Application.ScreenUpdating = False
     
    Dim xRow As Long
    Dim xDirect$, xFname$, InitialFoldr$
     
    InitialFoldr$ = Cells(1, 2).Value '<<< Startup folder to begin searching from'
     
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
 

 
 
 
' copy sheet b2b'
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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    

    End If
    

' copy sheet B2BA'

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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
     
     
     ' copy sheet CDNR'

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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
    
         ' copy sheet CDNRA'

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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
     
         ' copy sheet ISD'

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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
             ' copy sheet ISDA'

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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
     
     
     
             ' copy sheet TDS'
             
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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
    
             ' copy sheet TDSA'

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

    Selection.Copy
    
    Windows("Output.xlsx").Activate
    
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End If
    
    
    
             ' copy sheet TCS'

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





Sub deleteSubtotal()

End Sub

Dim MyTimer As Double
Dim StartTime As Double
Dim MinutesElapsed As String
Dim EstimatedTotalTime As String

'Remember time when macro starts'
        StartTime = Timer



Windows("Output.xlsx").Activate

Sheets("B2B").Select

Last = Cells(Rows.Count, "C").End(xlUp).Row

For I = Last To 2 Step -1

If Right(Cells(I, "C").Value, 6) = "-Total" Then

    Cells(I, "C").EntireRow.delete

End If

        MyTimer = Timer
             
        MinutesElapsed = Format((MyTimer - StartTime) / 86400, "hh:mm:ss")
        Application.StatusBar = "Combined Output file created. Progress of removing totals from B2B: " & Last - ActiveCell.Row & "Out of Total" & Last & " Rows Pending " & MinutesElapsed
        DoEvents

Next I


Sheets("B2BA").Select

Last = Cells(Rows.Count, "F").End(xlUp).Row

For I = Last To 2 Step -1

        StartTime = Timer

If Right(Cells(I, "F").Value, 6) = "-Total" Then

    Cells(I, "F").EntireRow.delete

End If

        MyTimer = Timer
         
        MinutesElapsed = Format((MyTimer - StartTime) / 86400, "hh:mm:ss")
        Application.StatusBar = "Combined Output file created. Progress of removing totals from B2BA: " & Last - ActiveCell.Row & "Out of Total" & Last & " Rows Pending " & MinutesElapsed
        DoEvents


Next I



Sheets("CDNR").Select

Last = Cells(Rows.Count, "D").End(xlUp).Row

For I = Last To 2 Step -1

        StartTime = Timer

If Right(Cells(I, "D").Value, 6) = "-Total" Then

    Cells(I, "D").EntireRow.delete

End If
Next I

        MyTimer = Timer
         
        MinutesElapsed = Format((MyTimer - StartTime) / 86400, "hh:mm:ss")
        Application.StatusBar = "Combined Output file created. Progress of removing totals from CDNR: " & Last - ActiveCell.Row & "Out of Total" & Last & " Rows Pending " & MinutesElapsed
        DoEvents


Sheets("CDNRA").Select

Last = Cells(Rows.Count, "H").End(xlUp).Row

For I = Last To 2 Step -1

        StartTime = Timer

If Right(Cells(I, "H").Value, 6) = "-Total" Then

    Cells(I, "H").EntireRow.delete

End If

        MyTimer = Timer
         
        MinutesElapsed = Format((MyTimer - StartTime) / 86400, "hh:mm:ss")
        Application.StatusBar = "Combined Output file created. Progress of removing totals from B2B: " & Last - 2 & " Rows Pending " & MinutesElapsed
        DoEvents

Next I




Sheets("ISD").Select

Last = Cells(Rows.Count, "E").End(xlUp).Row

For I = Last To 2 Step -1

If Right(Cells(I, "E").Value, 6) = "-Total" Then

    Cells(I, "E").EntireRow.delete

End If
Next I




Sheets("ISDA").Select

Last = Cells(Rows.Count, "H").End(xlUp).Row

For I = Last To 2 Step -1

If Right(Cells(I, "H").Value, 6) = "-Total" Then

    Cells(I, "H").EntireRow.delete

End If
Next I



Sheets("TDS").Select

Last = Cells(Rows.Count, "C").End(xlUp).Row

For I = Last To 2 Step -1

If Right(Cells(I, "C").Value, 6) = "-Total" Then

    Cells(I, "C").EntireRow.delete

End If
Next I



Sheets("TDSA").Select

Last = Cells(Rows.Count, "C").End(xlUp).Row

For I = Last To 2 Step -1

If Right(Cells(I, "C").Value, 6) = "-Total" Then

    Cells(I, "C").EntireRow.delete

End If
Next I

Sheets("TCS").Select

Last = Cells(Rows.Count, "C").End(xlUp).Row

For I = Last To 2 Step -1

If Right(Cells(I, "C").Value, 6) = "-Total" Then

    Cells(I, "C").EntireRow.delete

End If
Next I

End Sub




Sub OutputdeleteSubtotal()


    Windows("Output.xlsx").Activate
    Sheets("B2B").Select
    Call B2B_RemoveSubTotals
    
    Application.StatusBar = "B2B Processed. Pending B2BA and CDNR"
    
    Sheets("B2BA").Select
    Call B2BA_RemoveSubTotals
    
    Application.StatusBar = "B2B and B2BA Processed. Pending CDNR"
    
    Sheets("CDNR").Select
    Call CDNR_RemoveSubTotals
    
    Application.StatusBar = "Processed All Three"
    
    ActiveWorkbook.Save


End Sub



Sub B2B_RemoveSubTotals()

' Macro3 Macro'


    Range("A12").Select
    Selection.EntireColumn.Insert
    Rows("7:7").Select
    Selection.Insert Shift:=xlDown

    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    Range("A3").Select
    Application.Goto Reference:="R2C1:R1000000C1"
    Selection.FillDown
    Columns("A:A").Select
    Range("A2").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("7:7").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$7:$AJ$1000000").AutoFilter Field:=4, Criteria1:= _
        "=*-Total*", Operator:=xlAnd
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.ClearContents
    Selection.AutoFilter
    Range("A7").Select
    Selection.EntireColumn.Insert
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    Application.Goto Reference:="R1C1:R1000000C1"
    Selection.FillDown
    Rows("7:7").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("B2B").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("B2B").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "B7"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("B2B").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:B").Select
    Selection.delete Shift:=xlToLeft
    Rows("7:7").Select
    Selection.delete Shift:=xlUp
End Sub




Sub B2BA_RemoveSubTotals()

' Macro1 Macro'



    Range("A11").Select
    Selection.EntireColumn.Insert
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    Range("A3").Select
    Application.Goto Reference:="R2C1:R1000000C1"
    Selection.FillDown
    Columns("A:A").Select
    Range("A2").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("8:8").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown
    Selection.AutoFilter
    ActiveSheet.Range("$A$8:$R$1000001").AutoFilter Field:=7, Criteria1:= _
        "=*-Total*", Operator:=xlAnd
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.ClearContents
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    Application.Goto Reference:="R1C1:R1000000C1"
    Selection.FillDown
    Rows("8:8").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("B2BA").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("B2BA").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "B8"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("B2BA").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:B").Select
    Selection.delete Shift:=xlToLeft
    Rows("8:8").Select
    Selection.delete Shift:=xlUp
End Sub



Sub CDNR_RemoveSubTotals()

' Macro5 Macro'



    Range("A10").Select
    Selection.EntireColumn.Insert
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    Range("A3").Select
    Application.Goto Reference:="R2C1:R1000000C1"
    Selection.FillDown
    Columns("A:A").Select
    Range("A2").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("7:7").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("C7").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("F7").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("G7").Select
    ActiveCell.FormulaR1C1 = "6"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "7"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("J7").Select
    ActiveCell.FormulaR1C1 = "9"
    Range("K7").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("L7").Select
    ActiveCell.FormulaR1C1 = "11"
    Range("M7").Select
    ActiveCell.FormulaR1C1 = "12"
    Range("N7").Select
    ActiveCell.FormulaR1C1 = "13"
    Range("O7").Select
    ActiveCell.FormulaR1C1 = "14"
    Range("B8").Select
    ActiveWindow.ScrollColumn = 1
    Rows("7:7").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$7:$P$1000001").AutoFilter Field:=5, Criteria1:= _
        "=*-Total*", Operator:=xlAnd
    Range("A7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("A7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.ClearContents
    Selection.AutoFilter
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    Application.Goto Reference:="R1C1:R1000000C1"
    Selection.FillDown
    Rows("7:7").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("CDNR").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CDNR").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "B7"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("CDNR").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:B").Select
    Selection.delete Shift:=xlToLeft
    Rows("7:7").Select
    Selection.delete Shift:=xlUp
End Sub




Sub Clear_FilesInfo()

' Macro6 Macro'



    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("J5").Select
    Selection.ClearContents
End Sub


Sub SelectSourceFolder()
    Dim diaFolder As FileDialog

    ' Open the file dialog'
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    Cells(1, 1).Value = diaFolder.SelectedItems(1) & "\"
    
    Set diaFolder = Nothing
End Sub


Sub SelectOutputFile()
    Dim diaFolder As FileDialog

    ' Open the file dialog'
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    Cells(1, 14).Value = diaFolder.SelectedItems(1) & "\Output.xlsx"
    MsgBox diaFolder.SelectedItems(1)

    Set diaFolder = Nothing
End Sub
