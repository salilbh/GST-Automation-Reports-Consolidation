Attribute VB_Name = "Module3"

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
'
' Macro3 Macro
'

'
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
'
' Macro1 Macro
'

'
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
'
' Macro5 Macro
'

'
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


