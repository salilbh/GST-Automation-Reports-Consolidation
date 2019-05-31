Attribute VB_Name = "Module2"
Sub deleteSubtotal()

End Sub

Dim MyTimer As Double
Dim StartTime As Double
Dim MinutesElapsed As String
Dim EstimatedTotalTime As String

'Remember time when macro starts
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

