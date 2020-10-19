Private Sub CommandButton1_Click()
Dim r As Integer, c As Integer, total As Double
On Error Resume Next
For c = 2 To ActiveSheet.UsedRange.Columns.Count Step 2
    total = 0
    For r = 7 To ActiveSheet.UsedRange.Rows.Count - 4
        If Not IsEmpty(Cells(r, c).Value) Then
	    'get unit
            Select Case Right(Cells(r, c).Value, 2)
		'convert number part of string to double, make everything in GB, add to running total
                Case "GB" Or "gb" Or "Gb" Or "gB"
                    total = total + CDbl(Left(Cells(r, c), Len(Cells(r, c)) - 3))
                Case "MB" Or "mb" Or "Mb" Or "mB"
                    total = total + (CDbl(Left(Cells(r, c), Len(Cells(r, c)) - 3)) / 1024)
                Case "KB" Or "kb" Or "Kb" Or "kB"
                    total = total + (CDbl(Left(Cells(r, c), Len(Cells(r, c)) - 3)) / (1024 ^ 2))
            End Select
        End If
    Next r
    Cells(r + 1, c).Value = total & " GB"
    'Sum number of files
    Cells(r + 1, c + 1).Value = WorksheetFunction.Sum(Range(Cells(7, c + 1), Cells(r, c + 1)))
Next c
End Sub