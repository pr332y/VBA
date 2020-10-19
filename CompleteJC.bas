Attribute VB_Name = "Module1"
Option Explicit
Public job As String
Sub CompleteJC()
JCHours
PasteToSheets
SubtotalForLC
ThisWorkbook.Protect ("gofigure")
End Sub

Sub JCHours()
'takes the jobcost entries for a particular job across multiple years and places them in one file to be analyzed together
Dim path As String, filename As String, LastRow As Double, sheet As Worksheet, wbname As String, job As String, n As Integer
Application.ScreenUpdating = False
Application.DisplayAlerts = False
path = "K:\TERRY\JOBCOSTS\JobCost Summary\"
job = InputBox("Enter Job Number", "Job") 'delivery order not required
If job = "" Then Exit Sub 'Don't run subroutine if no input is entered
ThisWorkbook.SaveAs filename:="K:\TERRY\JOBCOSTS\Hours by Job\" & job & " All.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
wbname = ThisWorkbook.Name
ActiveSheet.Name = job
LastRow = 2 'actually new row
filename = Dir(path & "*Jobcost 20*" & "*.xlsm")
Do While filename <> ""
    'reset counter for new file
    n = 0
    Workbooks.Open (path & filename), True, , , "gofigure"
    For Each sheet In Worksheets
        'If the beginning of the sheet name doesn't match the input and the copy paste loop has already run, the rest of the workbook doesn't need to be searched
        If Not Left(sheet.Name, Len(job)) = job And n <> 0 Then Exit For
        If Left(sheet.Name, Len(job)) = job Then
            Sheets(sheet.Name).Activate
            Rows("2:" & Workbooks(filename).Sheets(sheet.Name).UsedRange.Rows.Count).Select
            Selection.Copy
            Windows(wbname).Activate 'workbook where macro is running
            Rows(LastRow).Select
            ActiveSheet.Paste
            LastRow = ActiveSheet.Range("G1").End(xlDown).Row + 1
            Windows(filename).Activate 'yearly workbook file
            'number of times sheets have been copied and pasted
            n = n + 1
        End If
    Next sheet
    Workbooks(filename).Close savechanges:=False
    filename = Dir()
Loop
ActiveSheet.Cells.RemoveSubtotal
'Sort by Delivery Order, LC and Work Date
ActiveSheet.Sort.SortFields.Clear
ActiveSheet.Sort.SortFields.Add Key:=Range("F2:F" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveSheet.Sort.SortFields.Add Key:=Range("G2:G" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveSheet.Sort.SortFields.Add Key:=Range("A2:A" & LastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:T" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
ThisWorkbook.Save
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub PasteToSheets()
'
' PasteToSheets Macro
'
Dim n As Integer, nn As Integer, job As String, DeliveryOrder As Double, sheetname As String, sheetnm As String, LastRow As Long, sheet As Worksheet
'Separate entries into sheets by change order
n = 1
sheetname = ""
job = ActiveSheet.Name
'If the delivery order is entered into the input, the data cannot be separated into sheets
If InStr(job, "-") <> 0 Then Exit Sub
LastRow = Sheets(job).UsedRange.Rows.Count
Do While n < LastRow
    n = n + 1
    DeliveryOrder = Sheets(job).Range("F" & n).Value
    sheetnm = job & "-" & DeliveryOrder
    If sheetname <> sheetnm Then
        sheetname = sheetnm
        If Not sheetExists(sheetname) Then
            Set sheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            sheet.Name = sheetname
            nn = 2
            Sheets(job).Rows(1).EntireRow.Copy
            Sheets(sheetname).Range("A1").Select
            Sheets(sheetname).Paste
        Else
            nn = Sheets(sheetnm).UsedRange.Rows.Count + 1
        End If
    End If
        Sheets(job).Rows(n).EntireRow.Copy
        Sheets(sheetname).Range("A" & nn).Select
        ActiveSheet.Paste
        nn = nn + 1
Loop
ThisWorkbook.Save
Sheets(job).Activate
End Sub
Private Function sheetExists(SheetToFind As String, Optional sheet As Worksheet) As Boolean
sheetExists = False
For Each sheet In Worksheets
    If SheetToFind = sheet.Name Then
        sheetExists = True
        Sheets(SheetToFind).Select
    End If
Next sheet
End Function
Sub SubtotalForLC()
'
' SubtotalForLC Macro
'
Dim sheet As Worksheet, n As Long, SubtotalLine As Long, hours As Range, DeliveryOrder As Double
job = ActiveSheet.Name
For Each sheet In Worksheets
    If Sheets.Count = 1 Or (Sheets.Count <> 1 And sheet.Name <> job) Then
        sheet.Select
        Cells.Select
    'Total hours and billing amt by each change in lc
        Selection.Subtotal GroupBy:=7, Function:=xlSum, TotalList:=Array(9, 13), _
            Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        For n = 2 To ActiveSheet.Range("G1").End(xlDown).Row
            If Not WorksheetFunction.IsNumber(Range("G" & n)) Then
        'formating the subtotal line
                sheet.Rows(n).EntireRow.Select
                Selection.Font.Bold = True
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            End If
        Next n
    End If
Next sheet
ThisWorkbook.Save
End Sub
