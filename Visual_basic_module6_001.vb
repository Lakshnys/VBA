Visual basic

Sub Autofilter_subtotal_and_Column_to_Output_File_module6()

Dim count_col, count_row As Integer
Dim orgi, output As Worksheet
Dim a As Integer
Dim b As Range
Dim v As Variant
Dim Dis As Workbook
Dim Org As Workbook
Dim Sno As Integer
Dim ws As Integer

Dim NextCol As Long
ws = 32
v = InputBox("Enter Value")

'Looping start

For i = 1 To 31 ' input value based on number of sheets to be considered for AutoFilteration
Sno = i 'Worksheet index no.
'open destination workbook

Set Dis = Workbooks.Open("C:\Users\VAS\Desktop\Test\001. AutoFilteration_VBA\Output.xlsx")
'Set Org = Workbooks.Open("C:\Users\VAS\Desktop\Test\001. AutoFilteration_VBA\Autofilteration_past_to_another_book.xlsm")

'ThisWorkbook.Sheets("Sheet5").Cells.ClearContents
'Worksheets("Jan_2020_1").Activate

'ThisWorkbook.Sheets(Sheet1).Cells.ClearContents
'Worksheets(Sheet2).Activate

ThisWorkbook.Worksheets(ws).Cells.ClearContents
ThisWorkbook.Worksheets(Sno).Activate

'Set orgi = Worksheets("Jan_2020_1")
'Set output = Worksheets("Sheet5")

'Set orgi = Worksheets(1)
'Set output = Worksheets(6)

count_col = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlToRight)))
count_row = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlDown)))

ActiveSheet.Range("A1").AutoFilter Field:=4, Criteria1:=v
Worksheets(Sno).Range(Cells(1, 1), Cells(count_row, count_col)).SpecialCells(xlCellTypeVisible).Copy
Application.DisplayAlerts = False
'Paste the copied values to ws sheet in the same workbook
Worksheets(ws).Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.DisplayAlerts = False
Application.CutCopyMode = False
Worksheets(Sno).ShowAllData
Worksheets(Sno).AutoFilterMode = False

'---------------'xxxxx----------------xxxxx----------------xxxxxx-------------------

Worksheets(ws).Activate
Range("A1").Select
a = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
'Set b = Range("E291:E2")
'MsgBox (a)
'a = a
'MsgBox (b)
'output.Range("E293").Select
Worksheets(ws).Cells(a + 1, 5).Select
'ActiveCell.FormulaR1C1 = "=SUBTOTAL(1,R[-291]C:R[-1]C)"
'ActiveCell.Formula = "=SUBTOTAL(1,E291:E2" & Finalrow & ")"
ActiveCell.Formula = "=SUBTOTAL(1, E291:E2)"
Application.DisplayAlerts = False
'output.Range("E293").Select
Worksheets(ws).Cells(a + 1, 4) = "Average"
Application.DisplayAlerts = False
'Cell.Value = "Test"
'Worksheets(6).Range(Cells(292, 4), Cells(292, 5)).Copy
'Copy SubTotal Value from Cell(292,5)
Worksheets(ws).Range(Cells(292, 5), Cells(292, 5)).Copy 'Copy SubTotal value cell
Application.DisplayAlerts = False

'----------------xxxxxx---------------xxxxxx----------------xxxxxx--------------------
'Dis.Sheets("Sheet1").Range("D5").PasteSpecial Paste:=xlPasteValues
'.Range("IV1") = Row 1
'.Range("IV2") = Row 2
'.Range("IV3") = Row 3
'.Range("IV4") = Row 4
'.Range("IV5") = Row 5
'.Range("IV6") = Row 6
'.Range("IV7") = Row 7
'.Range("IV8") = Row 8
'.Range("IV9") = Row 9
'.Range("IV10") = Row 10
'SubTotal value past to next empty cloumn

Dis.Sheets("Sheet1").Range("IV8").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValues
Application.DisplayAlerts = False
ActiveSheet.Range("B1").Select
Application.CutCopyMode = False
ActiveSheet.Range("A2").Select
Selection.Copy
Application.DisplayAlerts = False
'Date past to next empty column
Dis.Sheets("Sheet1").Range("IV7").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.DisplayAlerts = False
Application.CutCopyMode = False
Worksheets(ws).Range(Cells(2, 5), Cells(292, 5)).Copy
Application.DisplayAlerts = False

Dis.Sheets("Sheet5").Range("IV2").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValues
ActiveSheet.Range("B1").Select
Application.CutCopyMode = False
ActiveSheet.Range("A2").Select
Selection.Copy
Application.DisplayAlerts = False
'Date past to next empty column
Dis.Sheets("Sheet5").Range("IV1").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
ActiveSheet.Range("A1").Select
Application.DisplayAlerts = False
'Dis.Close Savechanges:=True
Dis.Save

Next

With Dis.Sheets("Sheet1").Cells
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

With Dis.Sheets("Sheet5").Cells
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

Dis.Save

Set orgi = Nothing
Set output = Nothing
Set Dis = Nothing
Set Org = Nothing

End Sub
