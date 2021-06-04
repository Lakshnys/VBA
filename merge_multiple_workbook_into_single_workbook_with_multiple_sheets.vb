Sub MergeWorkbooks_to_single_workbook_multisheet____()

Dim FolderPath As String
Dim File As String
Dim i As Long


FolderPath = "C:\Users\VAS\Desktop\Test\Multiple_files_to_Single_Workbook\"

File = Dir(FolderPath)

Do While File <> ""

    Workbooks.Open FolderPath & File
        ActiveWorkbook.Worksheets(1).Copy _
            after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            ActiveSheet.Name = Replace(File, ".xlsx", "")
            Workbooks(File).Close

    File = Dir()

Loop

For i = 1 To 31

    Worksheets(MonthName(i, True)).Move after:=Worksheets(Worksheets.Count)
    
Next


End Sub