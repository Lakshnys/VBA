Sub GetDataFromAllFilesInaFolder()
 'Tested and working 
 'Set a reference to Microsoft Scripting Runtime by using
    'Tools > References in the Visual Basic Editor (Alt+F11)
Dim objFile As Scripting.File
Dim objFolder As Scripting.Folder
Dim owbk As Workbook, twbk As Worksheet, ws As Worksheet
Dim cRow As Integer, fName As String, fol As String
Dim v As String, fv As String, u As String

With Application.FileDialog(msoFileDialogFolderPicker)
      .AllowMultiSelect = False
      .Show
      On Error Resume Next
      fol = .SelectedItems(1)
      Err.Clear
      On Error GoTo 0
    End With
    If fol = "" Then Exit Sub
   
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(fol)

For Each objFile In objFolder.Files
    Set twbk = ThisWorkbook.Sheets("Sheet1")
    cRow = twbk.Range("A" & Rows.Count).End(xlUp).Row
            Set owbk = Workbooks.Open(objFile)
                Set ws = owbk.Sheets(1)
                    u = ws.[A2].Value
                    'v = Application.WorksheetFunction.Text(u, "yyyy_mm_dd")
                    'v = Application.WorksheetFunction.Text(u, "dd_mm_yyyy")
                    'v = Format(Range("A1"), "dd_mm_yyyy")
                     v = Format(Date, "yyyymmdd")
                    MsgBox u
                    MsgBox v
                        twbk.Range("A" & cRow + 1).Value = v 'Change as need
                fv = v & ".xlsx"
                fName = objFolder & "\" & fv
                MsgBox fName
                ws.SaveAs Filename:=fName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                Windows(fv).Close False
            Kill objFile
Next objFile

Set ws = Nothing
Set owbk = Nothing
Set twbk = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

End Sub

