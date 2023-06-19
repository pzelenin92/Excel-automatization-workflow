Attribute VB_Name = "CopyExcel"
Option Explicit
Public oldwb As Workbook
Public newwb As Workbook
Public new_name As String
Public new_path As String
Public name As String
Public path As String



Sub CopyExcel()

'    Dim path As String
'    Dim new_path As String
    Dim body_name As String
'    Dim name As String
'    Dim new_name As String
    Dim fileExtension As String
    Dim newExtension As String
    Dim date_today As Date
    
    'Naming files
    body_name = "All Eurobonds ������ 1.1"
'    body_name = "test"
    fileExtension = ".xlsm"
'    fileExtension = ".xlsx"
    newExtension = ".xlsx"
    name = body_name & fileExtension
    
    date_today = Date
    new_name = body_name & " " & date_today & newExtension
    
    'Getting paths
    path = "C:\Work\Pavel\0.Eurobonds\AllEurobondstocopy\AllEurobondstemplate\"
    new_path = "C:\Work\Pavel\0.Eurobonds\AllEurobondstocopy\"
    
    'check for the existence of a master workbook
    If Dir(path & name) = "" Then
        MsgBox "File doesn't exist"
        Exit Sub
    End If
    
    'check for the existence of a new workbook
    If Dir(new_path & new_name) <> "" Then
        MsgBox "File already exists"
        Exit Sub
    End If
    
    'check for opended new workbook
'    Dim newwb As Workbook
    On Error Resume Next
    Set newwb = Workbooks(new_name)
    On Error GoTo 0

    If newwb Is Nothing Then
        Set newwb = Workbooks.Add
    End If
    
    
    'check for the existence of opended master workbook
'    Dim oldwb As Workbook
    On Error Resume Next
    Set oldwb = Workbooks(name)
    On Error GoTo 0

    If oldwb Is Nothing Then
        Set oldwb = Workbooks.Open(path & name)
    End If

    Application.OnTime Now + TimeValue("00:00:10"), "WaitForOpen"
    
End Sub



Sub WaitForOpen()
    Debug.Print "Wait for open"
    RefreshData
End Sub



Sub RefreshData()
    Application.Run "RefreshEntireWorkbook"
    Debug.Print "Refreshnuli"
    Application.OnTime Now + TimeValue("00:00:10"), "IntermSub"
End Sub



Sub IntermSub()
    Copyws
End Sub



Sub Copyws()
Debug.Print "Copyws"
'(oldwb As Workbook, newwb As Workbook, new_path As String, new_name As String)
        'copying worksheets
    Dim oldws As Worksheet
    Dim newws As Worksheet

    For Each oldws In oldwb.Worksheets
        If oldws.Visible Then
            Set newws = newwb.Worksheets.Add(After:=newwb.Worksheets(newwb.Worksheets.Count))
            newws.name = oldws.name
            Dim copiedrange As Range
            Set copiedrange = oldws.Cells(2, 1).CurrentRegion.SpecialCells(xlCellTypeVisible)

            Dim pastedrange As Range
            Set pastedrange = newws.Cells(1, 1)

            copiedrange.Copy
            pastedrange.PasteSpecial xlPasteFormats
            pastedrange.PasteSpecial xlPasteColumnWidths

    '        Dim newpastrange As Range
    '        Set newpastrange = newws.Range("A1").Resize(copiedrange.Rows.Count, copiedrange.Columns.Count)
    '        newpastrange.Value2 = copiedrange.Value2
            newws.Rows(2).RowHeight = oldws.Rows(2).RowHeight
            pastedrange.PasteSpecial xlPasteValues
            Application.Wait Now + TimeValue("00:00:01")
        End If
    Next oldws

    'autofilter
    newwb.Worksheets("USD-��� �������").Rows(2).AutoFilter

    'save new file
    newwb.SaveAs new_path & new_name
End Sub