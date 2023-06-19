Attribute VB_Name = "FindRepeatedCells"
Option Explicit



Sub FindRepeatedCells()


Dim ws As Worksheet
Set ws = Workbooks("file-name.xlsm").Worksheets("worksheet-name")


Dim rangenash As Range
Set rangenash = ws.Range("D3:D411")


Dim rangeaton As Range
Set rangeaton = ws.Range("D412:D1086")


Dim cellnash As Range
Dim cellaton As Range
For Each cellaton In rangeaton

    Set cellnash = rangenash.Find(cellaton.Value2, LookIn:=xlValues)
    If Not cellnash Is Nothing Then
        Dim firstcelladdress As String
        firstcelladdress = cellnash.Address
        Do
            cellnash.Interior.ColorIndex = 5
            Set cellnash = rangenash.FindNext(cellnash)
        Loop While firstcelladdress <> cellnash.Address
    End If
Next cellaton

End Sub