Attribute VB_Name = "NsfChecker"



Sub NsfChecker()

    Dim pathNSF As String
    pathNSF = "C:\path.xlsx"
    Dim pathMan As String
    pathMan = "C:\path.xlsm"
    
    FilesExistChecker pathNSF, pathMan
    
    Dim wbNSF As Workbook
    Set wbNSF = WbOpenChecker(pathNSF)
    
    Dim wbMan As Workbook
    Set wbMan = WbOpenChecker(pathMan)
    
    Dim counter As Integer
    counter = 3
    
    'Iterate through copied cells from email, get ISIN
    Dim wsBlank As Worksheet
    Set wsBlank = wbNSF.Worksheets("sheet-name")
    
    Dim wsCurName As String
    wsCurName = ""
    
    Do While wsBlank.Cells(counter, 6).value <> 0
        
        'Split worksheet name and get only surname
        Dim SplittedName As Variant
        SplittedName = Split(wsBlank.Cells(counter, 4))
        Dim CltName As String
        CltName = SplittedName(0)
        
    
        'Iterate through tabs and search for surname
        Dim ws As Worksheet
        For Each ws In wbMan.Worksheets
            If ws.name Like "*" & CltName & "*" Then

                'remember the name of tab
                wsCurName = ws.name 
    
                'Take isin
                Dim isin As String
                isin = wsBlank.Cells(counter, 6)
                
                'Take date to search with it
                Dim data As Date
                data = DateValue(wsBlank.Cells(counter, 3))
                
                Dim curWsMan As Worksheet
                Set curWsMan = wbMan.Worksheets(wsCurName)
                
                Dim daterange As Range
                Set daterange = curWsMan.Range("C:C")
                
                Dim findrange As Range
                Set findrange = daterange.Find(data, , xlValues)
                
                If Not findrange Is Nothing Then
                
                    Dim ValueCellAddress As String
                    'Save the first found adress which is Value
                    ValueCellAddress = findrange.Address
                    
                    Do
                        If curWsMan.Cells(findrange.row, 6).Value2 = isin And wsBlank.Cells(counter, 8).Value2 = curWsMan.Cells(findrange.row, 8).Value2 Then
                            
                            Dim FoundRow As Long
                            FoundRow = findrange.row
                            
                            'Sub which copies the last row
                            CopyPasteLastRow 
                            
                            'Sub which assign names from manager's file
                            AssignValues wsCurName, FoundRow 
                            
                            'Exit from loop, cause wi find tab (assuming only one valid tab exists)
                            Exit For 
                        Else
                            Set findrange = daterange.FindNext(findrange)
                        End If
                        
                    Loop While ValueCellAddress <> findrange.Address
                Else
                    MsgBox ("Value hasn't been found on the tab" & wsCurName)
                End If
            End If
        Next ws
        
        'let's check another value
        counter = counter + 3
        
    Loop

End Sub



Private Sub CopyPasteLastRow()

    Dim ws As Worksheet
    Set ws = Workbooks("file-name.xlsx").Worksheets("worksheet-name")
    
    Dim rg As Range
    Set rg = ws.Range("A2").CurrentRegion
    
    'Copying the last row to the next new row
    rg.Rows(rg.Rows.Count).Copy Destination:=rg.Rows(rg.Rows.Count + 1) 

End Sub



Private Sub AssignValues(WorksheetCurrentName, InputFoundRow)

    Dim wsNSF As Worksheet
    Set wsNSF = Workbooks("file-name.xlsx").Worksheets("worksheet-name")
    
    Dim wsMan As Worksheet
    Set wsMan = Workbooks("workbook-name.xlsm").Worksheets(WorksheetCurrentName)
    
    Dim lastrow As Long
    lastrow = wsNSF.Range("A2").CurrentRegion.Rows.Count
    
    'name and clf
    wsNSF.Range("A" & lastrow & ":B" & lastrow).Value2 = wsMan.Range("D" & InputFoundRow & ":E" & InputFoundRow).Value2
    
    'quantity
    wsNSF.Range("C" & lastrow).Value2 = wsMan.Range("I" & InputFoundRow).Value2
    
    'ISIN
    wsNSF.Range("D" & lastrow).Value2 = wsMan.Range("F" & InputFoundRow).Value2
    
    'Dates
    wsNSF.Range("F" & lastrow & ":G" & lastrow).Value2 = wsMan.Range("B" & InputFoundRow & ":C" & InputFoundRow).Value2
    
    'Price1
    wsNSF.Range("H" & lastrow).Value2 = wsMan.Range("M" & InputFoundRow).Value2
    
    'Price2
    wsNSF.Range("K" & lastrow).Value2 = wsMan.Range("J" & InputFoundRow).Value2
    
    'interest rate
    wsNSF.Range("L" & lastrow).Value2 = wsMan.Range("Q" & InputFoundRow).Value2
    
    'Accrued
    wsNSF.Range("M" & lastrow).Value2 = wsMan.Range("N" & InputFoundRow).Value2
    
    'Formula
    wsNSF.Range("N" & lastrow).FormulaR1C1 = "=BDP(RC[-10]&"" ISIN"",""PX_LAST"")/100*BDP(RC[-10]&"" ISIN"",""PAR_AMT"")"
    
    'Days
    wsNSF.Range("S" & lastrow).Value2 = wsMan.Range("P" & InputFoundRow).Value2 '����

End Sub



Private Function WbOpenChecker(FilePath As String) As Workbook
    Dim FileName As String
    FileName = Dir(FilePath)
    
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks(FileName)
    
    If wb Is Nothing Then
        Set wb = Workbooks.Open(FilePath, 0)
    End If
    
    On Error GoTo 0
    Set WbOpenChecker = wb
    
End Function



Private Function FilesExistChecker(sPathNSF As String, sPathMan As String) As Boolean

    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim filesdontexist As String
    Select Case fso.FileExists(sPathNSF)
        Case False: filesdontexist = "NSF"
    End Select
    
    Select Case fso.FileExists(sPathMan)
        Case False: filesdontexist = filedontexist & " Man"
    End Select
    
    If filesdontexist <> "" Then
        MsgBox ("There are no such file(s): " & filesdontexist)
        FilesExistChecker = False
    End If

End Function