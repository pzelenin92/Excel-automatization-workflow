Attribute VB_Name = "ReplaceFunc"



Function rplc(cell As Range) As String

Dim replace1 As String
replace1 = Replace(cell, " ", "_")
rplc = Replace(replace1, "/", "-")

End Function
