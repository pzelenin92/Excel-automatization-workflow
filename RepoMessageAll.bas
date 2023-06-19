Option Explicit



Sub RepoMessage_1_All()

Dim buyer, seller, issuer, Crncy, settledt1, settledt2, addons, US_Sec, clt_name, iss_name As String
Dim Quantity, price1, price2, Accrued1, Accrued2 As Double
Dim repo_template As String

Dim outlookApp As Outlook.Application
Dim msg As Outlook.MailItem

Set outlookApp = New Outlook.Application
Set msg = outlookApp.CreateItem(olMailItem)

'Assign string values to variables
buyer = [R20C44].value
seller = [R20C43].value
issuer = [R20C45].value
Crncy = [R20C46].value
settledt1 = [R20C9].value
settledt2 = [R20C38].value
addons = [R20C51].value
US_Sec = [R20C48].value
'clt_name = [R20C58].Value
'iss_name = [R20C19].Value

'Assign string values to variables
Quantity = Round([R20C32].value, 6)
price1 = Round([R20C33].value, 6)
Accrued1 = Round([R20C34].value, 6)
price2 = Round([R20C35].value, 6)
Accrued2 = Round([R20C36].value, 6)

'Repo Template
repo_template = _
    "<p style = 'color: red; font: 10pt Calibri'>" & US_Sec & "</p>" & _
    "<table width = 90% border = 1 bordercolor = black style = 'border-collapse: collapse; font: normal 10pt Calibri'>" & _
        "<tr valign = bottom style = 'padding: 0px 5px'>" & _
            "<th>¹ Ñäåëêè</th><th>Ìåñòî<br>ñîâåðøåíèÿ<br>ñäåëêè</th><th>Âèä<br>ñäåëêè</th><th>Ïîêóïàòåëü</th><th>Ïðîäàâåö</th><th>Ýìèòåíò</th><th>Êîë-âî ÖÁ</th><th>Öåíà</th><th>ÍÊÄ</th><th>Âàëþòà<br>öåíû</th><th>Âàëþòà<br>ñäåëêè</th><th>Ïðåäîïëàòà</th><th>Ïðåäïîñòàâêà</th><th>ÏÏÏ</th><th>Ñðîê îïëàòû</th><th>Ñðîê ïîñòàâêè</th><th>Äîï.<br>Óñëîâèÿ</th><th>Ïðèìå÷àíèå</th>" & _
        "</tr>" & _
        "<tr align = center valign = bottom style = 'padding: 0px 5px'>" & _
            "<td>1</td><td>Âíåáèðæåâàÿ</td><td>Ðåïî 1÷</td><td>" & buyer & "</td><td>" & seller & "</td><td>" & issuer & "</td><td>" & Quantity & "</td><td>" & price1 & "</td><td>" & Accrued1 & "</td><td>" & Crncy & "</td><td>" & Crncy & "</td><td></td><td></td><td>ÏÏÏ</td><td>" & settledt1 & "</td><td>" & settledt1 & "</td><td>" & addons & "</td><td></td>" & _
        "</tr>" & _
        "<tr align = center valign = bottom style = 'padding: 0px 5px'>" & _
            "<td>1</td><td>Âíåáèðæåâàÿ</td><td>Ðåïî 2÷</td><td>" & seller & "</td><td>" & buyer & "</td><td>" & issuer & "</td><td>" & Quantity & "</td><td>" & price2 & "</td><td>" & Accrued2 & "</td><td>" & Crncy & "</td><td>" & Crncy & "</td><td></td><td></td><td>ÏÏÏ</td><td>" & settledt2 & "</td><td>" & settledt2 & "</td><td>" & addons & "</td><td></td>" & _
        "</tr>" & _
    "</table>"
    
msg.To = [R3C21].value
msg.cc = [R2C21].value & ";" & [R4C19].value & ";" & [R2C19].value & ";" & [R20C57].value
'Msg.Subject = "Ñäåëêà Ðåïî " & clt_name & " " & iss_name
msg.Subject = [R20C59].value
msg.HTMLBody = repo_template
msg.Display

End Sub