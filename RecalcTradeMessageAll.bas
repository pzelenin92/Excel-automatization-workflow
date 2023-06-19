Option Explicit



Sub RecalcTradeMessage_All()
'Recalculates trade to special date and make an email to send

Dim direction, buyer, seller, issuer, Crncy, settledt, US_Sec As String
Dim Quantity, Price, accrued, full_price As Double
Dim trade_template As String

Dim outlookApp As Outlook.Application
Dim msg As Outlook.MailItem

Set outlookApp = New Outlook.Application
Set msg = outlookApp.CreateItem(olMailItem)

'Assign string values to variables
direction = [R19C36].value
buyer = [R19C40].value
seller = [R19C41].value
issuer = [R19C21].value
Crncy = [R19C24].value
settledt = [R19C2].value
US_Sec = [R19C35].value

'Assign string values to variables
Quantity = Round([R19C16].value, 6)
Price = Round([R19C17].value, 6)
accrued = Round([R19C18].value, 6)
full_price = Round([R19C19].value, 6)

'Template trade email
trade_template = _
    "<p style = 'color: red; font: 10pt Calibri'>" & US_Sec & "</p>" & _
    "<table width = 90% border = 1 bordercolor = black style = 'border-collapse: collapse; font: normal 10pt Calibri'>" & _
        "<tr valign = bottom style = 'padding: 0px 5px'>" & _
            "<th>№ Сделки</th><th>Место<br>совершения<br>сделки</th><th>Вид<br>сделки</th><th>Торговая<br>система</th><th>Покупатель</th><th>Продавец</th><th>Эмитент</th><th>Кол-во ЦБ</th><th>Цена</th><th>НКД</th><th>Цена полная</th><th>Валюта<br>цены</th><th>Валюта<br>сделки</th><th>Предоплата</th><th>Предпоставка</th><th>ППП</th><th>Срок оплаты</th><th>Срок поставки</th><th>Доп.<br>Условия</th><th>Примечание</th>" & _
        "</tr>" & _
        "<tr align = center valign = bottom style = 'padding: 0px 5px'>" & _
            "<td>1</td><td>Внебиржевая</td><td>" & direction & "</td><td></td><td>" & buyer & "</td><td>" & seller & "</td><td>" & issuer & "</td><td>" & Quantity & "</td><td>" & Price & "</td><td>" & accrued & "</td><td>" & full_price & "</td><td>" & Crncy & "</td><td>" & Crncy & "</td><td></td><td></td><td>ППП</td><td>" & settledt & "</td><td>" & settledt & "</td><td></td><td>С голоса</td>" & _
        "</tr>" & _
    "</table>"
    
msg.To = [R4C29].value & ";" & [R19C39].value
msg.cc = [R3C27].value & ";" & [R19C37].value
msg.Subject = "Сделка"
msg.HTMLBody = trade_template
msg.Display

End Sub