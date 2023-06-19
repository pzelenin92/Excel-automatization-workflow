Private screenshotname, screenshotpath, screenshotExtension As String



'#1
Sub InitScreenshotVars()

'Sub which asserts private variables which will be used further 
'in SaveScreenshot,RunAutoitWithShell,TradeMessageAttach Subs

screenshotname = [R18C39].value
screenshotpath = "C:\screenshot\path"
screenshotExtension = ".gif"
End Sub



'#2
Sub SaveScreenshot()

InitScreenshotVars 'Initialise private variables for this module

'Check if file exists
If Dir(screenshotpath & screenshotname & screenshotExtension) <> "" Then
    MsgBox "File exists"
    End
End If

BTCRun
RunAutoitWithShell

End Sub



'#3
Sub BTCRun()

'определение переменных и присвоение значений
Dim go, tabr, tabl, pause As String
go = "<GO>"
tabr = "<tabr>"
tabl = "<tabl>"
pause = "<pause>003<pause>"

Dim command_sxt, command_grab, security, tail_sxt, tail_grab, run_sxt, run_grab As String
command_sxt = "SXT"
command_grab = "GRAB"
security = [R18C16].value
tail_grab = go & [R15C2].value & tabr & [R18C27].value & go & 1 & go

Dim panel As Integer
panel = 2

'Определение sxt_tail в зависимости от calc_typ_des
If [R18C43].value = "*NO CALCULATIONS*" Then
    tail_sxt = go & pause & tabr & [R18C4].value & tabr & [R18C28].value & tabr & [R18C29].value & go
Else
    'Определение sxt_tail в зависимости от maturity эмитента
    If [R18C31].value = "AT MATURITY" Then
        tail_sxt = go & pause & tabr & [R18C4].value & tabr & [R18C28].value & tabr & tabr & [R18C29].value & go
    ElseIf [R18C31].value = "CALLABLE" Or [R18C31].value = "PUTABLE" Or [R18C31].value = "PERP/CALL" Or [R18C31].value = "CALL/SINK" Then
        tail_sxt = go & pause & tabr & [R18C4].value & tabr & [R18C28].value & tabr & tabr & tabr & [R18C29].value & go
    ElseIf [R18C31].value = "SINKABLE" Then
        tail_sxt = go & pause & tabr & [R18C4].value & tabr & [R18C28].value & tabr & tabr & tabr & tabr & [R18C29].value & go
    End If
End If

'run sxt
run_sxt = Application.Run("BTCRUNCmd", command_sxt, panel, tail_sxt, security)
'run_grab = Application.Run("BTCRUNCmd", command_grab, panel, tail_grab)
Application.Wait Now + TimeValue("00:00:10")

End Sub



'#4
Sub RunAutoitWithShell()

Dim run_shell, AutoItExe_path, AutoItSript_path, pathname, param1, param2, pathnameWithParams As String

AutoItExe_path = "C:\Path\to\AutoIt3_x64.exe"
AutoItSript_path = "C:\Path\to\AutoIt-scripts\SaveScreenshotBloomberg.au3"
pathname = AutoItExe_path & " " & AutoItSript_path

param1 = screenshotname
param2 = screenshotpath
pathnameWithParams = pathname & " " & param1 & " " & param2

run_shell = Shell(pathnameWithParams, 1)
Application.Wait Now + TimeValue("00:00:10")

End Sub



'#5
Sub TradeMessageAttach_Ola()

InitScreenshotVars 'Initialise private variables for this module

'Check if file exists. If not runs SaveScreenshot sub to get screenshot
If Dir(screenshotpath & screenshotname & screenshotExtension) = "" Then
    MsgBox "File doesn't exist"
    SaveScreenshot
End If

Dim direction, buyer, seller, issuer, Crncy, settledt, US_Sec As String
Dim Quantity, Price, accrued, full_price As Double
Dim trade_template As String

Dim outlookApp As Outlook.Application
Dim msg As Outlook.MailItem

Set outlookApp = New Outlook.Application
Set msg = outlookApp.CreateItem(olMailItem)

'Присвоение сринговых значений:
direction = [R18C33].value
buyer = [R18C25].value
seller = [R18C26].value
issuer = [R18C18].value
Crncy = [R18C20].value
settledt = [R18C6].value
US_Sec = [R18C35].value

'Присвоение стринговых значений:
Quantity = Round([R18C12].value, 6)
Price = Round([R18C13].value, 6)
accrued = Round([R18C14].value, 6)
full_price = Round([R18C15].value, 6)

'Шаблон тикета на сделку
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
    
msg.To = [R4C13].value
msg.cc = [R3C13].value
msg.Subject = "Сделка"
msg.HTMLBody = trade_template
msg.Attachments.Add screenshotpath & screenshotname & screenshotExtension
msg.Display

End Sub



'#6
Sub TradeMessageAttach_ALL()

InitScreenshotVars
If Dir(screenshotpath & screenshotname & screenshotExtension) = "" Then
    MsgBox "File doesn't exist"
    SaveScreenshot
End If

Dim direction, buyer, seller, issuer, Crncy, settledt, US_Sec As String
Dim Quantity, Price, accrued, full_price As Double
Dim trade_template As String

Dim outlookApp As Outlook.Application
Dim msg As Outlook.MailItem

Set outlookApp = New Outlook.Application
Set msg = outlookApp.CreateItem(olMailItem)

'Присвоение сринговых значений:
direction = [R18C33].value
buyer = [R18C25].value
seller = [R18C26].value
issuer = [R18C18].value
Crncy = [R18C20].value
settledt = [R18C6].value
US_Sec = [R18C35].value

'Присвоение стринговых значений:
Quantity = Round([R18C12].value, 6)
Price = Round([R18C13].value, 6)
accrued = Round([R18C14].value, 6)
full_price = Round([R18C15].value, 6)

'Шаблон тикета на сделку
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
    
msg.To = [R4C15].value & ";" & [R18C36].value
msg.cc = [R3C13].value & ";" & [R18C34].value
msg.Subject = "Сделка"
msg.HTMLBody = trade_template
msg.Attachments.Add screenshotpath & screenshotname & screenshotExtension
msg.Display

End Sub