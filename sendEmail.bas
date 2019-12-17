Attribute VB_Name = "sendEmail"
Option Explicit

Sub send_email()

'this sub will create and send a new outlook mail and add the artikkel-endringer workbook from the current week

Dim olApp As Outlook.Application
Dim olEmail As Outlook.MailItem
Dim weeknumber As String
Dim currentYear As String

weeknumber = Application.WorksheetFunction.WeekNum(Date)
currentYear = Year(Date)


Set olApp = New Outlook.Application
Set olEmail = olApp.CreateItem(olMailItem)

With olEmail
.BodyFormat = olFormatHTML
.HTMLBody = ""
.Subject = ""
.Attachments.Add (ThisWorkbook.Path & "Product_changes_week_" & weeknumber & "_" & currentYear & ".xlsx")
.To = ""
.CC = ""
.BCC = ""
.SentOnBehalfOfName = ""
'.Display
.Send

End With

End Sub
