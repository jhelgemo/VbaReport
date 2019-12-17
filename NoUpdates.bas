Attribute VB_Name = "NoUpdates"
Option Explicit

Sub send_email_no_updates()

  'this sub will create and send a new outlook email if there are no updated SKUs

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
.To = ""
.CC = ""
.BCC = ""
.SentOnBehalfOfName = ""
'.Display
.Send

End With

End Sub
