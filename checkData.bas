Attribute VB_Name = "checkData"
Option Explicit

Sub check_data()

'this sub will update sheet SRC from DB and check for updated statuses from previous week

Dim rowCount As Long
Dim i As Long
Dim j As Long
Dim k As Integer
Dim vlookupRange As Range
Dim vlookupcode As String
Dim vlookupresult As Long
Dim url As String
Dim productchangecount As Long
Dim weeknumber As String
Dim currentYear As String
Dim currentdate As String
Dim lastDateTocheck As String
Dim discontinuedProducts As Long
Dim newProducts As Long
Dim newJProducts As Long

' get current date and convert to string
currentdate = Format(Date, "yyyymmdd")
lastDateTocheck = currentdate - 6

' get current weeknumber and year
weeknumber = Application.WorksheetFunction.WeekNum(Date)
currentYear = Year(Date)


Application.DisplayAlerts = False
Application.ScreenUpdating = False

'reset productchangecount, delete the old "Report" sheet and create new "Report" sheet from template


productchangecount = 0
Worksheets("Report").Delete
With Worksheets("Report_template")
  .Copy after:=Worksheets("Report_template")
  Worksheets("Report_template (2)").Name = "Report"

End With




'update SRC

    ActiveWorkbook.RefreshAll

' check if status has changed from 10 or 15 to 20 and add to list of new products

newProducts = 0
Worksheets("15-20").Activate
Range("E2").Select
Worksheets("15-20").AutoFilter.ShowAllData
ActiveCell.FormulaR1C1 = "=VLOOKUP([@MMITNO],SRC!C[-4]:C[-1],4,FALSE)"
rowCount = Worksheets("15-20").UsedRange.Rows.Count
Range("E2").Select
Selection.AutoFill Destination:=Range("E2" & ":" & "E" & rowCount)

With Worksheets("15-20").Range("a1:E1")
        .AutoFilter Field:=4, Criteria1:="10", Criteria2:="15", Operator:=xlOr
        .AutoFilter Field:=5, Criteria1:="20"

End With

newProducts = Worksheets("15-20") _
.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

'check if there are any SKUS, if none then skip to "20-80"
If newProducts >= 1 Then

  Worksheets("15-20").Activate
  Range(ActiveSheet.ListObjects(1)).Resize(, 3).Copy
  Worksheets("Report").Range("A3").PasteSpecial Paste:=xlPasteValues, skipblanks:=True
  Application.CutCopyMode = False
  Worksheets("Report").Activate
  Worksheets("Report").Range("A1").Select

Else
  GoTo line2

End If

line2:

' add J products with status "20" within the last 6 days

Worksheets("SRC").Activate
Worksheets("SRC").Range("a2").Select
Worksheets("SRC").AutoFilter.ShowAllData
With Worksheets("SRC").Range("a1:g1")
        .AutoFilter Field:=6, Criteria1:=">=" & lastDateTocheck
        .AutoFilter Field:=4, Criteria1:="20"
        .AutoFilter Field:=7, Criteria1:="J"

End With

newJProducts = Worksheets("SRC") _
.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


If newJProducts >= 1 Then

  newProducts = newProducts + newJProducts
  rowCount = Worksheets("Report").Range("b2").End(xlDown).Row

  If Worksheets("Report").Range("a" & rowCount).Value <> "" Then
    rowCount = rowCount + 1
  End If

  Worksheets("SRC").Activate
  Range(ActiveSheet.ListObjects(1)).Resize(, 3).Copy
  Worksheets("Report").Range("A" & rowCount).PasteSpecial Paste:=xlPasteValues, skipblanks:=True
  Application.CutCopyMode = False
  Worksheets("Report").Activate
  Worksheets("Report").Range("A1").Select

ElseIf newProducts = 0 Then

  GoTo line3

End If


' test the url and add the product URL to new products if available
rowCount = Worksheets("Report").Range("a2").End(xlDown).Row

Worksheets("Report").Activate
k = 3

For k = k To rowCount

    url = "https://shop.wj.no/p/" & Range("A" & k).Value

    If URLrequest(url) = True Then

    Worksheets("Report").Range("D" & k).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=HYPERLINK(CONCAT(""https://shop.wj.no/p/"",[@Varenummer]),[@Varenummer])"


    End If

Next k



line3:

' check if status has changed from 50 to 80 and add to list of discontinued products

discontinuedProducts = 0
Worksheets("20-80").Activate
Range("E2").Select
Worksheets("20-80").AutoFilter.ShowAllData
ActiveCell.FormulaR1C1 = "=VLOOKUP([@MMITNO],SRC!C[-4]:C[-1],4,FALSE)"
rowCount = Worksheets("20-80").UsedRange.Rows.Count
Range("E2").Select
Selection.AutoFill Destination:=Range("E2" & ":" & "E" & rowCount)

With Worksheets("20-80").Range("a1:E1")
        .AutoFilter Field:=4, Criteria1:="20", Criteria2:="50", Operator:=xlOr
        .AutoFilter Field:=5, Criteria1:="80"

End With

discontinuedProducts = Worksheets("20-80") _
.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

If discontinuedProducts >= 1 Then

Worksheets("20-80").Activate
Range(ActiveSheet.ListObjects(1)).Resize(, 3).Copy
Worksheets("Report").Range("F3").PasteSpecial Paste:=xlPasteValues, skipblanks:=True
Worksheets("Report").Activate
Worksheets("Report").Range("A1").Select

End If



'this checks if any product changes exists. and calls the correct email procedure

productchangecount = newProducts + discontinuedProducts
If productchangecount = 0 Then
  Call NoUpdates.send_email_no_updates
  GoTo line99

Else
' copy "Report" to new workbook and save as "Product_changes_week_" followed by the current week and year

    Worksheets("Report").Copy
    ActiveWorkbook.SaveAs Filename:= ThisWorkbook.Path & "Product_changes_week_" & weeknumber & "_" & currentYear, CreateBackup:=True, local:=True
    ActiveWorkbook.Close False

' start email Procedure
    Call sendEmail.send_email

End If




Application.ScreenUpdating = True
line99:
'call the procedure to update SRC
Call updateData.update_data

End Sub

'this Function will call an WinHttpRequest and return a true or false value based on the http statustext
Function URLrequest(url As String) As Boolean

Dim Request As Object
    Dim rc As Variant

    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Request
      .Open "GET", url, False
      .Send
      rc = .Statustext
    End With
    Set Request = Nothing
    If rc = "OK" Then URLrequest = True

    Exit Function
EndNow:

End Function
