Attribute VB_Name = "updateData"
Option Explicit

Sub update_data()
'this sub will remove and update the sheets 15-20 and 50-80 with the currently updated SRC

' delete data
Sheets("15-20").Select
Range("a2").Select
Sheets("15-20").AutoFilter.ShowAllData
Range(ActiveSheet.ListObjects(1)).ClearContents


    Sheets("20-80").Select
    Range("a2").Select
    Sheets("20-80").AutoFilter.ShowAllData
    Range(ActiveSheet.ListObjects(1)).ClearContents



' update data
Sheets("SRC").Select
Range("a2").Select
Sheets("SRC").AutoFilter.ShowAllData

With Worksheets("SRC").Range("a1:h1")
        .AutoFilter Field:=4, Criteria1:="10", Criteria2:="15", Operator:=xlOr



End With


        Range(ActiveSheet.ListObjects(1)).Resize(, 4).Copy

    Sheets("15-20").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, skipblanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False





line2:


        Worksheets("SRC").Activate
        Range("a1").Select
        Worksheets("SRC").AutoFilter.ShowAllData
        With Worksheets("SRC").Range("a1:h1")

            .AutoFilter Field:=4, Criteria1:="20", Criteria2:="50", Operator:=xlOr
            .AutoFilter Field:=8, Criteria1:="<>" & "079?"


        End With


                Range(ActiveSheet.ListObjects(1)).Resize(, 4).Copy
                Worksheets("20-80").Activate
                Range("A2").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, skipblanks _
                :=False, Transpose:=False


    Application.CutCopyMode = False

line99:
ThisWorkbook.Save
End Sub
