Attribute VB_Name = "Program"
Option Explicit

Sub ImportOrder()
    Dim Found As Boolean
    Dim Result As Long
    Dim sPath As Variant
    Dim iRows As Long
    Dim i As Long


    For i = 0 To 30
        sPath = "\\BR3615GAPS\GAPS\Carrier\Carrier Order Entry\" & Format(Date - i, "mm-dd-yy") & ".xls"
        If FileExists(sPath) Then
            Found = True
            Exit For
        End If
    Next

    If Found And i > 0 Then
        Result = MsgBox("An order from " & Format(Date - i, "mmm dd, yyyy") & " was found." & vbCrLf & _
                        vbCrLf & "Would you like to continue?", vbYesNo)
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    If Found = True And i = 0 Or Result = vbYes And Found = True Then
        Sheets("Sheet1").Select
        Workbooks.Open sPath
        Sheets("ORDER PAGE").Select
        ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Sheet1").Range("A1")
        ActiveWorkbook.Close


        Columns("I:I").Delete
        Columns("G:G").Delete
        Columns("D:E").Delete
        Columns("A:A").Delete

        iRows = ActiveSheet.UsedRange.Rows.Count

        Range(Cells(1, 2), Cells(iRows, 2)).Cut
        Range("A1").Insert Shift:=xlToRight

        Range("A1:E1").ClearContents
        Range("A1:E1").ClearFormats


        Range(Cells(1, 1), Cells(1, 6)) = Array("PO", "Part", "Area", "Qty", "Cust", "Lookup")
        Range(Cells(2, 5), Cells(iRows, 5)).Value = "12148"

        Cells(2, 6).Formula = "=VLOOKUP(C2, Master!A:B, 2, FALSE)"
        Cells(2, 6).AutoFill Destination:=Range(Cells(2, 6), Cells(iRows, 6))
        Range(Cells(2, 3), Cells(iRows, 3)).Value = Range(Cells(2, 6), Cells(iRows, 6)).Value
        Columns("F:F").Delete

        i = 2
        Do While i < ActiveSheet.UsedRange.Rows.Count
            If InStr(Left(Cells(i, 2).Value, 1), "9") Then
                Rows(i).Delete
            Else
                i = i + 1
            End If
        Loop

        ActiveSheet.ListObjects.Add(xlSrcRange, Range(ActiveSheet.UsedRange.Address(False, False)), , xlYes).Name = "Table1"
        ActiveSheet.UsedRange.Columns.AutoFit

        Sheets("Sheet1").Copy
        sPath = Application.Dialogs(xlDialogSaveAs).Show
        ActiveWorkbook.Close
        ActiveSheet.Cells.Delete
        MsgBox "Complete!"
    ElseIf Found = False Then
        MsgBox "Order not found!"
    End If

    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
