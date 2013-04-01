Attribute VB_Name = "Program"
Option Explicit

Public Function FileExists(ByVal sPath As String) As Boolean
    'Remove trailing backslash
    If InStr(Len(sPath), sPath, "\") > 0 Then sPath = Left(sPath, Len(sPath) - 1)
    'Check to see if the directory exists and return true/false
    If Dir(sPath, vbDirectory) <> "" Then FileExists = True
End Function

Sub ImportOrder()
    Dim sPath As Variant
    Dim iRows As Long
    Dim i As Long

    'sPath = "\\BR3615GAPS\GAPS\Carrier\Carrier Order Entry\Carrier Order " & Format(Date, "mm-dd-yy") & ".xls"
    sPath = "\\BR3615GAPS\GAPS\Carrier\Carrier Order Entry\" & Format(Date, "mm-dd-yy") & ".xls"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    If FileExists(sPath) = True Then
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
    Else
        MsgBox "Order not found!"
    End If

    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
