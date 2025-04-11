Attribute VB_Name = "Module1"
Sub WriteDataFast(dataRange As Variant, startRow As Long, startCol As Long)
    Dim r As Long, c As Long
    Dim destSheet As Worksheet
    Set destSheet = ThisWorkbook.Sheets(1)

    For r = 1 To UBound(dataRange, 1)
        For c = 1 To UBound(dataRange, 2)
            destSheet.Cells(startRow + r - 1, startCol + c - 1).Value = dataRange(r, c)
        Next c
    Next r
End Sub
