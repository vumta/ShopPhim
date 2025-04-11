Attribute VB_Name = "Module1"
Sub CopyMatchingColumnsFromSourceToTarget()

    Dim srcWB As Workbook, tgtWB As Workbook
    Dim srcWS As Worksheet, tgtWS As Worksheet
    Dim srcPath As String, tgtPath As String
    Dim srcHeaders As Variant, tgtHeaders As Variant
    Dim srcColIndexes() As Long
    Dim tgtColIndexes() As Long
    Dim lastRowSrc As Long, i As Long, j As Long
    Dim mapCount As Long

    srcPath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Chọn file SOURCE")
    tgtPath = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Chọn file TARGET (.xlsm)")

    If srcPath = "False" Or tgtPath = "False" Then Exit Sub

    Set srcWB = Workbooks.Open(srcPath, ReadOnly:=True)
    Set tgtWB = Workbooks.Open(tgtPath)
    Set srcWS = srcWB.Sheets(1)
    Set tgtWS = tgtWB.Sheets(1)

    srcHeaders = srcWS.Range("1:1").Value2
    tgtHeaders = tgtWS.Range("2:2").Value2

    ReDim srcColIndexes(1 To UBound(tgtHeaders, 2))
    ReDim tgtColIndexes(1 To UBound(tgtHeaders, 2))
    mapCount = 0

    For j = 1 To UBound(tgtHeaders, 2)
        For i = 1 To UBound(srcHeaders, 2)
            If srcHeaders(1, i) = tgtHeaders(1, j) And srcHeaders(1, i) <> "" Then
                mapCount = mapCount + 1
                srcColIndexes(mapCount) = i
                tgtColIndexes(mapCount) = j
                Exit For
            End If
        Next i
    Next j

    If mapCount = 0 Then
        MsgBox "Không tìm thấy cột nào trùng tên."
        Exit Sub
    End If

    lastRowSrc = srcWS.Cells(srcWS.Rows.Count, 1).End(xlUp).Row
    Dim r As Long

    Application.ScreenUpdating = False

    For r = 2 To lastRowSrc
        For i = 1 To mapCount
            tgtWS.Cells(r + 2, tgtColIndexes(i)).Value = srcWS.Cells(r, srcColIndexes(i)).Value
        Next i
        If r Mod 1000 = 0 Then DoEvents
    Next r

    Application.ScreenUpdating = True

    MsgBox "Copy hoàn tất: " & (lastRowSrc - 1) & " dòng được chuyển!"

    srcWB.Close SaveChanges:=False
    tgtWB.Save
    tgtWB.Close

End Sub
