
' JsonConverter.bas
' https://github.com/VBA-tools/VBA-JSON

' --- Bắt đầu của JsonConverter ---
' (Giả lập đoạn mã, phần này phải thay bằng mã thực từ file JsonConverter.bas)
Function ParseJson(jsonString As String) As Object
    ' Nội dung xử lý JSON
End Function
' --- Kết thúc của JsonConverter ---



' SharePointGetFiles.bas

Option Explicit

Public Sub GetAllSharePointFiles()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    Dim ROOT_FOLDER As String
    ROOT_FOLDER = "Shared Documents/Reports"

    Dim rowCounter As Long
    rowCounter = 5 ' bắt đầu từ dòng 5

    ProcessFolder "/" & ROOT_FOLDER, ws, rowCounter
End Sub

Sub ProcessFolder(folderPath As String, ws As Worksheet, ByRef rowCounter As Long)
    Dim url As String
    url = "https://contoso.sharepoint.com/sites/yoursite/_api/web/GetFolderByServerRelativeUrl('" & folderPath & "')?$expand=Folders,Files"

    Dim Http As Object
    Set Http = CreateObject("MSXML2.XMLHTTP")
    Http.Open "GET", url, False
    Http.setRequestHeader "Accept", "application/json;odata=verbose"
    Http.Send

    Dim json As Object
    Set json = ParseJson(Http.responseText)
    
    Dim files, folders, file, subfolder
    On Error Resume Next
    Set files = json("d")("Files")
    Set folders = json("d")("Folders")
    On Error GoTo 0

    If Not files Is Nothing Then
        For Each file In files
            ws.Cells(rowCounter, 1).Value = file("Name")
            ws.Cells(rowCounter, 2).Value = file("ServerRelativeUrl")
            rowCounter = rowCounter + 1
        Next
    End If

    If Not folders Is Nothing Then
        For Each subfolder In folders
            Dim name As String
            name = subfolder("Name")
            If name <> "Forms" Then ' Skip system folder
                ProcessFolder subfolder("ServerRelativeUrl"), ws, rowCounter
            End If
        Next
    End If
End Sub
