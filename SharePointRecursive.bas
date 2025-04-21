
Attribute VB_Name = "SharePointRecursive"
Option Explicit

Const BASE_URL As String = "https://contoso.sharepoint.com/sites/YourSite"
Const ROOT_FOLDER As String = "Shared Documents"
Const START_ROW As Long = 5

Dim rowIndex As Long

Sub GetAllFilesFromSharePoint()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    rowIndex = START_ROW

    ' Xóa dữ liệu cũ nếu cần
    ws.Range("A" & rowIndex & ":B" & ws.Rows.Count).ClearContents

    ' Gọi hàm đệ quy
    ProcessFolder "/" & ROOT_FOLDER, ws

    MsgBox "Done!"
End Sub

Sub ProcessFolder(folderPath As String, ws As Worksheet)
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    url = BASE_URL & "/_api/web/GetFolderByServerRelativeUrl('" & EncodeURL(folderPath) & "')?$expand=Folders,Files"

    xhr.Open "GET", url, False
    xhr.setRequestHeader "Accept", "application/json;odata=verbose"
    xhr.Send

    If xhr.Status = 200 Then
        Dim response As Object
        Set response = JsonConverter.ParseJson(xhr.responseText)

        Dim files As Object, file As Object
        Set files = response("d")("Files")("results")
        For Each file In files
            ws.Cells(rowIndex, 1).Value = file("Name")
            ws.Cells(rowIndex, 2).Value = file("ServerRelativeUrl")
            rowIndex = rowIndex + 1
        Next

        Dim folders As Object, folder As Object
        Set folders = response("d")("Folders")("results")
        For Each folder In folders
            If folder("Name") <> "Forms" Then
                ProcessFolder folder("ServerRelativeUrl"), ws
            End If
        Next
    Else
        Debug.Print "Error: " & xhr.Status & " - " & xhr.statusText
    End If
End Sub

Function EncodeURL(url As String) As String
    url = Replace(url, " ", "%20")
    url = Replace(url, "'", "%27")
    EncodeURL = url
End Function
