Attribute VB_Name = "Module1"
Option Explicit

Public Sub Start()
    Dim ROOT_FOLDER As String
    ROOT_FOLDER = "Shared Documents/Test"
    
    Dim fullPath As String
    If Left(ROOT_FOLDER, 1) = "/" Then
        fullPath = ROOT_FOLDER
    Else
        fullPath = "/" & ROOT_FOLDER
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ProcessFolder fullPath, ws
End Sub

Sub ProcessFolder(folderPath As String, ws As Worksheet)
    ' Placeholder logic for processing the folder
    Debug.Print "Processing folder: " & folderPath
End Sub
