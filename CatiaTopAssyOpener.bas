Option Explicit

Public Sub CatiaTopAssemblyOpener()
    Dim catApp As Object
    Dim doc As Object
    Dim folderPath As String
    Dim topFile As String

    On Error GoTo ErrHandler

    folderPath = PickFolderExcel()
    If Len(folderPath) = 0 Then Exit Sub

    topFile = FindTopAssemblyFile(folderPath)
    If Len(topFile) = 0 Then
        MsgBox "No CATProduct found in selected folder.", vbExclamation
        Exit Sub
    End If

    Set catApp = GetObject(, "CATIA.Application")
    catApp.Visible = True
    Set doc = catApp.Documents.Open(topFile)

    If InStr(1, TypeName(doc), "ProductDocument", vbTextCompare) = 0 Then
        MsgBox "Opened file is not a ProductDocument.", vbExclamation
        Exit Sub
    End If

    MsgBox "Opened top assembly:" & vbCrLf & topFile, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error opening top assembly: " & Err.Description, vbCritical
End Sub

Private Function FindTopAssemblyFile(ByVal folderPath As String) As String
    Dim f As String
    Dim folderName As String
    Dim topCode As String

    folderName = GetLeafFolderName(folderPath)
    topCode = ExtractLeadingAlphaNum(folderName)

    ' Strict match first: <TopCode>-001*.CATProduct
    If Len(topCode) > 0 Then
        f = Dir(folderPath & "\" & topCode & "-001*.CATProduct", vbNormal)
        If Len(f) > 0 Then
            FindTopAssemblyFile = folderPath & "\" & f
            Exit Function
        End If
    End If

    ' Fallback if pattern not found
    f = Dir(folderPath & "\*-001*.CATProduct", vbNormal)
    If Len(f) > 0 Then
        FindTopAssemblyFile = folderPath & "\" & f
    Else
        FindTopAssemblyFile = ""
    End If
End Function

Private Function PickFolderExcel() As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select Assembly Folder"
    If fd.Show = -1 Then
        PickFolderExcel = fd.SelectedItems(1)
    Else
        PickFolderExcel = ""
    End If
End Function

Private Function GetLeafFolderName(ByVal folderPath As String) As String
    GetLeafFolderName = Mid$(folderPath, InStrRev(folderPath, "\") + 1)
End Function

Private Function ExtractLeadingAlphaNum(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9]" Then
            out = out & ch
        Else
            Exit For
        End If
    Next i

    ExtractLeadingAlphaNum = out
End Function
