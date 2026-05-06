Attribute VB_Name = "ExcelRewriteBootstrap"
Option Explicit

Public CATIA As Object

Public Sub EnsureCatiaSession()
    If CATIA Is Nothing Then
        On Error Resume Next
        Set CATIA = GetObject(, "CATIA.Application")
        On Error GoTo 0
    End If

    If CATIA Is Nothing Then
        Err.Raise vbObjectError + 6100, , "Could not connect to a running CATIA.Application session."
    End If
End Sub

Public Function NormalizeFolderPath(ByVal folderPath As String) As String
    NormalizeFolderPath = Trim$(folderPath)
    If Len(NormalizeFolderPath) = 0 Then
        Err.Raise vbObjectError + 6101, , "Export folder is required."
    End If
    If Right$(NormalizeFolderPath, 1) <> "\" Then
        NormalizeFolderPath = NormalizeFolderPath & "\"
    End If
End Function

Public Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        fileSystem.CreateFolder folderPath
    End If
End Sub

Public Function SafeFileName(ByVal rawValue As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    SafeFileName = Trim$(rawValue)
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In invalidChars
        SafeFileName = Replace(SafeFileName, CStr(item), "_")
    Next
End Function

Public Function HtmlEncode(ByVal rawValue As String) As String
    HtmlEncode = Replace(rawValue, "&", "&amp;")
    HtmlEncode = Replace(HtmlEncode, "<", "&lt;")
    HtmlEncode = Replace(HtmlEncode, ">", "&gt;")
    HtmlEncode = Replace(HtmlEncode, """", "&quot;")
End Function

Public Sub WriteTextFile(ByVal filePath As String, ByVal fileContents As String)
    Dim fileSystem As Object
    Dim textStream As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textStream = fileSystem.OpenTextFile(filePath, 2, True, 0)
    textStream.Write fileContents
    textStream.Close
End Sub
