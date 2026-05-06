Attribute VB_Name = "ExcelNativeExtraction"
Option Explicit

Public Sub ExportActivePVRNative(ByVal exportFolder As String)
    Dim catiaApp As Object
    Dim activeDoc As Object
    Dim rootProduct As Object
    Dim normalizedFolder As String
    Dim nodeRows As Collection
    Dim uniqueDocs As Object

    Call EnsureCatiaSession

    Set catiaApp = CATIA
    Set activeDoc = catiaApp.ActiveDocument
    If activeDoc Is Nothing Then
        Err.Raise vbObjectError + 5000, , "No active CATIA document is open."
    End If

    If Not _IsCatProductDocument(activeDoc.Name) Then
        Err.Raise vbObjectError + 5001, , "Active CATIA document must be a CATProduct."
    End If

    Set rootProduct = activeDoc.Product
    If rootProduct Is Nothing Then
        Err.Raise vbObjectError + 5002, , "The active CATIA document does not expose a Product object."
    End If

    normalizedFolder = _NormalizeFolderPath(exportFolder)
    Call _EnsureFolderExists(normalizedFolder)

    Set nodeRows = New Collection
    Set uniqueDocs = CreateObject("Scripting.Dictionary")

    Call _CopySyncReportIfPresent(normalizedFolder, _SafePartNumber(rootProduct))
    Call _CollectProductTree(rootProduct, 0, "", nodeRows, uniqueDocs)
    Call _SaveCollectedDocuments(uniqueDocs, normalizedFolder)
    Call _WriteMetadataReport(nodeRows, uniqueDocs, normalizedFolder, _SafePartNumber(rootProduct))
    Call _WriteDTExportReport(nodeRows, uniqueDocs, normalizedFolder, _SafePartNumber(rootProduct))
End Sub

Private Sub _CollectProductTree(ByVal currentProduct As Object, ByVal level As Long, ByVal parentPartNumber As String, ByRef nodeRows As Collection, ByVal uniqueDocs As Object)
    Dim node As Object
    Dim docInfo As Object
    Dim childProduct As Object
    Dim docKey As String

    Set node = _CreateNodeRow(currentProduct, level, parentPartNumber)
    nodeRows.Add node

    docKey = CStr(node("doc_key"))
    If Len(docKey) > 0 Then
        If Not uniqueDocs.Exists(docKey) Then
            Set docInfo = CreateObject("Scripting.Dictionary")
            docInfo.Add "doc_key", docKey
            docInfo.Add "part_number", node("part_number")
            docInfo.Add "revision", node("revision")
            docInfo.Add "doc_type", node("doc_type")
            docInfo.Add "doc_name", node("doc_name")
            docInfo.Add "extension", node("extension")
            docInfo.Add "document", _GetOwningDocument(currentProduct)
            docInfo.Add "saved_path", ""
            docInfo.Add "save_status", "Pending"
            docInfo.Add "save_error", ""
            uniqueDocs.Add docKey, docInfo
        End If
    End If

    On Error Resume Next
    For Each childProduct In currentProduct.Products
        Call _CollectProductTree(childProduct, level + 1, CStr(node("part_number")), nodeRows, uniqueDocs)
    Next
    On Error GoTo 0
End Sub

Private Function _CreateNodeRow(ByVal currentProduct As Object, ByVal level As Long, ByVal parentPartNumber As String) As Object
    Dim row As Object
    Dim docName As String
    Dim partNumber As String
    Dim revision As String
    Dim docType As String
    Dim extensionName As String

    docName = _GetReferenceDocumentName(currentProduct)
    partNumber = _SafePartNumber(currentProduct)
    revision = _SafeRevision(currentProduct, docName)
    docType = _InferDocumentType(docName, level)
    extensionName = _GetExtensionName(docName, docType)

    Set row = CreateObject("Scripting.Dictionary")
    row.Add "level", level
    row.Add "instance_name", _SafeInstanceName(currentProduct)
    row.Add "part_number", partNumber
    row.Add "revision", revision
    row.Add "doc_type", docType
    row.Add "doc_name", docName
    row.Add "extension", extensionName
    row.Add "parent_part_number", parentPartNumber
    row.Add "doc_key", UCase$(partNumber & "|" & revision & "|" & docType & "|" & docName)

    Set _CreateNodeRow = row
End Function

Private Sub _SaveCollectedDocuments(ByVal uniqueDocs As Object, ByVal exportFolder As String)
    Dim key As Variant
    Dim docInfo As Object
    Dim catiaDoc As Object
    Dim targetPath As String
    Dim originalAlerts As Boolean

    originalAlerts = CATIA.DisplayFileAlerts
    CATIA.DisplayFileAlerts = False

    For Each key In uniqueDocs.Keys
        Set docInfo = uniqueDocs(key)
        Set catiaDoc = Nothing
        On Error Resume Next
        Set catiaDoc = docInfo("document")
        On Error GoTo 0

        If catiaDoc Is Nothing Then
            docInfo("save_status") = "Skipped"
            docInfo("save_error") = "No document reference available."
        Else
            targetPath = _BuildUniqueSavePath(exportFolder, CStr(docInfo("part_number")), CStr(docInfo("revision")), CStr(docInfo("extension")))
            On Error Resume Next
            Err.Clear
            catiaDoc.SaveAs targetPath
            If Err.Number <> 0 Then
                docInfo("save_status") = "Failed"
                docInfo("save_error") = Err.Description
            Else
                docInfo("save_status") = "Saved"
                docInfo("saved_path") = targetPath
            End If
            On Error GoTo 0
        End If
    Next

    CATIA.DisplayFileAlerts = originalAlerts
End Sub

Private Sub _WriteMetadataReport(ByVal nodeRows As Collection, ByVal uniqueDocs As Object, ByVal exportFolder As String, ByVal topPartNumber As String)
    Dim htmlText As String
    Dim rowIndex As Long
    Dim node As Object
    Dim docInfo As Object

    htmlText = "<html><head><title>Metadata Package</title></head><body>"
    htmlText = htmlText & "<h1>MetadataPackage</h1>"
    htmlText = htmlText & "<p>Top Assembly: " & _HtmlEncode(topPartNumber) & "</p>"
    htmlText = htmlText & "<p>Generated: " & _HtmlEncode(CStr(Now)) & "</p>"
    htmlText = htmlText & "<table border='1' cellspacing='0' cellpadding='4'>"
    htmlText = htmlText & "<tr><th>Level</th><th>Part Number</th><th>Revision</th><th>Type</th><th>Instance</th><th>Document</th><th>Saved Path</th><th>Status</th></tr>"

    For rowIndex = 1 To nodeRows.Count
        Set node = nodeRows(rowIndex)
        Set docInfo = Nothing
        If uniqueDocs.Exists(node("doc_key")) Then Set docInfo = uniqueDocs(node("doc_key"))

        htmlText = htmlText & "<tr>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("level"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("part_number"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("revision"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("doc_type"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("instance_name"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("doc_name"))) & "</td>"
        If docInfo Is Nothing Then
            htmlText = htmlText & "<td></td><td>Not tracked</td>"
        Else
            htmlText = htmlText & "<td>" & _HtmlEncode(CStr(docInfo("saved_path"))) & "</td>"
            htmlText = htmlText & "<td>" & _HtmlEncode(CStr(docInfo("save_status"))) & "</td>"
        End If
        htmlText = htmlText & "</tr>"
    Next

    htmlText = htmlText & "</table></body></html>"
    Call _WriteTextFile(exportFolder & "MetadataPackage.html", htmlText)
End Sub

Private Sub _WriteDTExportReport(ByVal nodeRows As Collection, ByVal uniqueDocs As Object, ByVal exportFolder As String, ByVal topPartNumber As String)
    Dim htmlText As String
    Dim rowIndex As Long
    Dim node As Object
    Dim docInfo As Object

    htmlText = "<html><head><title>DT Export Report</title></head><body>"
    htmlText = htmlText & "<h1>DTExportReport</h1>"
    htmlText = htmlText & "<p>Top Assembly: " & _HtmlEncode(topPartNumber) & "</p>"
    htmlText = htmlText & "<p>Generated: " & _HtmlEncode(CStr(Now)) & "</p>"
    htmlText = htmlText & "<table border='1' cellspacing='0' cellpadding='4'>"
    htmlText = htmlText & "<tr><th>Part Number</th><th>Revision</th><th>Type</th><th>Document</th><th>Parent</th><th>Saved Path</th><th>Status</th><th>Error</th></tr>"

    For rowIndex = 1 To nodeRows.Count
        Set node = nodeRows(rowIndex)
        Set docInfo = Nothing
        If uniqueDocs.Exists(node("doc_key")) Then Set docInfo = uniqueDocs(node("doc_key"))

        htmlText = htmlText & "<tr>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("part_number"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("revision"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("doc_type"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("doc_name"))) & "</td>"
        htmlText = htmlText & "<td>" & _HtmlEncode(CStr(node("parent_part_number"))) & "</td>"
        If docInfo Is Nothing Then
            htmlText = htmlText & "<td></td><td>Not tracked</td><td></td>"
        Else
            htmlText = htmlText & "<td>" & _HtmlEncode(CStr(docInfo("saved_path"))) & "</td>"
            htmlText = htmlText & "<td>" & _HtmlEncode(CStr(docInfo("save_status"))) & "</td>"
            htmlText = htmlText & "<td>" & _HtmlEncode(CStr(docInfo("save_error"))) & "</td>"
        End If
        htmlText = htmlText & "</tr>"
    Next

    htmlText = htmlText & "</table></body></html>"
    Call _WriteTextFile(exportFolder & "DTExportReport.html", htmlText)
End Sub

Private Sub _CopySyncReportIfPresent(ByVal exportFolder As String, ByVal topPartNumber As String)
    Dim sourcePath As String
    Dim fileSystem As Object

    sourcePath = Environ$("temp") & "\PVRSync_ConfiguredStructure_Report_" & topPartNumber & ".xml"
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FileExists(sourcePath) Then
        fileSystem.CopyFile sourcePath, exportFolder & "PVRSync_ConfiguredStructure_Report_" & topPartNumber & ".xml", True
    End If
End Sub

Private Function _GetOwningDocument(ByVal currentProduct As Object) As Object
    Dim candidate As Object

    Set candidate = Nothing
    On Error Resume Next
    Set candidate = currentProduct.ReferenceProduct.Parent
    On Error GoTo 0

    If candidate Is Nothing Then
        On Error Resume Next
        Set candidate = CATIA.ActiveDocument
        On Error GoTo 0
    End If

    Set _GetOwningDocument = candidate
End Function

Private Function _GetReferenceDocumentName(ByVal currentProduct As Object) As String
    On Error Resume Next
    _GetReferenceDocumentName = CStr(currentProduct.ReferenceProduct.Parent.Name)
    If Len(_GetReferenceDocumentName) = 0 Then
        _GetReferenceDocumentName = CStr(currentProduct.Name)
    End If
    On Error GoTo 0
End Function

Private Function _SafePartNumber(ByVal currentProduct As Object) As String
    On Error Resume Next
    _SafePartNumber = Trim$(CStr(currentProduct.PartNumber))
    If Len(_SafePartNumber) = 0 Then
        _SafePartNumber = Trim$(CStr(currentProduct.ReferenceProduct.PartNumber))
    End If
    If Len(_SafePartNumber) = 0 Then
        _SafePartNumber = Trim$(CStr(currentProduct.Name))
    End If
    On Error GoTo 0
End Function

Private Function _SafeInstanceName(ByVal currentProduct As Object) As String
    On Error Resume Next
    _SafeInstanceName = Trim$(CStr(currentProduct.Name))
    If Len(_SafeInstanceName) = 0 Then
        _SafeInstanceName = _SafePartNumber(currentProduct)
    End If
    On Error GoTo 0
End Function

Private Function _SafeRevision(ByVal currentProduct As Object, ByVal docName As String) As String
    On Error Resume Next
    _SafeRevision = Trim$(CStr(currentProduct.Revision))
    If Len(_SafeRevision) = 0 Then
        _SafeRevision = Trim$(CStr(currentProduct.ReferenceProduct.Revision))
    End If
    On Error GoTo 0

    If Len(_SafeRevision) = 0 Then
        _SafeRevision = _ParseRevisionFromName(docName)
    End If
End Function

Private Function _ParseRevisionFromName(ByVal docName As String) As String
    Dim baseName As String

    baseName = docName
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    End If

    If Len(baseName) >= 2 Then
        _ParseRevisionFromName = Right$(baseName, 2)
    Else
        _ParseRevisionFromName = ""
    End If
End Function

Private Function _InferDocumentType(ByVal docName As String, ByVal level As Long) As String
    Dim upperName As String

    upperName = UCase$(docName)
    If level = 0 And InStr(upperName, "PVRREF") > 0 Then
        _InferDocumentType = "PVRREF"
    ElseIf Right$(upperName, 8) = ".CATPART" Then
        _InferDocumentType = "CATPart"
    ElseIf Right$(upperName, 11) = ".CATPRODUCT" Then
        _InferDocumentType = "CATProduct"
    Else
        _InferDocumentType = "Component"
    End If
End Function

Private Function _GetExtensionName(ByVal docName As String, ByVal docType As String) As String
    If InStrRev(docName, ".") > 0 Then
        _GetExtensionName = Mid$(docName, InStrRev(docName, ".") + 1)
    ElseIf docType = "CATPart" Then
        _GetExtensionName = "CATPart"
    Else
        _GetExtensionName = "CATProduct"
    End If
End Function

Private Function _BuildUniqueSavePath(ByVal exportFolder As String, ByVal partNumber As String, ByVal revision As String, ByVal extensionName As String) As String
    Dim fileSystem As Object
    Dim basePath As String
    Dim candidatePath As String
    Dim suffixIndex As Long

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    basePath = exportFolder & _SafeFileName(partNumber & " " & revision)
    If Right$(basePath, 1) = " " Then
        basePath = Left$(basePath, Len(basePath) - 1)
    End If
    candidatePath = basePath & "." & extensionName

    suffixIndex = 1
    Do While fileSystem.FileExists(candidatePath)
        candidatePath = basePath & "_" & CStr(suffixIndex) & "." & extensionName
        suffixIndex = suffixIndex + 1
    Loop

    _BuildUniqueSavePath = candidatePath
End Function

Private Function _SafeFileName(ByVal rawValue As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    _SafeFileName = Trim$(rawValue)
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In invalidChars
        _SafeFileName = Replace(_SafeFileName, CStr(item), "_")
    Next
End Function

Private Function _NormalizeFolderPath(ByVal folderPath As String) As String
    _NormalizeFolderPath = Trim$(folderPath)
    If Len(_NormalizeFolderPath) = 0 Then
        Err.Raise vbObjectError + 5003, , "Export folder is required."
    End If
    If Right$(_NormalizeFolderPath, 1) <> "\" Then
        _NormalizeFolderPath = _NormalizeFolderPath & "\"
    End If
End Function

Private Sub _EnsureFolderExists(ByVal folderPath As String)
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        fileSystem.CreateFolder folderPath
    End If
End Sub

Private Sub _WriteTextFile(ByVal filePath As String, ByVal fileContents As String)
    Dim fileSystem As Object
    Dim textStream As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textStream = fileSystem.OpenTextFile(filePath, 2, True, 0)
    textStream.Write fileContents
    textStream.Close
End Sub

Private Function _HtmlEncode(ByVal rawValue As String) As String
    _HtmlEncode = Replace(rawValue, "&", "&amp;")
    _HtmlEncode = Replace(_HtmlEncode, "<", "&lt;")
    _HtmlEncode = Replace(_HtmlEncode, ">", "&gt;")
    _HtmlEncode = Replace(_HtmlEncode, """", "&quot;")
End Function

Private Function _IsCatProductDocument(ByVal docName As String) As Boolean
    _IsCatProductDocument = (Right$(UCase$(docName), 11) = ".CATPRODUCT")
End Function
