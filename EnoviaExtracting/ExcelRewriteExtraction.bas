Attribute VB_Name = "ExcelRewriteExtraction"
Option Explicit

Public Sub RewriteExportActivePVR(ByVal exportFolder As String)
    Dim activeDoc As Object
    Dim rootProduct As Object
    Dim normalizedFolder As String
    Dim nodeRows As Collection
    Dim savedDocuments As Object
    Dim drawingDocuments As Object
    Dim topPartNumber As String

    Call EnsureCatiaSession

    Set activeDoc = CATIA.ActiveDocument
    If activeDoc Is Nothing Then
        Err.Raise vbObjectError + 6200, , "No active CATIA document is open."
    End If

    If Not IsCatProductDocument(activeDoc.Name) Then
        Err.Raise vbObjectError + 6201, , "Active CATIA document must be a CATProduct."
    End If

    Set rootProduct = activeDoc.Product
    If rootProduct Is Nothing Then
        Err.Raise vbObjectError + 6202, , "The active CATIA document does not expose a Product object."
    End If

    normalizedFolder = NormalizeFolderPath(exportFolder)
    Call EnsureFolderExists(normalizedFolder)

    Set nodeRows = New Collection
    Set savedDocuments = CreateObject("Scripting.Dictionary")
    Set drawingDocuments = CreateObject("Scripting.Dictionary")
    topPartNumber = GetPartNumber(rootProduct)

    Call CollectProductTree(rootProduct, 0, "", nodeRows, savedDocuments)
    Call BuildDrawingCandidates(nodeRows, drawingDocuments)
    Call SaveCollectedDocuments(savedDocuments, normalizedFolder)
    Call SaveCollectedDocuments(drawingDocuments, normalizedFolder)
    Call WriteRewriteStructureXml(nodeRows, normalizedFolder, topPartNumber)
    Call WriteRewritePvrSyncStyleReport(nodeRows, drawingDocuments, normalizedFolder, topPartNumber)
    Call WriteRewriteMetadataReport(nodeRows, savedDocuments, normalizedFolder, topPartNumber)
    Call WriteRewriteDTExportReport(nodeRows, savedDocuments, drawingDocuments, normalizedFolder, topPartNumber)
End Sub

Private Sub CollectProductTree(ByVal currentProduct As Object, ByVal level As Long, ByVal parentPartNumber As String, ByRef nodeRows As Collection, ByVal savedDocuments As Object)
    Dim row As Object
    Dim childProduct As Object
    Dim docKey As String

    Set row = CreateNodeRow(currentProduct, level, parentPartNumber)
    nodeRows.Add row

    docKey = CStr(row("doc_key"))
    If Len(docKey) > 0 Then
        If Not savedDocuments.Exists(docKey) Then
            savedDocuments.Add docKey, CreateSaveRecord(row, currentProduct)
        End If
    End If

    On Error Resume Next
    For Each childProduct In currentProduct.Products
        Call CollectProductTree(childProduct, level + 1, CStr(row("part_number")), nodeRows, savedDocuments)
    Next
    On Error GoTo 0
End Sub

Private Function CreateNodeRow(ByVal currentProduct As Object, ByVal level As Long, ByVal parentPartNumber As String) As Object
    Dim row As Object
    Dim docName As String
    Dim docType As String
    Dim extensionName As String

    docName = GetReferenceDocumentName(currentProduct)
    docType = InferDocumentType(docName, level)
    extensionName = GetExtensionName(docName, docType)

    Set row = CreateObject("Scripting.Dictionary")
    row.Add "level", level
    row.Add "instance_name", GetInstanceName(currentProduct)
    row.Add "part_number", GetPartNumber(currentProduct)
    row.Add "revision", GetRevision(currentProduct, docName)
    row.Add "doc_type", docType
    row.Add "doc_name", docName
    row.Add "extension", extensionName
    row.Add "parent_part_number", parentPartNumber
    row.Add "dataset_type", GetProductPropertyValue(currentProduct, "Dataset Type")
    row.Add "doc_organization", GetProductPropertyValue(currentProduct, "Revision Organization")
    row.Add "shareable", GetProductPropertyValue(currentProduct, "Shareable")
    row.Add "security_check", GetProductPropertyValue(currentProduct, "Security Check")
    row.Add "title", GetProductPropertyValue(currentProduct, "Title")
    row.Add "document_revision", GetProductPropertyValue(currentProduct, "Document Revision")
    row.Add "doc_key", UCase$(CStr(row("part_number")) & "|" & CStr(row("revision")) & "|" & docType & "|" & docName)
    Set CreateNodeRow = row
End Function

Private Function CreateSaveRecord(ByVal row As Object, ByVal currentProduct As Object) As Object
    Dim record As Object

    Set record = CreateObject("Scripting.Dictionary")
    record.Add "part_number", row("part_number")
    record.Add "revision", row("revision")
    record.Add "doc_type", row("doc_type")
    record.Add "doc_name", row("doc_name")
    record.Add "extension", row("extension")
    record.Add "document", GetOwningDocument(currentProduct)
    record.Add "saved_path", ""
    record.Add "save_status", "Pending"
    record.Add "save_error", ""

    Set CreateSaveRecord = record
End Function

Private Sub SaveCollectedDocuments(ByVal savedDocuments As Object, ByVal exportFolder As String)
    Dim key As Variant
    Dim record As Object
    Dim catiaDoc As Object
    Dim targetPath As String
    Dim originalAlerts As Boolean

    originalAlerts = CATIA.DisplayFileAlerts
    CATIA.DisplayFileAlerts = False

    For Each key In savedDocuments.Keys
        Set record = savedDocuments(key)
        Set catiaDoc = Nothing
        On Error Resume Next
        Set catiaDoc = record("document")
        On Error GoTo 0

        If catiaDoc Is Nothing Then
            record("save_status") = "Skipped"
            record("save_error") = "No document reference available."
        Else
            targetPath = BuildUniqueSavePath(exportFolder, CStr(record("part_number")), CStr(record("revision")), CStr(record("extension")))
            On Error Resume Next
            Err.Clear
            catiaDoc.SaveAs targetPath
            If Err.Number <> 0 Then
                record("save_status") = "Failed"
                record("save_error") = Err.Description
            Else
                record("save_status") = "Saved"
                record("saved_path") = targetPath
            End If
            On Error GoTo 0
        End If
    Next

    CATIA.DisplayFileAlerts = originalAlerts
End Sub

Private Sub BuildDrawingCandidates(ByVal nodeRows As Collection, ByVal drawingDocuments As Object)
    Dim rowIndex As Long
    Dim node As Object
    Dim drawingDoc As Object
    Dim drawingRecord As Object
    Dim drawingKey As String

    For rowIndex = 1 To nodeRows.Count
        Set node = nodeRows(rowIndex)
        Set drawingDoc = FindOpenDrawingDocument(CStr(node("part_number")))
        If Not drawingDoc Is Nothing Then
            drawingKey = UCase$(CStr(node("part_number")) & "|" & CStr(node("revision")) & "|CATDrawing|" & CStr(drawingDoc.Name))
            If Not drawingDocuments.Exists(drawingKey) Then
                Set drawingRecord = CreateObject("Scripting.Dictionary")
                drawingRecord.Add "part_number", node("part_number")
                drawingRecord.Add "revision", ExtractRevisionFromDocumentName(CStr(drawingDoc.Name))
                drawingRecord.Add "doc_type", "CATDrawing"
                drawingRecord.Add "doc_name", CStr(drawingDoc.Name)
                drawingRecord.Add "extension", "CATDrawing"
                drawingRecord.Add "document", drawingDoc
                drawingRecord.Add "saved_path", ""
                drawingRecord.Add "save_status", "Pending"
                drawingRecord.Add "save_error", ""
                drawingDocuments.Add drawingKey, drawingRecord
            End If
        End If
    Next
End Sub

Private Function GetOwningDocument(ByVal currentProduct As Object) As Object
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

    Set GetOwningDocument = candidate
End Function

Private Function GetReferenceDocumentName(ByVal currentProduct As Object) As String
    On Error Resume Next
    GetReferenceDocumentName = CStr(currentProduct.ReferenceProduct.Parent.Name)
    If Len(GetReferenceDocumentName) = 0 Then
        GetReferenceDocumentName = CStr(currentProduct.Name)
    End If
    On Error GoTo 0
End Function

Private Function GetPartNumber(ByVal currentProduct As Object) As String
    On Error Resume Next
    GetPartNumber = Trim$(CStr(currentProduct.PartNumber))
    If Len(GetPartNumber) = 0 Then
        GetPartNumber = Trim$(CStr(currentProduct.ReferenceProduct.PartNumber))
    End If
    If Len(GetPartNumber) = 0 Then
        GetPartNumber = Trim$(CStr(currentProduct.Name))
    End If
    On Error GoTo 0
End Function

Private Function GetInstanceName(ByVal currentProduct As Object) As String
    On Error Resume Next
    GetInstanceName = Trim$(CStr(currentProduct.Name))
    If Len(GetInstanceName) = 0 Then
        GetInstanceName = GetPartNumber(currentProduct)
    End If
    On Error GoTo 0
End Function

Private Function GetRevision(ByVal currentProduct As Object, ByVal docName As String) As String
    On Error Resume Next
    GetRevision = Trim$(CStr(currentProduct.Revision))
    If Len(GetRevision) = 0 Then
        GetRevision = Trim$(CStr(currentProduct.ReferenceProduct.Revision))
    End If
    On Error GoTo 0

    If Len(GetRevision) = 0 Then
        GetRevision = ParseRevisionFromName(docName)
    End If
End Function

Private Function ParseRevisionFromName(ByVal docName As String) As String
    Dim baseName As String

    baseName = docName
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    End If

    If Len(baseName) >= 2 Then
        ParseRevisionFromName = Right$(baseName, 2)
    Else
        ParseRevisionFromName = ""
    End If
End Function

Private Function ExtractRevisionFromDocumentName(ByVal docName As String) As String
    ExtractRevisionFromDocumentName = ParseRevisionFromName(docName)
End Function

Private Function InferDocumentType(ByVal docName As String, ByVal level As Long) As String
    Dim upperName As String

    upperName = UCase$(docName)
    If level = 0 And InStr(upperName, "PVRREF") > 0 Then
        InferDocumentType = "PVRREF"
    ElseIf Right$(upperName, 8) = ".CATPART" Then
        InferDocumentType = "CATPart"
    ElseIf Right$(upperName, 11) = ".CATPRODUCT" Then
        InferDocumentType = "CATProduct"
    Else
        InferDocumentType = "Component"
    End If
End Function

Private Function GetExtensionName(ByVal docName As String, ByVal docType As String) As String
    If InStrRev(docName, ".") > 0 Then
        GetExtensionName = Mid$(docName, InStrRev(docName, ".") + 1)
    ElseIf docType = "CATPart" Then
        GetExtensionName = "CATPart"
    Else
        GetExtensionName = "CATProduct"
    End If
End Function

Private Function BuildUniqueSavePath(ByVal exportFolder As String, ByVal partNumber As String, ByVal revision As String, ByVal extensionName As String) As String
    Dim fileSystem As Object
    Dim basePath As String
    Dim candidatePath As String
    Dim suffixIndex As Long

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    basePath = exportFolder & SafeFileName(partNumber & " " & revision)
    If Right$(basePath, 1) = " " Then
        basePath = Left$(basePath, Len(basePath) - 1)
    End If
    candidatePath = basePath & "." & extensionName

    suffixIndex = 1
    Do While fileSystem.FileExists(candidatePath)
        candidatePath = basePath & "_" & CStr(suffixIndex) & "." & extensionName
        suffixIndex = suffixIndex + 1
    Loop

    BuildUniqueSavePath = candidatePath
End Function

Private Function IsCatProductDocument(ByVal docName As String) As Boolean
    IsCatProductDocument = (Right$(UCase$(docName), 11) = ".CATPRODUCT")
End Function

Private Function FindOpenDrawingDocument(ByVal partNumber As String) As Object
    Dim docIndex As Long
    Dim candidate As Object
    Dim upperName As String

    Set FindOpenDrawingDocument = Nothing

    For docIndex = 1 To CATIA.Documents.Count
        Set candidate = CATIA.Documents.Item(docIndex)
        upperName = UCase$(candidate.Name)
        If Right$(upperName, 11) = ".CATDRAWING" Then
            If upperName Like UCase$(partNumber) & "*" & ".CATDRAWING" Then
                Set FindOpenDrawingDocument = candidate
                Exit Function
            End If
        End If
    Next
End Function

Private Function GetProductPropertyValue(ByVal currentProduct As Object, ByVal propertyName As String) As String
    Dim propertiesObject As Object
    Dim propertyObject As Object

    On Error Resume Next
    Set propertiesObject = currentProduct.ReferenceProduct.UserRefProperties
    If propertiesObject Is Nothing Then Set propertiesObject = currentProduct.UserRefProperties
    If propertiesObject Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If

    Set propertyObject = propertiesObject.Item(propertyName)
    If Err.Number <> 0 Then
        Err.Clear
        Set propertyObject = Nothing
    End If

    If Not propertyObject Is Nothing Then
        GetProductPropertyValue = Trim$(CStr(propertyObject.ValueAsString))
        If Len(GetProductPropertyValue) = 0 Then
            GetProductPropertyValue = Trim$(CStr(propertyObject.Value))
        End If
    End If
    On Error GoTo 0
End Function
