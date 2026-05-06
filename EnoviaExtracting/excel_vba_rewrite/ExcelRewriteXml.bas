Attribute VB_Name = "ExcelRewriteXml"
Option Explicit

Public Sub WriteRewriteStructureXml(ByVal nodeRows As Collection, ByVal exportFolder As String, ByVal topPartNumber As String)
    Dim xmlText As String
    Dim rowIndex As Long
    Dim node As Object

    xmlText = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    xmlText = xmlText & "<Structure TopPartNumber=""" & XmlEncode(topPartNumber) & """>" & vbCrLf

    For rowIndex = 1 To nodeRows.Count
        Set node = nodeRows(rowIndex)
        xmlText = xmlText & "  <Instance" & _
            " Level=""" & XmlEncode(CStr(node("level"))) & """" & _
            " InstanceName=""" & XmlEncode(CStr(node("instance_name"))) & """" & _
            " PartNumber=""" & XmlEncode(CStr(node("part_number"))) & """" & _
            " DocRev=""" & XmlEncode(CStr(node("revision"))) & """" & _
            " DocType=""" & XmlEncode(CStr(node("doc_type"))) & """" & _
            " DocumentName=""" & XmlEncode(CStr(node("doc_name"))) & """" & _
            " ParentPartNumber=""" & XmlEncode(CStr(node("parent_part_number"))) & """" & _
            " DatasetType=""" & XmlEncode(CStr(node("dataset_type"))) & """" & _
            " RevisionOrganization=""" & XmlEncode(CStr(node("doc_organization"))) & """" & _
            " Shareable=""" & XmlEncode(CStr(node("shareable"))) & """" & _
            " SecurityCheck=""" & XmlEncode(CStr(node("security_check"))) & """" & _
            " Title=""" & XmlEncode(CStr(node("title"))) & """" & _
            " DocumentRevision=""" & XmlEncode(CStr(node("document_revision"))) & """" & _
            " />" & vbCrLf
    Next

    xmlText = xmlText & "</Structure>" & vbCrLf
    Call WriteTextFile(exportFolder & "Conf.xml", xmlText)
End Sub

Public Sub WriteRewritePvrSyncStyleReport(ByVal nodeRows As Collection, ByVal drawingDocuments As Object, ByVal exportFolder As String, ByVal topPartNumber As String)
    Dim xmlText As String
    Dim rowIndex As Long
    Dim node As Object
    Dim drawingDoc As Object
    Dim selectedDwg As String

    xmlText = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    xmlText = xmlText & "<Report TopPartNumber=""" & XmlEncode(topPartNumber) & """>" & vbCrLf

    For rowIndex = 1 To nodeRows.Count
        Set node = nodeRows(rowIndex)
        selectedDwg = ""

        For Each drawingDoc In drawingDocuments.Items
            If UCase$(CStr(drawingDoc("part_number"))) = UCase$(CStr(node("part_number"))) Then
                selectedDwg = CStr(drawingDoc("part_number")) & CStr(drawingDoc("revision"))
                Exit For
            End If
        Next

        xmlText = xmlText & "  <Part" & _
            " PartNumber=""" & XmlEncode(CStr(node("part_number"))) & """" & _
            " RefPartNumber=""" & XmlEncode(CStr(node("part_number")) & CStr(node("revision"))) & """" & _
            " DocRev=""" & XmlEncode(CStr(node("revision"))) & """" & _
            " Type=""" & XmlEncode(CStr(node("doc_type"))) & """" & _
            " SelectedDwg=""" & XmlEncode(selectedDwg) & """" & _
            " PrimaryDocument="""" />" & vbCrLf
    Next

    xmlText = xmlText & "</Report>" & vbCrLf
    Call WriteTextFile(exportFolder & "PVRSync_ConfiguredStructure_Report_" & SafeFileName(topPartNumber) & ".xml", xmlText)
End Sub

Private Function XmlEncode(ByVal rawValue As String) As String
    XmlEncode = Replace(rawValue, "&", "&amp;")
    XmlEncode = Replace(XmlEncode, "<", "&lt;")
    XmlEncode = Replace(XmlEncode, ">", "&gt;")
    XmlEncode = Replace(XmlEncode, """", "&quot;")
End Function
