Attribute VB_Name = "ExcelRewriteReports"
Option Explicit

Public Sub WriteRewriteMetadataReport(ByVal nodeRows As Collection, ByVal savedDocuments As Object, ByVal exportFolder As String, ByVal topPartNumber As String)
    Dim htmlText As String
    Dim rowIndex As Long
    Dim node As Object
    Dim saveRecord As Object

    htmlText = "<html><head><title>Metadata Package</title></head><body>"
    htmlText = htmlText & "<h1>MetadataPackage</h1>"
    htmlText = htmlText & "<p>Top Assembly: " & HtmlEncode(topPartNumber) & "</p>"
    htmlText = htmlText & "<p>Generated: " & HtmlEncode(CStr(Now)) & "</p>"
    htmlText = htmlText & "<table border='1' cellspacing='0' cellpadding='4'>"
    htmlText = htmlText & "<tr><th>Level</th><th>Part Number</th><th>Revision</th><th>Type</th><th>Instance</th><th>Document</th><th>Dataset Type</th><th>Organization</th><th>Shareable</th><th>Title</th><th>Saved Path</th><th>Status</th></tr>"

    For rowIndex = 1 To nodeRows.Count
        Set node = nodeRows(rowIndex)
        Set saveRecord = Nothing
        If savedDocuments.Exists(node("doc_key")) Then Set saveRecord = savedDocuments(node("doc_key"))

        htmlText = htmlText & "<tr>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("level"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("part_number"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("revision"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("doc_type"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("instance_name"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("doc_name"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("dataset_type"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("doc_organization"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("shareable"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("title"))) & "</td>"
        If saveRecord Is Nothing Then
            htmlText = htmlText & "<td></td><td>Not tracked</td>"
        Else
            htmlText = htmlText & "<td>" & HtmlEncode(CStr(saveRecord("saved_path"))) & "</td>"
            htmlText = htmlText & "<td>" & HtmlEncode(CStr(saveRecord("save_status"))) & "</td>"
        End If
        htmlText = htmlText & "</tr>"
    Next

    htmlText = htmlText & "</table></body></html>"
    Call WriteTextFile(exportFolder & "MetadataPackage.html", htmlText)
End Sub

Public Sub WriteRewriteDTExportReport(ByVal nodeRows As Collection, ByVal savedDocuments As Object, ByVal drawingDocuments As Object, ByVal exportFolder As String, ByVal topPartNumber As String)
    Dim htmlText As String
    Dim rowIndex As Long
    Dim node As Object
    Dim saveRecord As Object
    Dim drawingRecord As Object
    Dim key As Variant

    htmlText = "<html><head><title>DT Export Report</title></head><body>"
    htmlText = htmlText & "<h1>DTExportReport</h1>"
    htmlText = htmlText & "<p>Top Assembly: " & HtmlEncode(topPartNumber) & "</p>"
    htmlText = htmlText & "<p>Generated: " & HtmlEncode(CStr(Now)) & "</p>"
    htmlText = htmlText & "<table border='1' cellspacing='0' cellpadding='4'>"
    htmlText = htmlText & "<tr><th>Part Number</th><th>Revision</th><th>Type</th><th>Document</th><th>Parent</th><th>Dataset Type</th><th>Organization</th><th>Shareable</th><th>Saved Path</th><th>Status</th><th>Error</th></tr>"

    For rowIndex = 1 To nodeRows.Count
        Set node = nodeRows(rowIndex)
        Set saveRecord = Nothing
        If savedDocuments.Exists(node("doc_key")) Then Set saveRecord = savedDocuments(node("doc_key"))

        htmlText = htmlText & "<tr>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("part_number"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("revision"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("doc_type"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("doc_name"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("parent_part_number"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("dataset_type"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("doc_organization"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(node("shareable"))) & "</td>"
        If saveRecord Is Nothing Then
            htmlText = htmlText & "<td></td><td>Not tracked</td><td></td>"
        Else
            htmlText = htmlText & "<td>" & HtmlEncode(CStr(saveRecord("saved_path"))) & "</td>"
            htmlText = htmlText & "<td>" & HtmlEncode(CStr(saveRecord("save_status"))) & "</td>"
            htmlText = htmlText & "<td>" & HtmlEncode(CStr(saveRecord("save_error"))) & "</td>"
        End If
        htmlText = htmlText & "</tr>"
    Next

    For Each key In drawingDocuments.Keys
        Set drawingRecord = drawingDocuments(key)
        htmlText = htmlText & "<tr>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(drawingRecord("part_number"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(drawingRecord("revision"))) & "</td>"
        htmlText = htmlText & "<td>CATDrawing</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(drawingRecord("doc_name"))) & "</td>"
        htmlText = htmlText & "<td></td><td></td><td></td><td></td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(drawingRecord("saved_path"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(drawingRecord("save_status"))) & "</td>"
        htmlText = htmlText & "<td>" & HtmlEncode(CStr(drawingRecord("save_error"))) & "</td>"
        htmlText = htmlText & "</tr>"
    Next

    htmlText = htmlText & "</table></body></html>"
    Call WriteTextFile(exportFolder & "DTExportReport.html", htmlText)
End Sub
