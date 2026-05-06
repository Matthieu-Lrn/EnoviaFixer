Attribute VB_Name = "BDIfunctions"


'****Mandatory for JSON webservices
Private Enum enumBDIObject
        BDIpart = 0
        BDIdoc = 1
        BDIrv = 2
End Enum

Public Function GetBDIPartCollection(ByVal sPartNumber As String, ByRef sPartRev As String, ByRef bExitByUser As Boolean, Optional bGetEff As Boolean = True)

Dim i As Integer, iIndex As Integer
Dim sPartKey As String
Dim oPartColl As New clsCollection
Dim sJsonInfo As String, sState As String
Dim vPartInfo As Variant, vOppPartInfo As Variant, vDocInfo As Variant
Dim oBDIpart As clsBDIPart
Dim oSortData As clsXML
Dim oElem As IXMLDOMElement
Dim oNode As IXMLDOMNode, oNextNode As IXMLDOMNode

'Get the list of all part revision. This list may not be sorted
sJsonInfo = ""
sJsonInfo = WebServiceAccessTool.JSON_BDIPartInfoFromPnRev(sPartNumber, sPartRev, bExitByUser, , oPasswordMgr.GetEncryptedPassword)
If Not oTracking Is Nothing Then Call oTracking.WebServiceResultAnalysis("JSON_BDIPartInfoFromPnRev", , sPartNumber, sPartRev)
Call mdlParseJSON.ParseJSON(sJsonInfo, vPartInfo, sState)

'Create a sorted list by rev
Set oSortData = New clsXML
Call oSortData.AddElement("RootElem")
For i = LBound(vPartInfo) To UBound(vPartInfo)
    Set oElem = oSortData.AddElement("part", oSortData.RootNode)
    Call oSortData.AddSingleAttribute(oElem, "INDEX", i)
    Call oSortData.AddSingleAttribute(oElem, "REV", vPartInfo(i)("rev"))
Next
Call oSortData.SortNodes(oSortData.RootNode, "REV", "DESC")

'Exit if selected by user
If bExitByUser Then GoTo exitFnctn

'Retrieve all the BDI parts
For Each oNode In oSortData.SelectNodes("./part", oSortData.RootNode)

    'Get the index from the sorted list
    iIndex = CInt(oSortData.GetAttributeValue(oNode, "INDEX"))
    
    'Create a new part
    Set oBDIpart = New clsBDIPart
    
    'Fill the attributes
    'oBDIpart.PartCompleteName = sPartKey 'PV&REv
    sPartKey = vPartInfo(iIndex)("num") & vPartInfo(iIndex)("rev")
    oBDIpart.PartCompleteName = sPartKey
    oBDIpart.Status = vPartInfo(iIndex)("status")
    oBDIpart.IsCI = vPartInfo(iIndex)("isCi")
    oBDIpart.PartId = vPartInfo(iIndex)("id")
    oBDIpart.PartNumber = vPartInfo(iIndex)("num")
    oBDIpart.PartTitle = vPartInfo(iIndex)("title")
    oBDIpart.PartRev = vPartInfo(iIndex)("rev")
    
    'Add the part to oPartColl
    If Not oPartColl.Exists(sPartKey) Then
        Call oPartColl.Add(sPartKey, oBDIpart)
    End If
    
    'At this point if the part rev is not a CI we don't need to retrieve the info on the documents. Thus we can exit
    If oBDIpart.IsCI = False Then GoTo exitFnctn
     
    'Get the primary document from the opposite part, if required.
    'For old parts in BDI it's possible that the primary doc can't be found for -002, -004.... We than need to retrieve the primary doc from the opposite part (-001, -003...)
    If TypeName(vPartInfo(iIndex)("primary_document")) <> "Dictionary" Then
        If TypeName(vPartInfo(iIndex)("rh_part")) = "Dictionary" Then
            If vPartInfo(iIndex)("rh_part").Exists("num") And vPartInfo(iIndex)("rh_part").Exists("rev") Then
                
                'Get the opposite part
                sJsonInfo = ""
                sJsonInfo = WebServiceAccessTool.JSON_BDIPartInfoFromPnRev(vPartInfo(iIndex)("rh_part")("num"), vPartInfo(iIndex)("rh_part")("rev"), bExitByUser, , oPasswordMgr.GetEncryptedPassword)
                If Not oTracking Is Nothing Then Call oTracking.WebServiceResultAnalysis("JSON_BDIPartInfoFromPnRev", , vPartInfo(iIndex)("rh_part")("num"), vPartInfo(iIndex)("rh_part")("rev"))
                Call mdlParseJSON.ParseJSON(sJsonInfo, vOppPartInfo, sState)
                
                'Copy the primary document info to the original part
                Set vPartInfo(iIndex)("primary_document") = vOppPartInfo(0)("primary_document")
            End If
        End If
    End If
    
    'Get the primary document
    If TypeName(vPartInfo(iIndex)("primary_document")) = "Dictionary" Then
        If vPartInfo(iIndex)("primary_document").Exists("id") Then oBDIpart.PrimaryDoc.DocID = vPartInfo(iIndex)("primary_document")("id")
        If vPartInfo(iIndex)("primary_document").Exists("num") Then oBDIpart.PrimaryDoc.DocNumber = vPartInfo(iIndex)("primary_document")("num")
        If vPartInfo(iIndex)("primary_document").Exists("rev") Then oBDIpart.PrimaryDoc.DocRev = vPartInfo(iIndex)("primary_document")("rev")
        If vPartInfo(iIndex)("primary_document").Exists("status") Then
            If vPartInfo(iIndex)("primary_document")("status").Exists("name") Then oBDIpart.PrimaryDoc.Status = UCase(vPartInfo(iIndex)("primary_document")("status")("name"))
        End If
    End If
    
    'Get the effectivity, when required
    If oBDIpart.IsCI And bGetEff And oBDIpart.PrimaryDoc.DocNumber <> "" And oBDIpart.PrimaryDoc.DocRev <> "" And oBDIpart.PrimaryDoc.Status = "CLOSED" Then
        Call GetEffectivityOfDocument(oBDIpart.PrimaryDoc.DocNumber, oBDIpart.PrimaryDoc.DocRev, oBDIpart.PrimaryDoc.Effectivity, bExitByUser)
    End If
 
NextPart:
Next
Set oSortData = Nothing

'****When the part is "CLOSED" but the primary doc <> "CLOSED"
'****we then retrieve the previous document revision. We will use this document to get the effectivity

'We first check that oPartColl is not empty
If oPartColl.Count >= 1 Then

    'Get the latest part revision. It's the first item in the collection
    For i = 1 To oPartColl.Count
        Set oBDIpart = oPartColl.GetItem(i)
        If oBDIpart.Status = "CLOSED" Then Exit For
    Next
    

    'Check that the BDI part has a primary document
    If oBDIpart.PrimaryDoc.DocNumber = "" Then Call Err.Raise(9999, , oBDIpart.PartNumber & " " & oBDIpart.PartRev & " doesn't have a primary document.")
    
    
    'Check status of part and primary doc
    If oBDIpart.Status = "CLOSED" And oBDIpart.PrimaryDoc.Status <> "CLOSED" And oBDIpart.PrimaryDoc.DocRev <> "--" Then
    
        'Get the list of all documents. Note that this list may not be sorted
        sJsonInfo = ""
        sJsonInfo = WebServiceAccessTool.JSON_BDIDocumentInfoFromPnRv(oBDIpart.PrimaryDoc.DocNumber, "", bExitByUser, , oPasswordMgr.GetEncryptedPassword)
        If Not oTracking Is Nothing Then Call oTracking.WebServiceResultAnalysis("JSON_BDIDocumentInfoFromPnRv", , oBDIpart.PrimaryDoc.DocNumber)
        Call mdlParseJSON.ParseJSON(sJsonInfo, vDocInfo, sState)
          
        'Create a sorted list by rev
        Set oSortData = New clsXML
        Call oSortData.AddElement("RootElem")
        For i = LBound(vDocInfo) To UBound(vDocInfo)
            Set oElem = oSortData.AddElement("doc", oSortData.RootNode)
            Call oSortData.AddSingleAttribute(oElem, "INDEX", i)
            Call oSortData.AddSingleAttribute(oElem, "REV", vDocInfo(i)("rev"))
        Next
        Call oSortData.SortNodes(oSortData.RootNode, "REV", "DESC")

        'Get the node of the document with status <> CLOSED
        Set oNode = oSortData.SelectSingleNode("./doc[@REV='" & oBDIpart.PrimaryDoc.DocRev & "']", oSortData.RootNode)
        
        'Get the document with previous revision
        Set oNextNode = oNode.nextSibling
        
        'Get the index from the sorted list
        iIndex = CInt(oSortData.GetAttributeValue(oNextNode, "INDEX"))
        
        'Web service doesn't return anything or return format is not good, we exit
        If TypeName(vDocInfo(iIndex)) <> "Dictionary" Then GoTo exitFnctn
        
        'Populate next doc info
        If vDocInfo(iIndex).Exists("id") Then oBDIpart.PrevDoc.DocID = vDocInfo(iIndex)("id")
        If vDocInfo(iIndex).Exists("num") Then oBDIpart.PrevDoc.DocNumber = vDocInfo(iIndex)("num")
        If vDocInfo(iIndex).Exists("rev") Then oBDIpart.PrevDoc.DocRev = vDocInfo(iIndex)("rev")
        If vDocInfo(iIndex).Exists("status") Then
            If vDocInfo(iIndex)("status").Exists("name") Then oBDIpart.PrevDoc.Status = UCase(vDocInfo(iIndex)("status")("name"))
        End If

        'Get the effectivity of the next doc
        If oBDIpart.IsCI And bGetEff And oBDIpart.PrimaryDoc.DocNumber <> "" And oBDIpart.PrimaryDoc.DocRev <> "" Then
            Call GetEffectivityOfDocument(oBDIpart.PrevDoc.DocNumber, oBDIpart.PrevDoc.DocRev, oBDIpart.PrevDoc.Effectivity, bExitByUser)
        End If
    End If
End If

'Exit
exitFnctn:
Set GetBDIPartCollection = oPartColl

End Function

Public Function GetEffectivityOfDocument(ByVal sDocNb As String, ByVal sDocRev As String, ByRef oEffectivity As clsCollection, ByRef bExitByUser As Boolean, Optional bUseDevtDB = False) As Boolean

Dim sResult As String
Dim vResult As Variant
Dim sPath As String
Dim sStart As String, sStop  As String, sState As String
Dim oItem As clsCollection

'Run web service
sResult = WebServiceAccessTool.JSON_DOCEffectivity(sDocNb, sDocRev, bExitByUser, , oPasswordMgr.GetEncryptedPassword)

'Transform string result in json
Call mdlParseJSON.ParseJSON(sResult, vResult, sState)

'Clear effectivity
oEffectivity.RemoveAll

'Scan results
For i = LBound(vResult.Item("ranges")) To UBound(vResult.Item("ranges"))
    
    sStart = vResult.Item("ranges")(i).Item("start").Item("tail")
    sStop = vResult.Item("ranges")(i).Item("end").Item("tail")

    'Add to clsCollection
    If UCase(sStart) <> "CANCELLED" Then
        Set oItem = New clsCollection
        Call oItem.Add("IN", sStart)
        Call oItem.Add("OUT", sStop)
        Call oItem.Add("RV", "N/A")
        
        Call oEffectivity.Add(CStr(sStart & "-" & sStop & " ." & oEffectivity.Count + 1), oItem)
    End If
Next

If oEffectivity.Count >= 1 Then
    GetEffectivityOfDocument = True
End If

End Function

'*** Enter Tail number and effectivity collection to
'*** determine if tail is part of effectivity or not.
'*** Effectivity collection is accesed from clsBDIpart.Effectivity
Public Function IsDocInTail(ByVal cEffectvty As clsCollection, ByVal sTail As String) As Boolean
    Dim oEff As clsCollection
    For Each oEff In cEffectvty.GetItems
        If IsNumeric(oEff.GetItem("IN")) And IsNumeric(oEff.GetItem("OUT")) And IsNumeric(sTail) Then
            If CLng(sTail) >= CLng(oEff.GetItem("IN")) And CLng(sTail) <= CLng(oEff.GetItem("OUT")) Then
                IsDocInTail = True
                Exit For
            End If
        Else
            If sTail = oEff.GetItem("IN") Or sTail = oEff.GetItem("OUT") Then
                IsDocInTail = True
                Exit For
            End If
        End If
    Next
End Function







