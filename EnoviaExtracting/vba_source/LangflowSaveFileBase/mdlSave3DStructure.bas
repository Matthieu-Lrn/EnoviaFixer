Attribute VB_Name = "mdlSave3DStructure"
Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Option Explicit
Private oConfXML As DOMDocument60
Private oPVRXML As DOMDocument60

Private oLogReport As DOMDocument60
Private oAttList As clsAttributesList
Private oRefProductList As clsCollection 'This is a collection of all the Reference Product found in the open PVR.
Private oLoadErrDoc As clsCollection
Private oSaveErrDoc As clsCollection
Private oPartList As clsCollection
Private sDestinationFolder As String

Public sLangflowDestinationFolder As String

Private bSaveXML As Boolean
Private sXMLPath As String

Private sTrackingLogText As String
Private sTrackingLogFilePath As String
Private bTrackingLogNeverSave As Boolean
Private bTrackingLogSaveEverytime As Boolean

Private sErrorLogFilePath As String
Private sEffectiveDocReportPath As String

Private oTopNode As clsPartInfo


Private Sub AddNewProduct(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sNewPN As String, ByVal sNewRev As String, oElem As IXMLDOMElement, ByVal sIteration As String)

Dim oParentRefProduct As Product
Dim oProductDoc As ProductDocument
Dim oProduct As Product
Dim oFSO
Dim sFileName As String
Dim sRelationID As String
Dim sErrorMsg As String
Dim dTimer As Double
Dim iIteration As Integer

'Initialize
Set oFSO = CreateObject("Scripting.FileSystemObject")
CATIA.DisplayFileAlerts = False

'No parent, we create a top product
If sParentPN = "" Then
   
  
    'Define file name
    sFileName = sNewPN & " " & sNewRev & "_" & sIteration & ".CATProduct"
    
    'Read template CATProduct
    Set oProductDoc = CATIA.Documents.Read(sCATProductTemplateFile)

    'Edit Part Number and revision
    oProductDoc.Product.PartNumber = sNewPN
    oProductDoc.Product.Revision = sNewRev

    'Since the "Organization" attribut is locked I can't edit it but (I don't understand why) I can delete it and recreate it
    Call oProductDoc.Product.ReferenceProduct.UserRefProperties.Remove("Organization")
    Call oProductDoc.Product.ReferenceProduct.UserRefProperties.CreateString("Organization", "")

    'Edit attributes
    Call AddAttributesToComponent(oProductDoc.Product, sNewRev)

    'Save
'   Call oProductDoc.SaveAs(sDestinationFolder & sNewPN & sNewRev & "_" & sIteration & ".CATProduct")
    Call oProductDoc.SaveAs(sDestinationFolder & sFileName)

    'Close the product
    Call CloseDocuments(oProductDoc)

    'Open the document
    Set oProductDoc = CATIA.Documents.Open(sDestinationFolder & sFileName)

    'Add to collection
    Call oRefProductList.Add(sNewPN & sNewRev, oProductDoc.Product.ReferenceProduct)
        
'We have a parent
Else

    'Get parent reference product
    Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)

    'Define file name
    sFileName = sNewPN & " " & sNewRev & "_" & sIteration & ".CATProduct"

    'Check if document already loaded in session
    Set oProductDoc = Nothing
    On Error Resume Next
    Set oProductDoc = CATIA.Documents.Item(sFileName)
    On Error GoTo 0

    'Create a new product document
    If oProductDoc Is Nothing Then

        'Read template CATProduct
        Set oProductDoc = CATIA.Documents.Read(sCATProductTemplateFile)
        
        'Edit Part Number and revision
        oProductDoc.Product.PartNumber = sNewPN
        oProductDoc.Product.Revision = sNewRev
        
        'Since the "Organization" attribut is locked I can't edit it but (I don't understand why) I can delete it and recreate it
        Call oProductDoc.Product.ReferenceProduct.UserRefProperties.Remove("Organization")
        Call oProductDoc.Product.ReferenceProduct.UserRefProperties.CreateString("Organization", "")

        'Edit attributes
        Call AddAttributesToComponent(oProductDoc.Product, sNewRev)

        'Save
        Call oProductDoc.SaveAs(sDestinationFolder & sFileName)
    End If
    
    'Add new instance under parent
    Set oProduct = oParentRefProduct.Products.AddExternalComponent(oProductDoc)

    'Add to oRefProductList
    If Not oRefProductList.Exists(sNewPN & sNewRev) Then
        Call oRefProductList.Add(sNewPN & sNewRev, oProduct.ReferenceProduct)
    End If

    'Get the RelationID of oElem
    sRelationID = oElem.Attributes.getNamedItem("RelationID").nodeValue

    'Transfer the instance name to all Conf instance with same RelationID
    Call TransferInstanceName(sRelationID, oProduct.Name)

    'Set status
    Call SetInstancesStatusOKorMoved(oConfXML, sParentPN, oProduct.Name, "OK")

    'Move instance
    Call MoveInstance(sParentPN, sParentRev, oProduct.Name, oElem)
    
End If
    
End Sub


Private Sub CopyFile(ByVal sFileName As String, ByVal SourcePath As String, ByVal TargetPath As String)

Dim oFSO

'Initialize
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Check if file exist
If Not oFSO.FileExists(SourcePath & sFileName) Then Exit Sub

'Copy file
oFSO.CopyFile SourcePath & sFileName, TargetPath

End Sub

Private Sub GettingPrimaryDocument(ByVal sTopPN As String)

Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim sPartNumber As String
Dim sRev As String
Dim iCount, iCountMax As Integer
Dim sString As String

Set oLogReport = New DOMDocument60
Call oLogReport.Load(sDestinationFolder & "PVRSync_ConfiguredStructure_Report_" & sTopPN & ".xml")

'Creating Primary Document attribut on all nodes
Set oNodes = oLogReport.selectNodes(".//Part")
For Each oNode In oNodes
    Call Add_Attribute(oLogReport, oNode, "PrimaryDocument", "")
Next

'Getting the nodes for NONE + non-CI
Set oNodes = oLogReport.selectNodes(".//Part[@IsCI='False' and (@Type='NONE' or @Type='CATProduct')]")
iCountMax = oNodes.Length
iCount = 0
For Each oNode In oNodes

    'Count
    iCount = iCount + 1
    
    'Get the Ref Part Number from XML
    sString = oNode.Attributes.getNamedItem("RefPartNumber").nodeValue

    'Extract the PN and REV
    sPartNumber = Left(sString, Len(sString) - 2)
    sRev = Right(sString, 2)
    
    'Progress
    Call frmProgress.progressBarRepaint("Step 3 of 8 - Retrieving primary document", 8, 3, "Retrieving primary document of " & sPartNumber & " " & sRev & " / " & iCount & " of " & iCountMax, iCountMax, iCount)
    
    'Get the drawing
    oNode.Attributes.getNamedItem("PrimaryDocument").nodeValue = GetPrimaryDocFromPart(sPartNumber, sRev)
Next

Call oLogReport.Save(sDestinationFolder & "PVRSync_ConfiguredStructure_Report_" & sTopPN & ".xml")
Set oLogReport = Nothing

End Sub

Private Sub ScanTopNode(ByVal oDMUParent As Product)

Dim oNode As IXMLDOMNode
Dim sRevision As String

'Add top node to collection
Set oNode = oConfXML.selectSingleNode(".//Instance[@PartNumber = '" & oDMUParent.PartNumber & "']")
sRevision = oNode.Attributes.getNamedItem("DocRev").nodeValue
Call oRefProductList.Add(oDMUParent.PartNumber & sRevision, oDMUParent.ReferenceProduct)

Call ScanExplodedStructure(oDMUParent)
End Sub

Private Sub TransferInstanceName(ByVal sRelationID As Long, ByVal sInstanceName As String)

Dim oConfNodeList As IXMLDOMNodeList
Dim oConfNode As IXMLDOMNode

Set oConfNodeList = oConfXML.selectNodes("//Instance[@RelationID='" & sRelationID & "']")
For Each oConfNode In oConfNodeList
    oConfNode.Attributes.getNamedItem("InstanceName").nodeValue = sInstanceName
Next

End Sub



Private Function CheckFileExist(ByVal sFileName As String) As Boolean

Dim oFSO

Set oFSO = CreateObject("Scripting.FileSystemObject")

If oFSO.FileExists(sFileName) = True Then
    CheckFileExist = True
Else
    CheckFileExist = False
End If

End Function


Private Sub CloseDocuments(ByVal oDoc As Document)

Dim i As Integer
Dim sAllDoc As String

'We first close the PVR doc
oDoc.Close

'We close all other opened documents
For i = CATIA.Documents.Count To 1 Step -1
    Set oDoc = CATIA.Documents.Item(i)
    
    If UCase(oDoc.Name) Like "*.CATPART" Or UCase(oDoc.Name) Like "*.CATPRODUCT" Or UCase(oDoc.Name) Like "*.CGR" Then
        oDoc.Close
    End If
Next

'Log list of all open documents
'sAllDoc = "List of loaded documents" & vbCrLf
'For i = 1 To CATIA.Documents.Count
'    sAllDoc = sAllDoc & " - " & CATIA.Documents.Item(i).Name & vbCrLf
'Next
'Call AddToTrackingLog(sAllDoc, False, True)

End Sub

Private Sub GenerateDTExportReport()

Dim sEntireFile As String
Dim sTextPart1 As String
Dim sTextPart2 As String
Dim sPartNumber As String
Dim sPartType As String
Dim sKey As String
Dim sRev As String
Dim i As Integer
Dim sIteration As String
Dim oNode As IXMLDOMNode
Dim sAnswer As String
Dim oFSO
Dim oTextStream

'Initialize
Set oFSO = CreateObject("Scripting.FileSystemObject")
   
'Read template file
Set oTextStream = oFSO.OpenTextFile(sDTExportTemplateFile, 1, False, 0)
sEntireFile = oTextStream.ReadAll
oTextStream.Close
    
'Split text
sTextPart1 = Split(sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(0) & "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>" & vbLf
sTextPart2 = Split(sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(1)

'Full text is the first part
sEntireFile = sTextPart1


For i = 1 To oRefProductList.Count
    
    'Extract Ref Product info
    sKey = oRefProductList.GetKey(i)
    sRev = Right(sKey, 2)
    sPartNumber = Left(sKey, Len(sKey) - 2)
    
    'Get attributes from Web Service
    sIteration = oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "DOCUMENT_ITERATION")
        
    'Get the Part Type from the XML
    Set oNode = Nothing
    Set oNode = oConfXML.selectSingleNode("//Instance[@PartNumber='" & sPartNumber & "' and @DocRev='" & sRev & "']")
    If Not oNode Is Nothing Then
        sPartType = oNode.Attributes.getNamedItem("DocType").nodeValue
        
        If sPartType <> "SPComponent" Then
            If sPartType = "Component" Then sPartType = "CATProduct"
            Call AddNewLineToDTExport(sEntireFile, sPartNumber, sRev, sPartType, sIteration)
        End If
    End If
Next

'Add second part of the text
sEntireFile = sEntireFile & sTextPart2

'Report date
sEntireFile = Replace(sEntireFile, "<b>xxx</b>", "<b>Report Generated: " & Format(Now(), "yyyy mm dd") & " - " & Format(Time, "h:m:s") & "</b>")

'Save file
Set oTextStream = oFSO.OpenTextFile(sDestinationFolder & "DTExportReport.html", 2, True, 0)
oTextStream.WriteLine sEntireFile
oTextStream.Close
Set oTextStream = Nothing
Set oFSO = Nothing

End Sub

Private Sub GenerateFileBaseStructure()

Dim sPartNumber As String
Dim sRev As String
Dim sParentPN As String
Dim sParentRev As String
Dim sExtension As String
Dim sDocType As String
Dim sIteration As String
Dim oNode As IXMLDOMNode
Dim oNodeList As IXMLDOMNodeList
Dim oProduct As Product
Dim i As Integer
Dim iCountMax As Integer
Dim iCount As Integer


'Initialize
Call oRefProductList.RemoveAll

'Delete child instances of all CATProducts
Set oNodeList = oConfXML.selectNodes(".//Instance[@DocType='CATProduct']/Instance")
For Each oNode In oNodeList
    Call oNode.parentNode.RemoveChild(oNode)
Next

'Clear attribute value
Set oNodeList = oConfXML.selectNodes(".//Instance")
For Each oNode In oNodeList
    oNode.Attributes.getNamedItem("SyncStatus").nodeValue = ""
    oNode.Attributes.getNamedItem("InstanceName").nodeValue = ""
Next

'Save XML
If bSaveXML = True Then
    oConfXML.Save sXMLPath & "Conf.xml"
End If

'Copy the Conf.xml in destination folder
Call CopyFile("Conf.xml", sXMLPath, sDestinationFolder)

'Get info of Top Product
Set oNode = oConfXML.selectSingleNode("./Instance")
sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue

'Get attributes from Web Service
sIteration = oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "DOCUMENT_ITERATION")

'Create Top Product
Call AddNewProduct("", "", sPartNumber, sRev, oNode, sIteration)
If bCancelAction = True Then Exit Sub

'Count max
iCountMax = oConfXML.selectNodes("/Instance//Instance[@SyncStatus='']").Length

'Loop thru all the instances with Status = ""
Do

    'Search nodes with Status = ""
    Set oNodeList = oConfXML.selectNodes("/Instance//Instance[@SyncStatus='']")
    If oNodeList.Length = 0 Then Exit Do
    
    'Progress bar
    If bCancelAction = True Then Exit Sub
    Call frmProgress.progressBarRepaint("Step 7 of 8 - Generate new structure", 8, 7, "Adding instance " & iCountMax - oNodeList.Length & " of " & iCountMax, iCountMax, iCountMax - oNodeList.Length)
    
    'Get the first node in the list
    Set oNode = oNodeList.Item(0)
    
    'Get node info
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    sDocType = oNode.Attributes.getNamedItem("DocType").nodeValue
    sParentPN = oNode.parentNode.Attributes.getNamedItem("PartNumber").nodeValue
    sParentRev = oNode.parentNode.Attributes.getNamedItem("DocRev").nodeValue
    
    'Get attributes from Web Service
    sIteration = oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "DOCUMENT_ITERATION")
    
    'For components create a new CATProduct
    If sDocType = "Component" Then
         Call AddNewProduct(sParentPN, sParentRev, sPartNumber, sRev, oNode, sIteration)
    
    'For CATPart and CATProduct add instance from a saved document
    Else
        sExtension = IIf(sDocType = "CATPart", "CATPart", "CATProduct")
        Call AddComponentFromFile(sParentPN, sParentRev, sPartNumber, sRev, sExtension, oNode, sIteration)
    End If
Loop


'Count max
iCountMax = oConfXML.selectNodes("//Instance[@DocType='CATPart']").Length
iCountMax = iCountMax + oConfXML.selectNodes(".//Instance[@DocType='CATProduct']").Length
iCountMax = iCountMax + oConfXML.selectNodes(".//Instance[@DocType='Component']").Length
iCount = 0


'Save all CATParts
Set oNodeList = oConfXML.selectNodes("//Instance[@DocType='CATPart']")
For Each oNode In oNodeList
    iCount = iCount + 1
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set oProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)
    Call frmProgress.progressBarRepaint("Step 7 of 7 - Saving new structure", 7, 7, "Saving document " & iCount & " of " & iCountMax, iCountMax, iCount)
    If oProduct.ReferenceProduct.Parent.Saved = False Then
        oProduct.ReferenceProduct.Parent.Save
    End If
Next

'Save all CATProducts
Set oNodeList = oConfXML.selectNodes(".//Instance[@DocType='CATProduct']")
For i = oNodeList.Length To 1 Step -1
    iCount = iCount + 1
    Set oNode = oNodeList.Item(i - 1)
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set oProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)
    Call frmProgress.progressBarRepaint("Step 7 of 7 - Saving new structure", 7, 7, "Saving document " & iCount & " of " & iCountMax, iCountMax, iCount)
    If oProduct.ReferenceProduct.Parent.Saved = False Then
        oProduct.ReferenceProduct.Parent.Save
    End If
Next

'Save all components
Set oNodeList = oConfXML.selectNodes(".//Instance[@DocType='Component']")
For i = oNodeList.Length To 1 Step -1
    iCount = iCount + 1
    Set oNode = oNodeList.Item(i - 1)
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set oProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)
    Call frmProgress.progressBarRepaint("Step 7 of 7 - Saving new structure", 7, 7, "Saving document " & iCount & " of " & iCountMax, iCountMax, iCount)
    If oProduct.ReferenceProduct.Parent.Saved = False Then
        oProduct.ReferenceProduct.Parent.Save
    End If
Next

'Save PVRREF
Set oNode = oConfXML.selectSingleNode(".//Instance[@DocType='PVRREF']")
sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
Set oProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)
oProduct.Parent.Save

Set oNodeList = oConfXML.selectNodes(".//Instance[@DocType='Component']")
For i = oNodeList.Length To 1 Step -1
    iCount = iCount + 1
    Set oNode = oNodeList.Item(i - 1)
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set oProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)
    If oProduct.ReferenceProduct.Parent.Saved = False Then
        oProduct.ReferenceProduct.Parent.Save
    End If
Next

Call MsgBox("Use Save Management to make sure all documents were saved.", vbExclamation)

End Sub


Private Sub AddNewLineToMetadata(ByRef sFullText As String, ByVal sPartNumber As String, ByVal sRev As String, ByVal sPartType As String)

Dim sReleaseDate, sFormatReleaseDate As String

'Get and format release date
sReleaseDate = ""
sReleaseDate = oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "REV_LAST_MOD_DATE")

If sReleaseDate <> "" Then
    sFormatReleaseDate = Left(sReleaseDate, 4)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(sReleaseDate, 5, 2)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(sReleaseDate, 7)
Else
    sFormatReleaseDate = ""
End If

'Part Type
If UCase(sPartType) = "PVRREF" Then sPartType = "CATProduct"

sFullText = sFullText & "<TR>"
sFullText = sFullText & "<A NAME=" & Chr(34) & sPartNumber & Chr(34) & ">"
sFullText = sFullText & "<TH>"
sFullText = sFullText & "<A HREF=" & Chr(34) & "#TOP" & Chr(34) & ">" & sPartNumber & "</A>"
sFullText = sFullText & "</TH>"
sFullText = sFullText & "</A>"
sFullText = sFullText & "<TD>"
sFullText = sFullText & "<PRE>Base Number          :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Base Number") & "</STRONG>"
sFullText = sFullText & "<BR>Dash Number          :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Dash Number") & "</STRONG>"
sFullText = sFullText & "<BR>Document Revision    :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "BA Document Revision") & "</STRONG>"
sFullText = sFullText & "<BR>Revision Status      :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Revision Status") & "</STRONG>"
sFullText = sFullText & "<BR>Title                :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Title") & "</STRONG>"
sFullText = sFullText & "<BR>Major Supplier Code  :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Major Supplier Code") & "</STRONG>"
sFullText = sFullText & "<BR>Release Date         :<STRONG> " & sFormatReleaseDate & "</STRONG>"
sFullText = sFullText & "<BR>Dataset Type         :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Dataset Type") & "</STRONG>"
sFullText = sFullText & "<BR>File Name            :<STRONG> " & sPartNumber & " " & sRev & "." & sPartType & "</STRONG>"
sFullText = sFullText & "<BR>Organization(Part)   :<STRONG> " & "" & "</STRONG>"
sFullText = sFullText & "<BR>Organization(Doc)    :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Revision Organization") & "</STRONG>"
sFullText = sFullText & "<BR>Shareable            :<STRONG> " & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Shareable") & "</STRONG>"

sFullText = sFullText & "</PRE>"
sFullText = sFullText & "</TD>"
sFullText = sFullText & "</TR>"



End Sub

Private Sub AddNewLineToDTExport(ByRef sFullText As String, ByVal sPartNumber As String, ByVal sRev As String, ByVal sPartType As String, ByVal sIteration As String)

Dim sSecurityCheck As String

'Get the security check
sSecurityCheck = Right(oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Security Check"), 2)

'Part Type
If UCase(sPartType) = "PVRREF" Then sPartType = "CATProduct"

sFullText = sFullText & "<tr>"
sFullText = sFullText & "<td>" & sPartNumber & "</td>"
sFullText = sFullText & "<td>" & sRev & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "DOCUMENT_ITERATION") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Revision Status") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Revision Organization") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Revision Project") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Shareable") & "</td>"
sFullText = sFullText & "<td>" & oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "Title") & "</td>"
sFullText = sFullText & "<td>" & sPartNumber & " " & sRev & "_" & sIteration & "." & sPartType & "</td>"
sFullText = sFullText & "<td>" & Left(sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "<td>" & Right(sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "</tr>"
sFullText = sFullText & vbLf
End Sub
Private Sub GenerateMetadataHTML()

Dim sEntireFile As String
Dim sTextPart1 As String
Dim sTextPart2 As String
Dim sPartNumber As String
Dim sPartType As String
Dim sKey As String
Dim sRev As String
Dim i As Integer
Dim oNode As IXMLDOMNode
Dim sAnswer As String
Dim oFSO
Dim oTextStream

'Initialize
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Read template file
Set oTextStream = oFSO.OpenTextFile(sMetadataTemplateFile, 1, False, 0)
sEntireFile = oTextStream.ReadAll
oTextStream.Close

'Split text
sTextPart1 = Split(sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(0) & "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>"
sTextPart2 = Split(sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(1)

'Full text is the first part
sEntireFile = sTextPart1

'Add all rows to full text
For i = 1 To oRefProductList.Count
        
    'Extract Ref Product info
    sKey = oRefProductList.GetKey(i)
    sRev = Right(sKey, 2)
    sPartNumber = Left(sKey, Len(sKey) - 2)
    
        
    'Get the Part Type from the XML
    Set oNode = Nothing
    Set oNode = oConfXML.selectSingleNode("//Instance[@PartNumber='" & sPartNumber & "' and @DocRev='" & sRev & "']")
    If Not oNode Is Nothing Then
        sPartType = oNode.Attributes.getNamedItem("DocType").nodeValue
        
        If sPartType <> "SPComponent" Then
            If sPartType = "Component" Then sPartType = "CATProduct"
            Call AddNewLineToMetadata(sEntireFile, sPartNumber, sRev, sPartType)
        End If
    End If
    
Next

'Add second part of the text
sEntireFile = sEntireFile & sTextPart2

'Report data
sEntireFile = Replace(sEntireFile, "<p>xxx</p>", "<p>Report Generated: " & Format(Now(), "yyyy mm dd") & " - " & Format(Time, "h:m:s") & "</p>")

'Save file
Set oTextStream = oFSO.OpenTextFile(sDestinationFolder & "MetadataPackage.html", 2, True, 0)
oTextStream.WriteLine sEntireFile
oTextStream.Close
Set oTextStream = Nothing
Set oFSO = Nothing

End Sub


Private Sub ScanAndSaveBlackBox()

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim oNodeList2 As IXMLDOMNodeList
Dim oNode2 As IXMLDOMNode
Dim sPartNumber As String
Dim sRev As String
Dim oRefProduct As Product
Dim oDoc As Document
Dim oSavedList As New clsCollection
Dim i As Integer
Dim sIteration As String
Dim iCount, iCountMax As Integer
Dim sAnswer As String
Dim sMessage As String

'Initialize
CATIA.DisplayFileAlerts = False

'Count
iCount = 0
iCountMax = oConfXML.selectNodes(".//Instance[@DocType = 'CATPart']").Length
iCountMax = iCountMax + oConfXML.selectNodes(".//Instance[@DocType = 'CATProduct']").Length

'Get the list of each CATPart and save the document
Set oNodeList = oConfXML.selectNodes(".//Instance[@DocType = 'CATPart']")
For i = oNodeList.Length To 1 Step -1
    
    'Get the node
    Set oNode = oNodeList.Item(i - 1)

    'Get attributes
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    
    'Count
    iCount = iCount + 1
    
    'Document is not already saved, we save it
    If Not oSavedList.Exists(sPartNumber & sRev) Then
    
        'Progress bar
        If bCancelAction = True Then Exit Sub
        Call frmProgress.progressBarRepaint("Step 4 of 8 - Saving black box documents", 8, 4, "Saving " & sPartNumber & " " & sRev & " / " & iCount & " of " & iCountMax, iCountMax, iCount)
    
        'Get attributes from Web Service
        sIteration = oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "DOCUMENT_ITERATION")
    
        'Get Reference Product
        Set oRefProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)
    
        'Set the revision
        oRefProduct.Revision = sRev
        
        'Get the Doc
        Set oDoc = oRefProduct.Parent

        'Save the document
        On Error Resume Next
        Err.Clear
        Call oDoc.SaveAs(sDestinationFolder & sPartNumber & " " & sRev & "_" & sIteration)
        
        'Error management
        If Err.Number <> 0 Then
            sMessage = "The tool can't save " & sPartNumber & " " & sRev & "_" & sIteration & "."
            sMessage = sMessage & vbCrLf & "Do you want to stop the process ?"
            sAnswer = MsgBox(sMessage, 20)
            
            'Exit process
            If sAnswer = vbYes Then
                bCancelAction = True
                Exit Sub
            'Add to error list and remove from XML
            Else
                Call oSaveErrDoc.Add(sPartNumber & sRev, sPartNumber & " " & sRev & "_" & sIteration)
                
                Set oNodeList2 = oConfXML.selectNodes(".//Instance[@PartNumber='" & sPartNumber & "' and @DocRev='" & sRev & "']")
                For Each oNode2 In oNodeList2
                    Call oNode2.parentNode.RemoveChild(oNode2)
                Next

            End If
        Else
            Call oSavedList.Add(sPartNumber & sRev, sPartNumber & sRev)
        End If
        On Error GoTo 0
    End If
Next

'Get the list of each CATProduct and save the document starting from the last one
Set oNodeList = oConfXML.selectNodes(".//Instance[@DocType = 'CATProduct']")
For i = oNodeList.Length To 1 Step -1
    
    'Get the node
    Set oNode = oNodeList.Item(i - 1)

    'Get attributes
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
    
    'Count
    iCount = iCount + 1
    
    'Document is not already saved, we save it
    If Not oSavedList.Exists(sPartNumber & sRev) Then
    
        'Progress bar
        If bCancelAction = True Then Exit Sub
        Call frmProgress.progressBarRepaint("Step 4 of 8 - Saving black box documents", 8, 4, "Saving " & sPartNumber & " " & sRev & " / " & iCount & " of " & iCountMax, iCountMax, iCount)
    
        'Get attributes from Web Service
        sIteration = oAttList.GetEnoviaAttributes(sPartNumber, sRev, False, "DOCUMENT_ITERATION")
    
        'Get Reference Product
        Set oRefProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)
    
        'Set the revision
        oRefProduct.Revision = sRev
        
        'Get the Doc
        Set oDoc = oRefProduct.Parent

        'Save the document
        On Error Resume Next
        Err.Clear
        Call oDoc.SaveAs(sDestinationFolder & sPartNumber & " " & sRev & "_" & sIteration)
        
        'Error management
        If Err.Number <> 0 Then
            sMessage = "The tool can't save " & sPartNumber & " " & sRev & "_" & sIteration & "."
            sMessage = sMessage & vbCrLf & "Do you want to stop the process ?"
            sAnswer = MsgBox(sMessage, 20)
            
            'Exit process
            If sAnswer = vbYes Then
                bCancelAction = True
                Exit Sub
            'Add to error list and remove from XML
            Else
                Call oSaveErrDoc.Add(sPartNumber & sRev, sPartNumber & " " & sRev & "_" & sIteration)
                
                Set oNodeList2 = oConfXML.selectNodes(".//Instance[@PartNumber='" & sPartNumber & "' and @DocRev='" & sRev & "']")
                For Each oNode2 In oNodeList2
                    Call oNode2.parentNode.RemoveChild(oNode2)
                Next

            End If
        Else
            Call oSavedList.Add(sPartNumber & sRev, sPartNumber & sRev)
        End If
        On Error GoTo 0
    End If
Next

CATIA.DisplayFileAlerts = True
End Sub

Sub LoadDocument()

Dim EnoviaDoc As EnoviaDocument
Dim EV5product As Product

Set EnoviaDoc = CATIA.Application

CATIA.DisplayFileAlerts = False
Set EV5product = EnoviaDoc.OpenPartDocument("G25002012-101", "--")

MsgBox ("Allo")

End Sub

'***************************************************************************
'*
'*                                  MAIN
'*
'***************************************************************************
Public Sub StartProcess()

Dim oPVRDoc As ProductDocument
Dim bExitByUser As Boolean
Dim sAnswer As String
Dim oWindow As Window

'Error Log General Settings
sErrorLogFilePath = Environ("temp") & "\PVRSync_ErrorLog.txt"

'XML General Settings
bSaveXML = True
sXMLPath = Environ("temp") & "\"

'Setting
bCancelAction = False

'A window must be open
If CATIA.Windows.Count = 0 Then
    Call MsgBox("Active window must be a PVR REF. Process aborted.", vbCritical, "PVR Sync Tool")
    GoTo endsub
End If

'Check the active window
Set oWindow = CATIA.ActiveWindow
If Not oWindow.Name Like "ENOVIA5\*PVRREF*.CATProduct" Then
    Call MsgBox("Active window must be a PVR REF. Process aborted.", vbCritical, "PVR Sync Tool")
    GoTo endsub
End If

'Initialize objects
Call InitializeObjects

'Get the PVR document and retrieve information on the document
Set oPVRDoc = CATIA.ActiveDocument

'Check with user
'If MsgBox("Are you sure you want to save this PVR ?", 36) = vbNo Then GoTo endsub

'Load the xml structure
Call oConfXML.Load(sXMLPath & "Conf.xml")

'Check if template CATProduct exist
If CheckFileExist(sCATProductTemplateFile) = False Then
    Call MsgBox("The template for CATProduct can't be found. Process aborted.", vbCritical)
    GoTo endsub
End If

'Select the destination folder
Call SelectDestinationFolder
If sDestinationFolder = "" Then
    Call MsgBox("Destination folder is not selected. Process aborted.", vbCritical)
    GoTo endsub
End If

'Copy the PVRSync_ConfiguredStructure_Report_xxxxxx in destination folder
Call CopyFile("PVRSync_ConfiguredStructure_Report_" & oPVRDoc.Product.PartNumber & ".xml", sXMLPath, sDestinationFolder)

'Put everything in design mode
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarInitialize("PVR Sync Tool")
Call frmProgress.progressBarRepaint("Step 1 of 8 - Loading parts", 8, 1)
oPVRDoc.Product.ApplyWorkMode DESIGN_MODE

'Scan the structure
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 2 of 8 - Scan 3D structure", 8, 2)
Call ScanTopNode(oPVRDoc.Product.ReferenceProduct)
Call ScanCATProducts

'Getting primary document info for each Non-CI NONE
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 3 of 8 - Retrieving primary document", 8, 3)
Call GettingPrimaryDocument(oPVRDoc.Product.PartNumber)

'Save all the blackbox documents in the destination folder
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 4 of 8 - Saving black box documents", 8, 4)
Call ScanAndSaveBlackBox

'Save
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
If bSaveXML = True Then
    oPVRXML.Save sXMLPath & "PVR.xml"
    oConfXML.Save sXMLPath & "Conf.xml"
End If

'Close all the documents
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call CloseDocuments(oPVRDoc)

'Generating Metadata report
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 5 of 8 - Generating Metadata report", 8, 5)
Call GenerateMetadataHTML

'Generating DTExport report
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 6 of 8 - Generating DTExport report", 8, 6)
Call GenerateDTExportReport

'Generate the structure and save it file base
If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call frmProgress.progressBarRepaint("Step 7 of 8 - Generate new structure", 8, 7)
Call GenerateFileBaseStructure

'Error log
Call GenerateErrorLog

'Exit
Set oConfXML = Nothing
endsub:
Unload frmProgress


End Sub


Sub InitializeObjects()

Dim sArray(0, 1) As String

'Initialize attribute class
Set oAttList = New clsAttributesList

'Initialize oConfXML
Set oConfXML = New DOMDocument60
oConfXML.setProperty "SelectionLanguage", "XPath"
oConfXML.async = True

'Initialize oPVRXML
Set oPVRXML = New DOMDocument60
oPVRXML.setProperty "SelectionLanguage", "XPath"
oPVRXML.async = True

'Initialize oRefProductList
Set oRefProductList = New clsCollection

'Initialize oPartList
Set oPartList = New clsCollection

'Load Error Document list
Set oLoadErrDoc = New clsCollection
Set oSaveErrDoc = New clsCollection

'Initialize oTopNode
Set oTopNode = New clsPartInfo


End Sub
Private Sub ScanCATProducts()

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim sPartNumber As String
Dim sRev As String
Dim oRefProduct As Product


Set oNodeList = oConfXML.selectNodes(".//Instance[@DocType = 'CATProduct']")
For Each oNode In oNodeList

    If bCancelAction = True Then Exit Sub
    
    'Get attributes
    sPartNumber = oNode.Attributes.getNamedItem("PartNumber").nodeValue
    sRev = oNode.Attributes.getNamedItem("DocRev").nodeValue
        
    'Get Reference Product
    Set oRefProduct = oRefProductList.GetItemByKey(sPartNumber & sRev)

    'Scan one CATProduct
    Call ScanCATProduct(oRefProduct, oNode)
Next

''Save
'If bSaveXML = True Then
'    oPVRXML.Save sXMLPath & "PVR.xml"
'    oConfXML.Save sXMLPath & "Conf.xml"
'End If

End Sub

Private Sub ScanCATProduct(ByVal oParentProduct As Product, ByVal oParentNode As IXMLDOMElement)

Dim oChild As Product
Dim oDoc As Document
Dim oChildPart As clsPartInfo
Dim oChildElem As IXMLDOMElement
Dim sString, sMessage As String
Dim sPN As String
Dim sAnswer As String
Dim i As Integer

'Scan all children products
For i = oParentProduct.Products.Count To 1 Step -1

    Set oChild = oParentProduct.Products.Item(i)
    Set oChildPart = New clsPartInfo
    
    'Retrieve Part Number
    On Error Resume Next
    sPN = ""
    sPN = oChild.PartNumber
    If sPN = "" Then
        oChild.ApplyWorkMode DEFAULT_MODE
    End If
    sPN = oChild.PartNumber
    On Error GoTo 0
    
    If sPN <> "" Then
        
        'Get Type
        oChildPart.PartType = GetType(oChild)
        If oChildPart.PartType = "Component" Then oChildPart.PartType = "SPComponent"
        
        'Get info from file name for CATPart and CATProduct
        If oChildPart.PartType = "CATPart" Or oChildPart.PartType = "CATProduct" Then
            
            'Get document
            Set oDoc = oChild.ReferenceProduct.Parent
    
            'Override Part Number with the one found in the file name
            If oDoc.FullName Like "ENOVIA5*" Then
                sPN = Left(Split(oDoc.Name, ".")(0), Trim(Len(Split(oDoc.Name, ".")(0)) - 2))
                oChildPart.PartRev = Right(Split(oDoc.Name, ".")(0), 2)
            Else
                sPN = Trim(Split(oDoc.Name, " ")(0))
                oChildPart.PartRev = Left(Split(oDoc.Name, " ")(1), 2)
            End If
        Else
            'Dummy revision for components
            oChildPart.PartRev = "--"
        End If
        
        oChildPart.PartNumber = sPN
    Else
        oChildPart.PartNumber = "G25XXXXXX-XXX"
        oChildPart.PartType = "CATPart"
        oChildPart.PartRev = "--"
        sString = "Part reference of instance " & oChild.Name & " can't be loaded and thus will not be saved in the destination folder"
        sString = sString & vbCrLf & "Press:"
        sString = sString & vbCrLf & " - Yes to remove the instance from the structure"
        sString = sString & vbCrLf & " - No to keep the instance in the structure"
        sString = sString & vbCrLf & " - Cancel to stop the process"
        
        sAnswer = MsgBox(sString, 67, "PVR Sync")
        If sAnswer = vbYes Then
            oParentProduct.Products.Remove (i)
        ElseIf sAnswer = vbCancel Then
            bCancelAction = True
            Exit Sub
        End If
    End If
    
    'Add to oRefProductList if it doesn't exist
    'Inside the top pvr we can have two components with the same part number if they belong to two different blackbox. However I can't have two entries
    'with the same key in oRefProductList. For that reason I will edit the Part Number of the "SPComponent" part type by adding a dummy string (oRefProductList.Count) to make it unique
    If oChildPart.PartType = "SPComponent" Then oChildPart.PartNumber = oChildPart.PartNumber & oRefProductList.Count
    If oChildPart.PartNumber <> "G25XXXXXX-XXX" Then
        
        'Add to list
        If Not oRefProductList.Exists(oChildPart.PartNumber & oChildPart.PartRev) Then
            Call oRefProductList.Add(oChildPart.PartNumber & oChildPart.PartRev, oChild.ReferenceProduct)
        End If
        
        'Add instance to structure
        Set oChildElem = Add_Element(oConfXML, "Instance", oParentNode)
        Call Add_Attribute(oConfXML, oChildElem, "PartNumber", oChildPart.PartNumber)
        Call Add_Attribute(oConfXML, oChildElem, "InstanceName", oChild.Name)
        Call Add_Attribute(oConfXML, oChildElem, "DocRev", oChildPart.PartRev)
        Call Add_Attribute(oConfXML, oChildElem, "DocType", oChildPart.PartType)

        'Recursive call on Components and CATProduct
        If oChildPart.PartType = "SPComponent" Or oChildPart.PartType = "CATProduct" Then
            Call ScanCATProduct(oChild, oChildElem)
        End If
    End If
    

Next

End Sub





Private Sub SelectDestinationFolder()

Dim sAnswer As String

If Len(Trim$(sLangflowDestinationFolder)) > 0 Then
    sDestinationFolder = Trim$(sLangflowDestinationFolder)
    If Right$(sDestinationFolder, 1) <> Chr$(92) Then sDestinationFolder = sDestinationFolder & Chr$(92)
    Exit Sub
End If
Do
    'Select destination folder
    sDestinationFolder = OpenDirectoryDialog("Select Destination Folder")
    
    'Check sDestinationFolder <> ""
    If sDestinationFolder = "" Then
        Exit Sub
    End If
    
    'Check destination folder doesn't contain a white space
    If sDestinationFolder Like "* *" Then
        sAnswer = MsgBox(sDestinationFolder & " is invalid. Please select a path without any white spaces.", 16, "Destination Folder")
        GoTo TryAgain
    End If

    'Destination folder is ok, exit
    sDestinationFolder = sDestinationFolder & "\"
    Exit Do
    
TryAgain:
Loop

End Sub

Private Sub SetInstancesStatusOKorMoved(ByVal oDoc As DOMDocument60, ByVal sParentPN As String, ByVal sInstanceName As String, ByVal sStatus As String)

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode

'All instances with same parent PN and instance name are to be tagged with sStatus
Set oNodeList = oDoc.selectNodes("//Instance[@PartNumber='" & sParentPN & "']/Instance[@InstanceName='" & sInstanceName & "' and @SyncStatus='']")
For Each oNode In oNodeList
    oNode.Attributes.getNamedItem("SyncStatus").nodeValue = sStatus
Next

End Sub

Private Sub SetInstancesStatusKO(ByVal oDoc As DOMDocument60, ByVal sPN As String, ByVal sRev As String)

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode


'All instances with sPN should be set to "KO"
Set oNodeList = oDoc.selectNodes("//Instance[@PartNumber='" & sPN & "' and @DocRev='" & sRev & "']")
For Each oNode In oNodeList
    oNode.Attributes.getNamedItem("Status").nodeValue = "KO"
Next

End Sub

Private Sub SetInstancesStatusDelete(ByVal sParentPN As String, ByVal sInstanceName As String)

Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim oChildNodeList As IXMLDOMNodeList
Dim oChildNode As IXMLDOMNode

'All instances with same parent PN and instance name are to be tagged with "Deleted"
Set oNodeList = oPVRXML.selectNodes("//Instance[@PartNumber='" & sParentPN & "']/Instance[@InstanceName='" & sInstanceName & "']")
For Each oNode In oNodeList

    'Set Status of active node
    oNode.Attributes.getNamedItem("SyncStatus").nodeValue = "Deleted"
    
    'Set Status to "Deleted" for all children (all levels) of the active node
    Set oChildNodeList = oNode.selectNodes(".//Instance")
    For Each oChildNode In oChildNodeList
        oChildNode.Attributes.getNamedItem("SyncStatus").nodeValue = "Deleted"
    Next
Next

End Sub

Private Function GetClosestInstance(ByVal oConfNodeList As IXMLDOMNodeList, ByVal oPVRNode As IXMLDOMNode) As IXMLDOMNode

Dim i As Integer
Dim oConfNode As IXMLDOMNode
Dim dClosestDistance As Double
Dim dDistance As Double
Dim iIndex As Integer
Dim dPVRCoord(1 To 3) As Double
Dim dConfCoord(1 To 3) As Double

'Initial set
dClosestDistance = 99999999

'We only have one node in the list
If oConfNodeList.Length = 1 Then
    Set GetClosestInstance = oConfNodeList.Item(0)
    Exit Function
End If

'Get coordinates of PVR instance
dPVRCoord(1) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue
dPVRCoord(2) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue
dPVRCoord(3) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue

'Find the instance which is closest to PVR one
For i = 1 To oConfNodeList.Length
    
    Set oConfNode = oConfNodeList.Item(i - 1)
    
    'Get coordinates of Conf instance
    dConfCoord(1) = oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue
    dConfCoord(2) = oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue
    dConfCoord(3) = oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue

    'Calculate distance
    dDistance = Sqr((dPVRCoord(1) - dConfCoord(1)) ^ 2 + (dPVRCoord(2) - dConfCoord(2)) ^ 2 + (dPVRCoord(3) - dConfCoord(3)) ^ 2)

    'Compare distance
    If dDistance < dClosestDistance Then
        dClosestDistance = dDistance
        iIndex = i - 1
    End If
Next

'Return
Set GetClosestInstance = oConfNodeList.Item(iIndex)

End Function

Private Sub AddComponent(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sNewPN As String, oConfElem As IXMLDOMElement, ByVal sRev As String)

Dim oRefProduct As Product
Dim oParentRefProduct As Product
Dim oNewProduct As Product
Dim sRelationID As String

'Get the parent reference product
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)

'Check if we can copy an existing component
Set oRefProduct = Nothing
If oRefProductList.Exists(sNewPN & sRev) Then
    Set oRefProduct = oRefProductList.GetItemByKey(sNewPN & sRev)
End If

'Create a new component
If oRefProduct Is Nothing Then

    Set oNewProduct = oParentRefProduct.Products.AddNewProduct(sNewPN)
    Call AddAttributesToComponent(oNewProduct, sRev)
    
    'Add reference product to oRefProductList
    Call oRefProductList.Add(sNewPN & sRev, oNewProduct.ReferenceProduct)
    
'Create a component from an existing one
Else
    
    'Create new component
    Set oNewProduct = oParentRefProduct.Products.AddComponent(oRefProduct)
End If

'Get the RelationID of oConfElem
sRelationID = oConfElem.Attributes.getNamedItem("RelationID").nodeValue

'Transfer the instance name to all Conf instance with same RelationID
Call TransferInstanceName(sRelationID, oNewProduct.Name)

'Set status
Call SetInstancesStatusOKorMoved(oConfXML, sParentPN, oNewProduct.Name, "OK")

'Move instance
Call MoveInstance(sParentPN, sParentRev, oNewProduct.Name, oConfElem)

End Sub

Private Sub AddAttributesToComponent(ByVal oProduct As Product, ByVal sRev As String)

Dim oProperty As StrParam
Static oAttToCreate As clsCollection
Dim i As Integer
Dim sParamName As String
Dim sParamNameInList As String
Dim sParamValue As String
Dim sCheckString As String
Dim dLocalTimer As Double

If oAttToCreate Is Nothing Then

    Set oAttToCreate = New clsCollection
    dLocalTimer = Timer
    'The Key is the name of the attribute to be added and the value is the name of the attribute in clsAttributesList
    Call oAttToCreate.Add("Title", "Title")
    Call oAttToCreate.Add("Dataset Type", "Dataset Type")
    Call oAttToCreate.Add("Design Authority Program", "Design Authority Program")
    Call oAttToCreate.Add("Supplier Name And CAGE Code", "Supplier Name And CAGE Code")
    Call oAttToCreate.Add("Major Supplier Code", "Major Supplier Code")
    Call oAttToCreate.Add("3D Only", "3D Only")
    Call oAttToCreate.Add("Color Coded", "Color Coded")
    Call oAttToCreate.Add("PCCN", "PCCN")
    Call oAttToCreate.Add("Material Specifications", "Material Specifications")
    Call oAttToCreate.Add("Material Description", "Material Description")
    Call oAttToCreate.Add("Material Type", "Material Type")
    Call oAttToCreate.Add("Material Form", "Material Form")
    Call oAttToCreate.Add("Size", "Size")
    Call oAttToCreate.Add("Thickness", "Thickness")
    Call oAttToCreate.Add("Inside Diameter", "Inside Diameter")
    Call oAttToCreate.Add("Outside Diameter", "Outside Diameter")
    Call oAttToCreate.Add("Length", "Length")
    Call oAttToCreate.Add("Width", "Width")
    Call oAttToCreate.Add("Wall", "Wall")
    Call oAttToCreate.Add("Alloy", "Alloy")
    Call oAttToCreate.Add("Final Condition", "Final Condition")
    Call oAttToCreate.Add("Density", "Density")
    Call oAttToCreate.Add("Eng. Make From", "Eng. Make From")
    Call oAttToCreate.Add("Form", "Form")
    Call oAttToCreate.Add("Grade/Composition", "Grade/Composition")
    Call oAttToCreate.Add("Material Class", "Material Class")
    Call oAttToCreate.Add("Mesh Cell Size", "Mesh Cell Size")
    Call oAttToCreate.Add("Standard Spec Die", "Standard Spec Die")
    Call oAttToCreate.Add("Type", "Type")
    Call oAttToCreate.Add("TD Material Code", "TD Material Code")
    Call oAttToCreate.Add("Finish Code", "Finish Code")
    Call oAttToCreate.Add("MFG Process", "MFG Process")
    Call oAttToCreate.Add("Defining Part", "Defining Part")
    Call oAttToCreate.Add("Material Specification Production", "Material Specification Production")
    Call oAttToCreate.Add("Material Description Production", "Material Description Production")
    Call oAttToCreate.Add("Organization", "Revision Organization")
    Call oAttToCreate.Add("Project", "Revision Project")

End If

For i = 1 To oAttToCreate.Count
    
    'Get the parameter name
    sParamName = oAttToCreate.GetKey(i)
    sParamNameInList = oAttToCreate.GetItemByIndex(i)
    
    'Check if parameter exits
    On Error Resume Next
    sCheckString = ""
    sCheckString = oProduct.ReferenceProduct.UserRefProperties.Item(sParamName).Name
    On Error GoTo 0
    
    'Parameter doesn't exist we need to create it
    If sCheckString = "" Then
        sParamValue = oAttList.GetEnoviaAttributes(oProduct.PartNumber, sRev, False, sParamNameInList)
        Set oProperty = oProduct.ReferenceProduct.UserRefProperties.CreateString(sParamName, sParamValue)
        
    'Parameter exist, just change the value
    Else
        sParamValue = oAttList.GetEnoviaAttributes(oProduct.PartNumber, sRev, False, sParamNameInList)
        Call oProduct.UserRefProperties.Item(sParamName).ValuateFromString(sParamValue)
    End If
Next

End Sub

Private Sub AddComponentFromFile(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sNewPN As String, ByVal sNewRev As String, sNewExtension As String, oConfElem As IXMLDOMElement, Optional sIteration As String = "")

Dim oPVRWindow As Window
Dim oNewDoc As Document
Dim oParentRefProduct As Product
Dim oNewProduct As Product
Dim sRelationID As String
Dim dTimer As Double
Dim dLoadTimer As Double
Dim bTimeout As Boolean
Dim EnoviaDoc As EnoviaDocument
Dim EV5product As Product
Dim sLoadStatus As String

'Initialize
Set oPVRWindow = CATIA.ActiveWindow
Set EnoviaDoc = CATIA.Application
sLoadStatus = "Already Loaded"

 
'Get the parent reference product
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)

'Check if document already loaded in session.
'If sIteration is = "" it means we are looking for a document loaded from Enovia
'If sIteration <> "" it means we are looking for a document saved file base for data transfer
Set oNewDoc = Nothing
On Error Resume Next
If sIteration = "" Then
    Set oNewDoc = CATIA.Documents.Item(sNewPN & sNewRev & "." & sNewExtension)
Else
    Set oNewDoc = CATIA.Documents.Item(sNewPN & " " & sNewRev & "_" & sIteration & "." & sNewExtension)
End If
On Error GoTo 0


'Check if document is in oLoadErrDoc
If oLoadErrDoc.Exists(sNewPN & sNewRev) Then
    sLoadStatus = "Load Error"
End If

'Open document from ENOVIA
If oNewDoc Is Nothing And sLoadStatus <> "Load Error" Then

    If sIteration = "" Then
        'Set Status
        sLoadStatus = "New Load"
    
        'Loading document
        dTimer = Timer
    
        On Error Resume Next
        Err.Clear
        Set EV5product = EnoviaDoc.OpenPartDocument(sNewPN, sNewRev)
    
        'Make sure the document is loaded
        dLoadTimer = Timer
        bTimeout = False
        Do
            Set oNewDoc = EV5product.ReferenceProduct.Parent
    
            If Not oNewDoc Is Nothing Then Exit Do
    
            DoEvents
            Sleep 150
            If Abs(Timer - dLoadTimer) > 5 Then
                bTimeout = True
                Exit Do
            End If
        Loop
        On Error GoTo 0
    
        'Timeout reached, add document to oLoadErrDoc
        If bTimeout = True Then
            Call oLoadErrDoc.Add(sNewPN & sNewRev, "AddComponentFromFile")
            Call SetInstancesStatusKO(oConfXML, sNewPN, sNewRev)
            sLoadStatus = "Load Error"
        End If
    
        'Swap windows
        oPVRWindow.Activate
    Else
        Set oNewDoc = CATIA.Documents.Read(sDestinationFolder & sNewPN & " " & sNewRev & "_" & sIteration & "." & sNewExtension)

    End If
    
End If

'We have a document, let's add an instance to the PVR structure
If sLoadStatus = "Already Loaded" Or sLoadStatus = "New Load" Then

    'Add instance in PVR
    Set oNewProduct = oParentRefProduct.Products.AddExternalComponent(oNewDoc)
    
    'Add to oRefProductList
    If Not oRefProductList.Exists(sNewPN & sNewRev) Then
        Call oRefProductList.Add(sNewPN & sNewRev, oNewProduct.ReferenceProduct)
    End If
        
    'Get the RelationID of oConfElem
    sRelationID = oConfElem.Attributes.getNamedItem("RelationID").nodeValue
    
    'Transfer the instance name to all Conf instance with same RelationID
    Call TransferInstanceName(sRelationID, oNewProduct.Name)
    
    'Set status
    Call SetInstancesStatusOKorMoved(oConfXML, sParentPN, oNewProduct.Name, "OK")
    
    'Move instance
    Call MoveInstance(sParentPN, sParentRev, oNewProduct.Name, oConfElem)
        

End If

'We need to close a newly open document
If sLoadStatus = "New Load" Then

    'Close document
    oNewDoc.Close
    
End If


End Sub

Private Sub MoveInstance(ByVal sParentPN As String, ByVal sParentRev As String, ByVal sInstanceName As String, ByVal oConfElem As IXMLDOMElement)

Dim oParentRefProduct As Product
Dim oChildProduct As Product
Dim oPosition
Dim oMatrix(11) As Variant

'Get the Parent Reference Product
Set oParentRefProduct = oRefProductList.GetItemByKey(sParentPN & sParentRev)

'Get child
Set oChildProduct = oParentRefProduct.Products.Item(sInstanceName)

'Retrieve position matrix
oMatrix(0) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue)
oMatrix(1) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue)
oMatrix(2) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue)
oMatrix(3) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue)
oMatrix(4) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue)
oMatrix(5) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue)
oMatrix(6) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue)
oMatrix(7) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue)
oMatrix(8) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue)
oMatrix(9) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue)
oMatrix(10) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue)
oMatrix(11) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue)

'Position instance
Set oPosition = oChildProduct.Position
oPosition.SetComponents oMatrix

End Sub

Private Function ComparePosition(ByVal oPVRElem As IXMLDOMElement, ByVal oConfElem As IXMLDOMElement) As String

Dim sReturnValue As String

'Initial set
sReturnValue = "Same Position"

'Check all position
If oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue), 6) Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue), 6) Then
    sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue), 6) Then
    sReturnValue = "Position Different"
    Exit Function
End If

ComparePosition = sReturnValue


End Function
Private Sub ScanExplodedStructure(ByVal oDMUParent As Product)

Dim oChild As Product
Dim sPartNumber As String
Dim sRevision As String
Dim oNode As IXMLDOMNode
Dim sType As String
Dim oDoc As Document

'Scan all children products
For Each oChild In oDMUParent.Products

    'Retrieve Part Number
    On Error Resume Next
    sPartNumber = ""
    sPartNumber = oChild.PartNumber
    If sPartNumber = "" Then
        oChild.ApplyWorkMode DESIGN_MODE
    End If
    sPartNumber = oChild.PartNumber
    On Error GoTo 0
    
    'Part Number was found
    If sPartNumber <> "" Then
        
        'Get Type
        sType = GetType(oChild)
        
        'Override Part Number with the one found in the file name
        If sType = "CATPart" Or sType = "CATProduct" Then
            
            'Get document
            Set oDoc = oChild.ReferenceProduct.Parent
    
            'Override Part Number with the one found in the file name
            If oDoc.FullName Like "ENOVIA5*" Then
                sPartNumber = Left(Split(oDoc.Name, ".")(0), Trim(Len(Split(oDoc.Name, ".")(0)) - 2))
            Else
                sPartNumber = Trim(Split(oDoc.Name, " ")(0))
            End If
        End If
        
        'Get the revision from the XML
        Set oNode = oConfXML.selectSingleNode(".//Instance[@PartNumber = '" & sPartNumber & "']")
        sRevision = oNode.Attributes.getNamedItem("DocRev").nodeValue
        
        'Add to collection
        If Not oRefProductList.Exists(sPartNumber & sRevision) Then
            Call oRefProductList.Add(sPartNumber & sRevision, oChild.ReferenceProduct)
        End If

    'Add to error log
    Else
        Call oLoadErrDoc.Add(oLoadErrDoc.Count + 1, "Document associated to " & oChild & " can't be loaded.")
    End If
    
    'Recursive call on Components
    If sType = "Component" Then
        Call ScanExplodedStructure(oChild)
    End If

Next
End Sub




Private Function PopulateChildrenFromWebService(ByVal oData As Variant) As clsCollection

Dim i, j As Integer
Dim sPN As String
Dim oChildren As clsCollection
Dim oItem As clsCollection

'Initialize
Set oChildren = New clsCollection

'Scan oData and transfer value to oChildren
For i = 1 To oData.Count
    
    For j = 2 To oData.Item(i).Count
    
        'Get Part Number
        sPN = oData.Item(i).Item(1)

        Set oItem = New clsCollection
        Call oItem.Add("Part Number", sPN)
        Call oItem.Add("Revision", "N/A")
        Call oItem.Add("Position Matrix", oData.Item(i).Item(j).Item(2))
        
        Call oChildren.Add("Instance" & oChildren.Count + 1, oItem)
    Next
Next

'Return
Set PopulateChildrenFromWebService = oChildren

End Function

Private Function GetExistingNode(ByVal oDoc As DOMDocument60, ByVal sPN As String) As IXMLDOMElement

Dim oNodeList As IXMLDOMNodeList

Set oNodeList = oDoc.selectNodes("//Instance[@PartNumber='" & sPN & "' and @SyncStatus='']")
If oNodeList.Length >= 1 Then
    Set GetExistingNode = oNodeList.Item(0)
Else
    Set GetExistingNode = Nothing
End If

End Function

Private Function GetType(ByVal oProduct As Product) As String

'CATPart
If oProduct.HasAMasterShapeRepresentation Then
    GetType = "CATPart"
'Component
ElseIf oProduct.ReferenceProduct.Parent.Name = oProduct.Parent.Parent.ReferenceProduct.Parent.Name Then
    GetType = "Component"
'CATProduct
Else
    GetType = "CATProduct"
End If
 
End Function

Private Function Add_Element(ByVal oDoc As DOMDocument60, ByVal sTagName As String, ByVal oParentElem As IXMLDOMElement, Optional oBrotherElem = Nothing) As IXMLDOMElement

Dim oElement As IXMLDOMElement

'Add new element
Set oElement = oDoc.CreateElement(sTagName)

'Append node
If oBrotherElem Is Nothing Then
    oParentElem.appendChild oElement
Else
    Call oParentElem.InsertBefore(oElement, oBrotherElem)
End If

'Return
Set Add_Element = oElement

End Function

Private Sub Add_Attribute(ByVal oDoc As DOMDocument60, ByRef oElement As IXMLDOMElement, ByVal sAttributeName As String, ByVal sAttributeValue As String)

Dim oAttribute As IXMLDOMAttribute

'Add attribute
Set oAttribute = oDoc.createAttribute(sAttributeName)
oAttribute.nodeValue = sAttributeValue
oElement.setAttributeNode oAttribute

End Sub

Private Sub Add_Comment(ByVal oDoc As DOMDocument60, ByRef oElement As IXMLDOMElement, ByVal sCommentValue As String)

Dim oComment As IXMLDOMComment

'Add comment
Set oComment = oDoc.createComment(sCommentValue)
oElement.appendChild oComment

End Sub
Private Sub AddToTrackingLog(ByVal sText As String, ByVal bOverwrite As Boolean, ByVal bSaveNow As Boolean)

'Add text to log text
If sTrackingLogText = "" Then
    sTrackingLogText = sText
Else
    sTrackingLogText = sTrackingLogText & vbCrLf & sText
End If

'Display in debug.print
Debug.Print sText


'Save log file
If bTrackingLogNeverSave = False And (bSaveNow = True Or bTrackingLogSaveEverytime = True) Then
    Call WriteTextFile(sTrackingLogFilePath, sTrackingLogText, bOverwrite)
    sTrackingLogText = ""
End If

End Sub

Private Sub GenerateErrorLog()

Dim sText As String
Dim i As Integer
Dim sAnswer As String

If oSaveErrDoc.Count > 0 Then
    
    sText = "Following document could not be save:"
    For i = 1 To oSaveErrDoc.Count
        sText = sText & vbCrLf & "- " & oSaveErrDoc.GetItemByIndex(i)
    Next
    
    Call WriteTextFile(sErrorLogFilePath, sText, True)
    
    Call MsgBox("Save errors were found. Refer to " & sErrorLogFilePath, 48)
End If


End Sub

Private Sub GenerateLogReport()

Dim oTopElem As IXMLDOMElement
Dim oElem As IXMLDOMElement
Dim oNode As IXMLDOMNode
Dim oNodeList As IXMLDOMNodeList
Dim i As Integer
Dim oPart As clsPartInfo

'Create new xml structure
Set oLogReport = New DOMDocument60
oLogReport.setProperty "SelectionLanguage", "XPath"
oLogReport.async = True

'Create top node
Set oTopElem = oLogReport.CreateElement("Parts")
oLogReport.appendChild oTopElem

'Create comment on top node
If oTopNode.PartIsCI Then
    Call Add_Comment(oLogReport, oTopElem, oTopNode.PartNumber & " is a CI")
Else
    Call Add_Comment(oLogReport, oTopElem, oTopNode.PartNumber & " is not a CI")
End If
If oTopNode.SyncFromBSF = True Then
    Call Add_Comment(oLogReport, oTopElem, "PVR updated using best so far structure from ENOVIA")
Else
    Call Add_Comment(oLogReport, oTopElem, "Structure configured for " & oTopNode.Project & "/" & oTopNode.Tail)
End If
If oTopNode.NonCIOption = "BSF" Then
    Call Add_Comment(oLogReport, oTopElem, "User choose to update the non CIs using the Best so Far from Enovia")
Else
    Call Add_Comment(oLogReport, oTopElem, "User choose to update the non CIs using the latest released revision")
End If

Call Add_Comment(oLogReport, oTopElem, "Report generated on " & Format(Now(), "yyyy mm dd") & " at " & Format(Time, "hh:mm:ss"))

'Scan all part in oPartList
For i = 1 To oPartList.Count
    
    'Get the part
    Set oPart = oPartList.GetItemByIndex(i)
    
    'Create new elem
    Set oElem = Add_Element(oLogReport, "Part", oTopElem)
    
    'Add attributes
    Call Add_Attribute(oLogReport, oElem, "PartNumber", oPart.PartNumber)
    Call Add_Attribute(oLogReport, oElem, "Revision", oPart.PartRev)
    Call Add_Attribute(oLogReport, oElem, "Status", oPart.PartStatus)
    Call Add_Attribute(oLogReport, oElem, "DocumentOrganization", oAttList.GetEnoviaAttributes(oPart.PartNumber, oPart.PartRev, False, "Document Organization"))
    Call Add_Attribute(oLogReport, oElem, "Title", oPart.PartTitle)
    Call Add_Attribute(oLogReport, oElem, "IsCI", CStr(oPart.PartIsCI))
    Call Add_Attribute(oLogReport, oElem, "Type", IIf(oPart.PartType = "Component", "NONE", oPart.PartType))
    Call Add_Attribute(oLogReport, oElem, "EffectiveDwg", IIf(oPart.EffectiveDwgNb = "N/A" Or oPart.EffectiveDwgNb = "", "N/A", oPart.EffectiveDwgNb & oPart.EffectiveDwgRev))
    Call Add_Attribute(oLogReport, oElem, "ProposedSource", IIf(oPart.ProposedSourceNb = "", "N/A", oPart.ProposedSourceNb & oPart.ProposedSourceRev))
    Call Add_Attribute(oLogReport, oElem, "SelectedDwg", IIf(oPart.SelectedDwgNb = "N/A", "N/A", oPart.SelectedDwgNb & oPart.SelectedDwgRev))
    Call Add_Attribute(oLogReport, oElem, "SelectedSource", IIf(oPart.SelectedSourceNb = "", "N/A", oPart.SelectedSourceNb & oPart.SelectedSourceRev))
    
    If oPart.Comment = "No comment" And (oPart.ProposedSourceNb & oPart.ProposedSourceRev) <> (oPart.SelectedSourceNb & oPart.SelectedSourceRev) Then
        Call Add_Attribute(oLogReport, oElem, "Comment", "Selected source is different from proposed source")
    Else
       Call Add_Attribute(oLogReport, oElem, "Comment", oPart.Comment)
    End If
Next


'Save report
oLogReport.Save sXMLPath & "PVRSync_ConfiguredStructure_Report_" & oTopNode.PartNumber & ".xml"


End Sub

Private Function GetDocRevOrderedList(ByVal sDocNumber As String) As clsCollection

Dim oDocs As clsCollection
Dim oTemp
Dim oOrderedList As New clsCollection
Dim oOrderedColl As New Collection
Dim i, j As Integer

'Retrieve all documents with same base number with web service
Set oTemp = WebServiceAccessTool.GetDocumentByBaseNumber(sDocNumber)
Set oDocs = New clsCollection
Call oDocs.InitializeWithDLLclsColObject(oTemp, oDocs)

'Error management
If oDocs.Count = 0 Then
    Set GetDocRevOrderedList = Nothing
    Exit Function
End If

'Remove from the list all document that don't have the same Part Number
For i = oDocs.Count To 1 Step -1
    If oDocs.GetItemByIndex(i).GetItemByKey("FIELD_PART_NUMBER") <> sDocNumber Then oDocs.RemoveByIndex i
Next

'Add all revision to oOrderedColl
For i = 1 To oDocs.Count
    
    'First item
    If oOrderedColl.Count = 0 Then
        oOrderedColl.Add (oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION"))
    
    'Other object
    Else
        For j = 1 To oOrderedColl.Count
            If oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION") < oOrderedColl.Item(j) Then
                oOrderedColl.Add oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION"), , j
                Exit For
            End If
        Next
        If oOrderedColl.Count < i Then oOrderedColl.Add oDocs.GetItemByIndex(i).GetItemByKey("FIELD_DOCUMENT_REVISION")
    End If
Next

'Transfer object from oOrderedColl to oOrderedList
For i = 1 To oOrderedColl.Count
    Call oOrderedList.Add(oOrderedColl.Item(i), oOrderedColl.Item(i))
Next

'Return
Set GetDocRevOrderedList = oOrderedList

End Function

