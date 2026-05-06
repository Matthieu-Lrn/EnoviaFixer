Attribute VB_Name = "SFB_mdlSave3DStructure"
Public Declare Sub SFB_Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Option Explicit
Private SFB_oConfXML As DOMDocument60
Private SFB_oPVRXML As DOMDocument60

Private SFB_oLogReport As DOMDocument60
Private SFB_oAttList As SFB_clsAttributesList
Private SFB_oRefProductList As SFB_clsCollection 'This is SFB_a collection of all the Reference Product found in the open PVR.
Private SFB_oLoadErrDoc As SFB_clsCollection
Private SFB_oSaveErrDoc As SFB_clsCollection
Private SFB_oPartList As SFB_clsCollection
Private SFB_sDestinationFolder As String

Private SFB_bSaveXML As Boolean
Private SFB_sXMLPath As String

Private SFB_sTrackingLogText As String
Private SFB_sTrackingLogFilePath As String
Private SFB_bTrackingLogNeverSave As Boolean
Private SFB_bTrackingLogSaveEverytime As Boolean

Private SFB_sErrorLogFilePath As String
Private SFB_sEffectiveDocReportPath As String

Private SFB_oTopNode As SFB_clsPartInfo


Private Sub SFB_AddNewProduct(ByVal SFB_sParentPN As String, ByVal SFB_sParentRev As String, ByVal sNewPN As String, ByVal sNewRev As String, SFB_oElem As IXMLDOMElement, ByVal SFB_sIteration As String)

Dim SFB_oParentRefProduct As Product
Dim SFB_oProductDoc As ProductDocument
Dim SFB_oProduct As Product
Dim SFB_oFSO
Dim SFB_sFileName As String
Dim SFB_sRelationID As String
Dim SFB_sErrorMsg As String
Dim SFB_dTimer As Double
Dim SFB_iIteration As Integer

'Initialize
Set SFB_oFSO = CreateObject("Scripting.FileSystemObject")
CATIA.DisplayFileAlerts = False

'No parent, we create SFB_a top product
If SFB_sParentPN = "" Then
   
  
    'Define file name
    SFB_sFileName = sNewPN & " " & sNewRev & "_" & SFB_sIteration & ".CATProduct"
    
    'Read template CATProduct
    Set SFB_oProductDoc = CATIA.Documents.Read(SFB_sCATProductTemplateFile)

    'Edit Part Number and revision
    SFB_oProductDoc.Product.PartNumber = sNewPN
    SFB_oProductDoc.Product.Revision = sNewRev

    'Since the "Organization" attribut is locked I can't edit it but (I don't understand why) I can delete it and recreate it
    Call SFB_oProductDoc.Product.ReferenceProduct.UserRefProperties.Remove("Organization")
    Call SFB_oProductDoc.Product.ReferenceProduct.UserRefProperties.CreateString("Organization", "")

    'Edit attributes
    Call SFB_AddAttributesToComponent(SFB_oProductDoc.Product, sNewRev)

    'Save
'   Call SFB_oProductDoc.SaveAs(SFB_sDestinationFolder & sNewPN & sNewRev & "_" & SFB_sIteration & ".CATProduct")
    Call SFB_oProductDoc.SaveAs(SFB_sDestinationFolder & SFB_sFileName)

    'Close the product
    Call SFB_CloseDocuments(SFB_oProductDoc)

    'Open the document
    Set SFB_oProductDoc = CATIA.Documents.Open(SFB_sDestinationFolder & SFB_sFileName)

    'Add to collection
    Call SFB_oRefProductList.Add(sNewPN & sNewRev, SFB_oProductDoc.Product.ReferenceProduct)
        
'We have SFB_a parent
Else

    'Get parent reference product
    Set SFB_oParentRefProduct = SFB_oRefProductList.GetItemByKey(SFB_sParentPN & SFB_sParentRev)

    'Define file name
    SFB_sFileName = sNewPN & " " & sNewRev & "_" & SFB_sIteration & ".CATProduct"

    'Check if document already loaded in session
    Set SFB_oProductDoc = Nothing
    On Error Resume Next
    Set SFB_oProductDoc = CATIA.Documents.Item(SFB_sFileName)
    On Error GoTo 0

    'Create SFB_a new product document
    If SFB_oProductDoc Is Nothing Then

        'Read template CATProduct
        Set SFB_oProductDoc = CATIA.Documents.Read(SFB_sCATProductTemplateFile)
        
        'Edit Part Number and revision
        SFB_oProductDoc.Product.PartNumber = sNewPN
        SFB_oProductDoc.Product.Revision = sNewRev
        
        'Since the "Organization" attribut is locked I can't edit it but (I don't understand why) I can delete it and recreate it
        Call SFB_oProductDoc.Product.ReferenceProduct.UserRefProperties.Remove("Organization")
        Call SFB_oProductDoc.Product.ReferenceProduct.UserRefProperties.CreateString("Organization", "")

        'Edit attributes
        Call SFB_AddAttributesToComponent(SFB_oProductDoc.Product, sNewRev)

        'Save
        Call SFB_oProductDoc.SaveAs(SFB_sDestinationFolder & SFB_sFileName)
    End If
    
    'Add new instance under parent
    Set SFB_oProduct = SFB_oParentRefProduct.Products.AddExternalComponent(SFB_oProductDoc)

    'Add to SFB_oRefProductList
    If Not SFB_oRefProductList.Exists(sNewPN & sNewRev) Then
        Call SFB_oRefProductList.Add(sNewPN & sNewRev, SFB_oProduct.ReferenceProduct)
    End If

    'Get the RelationID of SFB_oElem
    SFB_sRelationID = SFB_oElem.Attributes.getNamedItem("RelationID").nodeValue

    'Transfer the instance name to all Conf instance with same RelationID
    Call SFB_TransferInstanceName(SFB_sRelationID, SFB_oProduct.Name)

    'Set status
    Call SFB_SetInstancesStatusOKorMoved(SFB_oConfXML, SFB_sParentPN, SFB_oProduct.Name, "OK")

    'Move instance
    Call SFB_MoveInstance(SFB_sParentPN, SFB_sParentRev, SFB_oProduct.Name, SFB_oElem)
    
End If
    
End Sub


Private Sub SFB_CopyFile(ByVal SFB_sFileName As String, ByVal SourcePath As String, ByVal TargetPath As String)

Dim SFB_oFSO

'Initialize
Set SFB_oFSO = CreateObject("Scripting.FileSystemObject")

'Check if file exist
If Not SFB_oFSO.FileExists(SourcePath & SFB_sFileName) Then Exit Sub

'Copy file
SFB_oFSO.SFB_CopyFile SourcePath & SFB_sFileName, TargetPath

End Sub

Private Sub SFB_GettingPrimaryDocument(ByVal sTopPN As String)

Dim SFB_oNodes As IXMLDOMNodeList
Dim SFB_oNode As IXMLDOMNode
Dim SFB_sPartNumber As String
Dim SFB_sRev As String
Dim SFB_iCount, SFB_iCountMax As Integer
Dim SFB_sString As String

Set SFB_oLogReport = New DOMDocument60
Call SFB_oLogReport.Load(SFB_sDestinationFolder & "PVRSync_ConfiguredStructure_Report_" & sTopPN & ".xml")

'Creating Primary Document attribut on all nodes
Set SFB_oNodes = SFB_oLogReport.selectNodes(".//Part")
For Each SFB_oNode In SFB_oNodes
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oNode, "PrimaryDocument", "")
Next

'Getting the nodes for NONE + non-CI
Set SFB_oNodes = SFB_oLogReport.selectNodes(".//Part[@IsCI='False' and (@Type='NONE' or @Type='CATProduct')]")
SFB_iCountMax = SFB_oNodes.Length
SFB_iCount = 0
For Each SFB_oNode In SFB_oNodes

    'Count
    SFB_iCount = SFB_iCount + 1
    
    'Get the Ref Part Number from XML
    SFB_sString = SFB_oNode.Attributes.getNamedItem("RefPartNumber").nodeValue

    'Extract the PN and REV
    SFB_sPartNumber = Left(SFB_sString, Len(SFB_sString) - 2)
    SFB_sRev = Right(SFB_sString, 2)
    
    'Progress
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 3 of 8 - Retrieving primary document", 8, 3, "Retrieving primary document of " & SFB_sPartNumber & " " & SFB_sRev & " / " & SFB_iCount & " of " & SFB_iCountMax, SFB_iCountMax, SFB_iCount)
    
    'Get the drawing
    SFB_oNode.Attributes.getNamedItem("PrimaryDocument").nodeValue = SFB_GetPrimaryDocFromPart(SFB_sPartNumber, SFB_sRev)
Next

Call SFB_oLogReport.Save(SFB_sDestinationFolder & "PVRSync_ConfiguredStructure_Report_" & sTopPN & ".xml")
Set SFB_oLogReport = Nothing

End Sub

Private Sub SFB_ScanTopNode(ByVal oDMUParent As Product)

Dim SFB_oNode As IXMLDOMNode
Dim SFB_sRevision As String

'Add top node to collection
Set SFB_oNode = SFB_oConfXML.selectSingleNode(".//Instance[@PartNumber = '" & oDMUParent.PartNumber & "']")
SFB_sRevision = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
Call SFB_oRefProductList.Add(oDMUParent.PartNumber & SFB_sRevision, oDMUParent.ReferenceProduct)

Call SFB_ScanExplodedStructure(oDMUParent)
End Sub

Private Sub SFB_TransferInstanceName(ByVal SFB_sRelationID As Long, ByVal sInstanceName As String)

Dim SFB_oConfNodeList As IXMLDOMNodeList
Dim SFB_oConfNode As IXMLDOMNode

Set SFB_oConfNodeList = SFB_oConfXML.selectNodes("//Instance[@RelationID='" & SFB_sRelationID & "']")
For Each SFB_oConfNode In SFB_oConfNodeList
    SFB_oConfNode.Attributes.getNamedItem("InstanceName").nodeValue = sInstanceName
Next

End Sub



Private Function SFB_CheckFileExist(ByVal SFB_sFileName As String) As Boolean

Dim SFB_oFSO

Set SFB_oFSO = CreateObject("Scripting.FileSystemObject")

If SFB_oFSO.FileExists(SFB_sFileName) = True Then
    SFB_CheckFileExist = True
Else
    SFB_CheckFileExist = False
End If

End Function


Private Sub SFB_CloseDocuments(ByVal SFB_oDoc As Document)

Dim SFB_i As Integer
Dim SFB_sAllDoc As String

'We first close the PVR doc
SFB_oDoc.Close

'We close all other opened documents
For SFB_i = CATIA.Documents.Count To 1 Step -1
    Set SFB_oDoc = CATIA.Documents.Item(SFB_i)
    
    If UCase(SFB_oDoc.Name) Like "*.CATPART" Or UCase(SFB_oDoc.Name) Like "*.CATPRODUCT" Or UCase(SFB_oDoc.Name) Like "*.CGR" Then
        SFB_oDoc.Close
    End If
Next

'Log list of all open documents
'SFB_sAllDoc = "List of loaded documents" & vbCrLf
'For SFB_i = 1 To CATIA.Documents.Count
'    SFB_sAllDoc = SFB_sAllDoc & " - " & CATIA.Documents.Item(SFB_i).Name & vbCrLf
'Next
'Call SFB_AddToTrackingLog(SFB_sAllDoc, False, True)

End Sub

Private Sub SFB_GenerateDTExportReport()

Dim SFB_sEntireFile As String
Dim SFB_sTextPart1 As String
Dim SFB_sTextPart2 As String
Dim SFB_sPartNumber As String
Dim SFB_sPartType As String
Dim SFB_sKey As String
Dim SFB_sRev As String
Dim SFB_i As Integer
Dim SFB_sIteration As String
Dim SFB_oNode As IXMLDOMNode
Dim SFB_sAnswer As String
Dim SFB_oFSO
Dim SFB_oTextStream

'Initialize
Set SFB_oFSO = CreateObject("Scripting.FileSystemObject")
   
'Read template file
Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sDTExportTemplateFile, 1, False, 0)
SFB_sEntireFile = SFB_oTextStream.ReadAll
SFB_oTextStream.Close
    
'Split text
SFB_sTextPart1 = Split(SFB_sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(0) & "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>" & vbLf
SFB_sTextPart2 = Split(SFB_sEntireFile, "<tr><th>PART NUMBER</th><th>REVISION</th><th>ITERATION</th><th>STATUS</th><th>ORGANIZATION</th><th>PROJECT</th><th>SHAREABLE</th><th>TITLE</th><th>FILE NAME</th><th>RA</th><th>EC</th></tr>")(1)

'Full text is the first part
SFB_sEntireFile = SFB_sTextPart1


For SFB_i = 1 To SFB_oRefProductList.Count
    
    'Extract Ref Product info
    SFB_sKey = SFB_oRefProductList.GetKey(SFB_i)
    SFB_sRev = Right(SFB_sKey, 2)
    SFB_sPartNumber = Left(SFB_sKey, Len(SFB_sKey) - 2)
    
    'Get attributes from Web Service
    SFB_sIteration = SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "DOCUMENT_ITERATION")
        
    'Get the Part Type from the XML
    Set SFB_oNode = Nothing
    Set SFB_oNode = SFB_oConfXML.selectSingleNode("//Instance[@PartNumber='" & SFB_sPartNumber & "' and @DocRev='" & SFB_sRev & "']")
    If Not SFB_oNode Is Nothing Then
        SFB_sPartType = SFB_oNode.Attributes.getNamedItem("DocType").nodeValue
        
        If SFB_sPartType <> "SPComponent" Then
            If SFB_sPartType = "Component" Then SFB_sPartType = "CATProduct"
            Call SFB_AddNewLineToDTExport(SFB_sEntireFile, SFB_sPartNumber, SFB_sRev, SFB_sPartType, SFB_sIteration)
        End If
    End If
Next

'Add second part of the text
SFB_sEntireFile = SFB_sEntireFile & SFB_sTextPart2

'Report date
SFB_sEntireFile = Replace(SFB_sEntireFile, "<b>xxx</b>", "<b>Report Generated: " & Format(Now(), "yyyy mm dd") & " - " & Format(Time, "h:m:s") & "</b>")

'Save file
Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sDestinationFolder & "DTExportReport.html", 2, True, 0)
SFB_oTextStream.WriteLine SFB_sEntireFile
SFB_oTextStream.Close
Set SFB_oTextStream = Nothing
Set SFB_oFSO = Nothing

End Sub

Private Sub SFB_GenerateFileBaseStructure()

Dim SFB_sPartNumber As String
Dim SFB_sRev As String
Dim SFB_sParentPN As String
Dim SFB_sParentRev As String
Dim SFB_sExtension As String
Dim SFB_sDocType As String
Dim SFB_sIteration As String
Dim SFB_oNode As IXMLDOMNode
Dim SFB_oNodeList As IXMLDOMNodeList
Dim SFB_oProduct As Product
Dim SFB_i As Integer
Dim SFB_iCountMax As Integer
Dim SFB_iCount As Integer


'Initialize
Call SFB_oRefProductList.RemoveAll

'Delete child instances of all CATProducts
Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance[@DocType='CATProduct']/Instance")
For Each SFB_oNode In SFB_oNodeList
    Call SFB_oNode.parentNode.RemoveChild(SFB_oNode)
Next

'Clear attribute value
Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance")
For Each SFB_oNode In SFB_oNodeList
    SFB_oNode.Attributes.getNamedItem("SyncStatus").nodeValue = ""
    SFB_oNode.Attributes.getNamedItem("InstanceName").nodeValue = ""
Next

'Save XML
If SFB_bSaveXML = True Then
    SFB_oConfXML.Save SFB_sXMLPath & "Conf.xml"
End If

'Copy the Conf.xml in destination folder
Call SFB_CopyFile("Conf.xml", SFB_sXMLPath, SFB_sDestinationFolder)

'Get info of Top Product
Set SFB_oNode = SFB_oConfXML.selectSingleNode("./Instance")
SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue

'Get attributes from Web Service
SFB_sIteration = SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "DOCUMENT_ITERATION")

'Create Top Product
Call SFB_AddNewProduct("", "", SFB_sPartNumber, SFB_sRev, SFB_oNode, SFB_sIteration)
If SFB_bCancelAction = True Then Exit Sub

'Count max
SFB_iCountMax = SFB_oConfXML.selectNodes("/Instance//Instance[@SyncStatus='']").Length

'Loop thru all the instances with Status = ""
Do

    'Search nodes with Status = ""
    Set SFB_oNodeList = SFB_oConfXML.selectNodes("/Instance//Instance[@SyncStatus='']")
    If SFB_oNodeList.Length = 0 Then Exit Do
    
    'Progress bar
    If SFB_bCancelAction = True Then Exit Sub
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 7 of 8 - Generate new structure", 8, 7, "Adding instance " & SFB_iCountMax - SFB_oNodeList.Length & " of " & SFB_iCountMax, SFB_iCountMax, SFB_iCountMax - SFB_oNodeList.Length)
    
    'Get the first node in the list
    Set SFB_oNode = SFB_oNodeList.Item(0)
    
    'Get node info
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
    SFB_sDocType = SFB_oNode.Attributes.getNamedItem("DocType").nodeValue
    SFB_sParentPN = SFB_oNode.parentNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sParentRev = SFB_oNode.parentNode.Attributes.getNamedItem("DocRev").nodeValue
    
    'Get attributes from Web Service
    SFB_sIteration = SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "DOCUMENT_ITERATION")
    
    'For components create SFB_a new CATProduct
    If SFB_sDocType = "Component" Then
         Call SFB_AddNewProduct(SFB_sParentPN, SFB_sParentRev, SFB_sPartNumber, SFB_sRev, SFB_oNode, SFB_sIteration)
    
    'For CATPart and CATProduct add instance from SFB_a saved document
    Else
        SFB_sExtension = IIf(SFB_sDocType = "CATPart", "CATPart", "CATProduct")
        Call SFB_AddComponentFromFile(SFB_sParentPN, SFB_sParentRev, SFB_sPartNumber, SFB_sRev, SFB_sExtension, SFB_oNode, SFB_sIteration)
    End If
Loop


'Count max
SFB_iCountMax = SFB_oConfXML.selectNodes("//Instance[@DocType='CATPart']").Length
SFB_iCountMax = SFB_iCountMax + SFB_oConfXML.selectNodes(".//Instance[@DocType='CATProduct']").Length
SFB_iCountMax = SFB_iCountMax + SFB_oConfXML.selectNodes(".//Instance[@DocType='Component']").Length
SFB_iCount = 0


'Save all CATParts
Set SFB_oNodeList = SFB_oConfXML.selectNodes("//Instance[@DocType='CATPart']")
For Each SFB_oNode In SFB_oNodeList
    SFB_iCount = SFB_iCount + 1
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set SFB_oProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 7 of 7 - Saving new structure", 7, 7, "Saving document " & SFB_iCount & " of " & SFB_iCountMax, SFB_iCountMax, SFB_iCount)
    If SFB_oProduct.ReferenceProduct.Parent.Saved = False Then
        SFB_oProduct.ReferenceProduct.Parent.Save
    End If
Next

'Save all CATProducts
Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance[@DocType='CATProduct']")
For SFB_i = SFB_oNodeList.Length To 1 Step -1
    SFB_iCount = SFB_iCount + 1
    Set SFB_oNode = SFB_oNodeList.Item(SFB_i - 1)
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set SFB_oProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 7 of 7 - Saving new structure", 7, 7, "Saving document " & SFB_iCount & " of " & SFB_iCountMax, SFB_iCountMax, SFB_iCount)
    If SFB_oProduct.ReferenceProduct.Parent.Saved = False Then
        SFB_oProduct.ReferenceProduct.Parent.Save
    End If
Next

'Save all components
Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance[@DocType='Component']")
For SFB_i = SFB_oNodeList.Length To 1 Step -1
    SFB_iCount = SFB_iCount + 1
    Set SFB_oNode = SFB_oNodeList.Item(SFB_i - 1)
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set SFB_oProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)
    Call SFB_frmProgress.SFB_progressBarRepaint("Step 7 of 7 - Saving new structure", 7, 7, "Saving document " & SFB_iCount & " of " & SFB_iCountMax, SFB_iCountMax, SFB_iCount)
    If SFB_oProduct.ReferenceProduct.Parent.Saved = False Then
        SFB_oProduct.ReferenceProduct.Parent.Save
    End If
Next

'Save PVRREF
Set SFB_oNode = SFB_oConfXML.selectSingleNode(".//Instance[@DocType='PVRREF']")
SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
Set SFB_oProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)
SFB_oProduct.Parent.Save

Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance[@DocType='Component']")
For SFB_i = SFB_oNodeList.Length To 1 Step -1
    SFB_iCount = SFB_iCount + 1
    Set SFB_oNode = SFB_oNodeList.Item(SFB_i - 1)
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
    Set SFB_oProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)
    If SFB_oProduct.ReferenceProduct.Parent.Saved = False Then
        SFB_oProduct.ReferenceProduct.Parent.Save
    End If
Next

Call MsgBox("Use Save Management to make sure all documents were saved.", vbExclamation)

End Sub


Private Sub SFB_AddNewLineToMetadata(ByRef sFullText As String, ByVal SFB_sPartNumber As String, ByVal SFB_sRev As String, ByVal SFB_sPartType As String)

Dim SFB_sReleaseDate, sFormatReleaseDate As String

'Get and format release date
SFB_sReleaseDate = ""
SFB_sReleaseDate = SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "REV_LAST_MOD_DATE")

If SFB_sReleaseDate <> "" Then
    sFormatReleaseDate = Left(SFB_sReleaseDate, 4)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(SFB_sReleaseDate, 5, 2)
    sFormatReleaseDate = sFormatReleaseDate & "/"
    sFormatReleaseDate = sFormatReleaseDate & Mid(SFB_sReleaseDate, 7)
Else
    sFormatReleaseDate = ""
End If

'Part Type
If UCase(SFB_sPartType) = "PVRREF" Then SFB_sPartType = "CATProduct"

sFullText = sFullText & "<TR>"
sFullText = sFullText & "<A NAME=" & Chr(34) & SFB_sPartNumber & Chr(34) & ">"
sFullText = sFullText & "<TH>"
sFullText = sFullText & "<A HREF=" & Chr(34) & "#TOP" & Chr(34) & ">" & SFB_sPartNumber & "</A>"
sFullText = sFullText & "</TH>"
sFullText = sFullText & "</A>"
sFullText = sFullText & "<TD>"
sFullText = sFullText & "<PRE>Base Number          :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Base Number") & "</STRONG>"
sFullText = sFullText & "<BR>Dash Number          :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Dash Number") & "</STRONG>"
sFullText = sFullText & "<BR>Document Revision    :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "BA Document Revision") & "</STRONG>"
sFullText = sFullText & "<BR>Revision Status      :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Revision Status") & "</STRONG>"
sFullText = sFullText & "<BR>Title                :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Title") & "</STRONG>"
sFullText = sFullText & "<BR>Major Supplier Code  :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Major Supplier Code") & "</STRONG>"
sFullText = sFullText & "<BR>Release Date         :<STRONG> " & sFormatReleaseDate & "</STRONG>"
sFullText = sFullText & "<BR>Dataset Type         :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Dataset Type") & "</STRONG>"
sFullText = sFullText & "<BR>File Name            :<STRONG> " & SFB_sPartNumber & " " & SFB_sRev & "." & SFB_sPartType & "</STRONG>"
sFullText = sFullText & "<BR>Organization(Part)   :<STRONG> " & "" & "</STRONG>"
sFullText = sFullText & "<BR>Organization(Doc)    :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Revision Organization") & "</STRONG>"
sFullText = sFullText & "<BR>Shareable            :<STRONG> " & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Shareable") & "</STRONG>"

sFullText = sFullText & "</PRE>"
sFullText = sFullText & "</TD>"
sFullText = sFullText & "</TR>"



End Sub

Private Sub SFB_AddNewLineToDTExport(ByRef sFullText As String, ByVal SFB_sPartNumber As String, ByVal SFB_sRev As String, ByVal SFB_sPartType As String, ByVal SFB_sIteration As String)

Dim SFB_sSecurityCheck As String

'Get the security check
SFB_sSecurityCheck = Right(SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Security Check"), 2)

'Part Type
If UCase(SFB_sPartType) = "PVRREF" Then SFB_sPartType = "CATProduct"

sFullText = sFullText & "<tr>"
sFullText = sFullText & "<td>" & SFB_sPartNumber & "</td>"
sFullText = sFullText & "<td>" & SFB_sRev & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "DOCUMENT_ITERATION") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Revision Status") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Revision Organization") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Revision Project") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Shareable") & "</td>"
sFullText = sFullText & "<td>" & SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "Title") & "</td>"
sFullText = sFullText & "<td>" & SFB_sPartNumber & " " & SFB_sRev & "_" & SFB_sIteration & "." & SFB_sPartType & "</td>"
sFullText = sFullText & "<td>" & Left(SFB_sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "<td>" & Right(SFB_sSecurityCheck, 1) & "</td>"
sFullText = sFullText & "</tr>"
sFullText = sFullText & vbLf
End Sub
Private Sub SFB_GenerateMetadataHTML()

Dim SFB_sEntireFile As String
Dim SFB_sTextPart1 As String
Dim SFB_sTextPart2 As String
Dim SFB_sPartNumber As String
Dim SFB_sPartType As String
Dim SFB_sKey As String
Dim SFB_sRev As String
Dim SFB_i As Integer
Dim SFB_oNode As IXMLDOMNode
Dim SFB_sAnswer As String
Dim SFB_oFSO
Dim SFB_oTextStream

'Initialize
Set SFB_oFSO = CreateObject("Scripting.FileSystemObject")

'Read template file
Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sMetadataTemplateFile, 1, False, 0)
SFB_sEntireFile = SFB_oTextStream.ReadAll
SFB_oTextStream.Close

'Split text
SFB_sTextPart1 = Split(SFB_sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(0) & "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>"
SFB_sTextPart2 = Split(SFB_sEntireFile, "<TH bgcolor=" & Chr(34) & "#DDDDDD" & Chr(34) & ">Document Info</TH>")(1)

'Full text is the first part
SFB_sEntireFile = SFB_sTextPart1

'Add all rows to full text
For SFB_i = 1 To SFB_oRefProductList.Count
        
    'Extract Ref Product info
    SFB_sKey = SFB_oRefProductList.GetKey(SFB_i)
    SFB_sRev = Right(SFB_sKey, 2)
    SFB_sPartNumber = Left(SFB_sKey, Len(SFB_sKey) - 2)
    
        
    'Get the Part Type from the XML
    Set SFB_oNode = Nothing
    Set SFB_oNode = SFB_oConfXML.selectSingleNode("//Instance[@PartNumber='" & SFB_sPartNumber & "' and @DocRev='" & SFB_sRev & "']")
    If Not SFB_oNode Is Nothing Then
        SFB_sPartType = SFB_oNode.Attributes.getNamedItem("DocType").nodeValue
        
        If SFB_sPartType <> "SPComponent" Then
            If SFB_sPartType = "Component" Then SFB_sPartType = "CATProduct"
            Call SFB_AddNewLineToMetadata(SFB_sEntireFile, SFB_sPartNumber, SFB_sRev, SFB_sPartType)
        End If
    End If
    
Next

'Add second part of the text
SFB_sEntireFile = SFB_sEntireFile & SFB_sTextPart2

'Report data
SFB_sEntireFile = Replace(SFB_sEntireFile, "<p>xxx</p>", "<p>Report Generated: " & Format(Now(), "yyyy mm dd") & " - " & Format(Time, "h:m:s") & "</p>")

'Save file
Set SFB_oTextStream = SFB_oFSO.OpenTextFile(SFB_sDestinationFolder & "MetadataPackage.html", 2, True, 0)
SFB_oTextStream.WriteLine SFB_sEntireFile
SFB_oTextStream.Close
Set SFB_oTextStream = Nothing
Set SFB_oFSO = Nothing

End Sub


Private Sub SFB_ScanAndSaveBlackBox()

Dim SFB_oNodeList As IXMLDOMNodeList
Dim SFB_oNode As IXMLDOMNode
Dim SFB_oNodeList2 As IXMLDOMNodeList
Dim SFB_oNode2 As IXMLDOMNode
Dim SFB_sPartNumber As String
Dim SFB_sRev As String
Dim SFB_oRefProduct As Product
Dim SFB_oDoc As Document
Dim SFB_oSavedList As New SFB_clsCollection
Dim SFB_i As Integer
Dim SFB_sIteration As String
Dim SFB_iCount, SFB_iCountMax As Integer
Dim SFB_sAnswer As String
Dim SFB_sMessage As String

'Initialize
CATIA.DisplayFileAlerts = False

'Count
SFB_iCount = 0
SFB_iCountMax = SFB_oConfXML.selectNodes(".//Instance[@DocType = 'CATPart']").Length
SFB_iCountMax = SFB_iCountMax + SFB_oConfXML.selectNodes(".//Instance[@DocType = 'CATProduct']").Length

'Get the list of each CATPart and save the document
Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance[@DocType = 'CATPart']")
For SFB_i = SFB_oNodeList.Length To 1 Step -1
    
    'Get the node
    Set SFB_oNode = SFB_oNodeList.Item(SFB_i - 1)

    'Get attributes
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
    
    'Count
    SFB_iCount = SFB_iCount + 1
    
    'Document is not already saved, we save it
    If Not SFB_oSavedList.Exists(SFB_sPartNumber & SFB_sRev) Then
    
        'Progress bar
        If SFB_bCancelAction = True Then Exit Sub
        Call SFB_frmProgress.SFB_progressBarRepaint("Step 4 of 8 - Saving black box documents", 8, 4, "Saving " & SFB_sPartNumber & " " & SFB_sRev & " / " & SFB_iCount & " of " & SFB_iCountMax, SFB_iCountMax, SFB_iCount)
    
        'Get attributes from Web Service
        SFB_sIteration = SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "DOCUMENT_ITERATION")
    
        'Get Reference Product
        Set SFB_oRefProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)
    
        'Set the revision
        SFB_oRefProduct.Revision = SFB_sRev
        
        'Get the Doc
        Set SFB_oDoc = SFB_oRefProduct.Parent

        'Save the document
        On Error Resume Next
        Err.Clear
        Call SFB_oDoc.SaveAs(SFB_sDestinationFolder & SFB_sPartNumber & " " & SFB_sRev & "_" & SFB_sIteration)
        
        'Error management
        If Err.Number <> 0 Then
            SFB_sMessage = "The tool can't save " & SFB_sPartNumber & " " & SFB_sRev & "_" & SFB_sIteration & "."
            SFB_sMessage = SFB_sMessage & vbCrLf & "Do you want to stop the process ?"
            SFB_sAnswer = MsgBox(SFB_sMessage, 20)
            
            'Exit process
            If SFB_sAnswer = vbYes Then
                SFB_bCancelAction = True
                Exit Sub
            'Add to error list and remove from XML
            Else
                Call SFB_oSaveErrDoc.Add(SFB_sPartNumber & SFB_sRev, SFB_sPartNumber & " " & SFB_sRev & "_" & SFB_sIteration)
                
                Set SFB_oNodeList2 = SFB_oConfXML.selectNodes(".//Instance[@PartNumber='" & SFB_sPartNumber & "' and @DocRev='" & SFB_sRev & "']")
                For Each SFB_oNode2 In SFB_oNodeList2
                    Call SFB_oNode2.parentNode.RemoveChild(SFB_oNode2)
                Next

            End If
        Else
            Call SFB_oSavedList.Add(SFB_sPartNumber & SFB_sRev, SFB_sPartNumber & SFB_sRev)
        End If
        On Error GoTo 0
    End If
Next

'Get the list of each CATProduct and save the document starting from the last one
Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance[@DocType = 'CATProduct']")
For SFB_i = SFB_oNodeList.Length To 1 Step -1
    
    'Get the node
    Set SFB_oNode = SFB_oNodeList.Item(SFB_i - 1)

    'Get attributes
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
    
    'Count
    SFB_iCount = SFB_iCount + 1
    
    'Document is not already saved, we save it
    If Not SFB_oSavedList.Exists(SFB_sPartNumber & SFB_sRev) Then
    
        'Progress bar
        If SFB_bCancelAction = True Then Exit Sub
        Call SFB_frmProgress.SFB_progressBarRepaint("Step 4 of 8 - Saving black box documents", 8, 4, "Saving " & SFB_sPartNumber & " " & SFB_sRev & " / " & SFB_iCount & " of " & SFB_iCountMax, SFB_iCountMax, SFB_iCount)
    
        'Get attributes from Web Service
        SFB_sIteration = SFB_oAttList.GetEnoviaAttributes(SFB_sPartNumber, SFB_sRev, False, "DOCUMENT_ITERATION")
    
        'Get Reference Product
        Set SFB_oRefProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)
    
        'Set the revision
        SFB_oRefProduct.Revision = SFB_sRev
        
        'Get the Doc
        Set SFB_oDoc = SFB_oRefProduct.Parent

        'Save the document
        On Error Resume Next
        Err.Clear
        Call SFB_oDoc.SaveAs(SFB_sDestinationFolder & SFB_sPartNumber & " " & SFB_sRev & "_" & SFB_sIteration)
        
        'Error management
        If Err.Number <> 0 Then
            SFB_sMessage = "The tool can't save " & SFB_sPartNumber & " " & SFB_sRev & "_" & SFB_sIteration & "."
            SFB_sMessage = SFB_sMessage & vbCrLf & "Do you want to stop the process ?"
            SFB_sAnswer = MsgBox(SFB_sMessage, 20)
            
            'Exit process
            If SFB_sAnswer = vbYes Then
                SFB_bCancelAction = True
                Exit Sub
            'Add to error list and remove from XML
            Else
                Call SFB_oSaveErrDoc.Add(SFB_sPartNumber & SFB_sRev, SFB_sPartNumber & " " & SFB_sRev & "_" & SFB_sIteration)
                
                Set SFB_oNodeList2 = SFB_oConfXML.selectNodes(".//Instance[@PartNumber='" & SFB_sPartNumber & "' and @DocRev='" & SFB_sRev & "']")
                For Each SFB_oNode2 In SFB_oNodeList2
                    Call SFB_oNode2.parentNode.RemoveChild(SFB_oNode2)
                Next

            End If
        Else
            Call SFB_oSavedList.Add(SFB_sPartNumber & SFB_sRev, SFB_sPartNumber & SFB_sRev)
        End If
        On Error GoTo 0
    End If
Next

CATIA.DisplayFileAlerts = True
End Sub

Sub SFB_LoadDocument()

Dim SFB_EnoviaDoc As EnoviaDocument
Dim SFB_EV5product As Product

Set SFB_EnoviaDoc = CATIA.Application

CATIA.DisplayFileAlerts = False
Set SFB_EV5product = SFB_EnoviaDoc.OpenPartDocument("G25002012-101", "--")

MsgBox ("Allo")

End Sub

'***************************************************************************
'*
'*                                  MAIN
'*
'***************************************************************************
Public Sub SFB_StartProcess()

Dim SFB_oPVRDoc As ProductDocument
Dim SFB_bExitByUser As Boolean
Dim SFB_sAnswer As String
Dim SFB_oWindow As Window

'Error Log General Settings
SFB_sErrorLogFilePath = Environ("temp") & "\PVRSync_ErrorLog.txt"

'XML General Settings
SFB_bSaveXML = True
SFB_sXMLPath = Environ("temp") & "\"

'Setting
SFB_bCancelAction = False

'A window must be open
If CATIA.Windows.Count = 0 Then
    Call MsgBox("Active window must be SFB_a PVR REF. Process aborted.", vbCritical, "PVR Sync Tool")
    GoTo endsub
End If

'Check the active window
Set SFB_oWindow = CATIA.ActiveWindow
If Not SFB_oWindow.Name Like "ENOVIA5\*PVRREF*.CATProduct" Then
    Call MsgBox("Active window must be SFB_a PVR REF. Process aborted.", vbCritical, "PVR Sync Tool")
    GoTo endsub
End If

'Initialize objects
Call SFB_InitializeObjects

'Get the PVR document and retrieve information on the document
Set SFB_oPVRDoc = CATIA.ActiveDocument

'Check with user
'If MsgBox("Are you sure you want to save this PVR ?", 36) = vbNo Then GoTo endsub

'Load the xml structure
Call SFB_oConfXML.Load(SFB_sXMLPath & "Conf.xml")

'Check if template CATProduct exist
If SFB_CheckFileExist(SFB_sCATProductTemplateFile) = False Then
    Call MsgBox("The template for CATProduct can't be found. Process aborted.", vbCritical)
    GoTo endsub
End If

'Select the destination folder
Call SFB_SelectDestinationFolder
If SFB_sDestinationFolder = "" Then
    Call MsgBox("Destination folder is not selected. Process aborted.", vbCritical)
    GoTo endsub
End If

'Copy the PVRSync_ConfiguredStructure_Report_xxxxxx in destination folder
Call SFB_CopyFile("PVRSync_ConfiguredStructure_Report_" & SFB_oPVRDoc.Product.PartNumber & ".xml", SFB_sXMLPath, SFB_sDestinationFolder)

'Put everything in design mode
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_frmProgress.SFB_progressBarInitialize("PVR Sync Tool")
Call SFB_frmProgress.SFB_progressBarRepaint("Step 1 of 8 - Loading parts", 8, 1)
SFB_oPVRDoc.Product.ApplyWorkMode DESIGN_MODE

'Scan the structure
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_frmProgress.SFB_progressBarRepaint("Step 2 of 8 - Scan 3D structure", 8, 2)
Call SFB_ScanTopNode(SFB_oPVRDoc.Product.ReferenceProduct)
Call SFB_ScanCATProducts

'Getting primary document info for each Non-CI NONE
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_frmProgress.SFB_progressBarRepaint("Step 3 of 8 - Retrieving primary document", 8, 3)
Call SFB_GettingPrimaryDocument(SFB_oPVRDoc.Product.PartNumber)

'Save all the blackbox documents in the destination folder
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_frmProgress.SFB_progressBarRepaint("Step 4 of 8 - Saving black box documents", 8, 4)
Call SFB_ScanAndSaveBlackBox

'Save
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
If SFB_bSaveXML = True Then
    SFB_oPVRXML.Save SFB_sXMLPath & "PVR.xml"
    SFB_oConfXML.Save SFB_sXMLPath & "Conf.xml"
End If

'Close all the documents
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_CloseDocuments(SFB_oPVRDoc)

'Generating Metadata report
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_frmProgress.SFB_progressBarRepaint("Step 5 of 8 - Generating Metadata report", 8, 5)
Call SFB_GenerateMetadataHTML

'Generating DTExport report
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_frmProgress.SFB_progressBarRepaint("Step 6 of 8 - Generating DTExport report", 8, 6)
Call SFB_GenerateDTExportReport

'Generate the structure and save it file base
If SFB_bCancelAction = True Or SFB_GetAsyncKeyState(vbKeyEscape) <> 0 Then
    GoTo endsub
End If
Call SFB_frmProgress.SFB_progressBarRepaint("Step 7 of 8 - Generate new structure", 8, 7)
Call SFB_GenerateFileBaseStructure

'Error log
Call SFB_GenerateErrorLog

'Exit
Set SFB_oConfXML = Nothing
endsub:
Unload SFB_frmProgress


End Sub


Sub SFB_InitializeObjects()

Dim SFB_sArray(0, 1) As String

'Initialize attribute class
Set SFB_oAttList = New SFB_clsAttributesList

'Initialize SFB_oConfXML
Set SFB_oConfXML = New DOMDocument60
SFB_oConfXML.setProperty "SelectionLanguage", "XPath"
SFB_oConfXML.async = True

'Initialize SFB_oPVRXML
Set SFB_oPVRXML = New DOMDocument60
SFB_oPVRXML.setProperty "SelectionLanguage", "XPath"
SFB_oPVRXML.async = True

'Initialize SFB_oRefProductList
Set SFB_oRefProductList = New SFB_clsCollection

'Initialize SFB_oPartList
Set SFB_oPartList = New SFB_clsCollection

'Load Error Document list
Set SFB_oLoadErrDoc = New SFB_clsCollection
Set SFB_oSaveErrDoc = New SFB_clsCollection

'Initialize SFB_oTopNode
Set SFB_oTopNode = New SFB_clsPartInfo


End Sub
Private Sub SFB_ScanCATProducts()

Dim SFB_oNodeList As IXMLDOMNodeList
Dim SFB_oNode As IXMLDOMNode
Dim SFB_sPartNumber As String
Dim SFB_sRev As String
Dim SFB_oRefProduct As Product


Set SFB_oNodeList = SFB_oConfXML.selectNodes(".//Instance[@DocType = 'CATProduct']")
For Each SFB_oNode In SFB_oNodeList

    If SFB_bCancelAction = True Then Exit Sub
    
    'Get attributes
    SFB_sPartNumber = SFB_oNode.Attributes.getNamedItem("PartNumber").nodeValue
    SFB_sRev = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
        
    'Get Reference Product
    Set SFB_oRefProduct = SFB_oRefProductList.GetItemByKey(SFB_sPartNumber & SFB_sRev)

    'Scan one CATProduct
    Call SFB_ScanCATProduct(SFB_oRefProduct, SFB_oNode)
Next

''Save
'If SFB_bSaveXML = True Then
'    SFB_oPVRXML.Save SFB_sXMLPath & "PVR.xml"
'    SFB_oConfXML.Save SFB_sXMLPath & "Conf.xml"
'End If

End Sub

Private Sub SFB_ScanCATProduct(ByVal oParentProduct As Product, ByVal oParentNode As IXMLDOMElement)

Dim SFB_oChild As Product
Dim SFB_oDoc As Document
Dim SFB_oChildPart As SFB_clsPartInfo
Dim SFB_oChildElem As IXMLDOMElement
Dim SFB_sString, SFB_sMessage As String
Dim SFB_sPN As String
Dim SFB_sAnswer As String
Dim SFB_i As Integer

'Scan all children products
For SFB_i = oParentProduct.Products.Count To 1 Step -1

    Set SFB_oChild = oParentProduct.Products.Item(SFB_i)
    Set SFB_oChildPart = New SFB_clsPartInfo
    
    'SFB_Retrieve Part Number
    On Error Resume Next
    SFB_sPN = ""
    SFB_sPN = SFB_oChild.PartNumber
    If SFB_sPN = "" Then
        SFB_oChild.ApplyWorkMode DEFAULT_MODE
    End If
    SFB_sPN = SFB_oChild.PartNumber
    On Error GoTo 0
    
    If SFB_sPN <> "" Then
        
        'Get Type
        SFB_oChildPart.PartType = SFB_GetType(SFB_oChild)
        If SFB_oChildPart.PartType = "Component" Then SFB_oChildPart.PartType = "SPComponent"
        
        'Get info from file name for CATPart and CATProduct
        If SFB_oChildPart.PartType = "CATPart" Or SFB_oChildPart.PartType = "CATProduct" Then
            
            'Get document
            Set SFB_oDoc = SFB_oChild.ReferenceProduct.Parent
    
            'Override Part Number with the one found in the file name
            If SFB_oDoc.FullName Like "ENOVIA5*" Then
                SFB_sPN = Left(Split(SFB_oDoc.Name, ".")(0), Trim(Len(Split(SFB_oDoc.Name, ".")(0)) - 2))
                SFB_oChildPart.PartRev = Right(Split(SFB_oDoc.Name, ".")(0), 2)
            Else
                SFB_sPN = Trim(Split(SFB_oDoc.Name, " ")(0))
                SFB_oChildPart.PartRev = Left(Split(SFB_oDoc.Name, " ")(1), 2)
            End If
        Else
            'Dummy revision for components
            SFB_oChildPart.PartRev = "--"
        End If
        
        SFB_oChildPart.PartNumber = SFB_sPN
    Else
        SFB_oChildPart.PartNumber = "G25XXXXXX-XXX"
        SFB_oChildPart.PartType = "CATPart"
        SFB_oChildPart.PartRev = "--"
        SFB_sString = "Part reference of instance " & SFB_oChild.Name & " can't be loaded and thus will not be saved in the destination folder"
        SFB_sString = SFB_sString & vbCrLf & "Press:"
        SFB_sString = SFB_sString & vbCrLf & " - Yes to remove the instance from the structure"
        SFB_sString = SFB_sString & vbCrLf & " - No to keep the instance in the structure"
        SFB_sString = SFB_sString & vbCrLf & " - Cancel to stop the process"
        
        SFB_sAnswer = MsgBox(SFB_sString, 67, "PVR Sync")
        If SFB_sAnswer = vbYes Then
            oParentProduct.Products.Remove (SFB_i)
        ElseIf SFB_sAnswer = vbCancel Then
            SFB_bCancelAction = True
            Exit Sub
        End If
    End If
    
    'Add to SFB_oRefProductList if it doesn't exist
    'Inside the top pvr we can have two components with the same part number if they belong to two different blackbox. However I can't have two entries
    'with the same key in SFB_oRefProductList. For that reason I will edit the Part Number of the "SPComponent" part type by adding SFB_a dummy string (SFB_oRefProductList.Count) to make it unique
    If SFB_oChildPart.PartType = "SPComponent" Then SFB_oChildPart.PartNumber = SFB_oChildPart.PartNumber & SFB_oRefProductList.Count
    If SFB_oChildPart.PartNumber <> "G25XXXXXX-XXX" Then
        
        'Add to list
        If Not SFB_oRefProductList.Exists(SFB_oChildPart.PartNumber & SFB_oChildPart.PartRev) Then
            Call SFB_oRefProductList.Add(SFB_oChildPart.PartNumber & SFB_oChildPart.PartRev, SFB_oChild.ReferenceProduct)
        End If
        
        'Add instance to structure
        Set SFB_oChildElem = SFB_Add_Element(SFB_oConfXML, "Instance", oParentNode)
        Call SFB_Add_Attribute(SFB_oConfXML, SFB_oChildElem, "PartNumber", SFB_oChildPart.PartNumber)
        Call SFB_Add_Attribute(SFB_oConfXML, SFB_oChildElem, "InstanceName", SFB_oChild.Name)
        Call SFB_Add_Attribute(SFB_oConfXML, SFB_oChildElem, "DocRev", SFB_oChildPart.PartRev)
        Call SFB_Add_Attribute(SFB_oConfXML, SFB_oChildElem, "DocType", SFB_oChildPart.PartType)

        'Recursive call on Components and CATProduct
        If SFB_oChildPart.PartType = "SPComponent" Or SFB_oChildPart.PartType = "CATProduct" Then
            Call SFB_ScanCATProduct(SFB_oChild, SFB_oChildElem)
        End If
    End If
    

Next

End Sub





Private Sub SFB_SelectDestinationFolder()

Dim SFB_sAnswer As String
Do
    'Select destination folder
    SFB_sDestinationFolder = SFB_OpenDirectoryDialog("Select Destination Folder")
    
    'Check SFB_sDestinationFolder <> ""
    If SFB_sDestinationFolder = "" Then
        Exit Sub
    End If
    
    'Check destination folder doesn't contain SFB_a white space
    If SFB_sDestinationFolder Like "* *" Then
        SFB_sAnswer = MsgBox(SFB_sDestinationFolder & " is invalid. Please select SFB_a path without any white spaces.", 16, "Destination Folder")
        GoTo TryAgain
    End If

    'Destination folder is ok, exit
    SFB_sDestinationFolder = SFB_sDestinationFolder & "\"
    Exit Do
    
TryAgain:
Loop

End Sub

Private Sub SFB_SetInstancesStatusOKorMoved(ByVal SFB_oDoc As DOMDocument60, ByVal SFB_sParentPN As String, ByVal sInstanceName As String, ByVal sStatus As String)

Dim SFB_oNodeList As IXMLDOMNodeList
Dim SFB_oNode As IXMLDOMNode

'All instances with same parent PN and instance name are to be tagged with sStatus
Set SFB_oNodeList = SFB_oDoc.selectNodes("//Instance[@PartNumber='" & SFB_sParentPN & "']/Instance[@InstanceName='" & sInstanceName & "' and @SyncStatus='']")
For Each SFB_oNode In SFB_oNodeList
    SFB_oNode.Attributes.getNamedItem("SyncStatus").nodeValue = sStatus
Next

End Sub

Private Sub SFB_SetInstancesStatusKO(ByVal SFB_oDoc As DOMDocument60, ByVal SFB_sPN As String, ByVal SFB_sRev As String)

Dim SFB_oNodeList As IXMLDOMNodeList
Dim SFB_oNode As IXMLDOMNode


'All instances with SFB_sPN should be set to "KO"
Set SFB_oNodeList = SFB_oDoc.selectNodes("//Instance[@PartNumber='" & SFB_sPN & "' and @DocRev='" & SFB_sRev & "']")
For Each SFB_oNode In SFB_oNodeList
    SFB_oNode.Attributes.getNamedItem("Status").nodeValue = "KO"
Next

End Sub

Private Sub SFB_SetInstancesStatusDelete(ByVal SFB_sParentPN As String, ByVal sInstanceName As String)

Dim SFB_oNodeList As IXMLDOMNodeList
Dim SFB_oNode As IXMLDOMNode
Dim SFB_oChildNodeList As IXMLDOMNodeList
Dim SFB_oChildNode As IXMLDOMNode

'All instances with same parent PN and instance name are to be tagged with "Deleted"
Set SFB_oNodeList = SFB_oPVRXML.selectNodes("//Instance[@PartNumber='" & SFB_sParentPN & "']/Instance[@InstanceName='" & sInstanceName & "']")
For Each SFB_oNode In SFB_oNodeList

    'Set Status of active node
    SFB_oNode.Attributes.getNamedItem("SyncStatus").nodeValue = "Deleted"
    
    'Set Status to "Deleted" for all children (all levels) of the active node
    Set SFB_oChildNodeList = SFB_oNode.selectNodes(".//Instance")
    For Each SFB_oChildNode In SFB_oChildNodeList
        SFB_oChildNode.Attributes.getNamedItem("SyncStatus").nodeValue = "Deleted"
    Next
Next

End Sub

Private Function SFB_GetClosestInstance(ByVal SFB_oConfNodeList As IXMLDOMNodeList, ByVal oPVRNode As IXMLDOMNode) As IXMLDOMNode

Dim SFB_i As Integer
Dim SFB_oConfNode As IXMLDOMNode
Dim SFB_dClosestDistance As Double
Dim SFB_dDistance As Double
Dim SFB_iIndex As Integer
Dim SFB_dPVRCoord(1 To 3) As Double
Dim SFB_dConfCoord(1 To 3) As Double

'Initial set
SFB_dClosestDistance = 99999999

'We only have one node in the list
If SFB_oConfNodeList.Length = 1 Then
    Set SFB_GetClosestInstance = SFB_oConfNodeList.Item(0)
    Exit Function
End If

'Get coordinates of PVR instance
SFB_dPVRCoord(1) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue
SFB_dPVRCoord(2) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue
SFB_dPVRCoord(3) = oPVRNode.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue

'Find the instance which is closest to PVR one
For SFB_i = 1 To SFB_oConfNodeList.Length
    
    Set SFB_oConfNode = SFB_oConfNodeList.Item(SFB_i - 1)
    
    'Get coordinates of Conf instance
    SFB_dConfCoord(1) = SFB_oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue
    SFB_dConfCoord(2) = SFB_oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue
    SFB_dConfCoord(3) = SFB_oConfNode.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue

    'Calculate distance
    SFB_dDistance = Sqr((SFB_dPVRCoord(1) - SFB_dConfCoord(1)) ^ 2 + (SFB_dPVRCoord(2) - SFB_dConfCoord(2)) ^ 2 + (SFB_dPVRCoord(3) - SFB_dConfCoord(3)) ^ 2)

    'Compare distance
    If SFB_dDistance < SFB_dClosestDistance Then
        SFB_dClosestDistance = SFB_dDistance
        SFB_iIndex = SFB_i - 1
    End If
Next

'Return
Set SFB_GetClosestInstance = SFB_oConfNodeList.Item(SFB_iIndex)

End Function

Private Sub SFB_AddComponent(ByVal SFB_sParentPN As String, ByVal SFB_sParentRev As String, ByVal sNewPN As String, oConfElem As IXMLDOMElement, ByVal SFB_sRev As String)

Dim SFB_oRefProduct As Product
Dim SFB_oParentRefProduct As Product
Dim SFB_oNewProduct As Product
Dim SFB_sRelationID As String

'Get the parent reference product
Set SFB_oParentRefProduct = SFB_oRefProductList.GetItemByKey(SFB_sParentPN & SFB_sParentRev)

'Check if we can copy an existing component
Set SFB_oRefProduct = Nothing
If SFB_oRefProductList.Exists(sNewPN & SFB_sRev) Then
    Set SFB_oRefProduct = SFB_oRefProductList.GetItemByKey(sNewPN & SFB_sRev)
End If

'Create SFB_a new component
If SFB_oRefProduct Is Nothing Then

    Set SFB_oNewProduct = SFB_oParentRefProduct.Products.SFB_AddNewProduct(sNewPN)
    Call SFB_AddAttributesToComponent(SFB_oNewProduct, SFB_sRev)
    
    'Add reference product to SFB_oRefProductList
    Call SFB_oRefProductList.Add(sNewPN & SFB_sRev, SFB_oNewProduct.ReferenceProduct)
    
'Create SFB_a component from an existing one
Else
    
    'Create new component
    Set SFB_oNewProduct = SFB_oParentRefProduct.Products.SFB_AddComponent(SFB_oRefProduct)
End If

'Get the RelationID of oConfElem
SFB_sRelationID = oConfElem.Attributes.getNamedItem("RelationID").nodeValue

'Transfer the instance name to all Conf instance with same RelationID
Call SFB_TransferInstanceName(SFB_sRelationID, SFB_oNewProduct.Name)

'Set status
Call SFB_SetInstancesStatusOKorMoved(SFB_oConfXML, SFB_sParentPN, SFB_oNewProduct.Name, "OK")

'Move instance
Call SFB_MoveInstance(SFB_sParentPN, SFB_sParentRev, SFB_oNewProduct.Name, oConfElem)

End Sub

Private Sub SFB_AddAttributesToComponent(ByVal SFB_oProduct As Product, ByVal SFB_sRev As String)

Dim SFB_oProperty As StrParam
Static SFB_oAttToCreate As SFB_clsCollection
Dim SFB_i As Integer
Dim SFB_sParamName As String
Dim SFB_sParamNameInList As String
Dim SFB_sParamValue As String
Dim SFB_sCheckString As String
Dim SFB_dLocalTimer As Double

If SFB_oAttToCreate Is Nothing Then

    Set SFB_oAttToCreate = New SFB_clsCollection
    SFB_dLocalTimer = Timer
    'The Key is the name of the attribute to be added and the value is the name of the attribute in SFB_clsAttributesList
    Call SFB_oAttToCreate.Add("Title", "Title")
    Call SFB_oAttToCreate.Add("Dataset Type", "Dataset Type")
    Call SFB_oAttToCreate.Add("Design Authority Program", "Design Authority Program")
    Call SFB_oAttToCreate.Add("Supplier Name And CAGE Code", "Supplier Name And CAGE Code")
    Call SFB_oAttToCreate.Add("Major Supplier Code", "Major Supplier Code")
    Call SFB_oAttToCreate.Add("3D Only", "3D Only")
    Call SFB_oAttToCreate.Add("Color Coded", "Color Coded")
    Call SFB_oAttToCreate.Add("PCCN", "PCCN")
    Call SFB_oAttToCreate.Add("Material Specifications", "Material Specifications")
    Call SFB_oAttToCreate.Add("Material Description", "Material Description")
    Call SFB_oAttToCreate.Add("Material Type", "Material Type")
    Call SFB_oAttToCreate.Add("Material Form", "Material Form")
    Call SFB_oAttToCreate.Add("Size", "Size")
    Call SFB_oAttToCreate.Add("Thickness", "Thickness")
    Call SFB_oAttToCreate.Add("Inside Diameter", "Inside Diameter")
    Call SFB_oAttToCreate.Add("Outside Diameter", "Outside Diameter")
    Call SFB_oAttToCreate.Add("Length", "Length")
    Call SFB_oAttToCreate.Add("Width", "Width")
    Call SFB_oAttToCreate.Add("Wall", "Wall")
    Call SFB_oAttToCreate.Add("Alloy", "Alloy")
    Call SFB_oAttToCreate.Add("Final Condition", "Final Condition")
    Call SFB_oAttToCreate.Add("Density", "Density")
    Call SFB_oAttToCreate.Add("Eng. Make From", "Eng. Make From")
    Call SFB_oAttToCreate.Add("Form", "Form")
    Call SFB_oAttToCreate.Add("Grade/Composition", "Grade/Composition")
    Call SFB_oAttToCreate.Add("Material Class", "Material Class")
    Call SFB_oAttToCreate.Add("Mesh Cell Size", "Mesh Cell Size")
    Call SFB_oAttToCreate.Add("Standard Spec Die", "Standard Spec Die")
    Call SFB_oAttToCreate.Add("Type", "Type")
    Call SFB_oAttToCreate.Add("TD Material Code", "TD Material Code")
    Call SFB_oAttToCreate.Add("Finish Code", "Finish Code")
    Call SFB_oAttToCreate.Add("MFG Process", "MFG Process")
    Call SFB_oAttToCreate.Add("Defining Part", "Defining Part")
    Call SFB_oAttToCreate.Add("Material Specification Production", "Material Specification Production")
    Call SFB_oAttToCreate.Add("Material Description Production", "Material Description Production")
    Call SFB_oAttToCreate.Add("Organization", "Revision Organization")
    Call SFB_oAttToCreate.Add("Project", "Revision Project")

End If

For SFB_i = 1 To SFB_oAttToCreate.Count
    
    'Get the parameter name
    SFB_sParamName = SFB_oAttToCreate.GetKey(SFB_i)
    SFB_sParamNameInList = SFB_oAttToCreate.GetItemByIndex(SFB_i)
    
    'Check if parameter exits
    On Error Resume Next
    SFB_sCheckString = ""
    SFB_sCheckString = SFB_oProduct.ReferenceProduct.UserRefProperties.Item(SFB_sParamName).Name
    On Error GoTo 0
    
    'Parameter doesn't exist we need to create it
    If SFB_sCheckString = "" Then
        SFB_sParamValue = SFB_oAttList.GetEnoviaAttributes(SFB_oProduct.PartNumber, SFB_sRev, False, SFB_sParamNameInList)
        Set SFB_oProperty = SFB_oProduct.ReferenceProduct.UserRefProperties.CreateString(SFB_sParamName, SFB_sParamValue)
        
    'Parameter exist, just change the value
    Else
        SFB_sParamValue = SFB_oAttList.GetEnoviaAttributes(SFB_oProduct.PartNumber, SFB_sRev, False, SFB_sParamNameInList)
        Call SFB_oProduct.UserRefProperties.Item(SFB_sParamName).ValuateFromString(SFB_sParamValue)
    End If
Next

End Sub

Private Sub SFB_AddComponentFromFile(ByVal SFB_sParentPN As String, ByVal SFB_sParentRev As String, ByVal sNewPN As String, ByVal sNewRev As String, sNewExtension As String, oConfElem As IXMLDOMElement, Optional SFB_sIteration As String = "")

Dim SFB_oPVRWindow As Window
Dim SFB_oNewDoc As Document
Dim SFB_oParentRefProduct As Product
Dim SFB_oNewProduct As Product
Dim SFB_sRelationID As String
Dim SFB_dTimer As Double
Dim SFB_dLoadTimer As Double
Dim SFB_bTimeout As Boolean
Dim SFB_EnoviaDoc As EnoviaDocument
Dim SFB_EV5product As Product
Dim SFB_sLoadStatus As String

'Initialize
Set SFB_oPVRWindow = CATIA.ActiveWindow
Set SFB_EnoviaDoc = CATIA.Application
SFB_sLoadStatus = "Already Loaded"

 
'Get the parent reference product
Set SFB_oParentRefProduct = SFB_oRefProductList.GetItemByKey(SFB_sParentPN & SFB_sParentRev)

'Check if document already loaded in session.
'If SFB_sIteration is = "" it means we are looking for SFB_a document loaded from Enovia
'If SFB_sIteration <> "" it means we are looking for SFB_a document saved file base for data transfer
Set SFB_oNewDoc = Nothing
On Error Resume Next
If SFB_sIteration = "" Then
    Set SFB_oNewDoc = CATIA.Documents.Item(sNewPN & sNewRev & "." & sNewExtension)
Else
    Set SFB_oNewDoc = CATIA.Documents.Item(sNewPN & " " & sNewRev & "_" & SFB_sIteration & "." & sNewExtension)
End If
On Error GoTo 0


'Check if document is in SFB_oLoadErrDoc
If SFB_oLoadErrDoc.Exists(sNewPN & sNewRev) Then
    SFB_sLoadStatus = "Load Error"
End If

'Open document from ENOVIA
If SFB_oNewDoc Is Nothing And SFB_sLoadStatus <> "Load Error" Then

    If SFB_sIteration = "" Then
        'Set Status
        SFB_sLoadStatus = "New Load"
    
        'Loading document
        SFB_dTimer = Timer
    
        On Error Resume Next
        Err.Clear
        Set SFB_EV5product = SFB_EnoviaDoc.OpenPartDocument(sNewPN, sNewRev)
    
        'Make sure the document is loaded
        SFB_dLoadTimer = Timer
        SFB_bTimeout = False
        Do
            Set SFB_oNewDoc = SFB_EV5product.ReferenceProduct.Parent
    
            If Not SFB_oNewDoc Is Nothing Then Exit Do
    
            DoEvents
            SFB_Sleep 150
            If Abs(Timer - SFB_dLoadTimer) > 5 Then
                SFB_bTimeout = True
                Exit Do
            End If
        Loop
        On Error GoTo 0
    
        'Timeout reached, add document to SFB_oLoadErrDoc
        If SFB_bTimeout = True Then
            Call SFB_oLoadErrDoc.Add(sNewPN & sNewRev, "SFB_AddComponentFromFile")
            Call SFB_SetInstancesStatusKO(SFB_oConfXML, sNewPN, sNewRev)
            SFB_sLoadStatus = "Load Error"
        End If
    
        'Swap windows
        SFB_oPVRWindow.Activate
    Else
        Set SFB_oNewDoc = CATIA.Documents.Read(SFB_sDestinationFolder & sNewPN & " " & sNewRev & "_" & SFB_sIteration & "." & sNewExtension)

    End If
    
End If

'We have SFB_a document, let's add an instance to the PVR structure
If SFB_sLoadStatus = "Already Loaded" Or SFB_sLoadStatus = "New Load" Then

    'Add instance in PVR
    Set SFB_oNewProduct = SFB_oParentRefProduct.Products.AddExternalComponent(SFB_oNewDoc)
    
    'Add to SFB_oRefProductList
    If Not SFB_oRefProductList.Exists(sNewPN & sNewRev) Then
        Call SFB_oRefProductList.Add(sNewPN & sNewRev, SFB_oNewProduct.ReferenceProduct)
    End If
        
    'Get the RelationID of oConfElem
    SFB_sRelationID = oConfElem.Attributes.getNamedItem("RelationID").nodeValue
    
    'Transfer the instance name to all Conf instance with same RelationID
    Call SFB_TransferInstanceName(SFB_sRelationID, SFB_oNewProduct.Name)
    
    'Set status
    Call SFB_SetInstancesStatusOKorMoved(SFB_oConfXML, SFB_sParentPN, SFB_oNewProduct.Name, "OK")
    
    'Move instance
    Call SFB_MoveInstance(SFB_sParentPN, SFB_sParentRev, SFB_oNewProduct.Name, oConfElem)
        

End If

'We need to close SFB_a newly open document
If SFB_sLoadStatus = "New Load" Then

    'Close document
    SFB_oNewDoc.Close
    
End If


End Sub

Private Sub SFB_MoveInstance(ByVal SFB_sParentPN As String, ByVal SFB_sParentRev As String, ByVal sInstanceName As String, ByVal oConfElem As IXMLDOMElement)

Dim SFB_oParentRefProduct As Product
Dim SFB_oChildProduct As Product
Dim SFB_oPosition
Dim SFB_oMatrix(11) As Variant

'Get the Parent Reference Product
Set SFB_oParentRefProduct = SFB_oRefProductList.GetItemByKey(SFB_sParentPN & SFB_sParentRev)

'Get child
Set SFB_oChildProduct = SFB_oParentRefProduct.Products.Item(sInstanceName)

'SFB_Retrieve position matrix
SFB_oMatrix(0) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue)
SFB_oMatrix(1) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue)
SFB_oMatrix(2) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue)
SFB_oMatrix(3) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue)
SFB_oMatrix(4) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue)
SFB_oMatrix(5) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue)
SFB_oMatrix(6) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue)
SFB_oMatrix(7) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue)
SFB_oMatrix(8) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue)
SFB_oMatrix(9) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue)
SFB_oMatrix(10) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue)
SFB_oMatrix(11) = CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue)

'Position instance
Set SFB_oPosition = SFB_oChildProduct.Position
SFB_oPosition.SetComponents SFB_oMatrix

End Sub

Private Function SFB_ComparePosition(ByVal oPVRElem As IXMLDOMElement, ByVal oConfElem As IXMLDOMElement) As String

Dim SFB_sReturnValue As String

'Initial set
SFB_sReturnValue = "Same Position"

'Check all position
If oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position0").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position1").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position2").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position3").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position4").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position5").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position6").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position7").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue <> oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position8").nodeValue Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position9").nodeValue), 6) Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position10").nodeValue), 6) Then
    SFB_sReturnValue = "Position Different"
    Exit Function
ElseIf Round(CDbl(oPVRElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue), 6) <> Round(CDbl(oConfElem.childNodes.Item(0).Attributes.getNamedItem("Position11").nodeValue), 6) Then
    SFB_sReturnValue = "Position Different"
    Exit Function
End If

SFB_ComparePosition = SFB_sReturnValue


End Function
Private Sub SFB_ScanExplodedStructure(ByVal oDMUParent As Product)

Dim SFB_oChild As Product
Dim SFB_sPartNumber As String
Dim SFB_sRevision As String
Dim SFB_oNode As IXMLDOMNode
Dim SFB_sType As String
Dim SFB_oDoc As Document

'Scan all children products
For Each SFB_oChild In oDMUParent.Products

    'SFB_Retrieve Part Number
    On Error Resume Next
    SFB_sPartNumber = ""
    SFB_sPartNumber = SFB_oChild.PartNumber
    If SFB_sPartNumber = "" Then
        SFB_oChild.ApplyWorkMode DESIGN_MODE
    End If
    SFB_sPartNumber = SFB_oChild.PartNumber
    On Error GoTo 0
    
    'Part Number was found
    If SFB_sPartNumber <> "" Then
        
        'Get Type
        SFB_sType = SFB_GetType(SFB_oChild)
        
        'Override Part Number with the one found in the file name
        If SFB_sType = "CATPart" Or SFB_sType = "CATProduct" Then
            
            'Get document
            Set SFB_oDoc = SFB_oChild.ReferenceProduct.Parent
    
            'Override Part Number with the one found in the file name
            If SFB_oDoc.FullName Like "ENOVIA5*" Then
                SFB_sPartNumber = Left(Split(SFB_oDoc.Name, ".")(0), Trim(Len(Split(SFB_oDoc.Name, ".")(0)) - 2))
            Else
                SFB_sPartNumber = Trim(Split(SFB_oDoc.Name, " ")(0))
            End If
        End If
        
        'Get the revision from the XML
        Set SFB_oNode = SFB_oConfXML.selectSingleNode(".//Instance[@PartNumber = '" & SFB_sPartNumber & "']")
        SFB_sRevision = SFB_oNode.Attributes.getNamedItem("DocRev").nodeValue
        
        'Add to collection
        If Not SFB_oRefProductList.Exists(SFB_sPartNumber & SFB_sRevision) Then
            Call SFB_oRefProductList.Add(SFB_sPartNumber & SFB_sRevision, SFB_oChild.ReferenceProduct)
        End If

    'Add to error log
    Else
        Call SFB_oLoadErrDoc.Add(SFB_oLoadErrDoc.Count + 1, "Document associated to " & SFB_oChild & " can't be loaded.")
    End If
    
    'Recursive call on Components
    If SFB_sType = "Component" Then
        Call SFB_ScanExplodedStructure(SFB_oChild)
    End If

Next
End Sub




Private Function SFB_PopulateChildrenFromWebService(ByVal oData As Variant) As SFB_clsCollection

Dim SFB_i, SFB_j As Integer
Dim SFB_sPN As String
Dim SFB_oChildren As SFB_clsCollection
Dim SFB_oItem As SFB_clsCollection

'Initialize
Set SFB_oChildren = New SFB_clsCollection

'Scan oData and transfer value to SFB_oChildren
For SFB_i = 1 To oData.Count
    
    For SFB_j = 2 To oData.Item(SFB_i).Count
    
        'Get Part Number
        SFB_sPN = oData.Item(SFB_i).Item(1)

        Set SFB_oItem = New SFB_clsCollection
        Call SFB_oItem.Add("Part Number", SFB_sPN)
        Call SFB_oItem.Add("Revision", "SFB_N/A")
        Call SFB_oItem.Add("Position Matrix", oData.Item(SFB_i).Item(SFB_j).Item(2))
        
        Call SFB_oChildren.Add("Instance" & SFB_oChildren.Count + 1, SFB_oItem)
    Next
Next

'Return
Set SFB_PopulateChildrenFromWebService = SFB_oChildren

End Function

Private Function SFB_GetExistingNode(ByVal SFB_oDoc As DOMDocument60, ByVal SFB_sPN As String) As IXMLDOMElement

Dim SFB_oNodeList As IXMLDOMNodeList

Set SFB_oNodeList = SFB_oDoc.selectNodes("//Instance[@PartNumber='" & SFB_sPN & "' and @SyncStatus='']")
If SFB_oNodeList.Length >= 1 Then
    Set SFB_GetExistingNode = SFB_oNodeList.Item(0)
Else
    Set SFB_GetExistingNode = Nothing
End If

End Function

Private Function SFB_GetType(ByVal SFB_oProduct As Product) As String

'CATPart
If SFB_oProduct.HasAMasterShapeRepresentation Then
    SFB_GetType = "CATPart"
'Component
ElseIf SFB_oProduct.ReferenceProduct.Parent.Name = SFB_oProduct.Parent.Parent.ReferenceProduct.Parent.Name Then
    SFB_GetType = "Component"
'CATProduct
Else
    SFB_GetType = "CATProduct"
End If
 
End Function

Private Function SFB_Add_Element(ByVal SFB_oDoc As DOMDocument60, ByVal sTagName As String, ByVal oParentElem As IXMLDOMElement, Optional oBrotherElem = Nothing) As IXMLDOMElement

Dim SFB_oElement As IXMLDOMElement

'Add new element
Set SFB_oElement = SFB_oDoc.CreateElement(sTagName)

'Append node
If oBrotherElem Is Nothing Then
    oParentElem.appendChild SFB_oElement
Else
    Call oParentElem.InsertBefore(SFB_oElement, oBrotherElem)
End If

'Return
Set SFB_Add_Element = SFB_oElement

End Function

Private Sub SFB_Add_Attribute(ByVal SFB_oDoc As DOMDocument60, ByRef SFB_oElement As IXMLDOMElement, ByVal sAttributeName As String, ByVal sAttributeValue As String)

Dim SFB_oAttribute As IXMLDOMAttribute

'Add attribute
Set SFB_oAttribute = SFB_oDoc.createAttribute(sAttributeName)
SFB_oAttribute.nodeValue = sAttributeValue
SFB_oElement.setAttributeNode SFB_oAttribute

End Sub

Private Sub SFB_Add_Comment(ByVal SFB_oDoc As DOMDocument60, ByRef SFB_oElement As IXMLDOMElement, ByVal sCommentValue As String)

Dim SFB_oComment As IXMLDOMComment

'Add comment
Set SFB_oComment = SFB_oDoc.createComment(sCommentValue)
SFB_oElement.appendChild SFB_oComment

End Sub
Private Sub SFB_AddToTrackingLog(ByVal SFB_sText As String, ByVal bOverwrite As Boolean, ByVal bSaveNow As Boolean)

'Add text to log text
If SFB_sTrackingLogText = "" Then
    SFB_sTrackingLogText = SFB_sText
Else
    SFB_sTrackingLogText = SFB_sTrackingLogText & vbCrLf & SFB_sText
End If

'Display in debug.print
Debug.Print SFB_sText


'Save log file
If SFB_bTrackingLogNeverSave = False And (bSaveNow = True Or SFB_bTrackingLogSaveEverytime = True) Then
    Call SFB_WriteTextFile(SFB_sTrackingLogFilePath, SFB_sTrackingLogText, bOverwrite)
    SFB_sTrackingLogText = ""
End If

End Sub

Private Sub SFB_GenerateErrorLog()

Dim SFB_sText As String
Dim SFB_i As Integer
Dim SFB_sAnswer As String

If SFB_oSaveErrDoc.Count > 0 Then
    
    SFB_sText = "Following document could not be save:"
    For SFB_i = 1 To SFB_oSaveErrDoc.Count
        SFB_sText = SFB_sText & vbCrLf & "- " & SFB_oSaveErrDoc.GetItemByIndex(SFB_i)
    Next
    
    Call SFB_WriteTextFile(SFB_sErrorLogFilePath, SFB_sText, True)
    
    Call MsgBox("Save errors were found. Refer to " & SFB_sErrorLogFilePath, 48)
End If


End Sub

Private Sub SFB_GenerateLogReport()

Dim SFB_oTopElem As IXMLDOMElement
Dim SFB_oElem As IXMLDOMElement
Dim SFB_oNode As IXMLDOMNode
Dim SFB_oNodeList As IXMLDOMNodeList
Dim SFB_i As Integer
Dim SFB_oPart As SFB_clsPartInfo

'Create new xml structure
Set SFB_oLogReport = New DOMDocument60
SFB_oLogReport.setProperty "SelectionLanguage", "XPath"
SFB_oLogReport.async = True

'Create top node
Set SFB_oTopElem = SFB_oLogReport.CreateElement("Parts")
SFB_oLogReport.appendChild SFB_oTopElem

'Create comment on top node
If SFB_oTopNode.PartIsCI Then
    Call SFB_Add_Comment(SFB_oLogReport, SFB_oTopElem, SFB_oTopNode.PartNumber & " is SFB_a CI")
Else
    Call SFB_Add_Comment(SFB_oLogReport, SFB_oTopElem, SFB_oTopNode.PartNumber & " is not SFB_a CI")
End If
If SFB_oTopNode.SyncFromBSF = True Then
    Call SFB_Add_Comment(SFB_oLogReport, SFB_oTopElem, "PVR updated using best so far structure from ENOVIA")
Else
    Call SFB_Add_Comment(SFB_oLogReport, SFB_oTopElem, "Structure configured for " & SFB_oTopNode.Project & "/" & SFB_oTopNode.Tail)
End If
If SFB_oTopNode.NonCIOption = "BSF" Then
    Call SFB_Add_Comment(SFB_oLogReport, SFB_oTopElem, "User choose to update the non CIs using the Best so Far from Enovia")
Else
    Call SFB_Add_Comment(SFB_oLogReport, SFB_oTopElem, "User choose to update the non CIs using the latest released revision")
End If

Call SFB_Add_Comment(SFB_oLogReport, SFB_oTopElem, "Report generated on " & Format(Now(), "yyyy mm dd") & " at " & Format(Time, "hh:mm:ss"))

'Scan all part in SFB_oPartList
For SFB_i = 1 To SFB_oPartList.Count
    
    'Get the part
    Set SFB_oPart = SFB_oPartList.GetItemByIndex(SFB_i)
    
    'Create new elem
    Set SFB_oElem = SFB_Add_Element(SFB_oLogReport, "Part", SFB_oTopElem)
    
    'Add attributes
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "PartNumber", SFB_oPart.PartNumber)
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "Revision", SFB_oPart.PartRev)
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "Status", SFB_oPart.PartStatus)
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "DocumentOrganization", SFB_oAttList.GetEnoviaAttributes(SFB_oPart.PartNumber, SFB_oPart.PartRev, False, "Document Organization"))
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "Title", SFB_oPart.PartTitle)
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "IsCI", CStr(SFB_oPart.PartIsCI))
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "Type", IIf(SFB_oPart.PartType = "Component", "NONE", SFB_oPart.PartType))
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "EffectiveDwg", IIf(SFB_oPart.EffectiveDwgNb = "SFB_N/A" Or SFB_oPart.EffectiveDwgNb = "", "SFB_N/A", SFB_oPart.EffectiveDwgNb & SFB_oPart.EffectiveDwgRev))
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "ProposedSource", IIf(SFB_oPart.ProposedSourceNb = "", "SFB_N/A", SFB_oPart.ProposedSourceNb & SFB_oPart.ProposedSourceRev))
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "SelectedDwg", IIf(SFB_oPart.SelectedDwgNb = "SFB_N/A", "SFB_N/A", SFB_oPart.SelectedDwgNb & SFB_oPart.SelectedDwgRev))
    Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "SelectedSource", IIf(SFB_oPart.SelectedSourceNb = "", "SFB_N/A", SFB_oPart.SelectedSourceNb & SFB_oPart.SelectedSourceRev))
    
    If SFB_oPart.Comment = "No comment" And (SFB_oPart.ProposedSourceNb & SFB_oPart.ProposedSourceRev) <> (SFB_oPart.SelectedSourceNb & SFB_oPart.SelectedSourceRev) Then
        Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "Comment", "Selected source is different from proposed source")
    Else
       Call SFB_Add_Attribute(SFB_oLogReport, SFB_oElem, "Comment", SFB_oPart.Comment)
    End If
Next


'Save report
SFB_oLogReport.Save SFB_sXMLPath & "PVRSync_ConfiguredStructure_Report_" & SFB_oTopNode.PartNumber & ".xml"


End Sub

Private Function SFB_GetDocRevOrderedList(ByVal sDocNumber As String) As SFB_clsCollection

Dim SFB_oDocs As SFB_clsCollection
Dim SFB_oTemp
Dim SFB_oOrderedList As New SFB_clsCollection
Dim SFB_oOrderedColl As New Collection
Dim SFB_i, SFB_j As Integer

'SFB_Retrieve all documents with same base number with web service
Set SFB_oTemp = SFB_WebServiceAccessTool.GetDocumentByBaseNumber(sDocNumber)
Set SFB_oDocs = New SFB_clsCollection
Call SFB_oDocs.InitializeWithDLLclsColObject(SFB_oTemp, SFB_oDocs)

'Error management
If SFB_oDocs.Count = 0 Then
    Set SFB_GetDocRevOrderedList = Nothing
    Exit Function
End If

'Remove from the list all document that don't have the same Part Number
For SFB_i = SFB_oDocs.Count To 1 Step -1
    If SFB_oDocs.GetItemByIndex(SFB_i).GetItemByKey("FIELD_PART_NUMBER") <> sDocNumber Then SFB_oDocs.RemoveByIndex SFB_i
Next

'Add all revision to SFB_oOrderedColl
For SFB_i = 1 To SFB_oDocs.Count
    
    'First item
    If SFB_oOrderedColl.Count = 0 Then
        SFB_oOrderedColl.Add (SFB_oDocs.GetItemByIndex(SFB_i).GetItemByKey("FIELD_DOCUMENT_REVISION"))
    
    'Other object
    Else
        For SFB_j = 1 To SFB_oOrderedColl.Count
            If SFB_oDocs.GetItemByIndex(SFB_i).GetItemByKey("FIELD_DOCUMENT_REVISION") < SFB_oOrderedColl.Item(SFB_j) Then
                SFB_oOrderedColl.Add SFB_oDocs.GetItemByIndex(SFB_i).GetItemByKey("FIELD_DOCUMENT_REVISION"), , SFB_j
                Exit For
            End If
        Next
        If SFB_oOrderedColl.Count < SFB_i Then SFB_oOrderedColl.Add SFB_oDocs.GetItemByIndex(SFB_i).GetItemByKey("FIELD_DOCUMENT_REVISION")
    End If
Next

'Transfer object from SFB_oOrderedColl to SFB_oOrderedList
For SFB_i = 1 To SFB_oOrderedColl.Count
    Call SFB_oOrderedList.Add(SFB_oOrderedColl.Item(SFB_i), SFB_oOrderedColl.Item(SFB_i))
Next

'Return
Set SFB_GetDocRevOrderedList = SFB_oOrderedList

End Function

