Attribute VB_Name = "IDLExtract"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Const iXLHeaderRow = 5
Const sMacroTitle = "CATIA IDL"

Private Const xlNone = -4142
Private Const xlLeft = -4131
Private Const xlCenter = -4108
Private Const xlBottom = -4107
Private Const xlContext = -5002
Private Const xlContinuous = 1
Private Const xlAutomatic = -4105
Private Const xlThin = 2
Private Const xlHairline = 1

Private Const xlDiagonalDown = 5
Private Const xlDiagonalUp = 6
Private Const xlEdgeLeft = 7
Private Const xlEdgeTop = 8
Private Const xlEdgeBottom = 9
Private Const xlEdgeRight = 10
Private Const xlInsideVertical = 11
Private Const xlInsideHorizontal = 12
Private Const xlCalculationAutomatic = -4105
Private Const xlCalculationManual = -4135

Private Const xlValidateList = 3
Private Const xlValidAlertStop = 1
Private Const xlBetween = 1

Private sExtractBDIData As String 'User selection
Private sExtractSP As String 'User selection
Private sExtractCG As String 'User selection

Private oStruct As DOMDocument60 'XML structure for the BOM
Private oPartAttributes As clsCollection 'clsCollection for the attributes of each reference
Private oColumnsList As clsCollection 'clsCollection of all the columns displayed in the Complete Structure sheet
Private oColorCodeDashNbColl As clsCollection
Private oColorCodePartExclusionList As clsCollection
Private oPVRSyncReport As clsXML
Private iCntProdMax As Integer


Private oExcel As clsExcel

Private Function FormatInstanceName(ByVal sName As String, ByVal iLevel As Integer) As String

Dim j As Integer
Dim sNamePrefix As String

sNamePrefix = ""
For j = 1 To iLevel
    sNamePrefix = sNamePrefix + Space(2) + Chr(42) + Space(2)
Next

'Return
FormatInstanceName = sNamePrefix + sName

End Function
Private Sub CleanStructureForIntegrationReport(ByRef bError As Boolean)

Dim sDatasteType As String
Dim sMaterialSpec As String
Dim sPartNumber As String
Dim oItem As clsCollection
Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim i As Integer

'Initialize
bError = False

'Delete all the node where NHAQty is "N/A"
Do
    Set oNodeList = Nothing
    Set oNodeList = oStruct.SelectNodes("//Instance[@NHAQty='N/A']")

    If oNodeList.Length = 0 Then Exit Do
    
    Set oNode = oNodeList.Item(0)
    Call DeleteNode(oNode, False)
Loop

'Scan the list of all the reference and look for:
' - SP Collector
' - Skeleton when Material Spec is ""
' - COMPLETION CATALOG (REFERENCE)
' - STANDARD (REFERENCE DATA)
For i = 1 To oPartAttributes.Count

    Set oItem = oPartAttributes.GetItemByIndex(i)
    
    sPartNumber = oPartAttributes.GetKey(i)
    sDatasteType = oItem.GetItemByKey("DatasetType").GetItemByKey("Value")
    sMaterialSpec = oItem.GetItemByKey("MatSpec").GetItemByKey("Value")

    'Tag node to be deleted
    If (sDatasteType = "COLLECTOR") Or _
       (sDatasteType = "SKELETON" And sMaterialSpec = "") Or _
       (sDatasteType = "COMPLETION CATALOG (REFERENCE)") Or _
       (sDatasteType = "STANDARD (REFERENCE DATA)") Then
       Set oNodeList = oStruct.SelectNodes("//Instance[@PartNumber='" & sPartNumber & "']")
       For Each oNode In oNodeList
            Call Add_Attribute(oStruct, oNode, "ToDelete", "Yes")
       Next
    End If
    
Next

'Check if Top node is flagged to be deleted
Set oNode = oStruct.SelectSingleNode("./Instance")
If Not oNode.Attributes.getNamedItem("ToDelete") Is Nothing Then
    bError = True
    Exit Sub
End If


'Delete all the node where ToDelete is "Yes" but keep the children
Do
    Set oNodeList = Nothing
    Set oNodeList = oStruct.SelectNodes("//Instance[@ToDelete='Yes']")

    If oNodeList.Length = 0 Then Exit Do
    
    Set oNode = oNodeList.Item(0)
    Call DeleteNode(oNode, True)
Loop

'Since we deleted the SP node e need to edit the Level attribute in oStruct. We skip the first node.
Set oNodeList = oStruct.SelectNodes("//Instance")
For i = 2 To oNodeList.Length
    Set oNode = oNodeList.Item(i - 1)
    oNode.Attributes.getNamedItem("Level").nodeValue = CInt(oNode.parentNode.Attributes.getNamedItem("Level").nodeValue) + 1
Next


End Sub

Private Sub DeleteNode(ByVal oNode As IXMLDOMNode, Optional bKeepChildren = False)

Dim oChildList As IXMLDOMNodeList
Dim oChild As IXMLDOMNode
Dim oClone As IXMLDOMNode

'Don't keep children
If bKeepChildren = False Then

    Call oNode.parentNode.RemoveChild(oNode)
    
'Keep children
Else
    
    'Get all first level children
    Set oChildList = oNode.SelectNodes("./*")
    
    'Clone all the children, if we have any and insert them before the node to be deleted
    If oChildList.Length >= 1 Then
        For Each oChild In oChildList
            Set oClone = oChild.cloneNode(True)
            Call oNode.parentNode.InsertBefore(oClone, oNode)
        Next
    End If
    
    'Delele the node
    Call oNode.parentNode.RemoveChild(oNode)

End If

End Sub

Private Sub GetRefPartNumber(ByVal sPartNumber As String)

Dim oItem As clsCollection
Dim sMatSpec As String
Dim sDatasetType As String
Dim sNomenclature As String
Dim sDefiningPart As String

'Initialize
Set oItem = oPartAttributes.GetItemByKey(sPartNumber)
sMatSpec = oItem.GetItemByKey("MatSpec").GetItemByKey("Value")
sDatasetType = oItem.GetItemByKey("DatasetType").GetItemByKey("Value")
sNomenclature = oItem.GetItemByKey("Nomenclature").GetItemByKey("Value")
sDefiningPart = oItem.GetItemByKey("DefiningPart").GetItemByKey("Value")

'Set the reference part number
If sDatasetType = "SKELETON" And sMatSpec <> "" Then
    Call oItem.GetItemByKey("RefPartNumber").SetItemByKey("Value", sMatSpec)

ElseIf sDatasetType = "FLEXIBLE REPRESENTATION" Or sDatasetType = "CATALOG LIGHT VERSION" Then
    Call oItem.GetItemByKey("RefPartNumber").SetItemByKey("Value", sDefiningPart)

ElseIf sNomenclature <> "" Then
    Call oItem.GetItemByKey("RefPartNumber").SetItemByKey("Value", sNomenclature)
Else
    Call oItem.GetItemByKey("RefPartNumber").SetItemByKey("Value", sPartNumber)
End If

End Sub
'#######################################################################################
'This is the list of all the part along with all the attributes
'Note that some attributes are coming from:
' - Enovia attribute web service
' - BDI web service
' - Manually entered. Ex. Base Number, Dash Number ....
'#######################################################################################
Private Function RetrievePartAttributes(ByVal sKey As String) As clsCollection

Dim oItem As clsCollection
Dim oAttList As clsCollection

'Item already exist
If oPartAttributes.Exists(sKey) Then

    Set RetrievePartAttributes = oPartAttributes.GetItemByKey(sKey)

'Item does not exist, we need to add it
Else
    
    Set oAttList = New clsCollection
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("ColorCode", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Defining Part")
    Call oItem.Add("Value", "")
    Call oAttList.Add("DefiningPart", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Nomenclature")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Nomenclature", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Title")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Title", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "BA Document Revision")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Revision", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("CIStatus", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "DOCUMENT_ITERATION")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Iteration", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Revision Status")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Status", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Dataset Type")
    Call oItem.Add("Value", "")
    Call oAttList.Add("DatasetType", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("DatasetTypeDefPart", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Supplier Name And CAGE Code")
    Call oItem.Add("Value", "")
    Call oAttList.Add("SupplierName", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Material Specification Production")
    Call oItem.Add("Value", "")
    Call oAttList.Add("MatSpecProd", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Material Description Production")
    Call oItem.Add("Value", "")
    Call oAttList.Add("MatDescProd", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Alloy")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Alloy", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Final Condition")
    Call oItem.Add("Value", "")
    Call oAttList.Add("FinalCondition", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Material Form")
    Call oItem.Add("Value", "")
    Call oAttList.Add("MaterialForm", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Finish Code")
    Call oItem.Add("Value", "")
    Call oAttList.Add("FinishCode", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Size")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Size", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "MFG Process")
    Call oItem.Add("Value", "")
    Call oAttList.Add("MFGProcess", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "PCCN")
    Call oItem.Add("Value", "")
    Call oAttList.Add("PCCN", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "I/R/C/S")
    Call oItem.Add("Value", "")
    Call oAttList.Add("I/R/C/S", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Serialized")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Serialized", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Thickness")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Thickness", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Eng. Make From")
    Call oItem.Add("Value", "")
    Call oAttList.Add("EngMakeFrom", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Grade/Composition")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Grade/Composition", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Material Class")
    Call oItem.Add("Value", "")
    Call oAttList.Add("MaterialClass", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Type")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Type", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Length")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Length", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Width")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Width", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Document Organization")
    Call oItem.Add("Value", "")
    Call oAttList.Add("DocOrg", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Revision Organization")
    Call oItem.Add("Value", "")
    Call oAttList.Add("RevOrg", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "Material Specifications")
    Call oItem.Add("Value", "")
    Call oAttList.Add("MatSpec", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("RVNumber", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("E-Item", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("FirstAC", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("MfgSupp", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("EDRN", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("BaseNumber", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("StringDashNumber", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("NumericDashNumber", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Suffix", oItem)

    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("RefPartNumber", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("Error", oItem)
    
    Set oItem = New clsCollection
    Call oItem.Add("EnoviaName", "")
    Call oItem.Add("Value", "")
    Call oAttList.Add("ReportIn", oItem)
        
    Call oPartAttributes.Add(sKey, oAttList)
    
    Set RetrievePartAttributes = oAttList

End If
End Function

'##################################################################################################
'This is the list of all the columns that will be displayed in the complete structure sheet
'##################################################################################################
Private Sub PopulateColumnsList()

'Initialize
Set oColumnsList = New clsCollection

'Define all the columns along with the source to retrieve the information
' - Instance : Info will be stored in the XML structure
' - Part :  Info will be stored in oPartAttributes

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "ITEM #")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("Item", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "INSTANCE NAME")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("InstanceName", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "LEVEL")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("Level", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "PART NUMBER")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("PartNumber", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "NHA")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("NHA", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "FML MATERIAL" & Chr(10) & "FIND NUMBER")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("ColorCode", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "DEFINING PART")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("DefiningPart", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "NOMENCLATURE")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Nomenclature", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "TITLE")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Title", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "REVISION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Revision", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "BDI CI STATUS")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("CIStatus", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "ITERATION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Iteration", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "STATUS")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Status", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "DATASET TYPE")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("DatasetType", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "DATASET TYPE OF" & Chr(10) & "DEFINING PART")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("DatasetTypeDefPart", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "SUPPLIER NAME" & Chr(10) & "AND CAGE CODE")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("SupplierName", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "MATERIAL SPECIFICATION" & Chr(10) & "PRODUCTION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("MatSpecProd", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "MATERIAL DESCRIPTION" & Chr(10) & "PRODUCTION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("MatDescProd", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "ALLOY")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Alloy", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "FINAL CONDITION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("FinalCondition", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "MATERIAL" & Chr(10) & "FORM")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("MaterialForm", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "FINISH" & Chr(10) & "CODE")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("FinishCode", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "SIZE")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Size", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "MFG" & Chr(10) & "PROCESS")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("MFGProcess", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "PCCN")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("PCCN", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "I/R/C/S")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("I/R/C/S", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "SERIALIZED")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Serialized", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "THICKNESS")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Thickness", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "ENG. MAKE FROM")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("EngMakeFrom", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "GRADE/COMPOSITION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Grade/Composition", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "MATERIAL CLASS")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("MaterialClass", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "TYPE")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Type", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "LENGTH")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Length", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "WIDTH")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("Width", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "DOCUMENT" & Chr(10) & "ORGANIZATION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("DocOrg", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "REVISION" & Chr(10) & "ORGANIZATION")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("RevOrg", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "MATERIAL SPECIFICATIONS")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("MatSpec", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "RV NUMBER")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("RVNumber", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "E-ITEM")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("E-Item", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "FIRST A/C")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("FirstAC", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "MANUFACTURING SUPPLIER")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("MfgSupp", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "EDRN NUMBER")
Call oItem.Add("Source", "Part")
Call oColumnsList.Add("EDRN", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "CATIA CGX (IN)")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("CGX", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "CATIA CGY (IN)")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("CGY", oItem)

Set oItem = New clsCollection
Call oItem.Add("DisplayName", "CATIA CGZ (IN)")
Call oItem.Add("Source", "Instance")
Call oColumnsList.Add("CGZ", oItem)

End Sub

Public Sub oMain()
   
Dim sMsg As String, sAnswer As String

'Ask question
sMsg = "Do want want to extract a new IDL ?"
sMsg = sMsg & vbCrLf & " - Click Yes to extract a new IDL"
sMsg = sMsg & vbCrLf & " - Click No to connect to an existing IDL"
sMsg = sMsg & vbCrLf & " - Click Cancel to exit the process"
sAnswer = MsgBox(sMsg, 67, "IDL Extract")

'Create a new extract
If sAnswer = vbYes Then
    Call ExtractNewIDL

'Connect to an existing extract
ElseIf sAnswer = vbNo Then

    'Get Excel
    Set oExcel = New clsExcel
    Call oExcel.GetNewExcel
    
    'Ask user to select the IDL Extract file
    Call oExcel.OpenExcelFile(, "Select the IDL extract to connect to.")
    
    'Define message
    sMsg = "No 'COMPLETE STRUCTURE' sheet were found in the selected workbook. Process aborted."
    
    'No COMPLETE STRUCTURE" sheet found
    If oExcel.WorksheetExists("COMPLETE STRUCTURE") = False Then
        Call MsgBox(sMsg, vbCritical, "Import Color Coding Table")
        Call oExcel.CloseExcelFile
        Set oExcel = Nothing
        GoTo ExitTheSub
    End If
    
    'Show Excel file
    Call oExcel.ShowExcelWindow
    
    frmIDLExtract.Show modeless
Else
    GoTo ExitTheSub
End If

Exit Sub
ExitTheSub:
Call frmKBEMain.resetToolbar

End Sub


Public Sub HightlightInExcel()

Dim oIDLTopInstance As Product
Dim sIDLTopInstancePartNumber As String
Dim oCatiaAddresses As clsCollection
Dim oIDLAddresses As clsCollection
Dim oSelectedInstances As Collection
Dim sError As String
Dim iRow As Integer
Dim sInstanceName As String
Dim sParentIsntanceName As String
Dim sString As String
Dim sSplitString As Variant
Dim iMinRow As Integer


'Get the instance part number at the top of the IDL, (row 6, column 4) in Excel file
sIDLTopInstancePartNumber = Trim(Replace(oExcel.App.Cells(6, 4).Value, "*", ""))

'For each selected instances, build the parent list and transfer it to oCatiaAddresses
Set oSelection = CATIA.ActiveDocument.Selection
Set oCatiaAddresses = New clsCollection
If oSelection.Count > 0 Then
    For i = 1 To oSelection.Count
        Call oCatiaAddresses.Add(i, BuildAddress(oSelection.Item(i).Value, sIDLTopInstancePartNumber))
    Next
Else
    Call MsgBox("No instances were selected in CATIA.", vbCritical)
    Exit Sub
End If

'Transfer CATIA selection in oSelectedInstances
Set oSelectedInstances = New Collection
If oSelection.Count > 0 Then
    For i = 1 To oSelection.Count
        Call oSelectedInstances.Add(oSelection.Item(i).Value)
    Next
End If


'In CATIA retrieve instance with name = sIDLTopInstancePartNumber
oSelection.Clear
oSelection.Search "CATProductSearch.Product.PartNumber='" & sIDLTopInstancePartNumber & "',all"

'IDL top instance can't be found
If oSelection.Count = 0 Then
    Call MsgBox("IDL top instance " & sIDLTopInstancePartNumber & " can't be found in the active window", vbCritical)
    Exit Sub

'Found at least one instance with name sIDLTopInstancePartNumber
Else
    Set oIDLTopInstance = oSelection.Item(1).Value
End If

'Reselect instances selected by user
oSelection.Clear
For i = 1 To oSelectedInstances.Count
    oSelection.Add oSelectedInstances(i)
Next

'Make sure that sIDLTopInstancePartNumber is found in each address of oCatiaAddresses
'if so trim everything left of sIDLTopInstancePartNumber, otherwise delete from oCatiaAddresses
sError = ""
For i = oCatiaAddresses.Count To 1 Step -1
    
    'Selected instance in CATIA is the top instance in the IDL
    If oCatiaAddresses.GetItemByIndex(i) = sIDLTopInstancePartNumber Then
        'Do Nothing
    
    'Address of the selected instance in CATIA contains sIDLTopInstancePartNumber
    ElseIf oCatiaAddresses.GetItemByIndex(i) Like "*" & sIDLTopInstancePartNumber & "*" Then
        Call oCatiaAddresses.SetItemByIndex(i, sIDLTopInstancePartNumber & "||" & Split(oCatiaAddresses.GetItemByIndex(i), sIDLTopInstancePartNumber & "||")(1))
    
    'Address of the selected instance in CATIA doesn't contains sIDLTopInstancePartNumber
    Else
        sError = sError & " - " & Trim(Replace(oCatiaAddresses.GetItemByIndex(i), "||", "/")) & vbCrLf
        Call oCatiaAddresses.RemoveByIndex(i)
    End If
Next

'Get addresses from all rows in IDL
Set oIDLAddresses = New clsCollection
iRow = 6
Do While oExcel.App.Cells(iRow, 2).Value <> ""
    
    'First item in oIDLAddresses
    If iRow = 6 Then
        
        'Get instance name
        sInstanceName = Trim(Replace(oExcel.App.Cells(iRow, 4).Value, "*", ""))
        
        'Add to oIDLAddresses
        Call oIDLAddresses.Add(sInstanceName, iRow)
        
    'Other items in oIDLAddresses
    Else
    
        'Get instance name
        sInstanceName = Trim(Replace(oExcel.App.Cells(iRow, 2).Value, "*", ""))
        
        
        'Get address of previous item in oIDLAddresses
        sString = oIDLAddresses.GetKey(oIDLAddresses.Count)
        
        'Only keep portion of the previous address
        sSplitString = Split(sString, "||")
        ReDim Preserve sSplitString(oExcel.App.Cells(iRow, 3).Value - 2)
        sString = Join(sSplitString, "||")
        
        'Add to oIDLAddresses
        If Not oIDLAddresses.Exists(sString & "||" & sInstanceName) Then
            Call oIDLAddresses.Add(sString & "||" & sInstanceName, iRow)
        End If
    End If

    iRow = iRow + 1
Loop


'Reset color in column A
With oExcel.App.Columns("A:A").Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

'Loop oCatiaAddresses and check if they can be found in oIDLAddresses
iMinRow = 1000
For i = 1 To oCatiaAddresses.Count

    'Change color in column 1
    If oIDLAddresses.Exists(oCatiaAddresses.GetItemByIndex(i)) Then
        With oExcel.App.Cells(oIDLAddresses.GetItemByKey(oCatiaAddresses.GetItemByIndex(i)), 1).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With

        'Find smallest highlighted row in Excel
        If oIDLAddresses.GetItemByKey(oCatiaAddresses.GetItemByIndex(i)) < iMinRow Then
            iMinRow = oIDLAddresses.GetItemByKey(oCatiaAddresses.GetItemByIndex(i))
        End If
    'Instance not found, report error
    Else
        sError = sError & " - " & Trim(Replace(oCatiaAddresses.GetItemByIndex(i), "||", "/")) & vbCrLf
    End If
Next

'Scroll in Excel
If iMinRow <> 1000 Then
    oExcel.App.ActiveWindow.ScrollRow = iMinRow
End If

'Display error message
If sError <> "" Then
    Call MsgBox("The following instance were not found in the IDL:" & vbCrLf & sError, vbExclamation)
End If

End Sub

Private Function BuildAddress(ByVal iObject As Variant, ByVal sIDLTopInstancePartNumber As String) As String
    
    'The parent list is a list of all the instances starting at the selection all the way up to the top product in the active window
    'or an instance with part number = sIDLTopInstancePartNumber

    
    'Default Mode
    iObject.ApplyWorkMode DEFAULT_MODE
    
    'iObject.PartNumber = sIDLTopInstancePartNumber
    If iObject.PartNumber = sIDLTopInstancePartNumber Then
        BuildAddress = iObject.PartNumber
        Exit Function
    Else
        BuildAddress = iObject.Name
    End If
    
    'Loop
    Do
        Set iObject = iObject.Parent
        
        'We don't care about the "Products" or "ProductDocuments" objects
        If TypeName(iObject) = "Product" Then
            If Not GetType(iObject) = "Component" Then
                
                'We reach an instance where part number = sIDLTopInstancePartNumber
                If iObject.PartNumber = sIDLTopInstancePartNumber Then
                    BuildAddress = iObject.PartNumber & "||" & BuildAddress
                    Exit Function
                Else
                    BuildAddress = iObject.Name & "||" & BuildAddress
                End If
            End If
        End If
    Loop Until TypeName(iObject) = "Application"
    
End Function


Public Sub HightlightInCATIA()

Dim oSelectedRows As clsCollection
Dim sAddress As String
Dim oIDLTopInstance As Product
Dim sIDLTopInstancePartNumber As String
Dim oInstance As Product
Dim sError As String

'Initial set
Set oSelectedRows = New clsCollection

'Get Excel selection
Set oSelectedRange = oExcel.App.Selection

'Find the unique rows from Excel selection and add an item in oSelectedRows
For Each oRng In oSelectedRange.Areas
    For i = oRng.Row To oRng.Row + oRng.Rows.Count - 1
        If oExcel.App.Cells(i, 1).EntireRow.Hidden = False Then
            If Not oSelectedRows.Exists(i) Then
                Call oSelectedRows.Add(i, "")
            End If
        End If
    Next
Next

'Scan all the row from oSelectedRows and get the full address of the instance.
For i = oSelectedRows.Count To 1 Step -1
    
    'Get address
    sAddress = ConcatenateAddress(oSelectedRows.GetKey(i))
    
    'Add address to collection if <> ""
    If sAddress <> "" Then
        Call oSelectedRows.SetItemByIndex(i, sAddress)
    Else
        Call oSelectedRows.RemoveByIndex(i)
    End If
Next

'Check if we have at least one row with an address
If oSelectedRows.Count = 0 Then
    Call MsgBox("None of the selected row in Excel have an instance name.", vbCritical)
    Exit Sub
End If

'Get the instance part number at the top of the IDL, (row 6, column 4) in Excel file
sIDLTopInstancePartNumber = Trim(Replace(oExcel.App.Cells(6, 4).Value, "*", ""))

'In CATIA retrieve instance with part number = sIDLTopInstancePartNumber
Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear
oSelection.Search "CATProductSearch.Product.PartNumber='" & sIDLTopInstancePartNumber & "',all"

'IDL top instance can't be found
If oSelection.Count = 0 Then
    Call MsgBox("IDL top instance " & sIDLTopInstancePartNumber & " can't be found in the active window", vbCritical)
    Exit Sub

'Found at least one instance with name sIDLTopInstancePartNumber
Else
    Set oIDLTopInstance = oSelection.Item(1).Value
End If


'Scan oSelectedRows and select all instances
oSelection.Clear
sError = ""
For i = 1 To oSelectedRows.Count

    Set oInstance = Nothing
    Set oInstance = GetInstanceByAddress(oIDLTopInstance, oSelectedRows.GetItemByIndex(i))

    If Not oInstance Is Nothing Then
        oSelection.Add oInstance
    Else
        sError = sError & " - " & Trim(Replace(oSelectedRows.GetItemByIndex(i), "||", "/")) & vbCrLf
    End If
Next

'Reset color in column A
With oExcel.App.Columns("A:A").Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

'Reframe on and Center graph
CATIA.Application.StartCommand ("Reframe on")
CATIA.Application.StartCommand ("Center graph")

'Display error message
If sError <> "" Then
    Call MsgBox("The following instance were not found in CATIA:" & vbCrLf & sError, vbExclamation)
End If

End Sub

Private Function GetInstanceByAddress(ByVal oIDLTopInstance As Product, ByVal sAddress As String) As Product

Dim oParentInstance As Product
Dim oChildInstance As Product
Dim oComponent As Product
Dim sChildInstanceName As String

'oIDLTopInstance
Set oParentInstance = oIDLTopInstance
If sAddress Like "*||*" Then
    sAddress = Split(sAddress, "||", 2)(1)
Else
    Set GetInstanceByAddress = oParentInstance
    Exit Function
End If

Do
    'Get instance name
    If sAddress Like "*||*" Then
        sChildInstanceName = Split(sAddress, "||")(0)
    Else
        sChildInstanceName = sAddress
    End If
    
    'Select child instance
    On Error Resume Next
    Set oChildInstance = Nothing
    Set oChildInstance = oParentInstance.Products.Item(sChildInstanceName)
    If oChildInstance Is Nothing Then
    
        'Scan all components under oParentInstance and try to find an instance with sChildInstanceName
        For Each oComponent In oParentInstance.Products
        
            'Component found
            If GetType(oComponent) = "Component" Then
                Set oChildInstance = oComponent.Products.Item(sChildInstanceName)
                If Not oChildInstance Is Nothing Then Exit For
            End If
        Next
    
        'Child instance was not found
        If oChildInstance Is Nothing Then
            Set GetInstanceByAddress = oChildInstance
            Exit Function
        End If
        
    End If
    On Error GoTo 0
    
    'Check if we've reach the last instance in the address
    If Not sAddress Like "*||*" Then Exit Do
    
    'Remove active instance from address name
    sAddress = Split(sAddress, "||", 2)(1)
    
    'Child instance becomes the parent instance
    Set oParentInstance = oChildInstance
Loop

Set GetInstanceByAddress = oChildInstance

End Function
Private Function ConcatenateAddress(ByVal iRow As Integer) As String

Dim sInstanceName As String
Dim i As Integer
Dim iLevel As Integer
Dim sFullAddress As String

'Get instance name and level on iRow
If iRow <> 6 Then
    sInstanceName = Trim(Replace(oExcel.App.Cells(iRow, 2).Value, "*", ""))
    iLevel = oExcel.App.Cells(iRow, 3).Value
    sFullAddress = sInstanceName

'On top node used part number instead of instance name
Else
    sInstanceName = Trim(Replace(oExcel.App.Cells(iRow, 4).Value, "*", ""))
    iLevel = oExcel.App.Cells(iRow, 3).Value
    sFullAddress = sInstanceName
End If

'Check sInstanceName and iLevel
If sInstanceName = "" Or iLevel < 1 Then
    ConcatenateAddress = ""
    Exit Function
End If

'Scan all row on report going up
For i = iRow - 1 To 6 Step -1
    
    'If level on active row = iLevel - 1 then we found the parent
    If i = 6 Then
        sInstanceName = Trim(Replace(oExcel.App.Cells(i, 4).Value, "*", ""))
        sFullAddress = sInstanceName & "||" & sFullAddress
    ElseIf oExcel.App.Cells(i, 3).Value = iLevel - 1 Then
        sInstanceName = Trim(Replace(oExcel.App.Cells(i, 2).Value, "*", ""))
        sFullAddress = sInstanceName & "||" & sFullAddress
        iLevel = iLevel - 1
    End If
Next

ConcatenateAddress = sFullAddress
End Function

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Public Sub ExtractNewIDL()

    Dim oSel
    Dim oProduct As Product
    Dim sStartProcessTime
    Dim oItem As clsCollection
    Dim sString As String
    Dim oElem As IXMLDOMElement
    Dim bError As Boolean
    Dim sAnswer As String
    Dim sXMLPath As String

    'Initialize collection
    Set oColorCodeDashNbColl = New clsCollection
    Set oColorCodePartExclusionList = New clsCollection
    Set oPartAttributes = New clsCollection
    Call PopulateColumnsList
    
    'Initialize oStruct and create root element
    Set oStruct = Nothing
    Set oStruct = New DOMDocument60
    oStruct.setProperty "SelectionLanguage", "XPath"
    oStruct.async = True

    'Ask user to select the assembly to be analyzed
    Set oSel = CATIA.ActiveDocument.Selection
    oSel.Clear
    
    If MsgBox("Please select the Top Assembly to Analyze", 65, sMacroTitle) <> vbOK Then
        Call frmKBEMain.resetToolbar
        Exit Sub
    End If
    
    Dim InputObjectType(0), Status
    InputObjectType(0) = "Product"
    Status = oSel.SelectElement2(InputObjectType, "Select Product to analyze", False)
    
    'Get the selected product
    Set oProduct = oSel.Item(1).Value
    
    'Active document is a PVRREF that was just sync with PVRSync. User wants to use the revision defined in the xml report
    If CATIA.ActiveDocument.Name Like "*PVRREF*" Then
        sXMLPath = Environ("temp") & "\PVRSync_ConfiguredStructure_Report_" & oProduct.PartNumber & ".xml"
        If CheckFileExist(sXMLPath) Then
            sAnswer = MsgBox("Following file was found:" & vbCrLf & "Path = " & sXMLPath & vbCrLf & "Last Modified Date = " & GetFileDateCreated(sXMLPath) & _
                             vbCrLf & "Do you want to retrieve the document revision from it?", 36, "IDL Extract")
            If sAnswer = vbYes Then
                Set oPVRSyncReport = New clsXML
                Call oPVRSyncReport.Load(Environ("temp") & "\PVRSync_ConfiguredStructure_Report_" & oProduct.PartNumber & ".xml")
            End If
        End If
    End If

    'Active Document is a PVR and user didn't want to use the revisions from the xml report
    If oPVRSyncReport Is Nothing And (CATIA.ActiveDocument.Name Like "*PVRREF*" Or CATIA.ActiveDocument.Name Like "*PVRPMY*") Then
        Set oItem = RetrievePartAttributes("General")
        Call oItem.GetItemByKey("Error").SetItemByKey("Value", "The IDL was extracted from a PVR document. Please note that the tool only extract attributes and BDI info of the latest revision of each document.")
    End If

    'Extract RV Number, E-Item and BDI CI Status
    sExtractBDIData = MsgBox("Do you want to extract BDI CI Status, RV number, E-Item, Manufacturing Supplier and EDRN Number ?" & vbCrLf & _
                     "Note that it may take 30% longer to extract IDL if answer is Yes.", 35)
    If sExtractBDIData = vbCancel Then
        Call frmKBEMain.resetToolbar
        Exit Sub
    End If
                     
    'Extract SP Collector content
    sExtractSP = MsgBox("Do you want to extract the content of the SP Collector ?", 35)
    If sExtractSP = vbCancel Then
        Call frmKBEMain.resetToolbar
        Exit Sub
    End If
    
    'Extract CoG of the detail parts
    sExtractCG = MsgBox("Do you want to extract the CoG of detail parts ?", 35)
    If sExtractCG = vbCancel Then
        Call frmKBEMain.resetToolbar
        Exit Sub
    End If
    
    'Select all part and product in order to count them
    If oSel.Count = 0 Then Exit Sub
    Call frmProgress.progressBarInitialize("Extract IDL")
    Call frmProgress.progressBarRepaint("Counting number of instances.")
    Set oProduct = oSel.Item(1).Value
    oSel.Search "('Product Structure'.Product + 'Product Structure'.Part),sel"
    iCntProdMax = oSel.Count
    oSel.Clear

    'Remove the content of all SP
    If sExtractSP = vbNo Then
        oSel.Add oProduct
        oSel.Search "CATProductSearch.Product.PartNumber_CAP=*SP*,sel"
        iCntProdMax = iCntProdMax + oSel.Count
        oSel.Search "('Product Structure'.Product + 'Product Structure'.Part),sel"
        iCntProdMax = iCntProdMax - oSel.Count
        oSel.Clear
    
    'Remove all component under SP
    Else
        oSel.Add oProduct
        oSel.Search "CATProductSearch.Product.PartNumber_CAP=*SP*,sel"
        iCntProdMax = iCntProdMax + oSel.Count
        oSel.Search "'Product Structure'.Assembly,sel"
        iCntProdMax = iCntProdMax - oSel.Count
        oSel.Clear
        
    End If

    'Put assembly in default mode
    If bCancelAction = True Then Exit Sub
    sStartProcessTime = Now
    Call frmProgress.progressBarRepaint("Loading data to access attributes... please wait.")
    oProduct.ApplyWorkMode DEFAULT_MODE
    
    'Clear cache
    Call StartWebServiceTool
    WebServiceAccessTool.ClearCache
    
    'Add the selected product to the structure
    Set oElem = oStruct.CreateElement("Instance")
    oStruct.appendChild oElem
    Call Add_Attribute(oStruct, oElem, "Item", oStruct.SelectNodes("//Instance").Length)
    Call Add_Attribute(oStruct, oElem, "InstanceName", FormatInstanceName(oProduct.Name, 1))
    Call Add_Attribute(oStruct, oElem, "Level", "1")
    Call Add_Attribute(oStruct, oElem, "PartNumber", oProduct.PartNumber)
    Call Add_Attribute(oStruct, oElem, "CGX", "")
    Call Add_Attribute(oStruct, oElem, "CGY", "")
    Call Add_Attribute(oStruct, oElem, "CGZ", "")
    Call Add_Attribute(oStruct, oElem, "TotalQty", 1)
    Call Add_Attribute(oStruct, oElem, "NHAQty", 1)
    Call Add_Attribute(oStruct, oElem, "NHA", "N/A")
    
    'Extract attribut
    Debug.Print "#1"
    Call ExtractAttribute(oProduct, oProduct.PartNumber)

    'Scan all level of the selected assembly
    Debug.Print "#1"
    Call ScanStructure(oProduct, 2, oElem)

    'Launch Excel
    Set oExcel = New clsExcel
    Call oExcel.GetNewExcel
    Call oExcel.NewExcelFile
    oExcel.App.Calculation = xlCalculationManual

    'Display Complete structure in Excel
    Call DisplayCompleteStructure(oProduct.PartNumber)

    'Display Complete BOM in Excel
    Call DisplayCompleteBOM(oProduct.PartNumber)

    'Display Integration structure in Excel
    Call CleanStructureForIntegrationReport(bError)
    If bError = False Then
        Call DisplayIntegrationStructure(oProduct.PartNumber)
    End If
    
    oExcel.App.Sheets("COMPLETE STRUCTURE").Activate
    
    'Display Error Messages in Excel
    Call DisplayErrorMessages

    Call oExcel.ShowExcelWindow
    oExcel.App.Calculation = xlCalculationAutomatic
    frmProgress.Hide
    
    'Save file
    oExcel.SaveWorkbook ("C:\Temp\IDL_Extract_" & oProduct.PartNumber & "_" & Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) & ".xlsx")
    
    'Show Excel
    oExcel.App.Visible = True
    'Log File
    Call AddToLogFile("IDL Extract", , oProduct.PartNumber)
    
    Set oStruct = Nothing
    Set oPartAttributes = Nothing
    Set oColumnsList = Nothing
    Set oColorCodeDashNbColl = Nothing
    Set oColorCodePartExclusionList = Nothing

    frmIDLExtract.Show modeless
    

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

Private Function GetFileDateCreated(ByVal sFileName As String) As String

Dim oFSO
Dim oFile

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.GetFile(sFileName)

GetFileDateCreated = oFile.DateLastModified

End Function

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Sub ScanStructure(ByVal oProduct As Product, ByVal iLevel As Long, ByVal oXMLParent As IXMLDOMElement)

Dim sType As String
Dim oChild As Product
Dim i As Long
Dim sPartNumber As String
Dim oElem As IXMLDOMElement
Dim oBrother As IXMLDOMNode
Dim oCousin As IXMLDOMNode

Dim iNbItem As Long

For i = 1 To oProduct.Products.Count
   
    'Get child item
    Set oChild = oProduct.Products.Item(i)
    
    'Refresh Progress Bar
    If bCancelAction = True Then Exit Sub
    iNbItem = oStruct.SelectNodes("//Instance").Length
    Call frmProgress.progressBarRepaint("[" & iNbItem & "/" & iCntProdMax & "] Analyzing " & oChild.Name & "...", iCntProdMax, iNbItem)
    
    'Retrieve Part Number
    On Error Resume Next
    sPartNumber = "Error Retrieving Part Number"
    sPartNumber = oChild.PartNumber
    If sPartNumber = "Error Retrieving Part Number" Then
        oChild.ApplyWorkMode DEFAULT_MODE
    End If
    sPartNumber = oChild.PartNumber
    On Error GoTo 0
    
    'Get Type
    If sPartNumber <> "Error Retrieving Part Number" Then
        sType = GetType(oChild)
    End If

    'Extract attribut
    If sType <> "Component" Or sPartNumber = "Error Retrieving Part Number" Then
        Call ExtractAttribute(oChild, sPartNumber)
    End If
    
    'Check if a child with same part number under the same parent is already extracted
    Set oBrother = Nothing
    Set oBrother = oXMLParent.SelectSingleNode("./Instance[@PartNumber='" & sPartNumber & "' and @NHAQty!='N/A']")
    
    'Check if a child with same part number is already extracted
    Set oCousin = Nothing
    Set oCousin = oStruct.SelectSingleNode("//Instance[@PartNumber='" & sPartNumber & "' and @TotalQty!='N/A']")
    
    'Add child to structure
    If sType = "CATPart" Or sType = "CATProduct" Then
        Set oElem = Add_Element(oStruct, "Instance", oXMLParent)
        Call Add_Attribute(oStruct, oElem, "Item", oStruct.SelectNodes("//Instance").Length)
        Call Add_Attribute(oStruct, oElem, "InstanceName", FormatInstanceName(oChild.Name, iLevel))
        Call Add_Attribute(oStruct, oElem, "Level", iLevel)
        Call Add_Attribute(oStruct, oElem, "PartNumber", sPartNumber)
        Call Add_Attribute(oStruct, oElem, "CGX", "")
        Call Add_Attribute(oStruct, oElem, "CGY", "")
        Call Add_Attribute(oStruct, oElem, "CGZ", "")
        Call Add_Attribute(oStruct, oElem, "NHA", oElem.parentNode.Attributes.getNamedItem("PartNumber").nodeValue)
        If oBrother Is Nothing Then
            Call Add_Attribute(oStruct, oElem, "NHAQty", 1)
        Else
            Call Add_Attribute(oStruct, oElem, "NHAQty", "N/A")
            Call IncreaseQty(oBrother, "NHAQty")
        End If
        
        If oCousin Is Nothing Then
            Call Add_Attribute(oStruct, oElem, "TotalQty", 1)
        Else
            Call Add_Attribute(oStruct, oElem, "TotalQty", "N/A")
            Call IncreaseQty(oCousin, "TotalQty")
        End If
    End If
    
    'Extract CG
    If sExtractCG = vbYes And sType = "CATPart" Then
        Call ExtractCoG(oElem, oChild)
    End If
    
    'Scan next level
    If sPartNumber <> "Error Retrieving Part Number" Then
      
        If sType = "CATProduct" And (sExtractSP = vbYes Or Not (sPartNumber Like "*SP##*")) Then

            Call ScanStructure(oChild, iLevel + 1, oElem)
            
        ElseIf sType = "Component" Then
            Call ScanStructure(oChild, iLevel, oXMLParent)

        End If
    End If

Next

End Sub

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Sub ExtractCoG(ByRef oElem As IXMLDOMElement, ByVal oProduct As Product)

Dim oInertia 'As Inertia
Dim oArray(2) As Variant

'Get Inertia
Set oInertia = oProduct.GetTechnologicalObject("Inertia")

'Get CoG
On Error Resume Next
Call oInertia.GetCOGPosition(oArray)
On Error GoTo 0

'Put in Design mode if required
If IsEmpty(oArray(0)) And Not oProduct.PartNumber Like "*ARM*" Then
    oProduct.ApplyWorkMode DESIGN_MODE
End If

'Add data to oStruct
If Not IsEmpty(oArray(0)) Then
    oElem.Attributes.getNamedItem("CGX").nodeValue = Round(oArray(0) * 1000 / 25.4, 1)
    oElem.Attributes.getNamedItem("CGY").nodeValue = Round(oArray(1) * 1000 / 25.4, 1)
    oElem.Attributes.getNamedItem("CGZ").nodeValue = Round(oArray(2) * 1000 / 25.4, 1)
Else
    oElem.Attributes.getNamedItem("CGX").nodeValue = "Not Found"
    oElem.Attributes.getNamedItem("CGY").nodeValue = "Not Found"
    oElem.Attributes.getNamedItem("CGZ").nodeValue = "Not Found"
End If

End Sub


'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Sub ExtractAttribute(ByVal oProduct As Product, ByVal sPartNumber As String)

Dim Attributes As Variant
Dim DefiningPartAttributes As Variant
Dim sRefPartNumber As String
Dim i As Long
Dim sDefiningPart As String
Dim oItem As clsCollection
Dim sString As String
Dim sRev As String
Dim oNode As IXMLDOMNode
Dim sOriginalPartNumber As String
Dim sComment As String

'Check if attributes of sPartNumber were already extracted
If oPartAttributes.Exists(sPartNumber) Then Exit Sub

'If we can't retrieve the part number we just log the error
If sPartNumber = "Error Retrieving Part Number" Then
    Set oItem = RetrievePartAttributes(sPartNumber)
    Call oItem.GetItemByKey("Error").SetItemByKey("Value", "Error Retrieving Part Number")
    Exit Sub
End If

'We get the revision from the xml file (if required)
If Not oPVRSyncReport Is Nothing Then
    Set oNode = oPVRSyncReport.SelectSingleNode("//Part[@PartNumber='" & sPartNumber & "']")
    If Not oNode Is Nothing Then
        sRev = oPVRSyncReport.GetAttributeValue(oNode, "Revision")
    End If
    If sRev = "NA" Then sRev = "??"
End If

'Compare Part Number with Original Part Number
If Not oPVRSyncReport Is Nothing Then
    Set oNode = oPVRSyncReport.SelectSingleNode("//Part[@PartNumber='" & sPartNumber & "']")
    If Not oNode Is Nothing Then
        sOriginalPartNumber = oPVRSyncReport.GetAttributeValue(oNode, "OriginalPartNumber")
        sComment = oPVRSyncReport.GetAttributeValue(oNode, "Comment")
        If sOriginalPartNumber <> sPartNumber Then
            Set oItem = RetrievePartAttributes(sPartNumber)
            Call oItem.GetItemByKey("Error").SetItemByKey("Value", "For " & sPartNumber & " the original part was " & sOriginalPartNumber & ". " & sComment)
        End If
    End If
End If

'Extract the attributes from Enovia
If sRev <> "" Then
    Call GetAttributesOfDocRevision(Attributes, sPartNumber, , sRev)
Else
    sRev = GetAttributesOfDocRevision(Attributes, sPartNumber, oProduct)
End If

'Check if we have attribute and if the part was found in Enovia
If IsEmpty(Attributes) Then
    Set oItem = RetrievePartAttributes(sPartNumber)
    Call oItem.GetItemByKey("Revision").SetItemByKey("Value", sRev)
    If sRev = "??" Then
        Call oItem.GetItemByKey("Error").SetItemByKey("Value", "Missing revision for " & sPartNumber & ". Attributes can't be extracted.")
    Else
        Call oItem.GetItemByKey("Error").SetItemByKey("Value", "Document " & sPartNumber & " " & sRev & " not found in Enovia. Attributes can't be extracted.")
    End If
    Exit Sub
ElseIf Attributes.Count = 0 Then
    Set oItem = RetrievePartAttributes(sPartNumber)
    Call oItem.GetItemByKey("Revision").SetItemByKey("Value", sRev)
    If sRev = "??" Then
        Call oItem.GetItemByKey("Error").SetItemByKey("Value", "Missing revision for " & sPartNumber & ". Attributes can't be extracted.")
    Else
        Call oItem.GetItemByKey("Error").SetItemByKey("Value", "Document " & sPartNumber & " " & sRev & " not found in Enovia. Attributes can't be extracted.")
    End If
    Exit Sub
End If

'Transfer attributes value to oPartAttributes
Set oItem = RetrievePartAttributes(sPartNumber)
For i = 1 To oItem.Count
    If oItem.GetItemByIndex(i).GetItemByKey("EnoviaName") <> "" Then
        Call oItem.GetItemByIndex(i).SetItemByKey("Value", Attributes.GetItem(oItem.GetItemByIndex(i).GetItemByKey("EnoviaName")))
    End If
Next

'Edit Base and Dash Number
If sPartNumber Like "G########-*" Then
    
    'Base number
    Call oItem.GetItemByKey("BaseNumber").SetItemByKey("Value", Left(sPartNumber, 9))

    'Dash number
    sString = Split(sPartNumber, "-")(1)
    If sString Like "###*" Then
        Call oItem.GetItemByKey("StringDashNumber").SetItemByKey("Value", Left(sString, 3))
        Call oItem.GetItemByKey("NumericDashNumber").SetItemByKey("Value", CInt(Left(sString, 3)))
        
        'Suffix
        sString = oItem.GetItemByKey("BaseNumber").GetItemByKey("Value")
        sString = sString & "-"
        sString = sString & oItem.GetItemByKey("StringDashNumber").GetItemByKey("Value")
        sString = Replace(sPartNumber, sString, "")
        Call oItem.GetItemByKey("Suffix").SetItemByKey("Value", sString)
    Else
        Call oItem.GetItemByKey("StringDashNumber").SetItemByKey("Value", "N/A")
        Call oItem.GetItemByKey("NumericDashNumber").SetItemByKey("Value", 0)
    End If
Else
    Call oItem.GetItemByKey("BaseNumber").SetItemByKey("Value", "N/A")
    Call oItem.GetItemByKey("StringDashNumber").SetItemByKey("Value", "N/A")
    Call oItem.GetItemByKey("NumericDashNumber").SetItemByKey("Value", 0)
End If

'Edit the reference part number
Call GetRefPartNumber(sPartNumber)

'Check the defining part
sDefiningPart = ""
If Attributes.GetItem("Dataset Type") = "FLEXIBLE REPRESENTATION" Or Attributes.GetItem("Dataset Type") = "CATALOG LIGHT VERSION" Then
    sDefiningPart = Attributes.GetItem("Defining Part")
    sDefiningPart = UCase(sDefiningPart)
End If

'Report when Defining Part attribute like "*(Don't use this part)*"
If sDefiningPart Like "*(DON'T USE THIS PART)*" Then
    sDefiningPart = Replace(sDefiningPart, "(DON'T USE THIS PART)", "")
    Call oItem.GetItemByKey("Error").SetItemByKey("Value", "For " & sPartNumber & " the Defining Part attributes was tagged with: '(Don't use this part)'")
End If

If sDefiningPart <> "" Then
    Call GetAttributesOfDocRevision(DefiningPartAttributes, sDefiningPart, oProduct)
    If Not IsEmpty(DefiningPartAttributes) Then
        Call oItem.GetItemByKey("DatasetTypeDefPart").SetItemByKey("Value", DefiningPartAttributes.GetItem("Dataset Type"))
    End If
End If

'Get color code
Call oItem.GetItemByKey("ColorCode").SetItemByKey("Value", IdentifyColorCodedParts(sPartNumber, _
                                                                                   Attributes.GetItem("Dataset Type"), _
                                                                                   oItem.GetItemByKey("DatasetTypeDefPart").GetItemByKey("Value"), _
                                                                                   Replace(sDefiningPart, "(Don't use this part)", ""), _
                                                                                   oItem.GetItemByKey("Nomenclature").GetItemByKey("Value")))


'Get info from BDI
If sExtractBDIData = vbYes Then
    Call GetBDIInfo(sPartNumber)
End If

'Edit the ReportIn value. This value defines where the part will be reported in the COMPLETE BOM report
' - Part
' - Consumable
' - Other (only if Dataset Type <> COMPLETION CATALOG (REFERENCE) or STANDARD (REFERENCE DATA)
If oItem.GetItemByKey("MatSpec").GetItemByKey("Value") <> "" And oItem.GetItemByKey("DatasetType").GetItemByKey("Value") = "SKELETON" Then
    Call oItem.GetItemByKey("ReportIn").SetItemByKey("Value", "Consumable")
ElseIf oItem.GetItemByKey("DatasetType").GetItemByKey("Value") = "SKELETON" Or oItem.GetItemByKey("DatasetType").GetItemByKey("Value") = "COLLECTOR" Then
    Call oItem.GetItemByKey("ReportIn").SetItemByKey("Value", "Other")
ElseIf Not (oItem.GetItemByKey("DatasetType").GetItemByKey("Value") = "COMPLETION CATALOG (REFERENCE)" Or oItem.GetItemByKey("DatasetType").GetItemByKey("Value") = "STANDARD (REFERENCE DATA)") Then
    Call oItem.GetItemByKey("ReportIn").SetItemByKey("Value", "Part")
End If

End Sub
'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Sub GetBDIInfo(ByVal sPartNumber As String)

    Dim oItem As clsCollection
    Dim sRevOrg As String
    Dim sDatasetType As String
    Dim sSuffix As String
    Dim sBasenumber As String
    Dim iDashNumber As Integer
    Dim sDwgLatestRev As String
    Dim iDwgLatestRev As Integer 'The revision is transformed in an integer in order to find the hightest one
    Dim sDocLatestRev As String
    Dim iDocLatestRev As Integer 'The revision is transformed in an integer in order to find the hightest one
    Dim bARMFound As Boolean
    Dim sRev As String
    Dim sEDRNBox4 As String
    Dim sEDRNBox5 As String
    Dim oResults 'As clsCollection - clsCollection in DDP Toolbar is different from clsCollection in BDIWebServices
    Dim i As Integer

    
    'Initialize
    Set oItem = RetrievePartAttributes(sPartNumber)
    sRevOrg = oItem.GetItemByKey("RevOrg").GetItemByKey("Value")
    sDatasetType = oItem.GetItemByKey("DatasetType").GetItemByKey("Value")
    sBasenumber = oItem.GetItemByKey("BaseNumber").GetItemByKey("Value")
    iDashNumber = oItem.GetItemByKey("NumericDashNumber").GetItemByKey("Value")
    sSuffix = oItem.GetItemByKey("Suffix").GetItemByKey("Value")
    
    'We are sure that in the following conditions that:
    ' - CI = No
    ' - RV Number = Not Applicable
    ' - E-Item = Not Applicable
    ' - FirstAC = Not Applicable
    ' - MfgSupp = Not Applicable
    If sRevOrg Like "*STANDARD*" Or _
       sDatasetType = "SKELETON" Or _
       sDatasetType = "COLLECTOR" Or _
       sDatasetType = "KBE" Then
        Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", "No")
        Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", "Not Applicable")
        Exit Sub
    End If

    'For all non G######## parts
    If sBasenumber = "N/A" Then
        Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", "No")
        Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", "Not Applicable")
        Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", "Not Applicable")
        Exit Sub
    End If

    'Look for all children of a Fab-Assy
    'Scan all elements in oPartAttributes. If there is one item with the same base number and a matching dash (even or odd)
    If oPartAttributes.Count > 1 Then
        For i = oPartAttributes.Count - 1 To 1 Step -1

            'Compare Base Number
            If oPartAttributes.GetItemByIndex(i).GetItemByKey("BaseNumber").GetItemByKey("Value") = sBasenumber Then

                'Compare Dash Number
                If IsEven(oPartAttributes.GetItemByIndex(i).GetItemByKey("NumericDashNumber").GetItemByKey("Value")) = IsEven(iDashNumber) And iDashNumber > 100 Then
                    Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", "No")
                    Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", "Not Applicable")
                    Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", "Not Applicable")
                    Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", "Not Applicable")
                    Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", "Not Applicable")
                    Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", "Not Applicable")
                    Exit Sub
                End If
            End If
        Next
    End If


    'Look for MBPs
    If oPartAttributes.Count > 1 Then
        For i = oPartAttributes.Count - 1 To 1 Step -1

            'Compare part number
            If sPartNumber Like oPartAttributes.GetItemByIndex(i).GetItemByKey("BaseNumber").GetItemByKey("Value") & "-" & oPartAttributes.GetItemByIndex(i).GetItemByKey("StringDashNumber").GetItemByKey("Value") & "MBP*" Then
                Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", "No")
                Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", "Not Applicable")
                Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", "Not Applicable")
                Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", "Not Applicable")
                Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", "Not Applicable")
                Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", "Not Applicable")
                Exit Sub
            End If
        Next
    End If


    'Get all documents in Enovia where the Part Number is like sBasenumber*
    Set oResults = WebServiceAccessTool.GetDocumentByBaseNumber(sBasenumber)

    'Initial set
    sDwgLatestRev = ""
    iDwgLatestRev = 0
    sDocLatestRev = ""
    iDocLatestRev = 0
    bARMFound = False

    'We scan all the documents in oResults.
    ' - Find the drawing with the highest revision. A drawing is
    '   - oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sBasenumber
    ' - Find the document (except drawing) with the highest revision. The acceptable documents are:
    '   - oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber
    '   - oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") like sPartNumber & "SP##"
    '   - oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber & "PVRREF"
    '   - oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber & "PVRPMY"
    '   - oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") like sPartNumber & "SKEL##"
    ' - We check if we have a ARM
    For i = 1 To oResults.Count

        'Drawing
        If oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sBasenumber Then
            sRev = oResults.GetItem(i).GetItem("FIELD_DOCUMENT_REVISION")
            If 5 * Asc(Left(sRev, 1)) + Asc(Right(sRev, 1)) > iDwgLatestRev Then
                sDwgLatestRev = sRev
                iDwgLatestRev = 5 * Asc(Left(sDwgLatestRev, 1)) + Asc(Right(sDwgLatestRev, 1))
            End If
        End If

        'Document
        If oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber Or _
           oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") Like sPartNumber & "SP##" Or _
           oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber & "PVRREF" Or _
           oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber & "PVRPMY" Or _
           oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") Like sPartNumber & "SKEL##" Then
            sRev = oResults.GetItem(i).GetItem("FIELD_DOCUMENT_REVISION")
            If 5 * Asc(Left(sRev, 1)) + Asc(Right(sRev, 1)) > iDocLatestRev Then
                sDocLatestRev = sRev
                iDocLatestRev = 5 * Asc(Left(sDocLatestRev, 1)) + Asc(Right(sDocLatestRev, 1))
            End If
        End If

        'ARM
        If oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") Like sPartNumber & "ARM##" Then bARMFound = True
    Next

    'We have a drawing
    If sDwgLatestRev <> "" Then
        sEDRNBox4 = sBasenumber
        sEDRNBox5 = sDwgLatestRev
        
    'We have an ARM or a detail with odd dash number
    ElseIf bARMFound = True Or (iDashNumber > 100 And Not IsEven(iDashNumber)) Then
        sEDRNBox4 = sPartNumber
        sEDRNBox5 = sDocLatestRev
        
    'We have a detail with even dash number
    ElseIf iDashNumber > 100 And IsEven(iDashNumber) Then
    
        'Check if an EDRN can be found with sBaseNumber + "-" + (iDashNumber-1) + sSuffix + sDocLatestRev
        oResults.RemoveAll
        Set oResults = WebServiceAccessTool.GetBDIFullData(sBasenumber & "-" & Trim(iDashNumber - 1) & sSuffix, sDocLatestRev)
        If oResults.Count > 0 Then
            sEDRNBox4 = sBasenumber & "-" & Trim(iDashNumber - 1) & sSuffix
            sEDRNBox5 = sDocLatestRev
            GoTo ExitTheIf
        End If
    
        'Check if an EDRN can be found with sPartNumber + sDocLatestRev
        oResults.RemoveAll
        Set oResults = WebServiceAccessTool.GetBDIFullData(sPartNumber, sDocLatestRev)
        If oResults.Count > 0 Then
            sEDRNBox4 = sPartNumber
            sEDRNBox5 = sDocLatestRev
            GoTo ExitTheIf
        End If
    
        'No EDRN was found
        Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", "EDRN not found")
        Exit Sub
    
    'We have an assembly
    Else
        
        'Even dash number
        If IsEven(iDashNumber) Then
            
            'Check if an EDRN can be found with sBaseNumber + "-" + (iDashNumber-1) + sSuffix + sDocLatestRev
            oResults.RemoveAll
            Set oResults = WebServiceAccessTool.GetBDIFullData(sBasenumber & "-" & IIf(iDashNumber < 10, "00" & iDashNumber - 1, "0" & iDashNumber - 1) & sSuffix, sDocLatestRev)
            If oResults.Count > 0 Then
                sEDRNBox4 = sBasenumber & "-" & IIf(iDashNumber < 10, "00" & iDashNumber - 1, "0" & iDashNumber - 1) & sSuffix
                sEDRNBox5 = sDocLatestRev
                GoTo ExitTheIf
            End If
        End If
        
        'Check if an EDRN can be found with sPartNumber + sDocLatestRev
        oResults.RemoveAll
        Set oResults = WebServiceAccessTool.GetBDIFullData(sPartNumber, sDocLatestRev)
        If oResults.Count > 0 Then
            sEDRNBox4 = sPartNumber
            sEDRNBox5 = sDocLatestRev
            GoTo ExitTheIf
        End If
        
        'Check if an EDRN can be found with sBasenumber + "--"
        oResults.RemoveAll
        Set oResults = WebServiceAccessTool.GetBDIFullData(sBasenumber, "--")
        If oResults.Count > 0 Then
            sEDRNBox4 = sBasenumber
            sEDRNBox5 = "--"
            GoTo ExitTheIf
        End If
        
        'No EDRN was found
        Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", "EDRN not found")
        Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", "EDRN not found")
        Exit Sub
    End If
    
ExitTheIf:

    'Loop through all the EDRN, starting from latest rev and down. We are looking for the first EDNR where
    'sPartNumber will be found as an "RelatedParts"
    Do
        'Get EDRN data from web service
        oResults.RemoveAll
        Set oResults = WebServiceAccessTool.GetBDIFullData(sEDRNBox4, sEDRNBox5)

        'Scan oResults and search for sPartNumber as a "RelatedParts"
        If oResults.Count > 0 Then
            If oResults.Exists("RelatedParts") Then
                For i = 1 To oResults.GetItem("RelatedParts").Count
                    If oResults.GetItem("RelatedParts").GetItem(i).GetItem("PARTNUMBER") = sPartNumber Then
                        Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", IIf(oResults.GetItem("RelatedParts").GetItem(i).GetItem("CI") = "false", "No", "Yes"))
                        Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", IIf(oResults.GetItem("EDRN").GetItem("FIELD_Edrn_9_ModsumNo") = "", "Empty info in EDRN", oResults.GetItem("EDRN").GetItem("FIELD_Edrn_9_ModsumNo")))
                        Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", IIf(oResults.GetItem("EDRN").GetItem("FIELD_eItemNoAndTitle") = "", "Empty info in EDRN", oResults.GetItem("EDRN").GetItem("FIELD_eItemNoAndTitle")))
                        Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", IIf(oResults.GetItem("dps_attributes").GetItem("FIELD_FIRST_AIRCRAFT") = "", "Empty info in EDRN", oResults.GetItem("dps_attributes").GetItem("FIELD_FIRST_AIRCRAFT")))
                        Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", IIf(oResults.GetItem("dps_attributes").GetItem("FIELD_MANUFACTURING_SUPPLIER") = "", "Empty info in EDRN", oResults.GetItem("dps_attributes").GetItem("FIELD_MANUFACTURING_SUPPLIER")))
                        Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", sEDRNBox4 & sEDRNBox5)
                        Exit Do
                    End If
                Next
            End If
        End If

        'If the revision is "--" and we didn't find anything
        If sEDRNBox5 = "--" Then
            Call oItem.GetItemByKey("CIStatus").SetItemByKey("Value", "EDRN not found")
            Call oItem.GetItemByKey("RVNumber").SetItemByKey("Value", "EDRN not found")
            Call oItem.GetItemByKey("E-Item").SetItemByKey("Value", "EDRN not found")
            Call oItem.GetItemByKey("FirstAC").SetItemByKey("Value", "EDRN not found")
            Call oItem.GetItemByKey("MfgSupp").SetItemByKey("Value", "EDRN not found")
            Call oItem.GetItemByKey("EDRN").SetItemByKey("Value", "EDRN not found")
            Exit Do
        End If

        'Nothing was found we check for previous rev
        sEDRNBox5 = GetPreviousRev(sEDRNBox5)

    Loop

End Sub


Private Function IsEven(ByVal iNumber As Integer) As Boolean

If iNumber Mod 2 = 0 Then
  IsEven = True
Else
    IsEven = False
End If

End Function


'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Function GetPreviousRev(ByVal sRev As String) As String

    Dim sFirstChar As String
    Dim sSecondChar As String
    
    sFirstChar = Left(sRev, 1)
    sSecondChar = Right(sRev, 1)
    
    'Special case
    If sRev = "-A" Then
       GetPreviousRev = "--"
       Exit Function
    End If
    
    If Asc(sSecondChar) = 65 Then
        sSecondChar = "Y"
        If sFirstChar = "A" Then
            sFirstChar = "-"
        Else
            sFirstChar = Chr(Asc(sFirstChar) - 1)
        End If
        GetPreviousRev = sFirstChar & sSecondChar

    Else
        sSecondChar = Chr(Asc(sSecondChar) - 1)
        GetPreviousRev = sFirstChar & sSecondChar
    End If

    
End Function

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Function IdentifyColorCodedParts(ByVal sPartNumber As String, ByVal sDatasetType As String, ByVal sDefiningPartDatasetType As String, ByVal sDefiningPart As String, ByVal sNomenclature As String) As String

Dim sDashNumber As String

'Define Part Number and Dataset Type
If (sDatasetType = "FLEXIBLE REPRESENTATION" Or sDatasetType = "CATALOG LIGHT VERSION") And sDefiningPart <> "" Then
    sPartNumber = sDefiningPart
    sDatasetType = sDefiningPartDatasetType
ElseIf sNomenclature <> "" Then
    sPartNumber = sNomenclature
End If


'System Parts
If (sPartNumber Like "*-XXXX" Or sPartNumber Like "*-XXXX-YYYY") And Not sPartNumber Like "G25*" And Not sPartNumber Like "G025*-?-????" And _
   (sDatasetType = "PROCUREMENT CONTROL DATASET" Or sDatasetType = "SYSTEM VENDOR COMPONENT") Then
    IdentifyColorCodedParts = "System"

'*-XXXX
ElseIf sPartNumber Like "*-XXXX" And Not sPartNumber Like "G025*-?-????" Then
    IdentifyColorCodedParts = "X"

'Other System Parts
ElseIf sPartNumber Like "*-????" And Not sPartNumber Like "G25*" And Not sPartNumber Like "G025*-?-????" And _
      (sDatasetType = "PROCUREMENT CONTROL DATASET" Or sDatasetType = "SYSTEM VENDOR COMPONENT") Then
    
    'Extract dash number
    sDashNumber = Split(sPartNumber, "-")(UBound(Split(sPartNumber, "-")))
    
    'Populate oColorCodePartExclusionList
    If oColorCodePartExclusionList.Count = 0 Then Call PopulateColorCodePartExclusionList
    
    'Check if sPartNumber is found in oColorCodePartExclusionList
    If oColorCodePartExclusionList.Exists(sPartNumber) Then
        IdentifyColorCodedParts = ""
        Exit Function
    End If
    
    'Populate oColorCodeDashNbColl
    If oColorCodeDashNbColl.Count = 0 Then Call PopulateColorCodeDashNumberColl

    'Check if dash number is in oColorCodeDashNbColl
    If oColorCodeDashNbColl.Exists("-" & sDashNumber) Then
        IdentifyColorCodedParts = "System"
    End If

'Other G25 color, but -XXXX was replaced with color code
ElseIf sPartNumber Like "*-????" And Not sPartNumber Like "G025*-?-????" Then

    'Extract dash number
    sDashNumber = Split(sPartNumber, "-")(UBound(Split(sPartNumber, "-")))

    'Populate oColorCodePartExclusionList
    If oColorCodePartExclusionList.Count = 0 Then Call PopulateColorCodePartExclusionList
    
    'Check if sPartNumber is found in oColorCodePartExclusionList
    If oColorCodePartExclusionList.Exists(sPartNumber) Then
        IdentifyColorCodedParts = ""
        Exit Function
    End If

    'Populate oColorCodeDashNbColl
    If oColorCodeDashNbColl.Count = 0 Then Call PopulateColorCodeDashNumberColl
    
    'Check if dash number is in oColorCodeDashNbColl
    If oColorCodeDashNbColl.Exists("-" & sDashNumber) Then
        IdentifyColorCodedParts = "X"
    End If
End If

End Function

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Sub PopulateColorCodeDashNumberColl()

    Dim sTextRow As String, iFileNo As Integer
    Dim iLine As Integer
    
    'Open file
    iFileNo = FreeFile
    Open sColorCodeDashNumber For Input As iFileNo

    'Scan file and transfer color code to oColorCodeDashNbColl
    iLine = 0
    Do While Not EOF(iFileNo)
        
        iLine = iLine + 1
        Line Input #iFileNo, sTextRow
        
        'The first 2 line in the file should be skipped
        If iLine >= 3 Then
            Call oColorCodeDashNbColl.Add(Split(sTextRow, ",")(3), Split(sTextRow, ",")(3))
        End If
    Loop
    Close #iFileNo

End Sub

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Sub PopulateColorCodePartExclusionList()

    Dim sTextRow As String, iFileNo As Integer
    Dim iLine As Integer
    
    'Open file
    iFileNo = FreeFile
    Open sColorCodePartExclusionList For Input As iFileNo

    'Scan file and transfer part number to oColorCodePartExclusionList
    iLine = 0
    Do While Not EOF(iFileNo)
        
        iLine = iLine + 1
        Line Input #iFileNo, sTextRow
        
        'The first line in the file should be skipped
        If iLine >= 2 Then
            Call oColorCodePartExclusionList.Add(sTextRow, sTextRow)
        End If
    Loop
    Close #iFileNo

End Sub

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Private Function GetType(ByVal oProduct As Product) As String

'CATPart
If oProduct.HasAMasterShapeRepresentation Then
    GetType = "CATPart"
ElseIf oProduct.PartNumber Like "*SP##*" Then
    GetType = "CATProduct"

'I can only identify the components that are directly under an SP node. For all the other ones (ex. found
'under electrical assemblies) i can't differiantiate them from none converted to component in PVR.
'ElseIf oProduct.ReferenceProduct.Parent.Name = oProduct.Parent.Parent.ReferenceProduct.Parent.Name Then
ElseIf oProduct.ReferenceProduct.Parent.Name Like "*SP##*" Then
    GetType = "Component"
Else
    GetType = "CATProduct"
End If
 
End Function

'######################################################################################
'######################################################################################
'######################################################################################
Sub DisplayCompleteBOM(ByVal sProductPartNumber As String)

    Dim iRow, iRowMax, iColMax
    Dim iTotalRefs As Long
    Dim oPart() As Variant
    Dim oConsumable() As Variant
    Dim oOther() As Variant
    Dim oInstances As IXMLDOMNodeList
    Dim oInstance As IXMLDOMNode
    Dim sReportIn As String
    Dim sTitle As String
    Dim iQty As Integer
    Dim sPartNumber As String
    Dim sRefPartNumber As String

    iXLRow = iXLHeaderRow

    'Add new sheet in Excel workbook
    Call oExcel.App.Sheets.Add
    Set oSheet = oExcel.App.ActiveSheet
    oExcel.App.ActiveSheet.Name = "COMPLETE BOM"

    'Format columns to text
    oExcel.App.Columns(2).NumberFormat = "@"
    oExcel.App.Columns(5).NumberFormat = "@"
    oExcel.App.Columns(8).NumberFormat = "@"

    'Header
    oSheet.Cells(iXLHeaderRow, 1).Value = "ITEM"
    oSheet.Cells(iXLHeaderRow, 2).Value = "PARTS"
    oSheet.Cells(iXLHeaderRow, 3).Value = "QTY"
    oSheet.Cells(iXLHeaderRow, 4).Value = "TITLE"
    oSheet.Cells(iXLHeaderRow, 5).Value = "CONSUMABLES"
    oSheet.Cells(iXLHeaderRow, 6).Value = "QTY"
    oSheet.Cells(iXLHeaderRow, 7).Value = "TITLE"
    oSheet.Cells(iXLHeaderRow, 8).Value = "OTHERS"
    oSheet.Cells(iXLHeaderRow, 9).Value = "QTY"
    oSheet.Cells(iXLHeaderRow, 10).Value = "TITLE"
    oSheet.Range(oExcel.App.Cells(iXLHeaderRow, 1), oExcel.App.Cells(iXLHeaderRow, 10)).Font.Bold = True
    
    ReDim oPart(1 To 3, 1 To 1)
    ReDim oConsumable(1 To 3, 1 To 1)
    ReDim oOther(1 To 3, 1 To 1)

    'Get all the instance with a total qty <> "N/A" and transfer in the appropriate array
    Set oInstances = oStruct.SelectNodes("//Instance[@TotalQty!='N/A']")
    For i = 1 To oInstances.Length
        
        Set oInstance = oInstances.Item(i - 1)
        sPartNumber = oInstance.Attributes.getNamedItem("PartNumber").nodeValue
        sTitle = oPartAttributes.GetItemByKey(sPartNumber).GetItemByKey("Title").GetItemByKey("Value")
        iQty = CInt(oInstance.Attributes.getNamedItem("TotalQty").nodeValue)
        sReportIn = oPartAttributes.GetItemByKey(sPartNumber).GetItemByKey("ReportIn").GetItemByKey("Value")
        sRefPartNumber = oPartAttributes.GetItemByKey(sPartNumber).GetItemByKey("RefPartNumber").GetItemByKey("Value")

        'Transfer the data in the appropriate array
        Select Case sReportIn
            Case "Part"
                If oPart(1, UBound(oPart, 2)) <> "" Then
                    ReDim Preserve oPart(1 To 3, 1 To UBound(oPart, 2) + 1)
                End If
                oPart(1, UBound(oPart, 2)) = sRefPartNumber
                oPart(2, UBound(oPart, 2)) = iQty
                oPart(3, UBound(oPart, 2)) = sTitle
                
            Case "Consumable"
                If oConsumable(1, UBound(oConsumable, 2)) <> "" Then
                    ReDim Preserve oConsumable(1 To 3, 1 To UBound(oConsumable, 2) + 1)
                End If
                oConsumable(1, UBound(oConsumable, 2)) = sRefPartNumber
                oConsumable(2, UBound(oConsumable, 2)) = iQty
                oConsumable(3, UBound(oConsumable, 2)) = sTitle
                
            Case "Other"
                If oOther(1, UBound(oOther, 2)) <> "" Then
                    ReDim Preserve oOther(1 To 3, 1 To UBound(oOther, 2) + 1)
                End If
                oOther(1, UBound(oOther, 2)) = sRefPartNumber
                oOther(2, UBound(oOther, 2)) = iQty
                oOther(3, UBound(oOther, 2)) = sTitle
        
        End Select
    Next

    'Display data
    If oPart(1, UBound(oPart, 2)) <> "" Then
        oSheet.Cells(iXLHeaderRow + 1, 2).Resize(UBound(oPart, 2), 3).Value = oExcel.App.Transpose(oPart)
    End If
    
    If oConsumable(1, UBound(oConsumable, 2)) <> "" Then
        oSheet.Cells(iXLHeaderRow + 1, 5).Resize(UBound(oConsumable, 2), 3).Value = oExcel.App.Transpose(oConsumable)
    End If
    
    If oOther(1, UBound(oOther, 2)) <> "" Then
        oSheet.Cells(iXLHeaderRow + 1, 8).Resize(UBound(oOther, 2), 3).Value = oExcel.App.Transpose(oOther)
    End If
    
    'Max row and column
    iRowMax = iXLHeaderRow + oExcel.App.Max(UBound(oPart, 2), UBound(oConsumable, 2), UBound(oOther, 2))
    iColMax = 10
    
    'Display item column
    If iRowMax >= iXLHeaderRow + 1 Then
        For iRow = iXLHeaderRow + 1 To iRowMax
            oSheet.Cells(iRow, 1).Value = iRow - iXLHeaderRow
        Next
    End If
    
    ' Format Columns:
    If IsEmpty(iRowMax) Then Exit Sub
    With oSheet.Range(oSheet.Cells(iXLHeaderRow, 1), oSheet.Cells(iRowMax, iColMax)).Columns
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    oSheet.Range(oSheet.Cells(iXLHeaderRow, 1), oSheet.Cells(iXLHeaderRow, iColMax)).AutoFilter

    ' Format Rows:
    With oSheet.Range(oSheet.Cells(iXLHeaderRow, 1), oSheet.Cells(iRowMax, iColMax))
        .Font.Size = 8
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
    End With

    'Count the number of reference
    iTotalRefs = 0
    If oPart(1, UBound(oPart, 2)) <> "" Then iTotalRefs = iTotalRefs + UBound(oPart, 2)
    If oConsumable(1, UBound(oConsumable, 2)) <> "" Then iTotalRefs = iTotalRefs + UBound(oConsumable, 2)
    If oOther(1, UBound(oOther, 2)) <> "" Then iTotalRefs = iTotalRefs + UBound(oOther, 2)
    
    
    'Display
    oSheet.Cells(1, 2) = "BOM of " & sProductPartNumber
    oSheet.Cells(2, 2) = "Total of " & iTotalRefs & " References"
    oSheet.Cells(3, 2) = Format(Now, "YYYY-mm-dd HH:MM")

    oSheet.Cells.EntireColumn.AutoFit
    oSheet.Columns(1).ColumnWidth = 5

End Sub

'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Sub DisplayCompleteStructure(ByVal sProductPartNumber As String)
    Dim i As Long
    Dim j As Long
    Dim iColumn As Integer
    Dim oSheet
    Dim oData() As Variant
    Dim oInstances As IXMLDOMNodeList
    Dim oInstance As IXMLDOMNode
    Dim oItem As clsCollection
    Dim sPartNumber As String

    'Progress bar
    Call frmProgress.progressBarRepaint("Formating Excel sheets.....Please wait....")
   
    'Transfer data to oData
    Set oInstances = oStruct.SelectNodes("//Instance")
    ReDim oData(1 To oInstances.Length, 1 To oColumnsList.Count)
    
    For i = 1 To oInstances.Length
        
        Set oInstance = oInstances.Item(i - 1)
        sPartNumber = oInstance.Attributes.getNamedItem("PartNumber").nodeValue
        
        
        For j = 1 To oColumnsList.Count
        
            'The info should be retrieve from the XML Structure
            If oColumnsList.GetItemByIndex(j).GetItemByKey("Source") = "Instance" Then
                oData(i, j) = oInstance.Attributes.getNamedItem(oColumnsList.GetKey(j)).nodeValue
            
            'The info should be retrieve from the the collection oPartAttributes
            Else
                oData(i, j) = oPartAttributes.GetItemByKey(sPartNumber).GetItemByKey(oColumnsList.GetKey(j)).GetItemByKey("Value")
            End If
        Next
    Next
    
    'Add new sheet in Excel workbook
    Call oExcel.App.Sheets.Add
    Set oSheet = oExcel.App.ActiveSheet
    oExcel.App.ActiveSheet.Name = "COMPLETE STRUCTURE"

    'Format columns to text
    oExcel.App.Columns(4).Resize(, 4).NumberFormat = "@"

    'Display values in Excel
    oExcel.App.Range(oExcel.App.Cells(iXLHeaderRow + 1, 1), oExcel.App.Cells(iXLHeaderRow + UBound(oData, 1), UBound(oData, 2))).Value = oData
    
    ' Format Columns:
    With oSheet.Range(oSheet.Cells(iXLHeaderRow, 1), oSheet.Cells(iXLHeaderRow + 1 + UBound(oData, 1), UBound(oData, 2))).Columns
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    oSheet.Columns(2).HorizontalAlignment = xlLeft

    ' Display Header:
    For iColumn = 1 To oColumnsList.Count
        oSheet.Cells(iXLHeaderRow, iColumn) = oColumnsList.GetItemByIndex(iColumn).GetItemByKey("DisplayName")
    Next


    ' Format Header:
    With oSheet.Range(oSheet.Cells(iXLHeaderRow, 1), oSheet.Cells(iXLHeaderRow, UBound(oData, 2)))
        .Font.Size = 8
        .Font.Bold = True
        .RowHeight = 30
        .AutoFilter
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With

    ' Format Rows:
    With oSheet.Range(oSheet.Cells(iXLHeaderRow + 1, 1), oSheet.Cells(iXLHeaderRow + UBound(oData, 1), UBound(oData, 2)))
        .Font.Size = 8
        .RowHeight = 15
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
    End With


    oSheet.Cells(iXLHeaderRow + 1, 3).Select
    oSheet.Activate
    oExcel.App.ActiveWindow.FreezePanes = True

    oSheet.Cells(1, 1) = ""
    oSheet.Cells(1, 2) = "Assembly " & sProductPartNumber
    oSheet.Cells(3, 2) = Format(Now, "YYYY-mm-dd HH:MM")
    oSheet.Cells(2, 2) = "Total of " & UBound(oData, 1) & " instances"
    oSheet.Cells.EntireColumn.AutoFit
    oSheet.Columns(1).ColumnWidth = 5
End Sub

Sub DisplayErrorMessages()

    Dim i As Long
    Dim oSheet
    Dim sErrorMsg As String
    Dim oData() As Variant
    Dim oNodeList As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim sRev As String
    
    ReDim oData(1 To 1)
    
    'Initialize
    oData(1) = "The error/warning messages are:"
    
    'Scan all the references and get the error messages
    For i = 1 To oPartAttributes.Count
    
        Set oItem = oPartAttributes.GetItemByIndex(i)
        
        sPartNumber = oPartAttributes.GetKey(i)
        sErrorMsg = oItem.GetItemByKey("Error").GetItemByKey("Value")
        sRev = oItem.GetItemByKey("Revision").GetItemByKey("Value")
        
        'Get all the instances where we can't retrieve the part number
        If sPartNumber = "Error Retrieving Part Number" Then
        
            Set oNodeList = oStruct.SelectNodes("//Instance[@PartNumber='Error Retrieving Part Number']")
            For Each oNode In oNodeList
                ReDim Preserve oData(1 To UBound(oData) + 1)
                oData(UBound(oData)) = "Impossible to retrieve reference of instance " & Trim(Replace(oNode.Attributes.getNamedItem("InstanceName").nodeValue, "*", ""))
            Next

        ElseIf sErrorMsg <> "" Then
                ReDim Preserve oData(1 To UBound(oData) + 1)
                oData(UBound(oData)) = sErrorMsg
        End If
    Next

    'We have errors we display them
    If UBound(oData) > 1 Then
    
        'Add new sheet in Excel workbook
        Call oExcel.App.Sheets.Add
        Set oSheet = oExcel.App.ActiveSheet
        oExcel.App.ActiveSheet.Name = "ErrorMsg"
    
   
        oSheet.Cells(1, 1).Resize(UBound(oData)).Value = oExcel.App.Transpose(oData)
    End If
End Sub


'######################################################################################
'######################################################################################
'######################################################################################
'######################################################################################
Sub DisplayIntegrationStructure(ByVal sProductPartNumber As String)
    Dim i As Long
    Dim j As Long
    Dim oSheet
    Dim iXLRow As Long
    Dim iNbLevel As Long
    Dim oInstances As IXMLDOMNodeList
    Dim oInstance As IXMLDOMNode
    Dim iMaxLevel As Integer
    Dim iLevel As Integer
    Dim oItem As clsCollection
    

    'Add new sheet in Excel workbook
    Call oExcel.App.Sheets.Add
    Set oSheet = oExcel.App.ActiveSheet
    oExcel.App.ActiveSheet.Name = "INTEGRATION STRUCTURE"

    'From oStruct get the highest level
    iMaxLevel = 0
    Set oInstances = oStruct.SelectNodes("//Instance")
    For Each oInstance In oInstances
        If CInt(oInstance.Attributes.getNamedItem("Level").nodeValue) > iMaxLevel Then
            iMaxLevel = CInt(oInstance.Attributes.getNamedItem("Level").nodeValue)
        End If
    Next

    'Transfer data to oData
    Set oInstances = oStruct.SelectNodes("//Instance")
    ReDim oData(1 To oInstances.Length, 1 To iMaxLevel + 12)
    For i = 1 To oInstances.Length
        
        Set oInstance = oInstances.Item(i - 1)
        sPartNumber = oInstance.Attributes.getNamedItem("PartNumber").nodeValue
        iLevel = CInt(oInstance.Attributes.getNamedItem("Level").nodeValue)
        Set oItem = oPartAttributes.GetItemByKey(sPartNumber)
        
        oData(i, 1) = i
        oData(i, 2) = oItem.GetItemByKey("Title").GetItemByKey("Value")
        oData(i, 2 + iLevel) = oItem.GetItemByKey("RefPartNumber").GetItemByKey("Value")
        oData(i, 3 + iMaxLevel) = CInt(oInstance.Attributes.getNamedItem("NHAQty").nodeValue)
        oData(i, 4 + iMaxLevel) = oItem.GetItemByKey("DatasetType").GetItemByKey("Value")
        oData(i, 5 + iMaxLevel) = oItem.GetItemByKey("MatSpecProd").GetItemByKey("Value")
        oData(i, 6 + iMaxLevel) = oItem.GetItemByKey("Alloy").GetItemByKey("Value")
        oData(i, 7 + iMaxLevel) = oItem.GetItemByKey("FinalCondition").GetItemByKey("Value")
        oData(i, 8 + iMaxLevel) = oItem.GetItemByKey("EngMakeFrom").GetItemByKey("Value")
        oData(i, 9 + iMaxLevel) = oItem.GetItemByKey("Grade/Composition").GetItemByKey("Value")
        oData(i, 10 + iMaxLevel) = oItem.GetItemByKey("MaterialClass").GetItemByKey("Value")
        oData(i, 11 + iMaxLevel) = oItem.GetItemByKey("Type").GetItemByKey("Value")
        oData(i, 12 + iMaxLevel) = oItem.GetItemByKey("SupplierName").GetItemByKey("Value")
                
    Next


    'Format columns with part number to text
    For i = 3 To 3 + iMaxLevel - 1
        oExcel.App.Columns(i).NumberFormat = "@"
    Next
    
    'Display values in Excel
    oExcel.App.Range(oExcel.App.Cells(iXLHeaderRow + 1, 1), oExcel.App.Cells(iXLHeaderRow + UBound(oData, 1), UBound(oData, 2))).Value = oData
    
    'Format Columns:
    With oSheet.Range(oSheet.Cells(iXLHeaderRow, 1), oSheet.Cells(iXLHeaderRow + UBound(oData, 1), UBound(oData, 2))).Columns
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    oSheet.Columns(2).HorizontalAlignment = xlLeft


    'Display Header
    oSheet.Cells(iXLHeaderRow, 1) = "ITEM"
    oSheet.Cells(iXLHeaderRow, 2) = "TITLE"
    For i = 1 To iMaxLevel
        oSheet.Cells(iXLHeaderRow, i + 2) = "LEVEL" & i
    Next
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 3) = "QTY PER NHA"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 4) = "DATASET TYPE"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 5) = "MATERIAL SPECIFICATION PRODUCTION"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 6) = "ALLOY"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 7) = "FINAL CONDITION"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 8) = "ENG MAKE FROM"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 9) = "GRADE/COMPOSITION"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 10) = "CLASS"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 11) = "TYPE"
    oSheet.Cells(iXLHeaderRow, iMaxLevel + 12) = "MAJOR SUPPLIER CODE"
    

    'Format Header:
    With oSheet.Range(oSheet.Cells(iXLHeaderRow, 1), oSheet.Cells(iXLHeaderRow, UBound(oData, 2)))
        .Font.Size = 8
        .RowHeight = 30
        .Font.Bold = True
        .AutoFilter
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With

    ' Format Rows:
    With oSheet.Range(oSheet.Cells(iXLHeaderRow + 1, 1), oSheet.Cells(iXLHeaderRow + UBound(oData, 1), UBound(oData, 2)))
        .Font.Size = 8
        .RowHeight = 15
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
    End With


    oSheet.Cells(iXLHeaderRow + 1, 3).Select
    oSheet.Activate
    oExcel.App.ActiveWindow.FreezePanes = True

    oSheet.Cells(1, 1) = ""
    oSheet.Cells(1, 2) = "Assembly " & sProductPartNumber
    oSheet.Cells(3, 2) = Format(Now, "YYYY-mm-dd HH:MM")
    oSheet.Cells(2, 2) = "Total of " & UBound(oData, 1) & " items"
    oSheet.Cells.EntireColumn.AutoFit
    oSheet.Columns(1).ColumnWidth = 5
End Sub

Private Sub Class_Terminate()
    Set oExcel = Nothing
End Sub

Private Sub IncreaseQty(ByVal oElem As IXMLDOMElement, ByVal sNodename As String)

oElem.Attributes.getNamedItem(sNodename).nodeValue = CInt(oElem.Attributes.getNamedItem(sNodename).nodeValue) + 1

End Sub
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

