Attribute VB_Name = "PLExtract"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private bStatus As Boolean
Private sConsumableCageCode As clsCollection
Private oExcel As clsExcel
Dim sLogDump As String
Dim oWngColl As clsCollection


Public Sub LaunchExtraction(ByVal oNHA_LHS As Product, ByVal oNHA_RHS As Product)

Dim iRow As Integer, iColumn As Integer, i As Integer
Dim oSP As Product, oChild As Product
Dim sString As String
Dim oData As clsXML

'Get Cage Code from Consumable
Set sConsumableCageCode = New clsCollection
Call GetConsumableCageCode

'Log dump file
sLogDump = ""

'Initialize
Set oData = New clsXML
Call oData.AddElement("Root")

'Extract NHA LHS
Set oWngColl = New clsCollection
Call ExtractAttributValues(oData, oNHA_LHS, "LHS", "01_NHA")

'Clear cache
Call StartWebServiceTool
WebServiceAccessTool.ClearCache

'Extract NHA RHS
If Not oNHA_RHS Is Nothing Then
    Call ExtractAttributValues(oData, oNHA_RHS, "RHS", "02_NHA")
End If


'Scan Child Parts
Set oSP = Nothing
bStatus = True
Call frmProgress.progressBarInitialize("Extracting Data")

For i = 1 To oNHA_LHS.Products.Count

    If bStatus = False Then Exit For
    
    Set oChild = oNHA_LHS.Products.Item(i)
    
    Call CheckInstance(oChild)
        
    If bStatus = False Then Exit Sub
    
    Call frmProgress.progressBarRepaint("Extracting " & oChild.PartNumber, oNHA_LHS.Products.Count, i)

    If oChild.PartNumber Like "*SP*" Then
    
        Call ScanSPNode(oData, oChild, "LHS")
    Else
    
        If oChild.ReferenceProduct.Parent.Name Like "*CATProduct" Then
            Call ExtractAttributValues(oData, oChild, "LHS", "03_Assembly")
        Else
            Call ExtractAttributValues(oData, oChild, "LHS", "04_Detail")
        End If
    End If
    
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit For
Next

If Not oNHA_RHS Is Nothing Then
    For i = 1 To oNHA_RHS.Products.Count
    
        If bStatus = False Then Exit For
        
        Set oChild = oNHA_RHS.Products.Item(i)
        
        Call CheckInstance(oChild)
            
        If bStatus = False Then Exit Sub
        
        Call frmProgress.progressBarRepaint("Extracting " & oChild.PartNumber, oNHA_RHS.Products.Count, i)
    
        If oChild.PartNumber Like "*SP*" Then
        
            Call ScanSPNode(oData, oChild, "RHS")
        Else
        
            If oChild.ReferenceProduct.Parent.Name Like "*CATProduct" Then
                Call ExtractAttributValues(oData, oChild, "RHS", "03_Assembly")
            Else
                Call ExtractAttributValues(oData, oChild, "RHS", "04_Detail")
            End If
        End If
        
        If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit For
    Next
End If

Call oData.SortNodes(oData.RootNode, "Sort")

'Log File
Call AddToLogFile("Extract Part List", , , , , , , sLogDump)

'Display in Excel
Call DisplayInExcel(oData, oNHA_RHS)

'Kill Excel object
Set oData = Nothing

'Error message
If oWngColl.Count > 0 Then
    For i = 1 To oWngColl.Count
        If i = 1 Then
            sString = oWngColl.GetItemByIndex(i)
        Else
            sString = sString & vbCrLf & oWngColl.GetItemByIndex(i)
        End If
    Next
    Call MsgBox(sString, vbCritical)
End If

Call oExcel.ShowExcelWindow
Set oExcel = Nothing

Call frmKBEMain.resetToolbar
End Sub

Private Sub ScanSPNode(ByRef oData As clsXML, ByVal oProduct As Product, ByVal sSide As String)

Dim oChild As Product
Dim sChildType As String
Dim i As Integer

For i = 1 To oProduct.Products.Count
    
    If bStatus = False Then Exit Sub
    
    Set oChild = oProduct.Products.Item(i)

    Call CheckInstance(oChild)
    
    If bStatus = False Then Exit Sub
    
    'Child is a component
    If oChild.ReferenceProduct.Parent.Name = oChild.Parent.Parent.ReferenceProduct.Parent.Name Then
        Call ScanSPNode(oData, oChild, sSide)
    Else
        Call ExtractAttributValues(oData, oChild, sSide, "05_Std")
    End If
        
    If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then Exit Sub
Next

End Sub

Private Sub ExtractAttributValues(ByRef oData As clsXML, ByVal oProduct As Product, ByVal sSide As String, ByVal sType As String)

Dim sPartNumber As String, sDefiningPart As String, sMaterialSpec As String, sRevision As String, sDatasetType As String
Dim oParameters As Parameters
Dim oStrParam As StrParam
Dim i As Integer
Dim sParamValue As String, sString As String
Dim bRetrieveFromWebService As Boolean
Dim params()
Dim SysService
Dim sAttributes As Variant
Dim oNode As IXMLDOMNode
Dim sAttNames(1 To 10) As String, sAttValues(1 To 10) As String

'*** Initial Set
bRetrieveFromWebService = False

'Initialize attribute name
sAttNames(1) = "Sort": sAttNames(2) = "QtyRHS": sAttNames(3) = "QtyLHS": sAttNames(4) = "CageCode": sAttNames(5) = "PartNumber"
sAttNames(6) = "Description": sAttNames(7) = "Size": sAttNames(8) = "MatlSpec": sAttNames(9) = "FinishCode": sAttNames(10) = "PCCN":

'*** Parameters
Set oParameters = oProduct.ReferenceProduct.UserRefProperties

'*** Check the value of the "Dataset Type"
On Error Resume Next
sMaterialSpec = ""
sDefiningPart = ""
Set oStrParam = Nothing
Set oStrParam = oParameters.GetItem("Dataset Type")
If Not oStrParam Is Nothing Then

    'When Dataset Type = "SKELETON" we retrieve the "Material Specifications"
    'If "Material Specifications" = "" then we don't have a consumable and we exit Sub
    If oStrParam.Value = "SKELETON" Then
        Set oStrParam = Nothing
        Set oStrParam = oParameters.GetItem("Material Specifications")
        If Not oStrParam Is Nothing Then
            sMaterialSpec = oStrParam.Value
            If sMaterialSpec = "" Then Exit Sub
        End If
    
    '"COMPLETION CATALOG (REFERENCE)" or "STANDARD (REFERENCE DATA)" should not be in the PL
    ElseIf oStrParam.Value = "COMPLETION CATALOG (REFERENCE)" Or oStrParam.Value = "STANDARD (REFERENCE DATA)" Then
        Exit Sub
    
    'When Dataset Type = "CATALOG LIGHT VERSION" we retrieve the "Defining Part"
    ElseIf oStrParam.Value = "CATALOG LIGHT VERSION" Or oStrParam.Value = "FLEXIBLE REPRESENTATION" Then
        sDatasetType = oStrParam.Value
        Set oStrParam = Nothing
        Set oStrParam = oParameters.GetItem("Defining Part")
        If Not oStrParam Is Nothing Then
            
            sDefiningPart = oStrParam.Value
            bRetrieveFromWebService = True
            
            'Defining part attribut must be non-null
            If sDefiningPart = "" Then
                If Not oWngColl.Exists(oProduct.PartNumber) Then
                    sString = "- Dataset Type of " & oProduct.PartNumber & " is " & sDatasetType & ". Defining Part should not be empty."
                    Call oWngColl.Add(oProduct.PartNumber, sString)
                End If
            End If
            
            'Defining Part: "(Don't use this part)"
            If UCase(sDefiningPart) Like "*(DON*T USE THIS PART)*" Then
                If Not oWngColl.Exists("DONTUSE") Then
                    sString = "- Parts tagged '(Don't use this part)' were not extracted in the PL"
                    Call oWngColl.Add("DONTUSE", sString)
                End If
                Exit Sub
            End If
        
        'Defining part attribut not found
        Else
            If Not oWngColl.Exists(oProduct.PartNumber) Then
                sString = "Dataset Type of " & oProduct.PartNumber & " is " & sDatasetType & ". Defining Part should not be empty."
                Call oWngColl.Add(oProduct.PartNumber, sString)
            End If
        End If
    End If
End If
On Error GoTo 0


'**** Define the Part Number. sPartNumber =
'*** - sMaterialSpec
'*** - DefiningPart
'*** - Nomenclature
'*** - Part Number
If sMaterialSpec <> "" Then
    sPartNumber = sMaterialSpec
ElseIf sDefiningPart <> "" Then
    sPartNumber = sDefiningPart
ElseIf oProduct.Nomenclature <> "" Then
    sPartNumber = oProduct.Nomenclature
Else
    sPartNumber = oProduct.PartNumber
End If

'Debug.Print sPartNumber

'*************
'Part already extracted we just update the qty
If oData.Exists("./Part[@PartNumber='" & sPartNumber & "']", oData.RootNode) Then
    
    Set oNode = oData.SelectSingleNode("./Part[@PartNumber='" & sPartNumber & "']", oData.RootNode)
    
    If sMaterialSpec = "" Then
        Call oData.EditAttributeValue(oNode, "Qty" & sSide, CInt(oData.GetAttributeValue(oNode, "Qty" & sSide)) + 1)
    Else
        Call oData.EditAttributeValue(oNode, "Qty" & sSide, "A/R")
    End If
    
    Exit Sub
End If

'*** Retrieve the attributes of the Defining Part, if required
If bRetrieveFromWebService = True Then
    Call GetAttributesOfDocRevision(sAttributes, sPartNumber)
End If

'*** Sort attribut
sAttValues(1) = sType & sPartNumber

'*** Qty
If sSide = "LHS" Then
    sAttValues(2) = 0: sAttValues(3) = 1
Else
    sAttValues(2) = 1: sAttValues(3) = 0
End If

'*** Part Number
sAttValues(5) = sPartNumber

'*** Cage Code
If sMaterialSpec <> "" Then
    sAttValues(4) = sConsumableCageCode.GetItemByKey(sMaterialSpec)
ElseIf bRetrieveFromWebService = False Then
    On Error Resume Next
    Set oStrParam = Nothing
    Set oStrParam = oParameters.GetItem("Supplier Name And CAGE Code")
    If Not oStrParam Is Nothing Then
        If oStrParam.Value = "-" Then
            sAttValues(4) = "-"
        Else
            sAttValues(4) = Split(oStrParam.Value, " ")(UBound(Split(oStrParam.Value, " ")))
        End If
    Else
        sAttValues(4) = ""
    End If
    On Error GoTo 0
Else
    If Not IsEmpty(sAttributes) Then
        If Not sAttributes.Count = 0 Then
            If sAttributes.GetItem("Supplier Name And CAGE Code") = "-" Then
                sAttValues(4) = sAttributes("Supplier Name And CAGE Code")
            Else
                sAttValues(4) = Split(sAttributes.GetItem("Supplier Name And CAGE Code"), " ")(UBound(Split(sAttributes.GetItem("Supplier Name And CAGE Code"), " ")))
            End If
        End If
    Else
        sAttValues(4) = ""
    End If
End If

'*** Title
If bRetrieveFromWebService = False Then
    On Error Resume Next
    Set oStrParam = Nothing
    Set oStrParam = oParameters.GetItem("Title")
    If Not oStrParam Is Nothing Then
        sAttValues(6) = oStrParam.Value
    Else
        sAttValues(6) = ""
    End If
    On Error GoTo 0
Else
    If Not IsEmpty(sAttributes) Then
        If Not sAttributes.Count = 0 Then
            sAttValues(6) = sAttributes.GetItem("Title")
        Else
            sAttValues(6) = "Defining Part not found in Enovia. Attribute can't be extracted."
        End If
    Else
        sAttValues(6) = "Defining Part not found in Enovia. Attribute can't be extracted."
        
        If sLogDump = "" Then
            sLogDump = "Defining Part " & sPartNumber & " not found in Enovia"
        Else
            sLogDump = sLogDump & " | " & "Defining Part " & sPartNumber & " not found in Enovia"
        End If
    End If
End If

'*** Size
If bRetrieveFromWebService = False Then
    On Error Resume Next
    Set oStrParam = Nothing
    Set oStrParam = oParameters.GetItem("Size")
    If Not oStrParam Is Nothing Then
        sAttValues(7) = oStrParam.Value
    Else
        sAttValues(7) = ""
    End If
    On Error GoTo 0
Else
    If Not IsEmpty(sAttributes) Then
        If Not sAttributes.Count = 0 Then
            sAttValues(7) = sAttributes.GetItem("Size")
        End If
    Else
        sAttValues(7) = ""
    End If
End If

'*** Material Specification Production
If bRetrieveFromWebService = False Then
    On Error Resume Next
    Set oStrParam = Nothing
    Set oStrParam = oParameters.GetItem("Material Specification Production")
    If Not oStrParam Is Nothing Then
        sAttValues(8) = oStrParam.Value
    Else
        sAttValues(8) = ""
    End If
    On Error GoTo 0
Else
    If Not IsEmpty(sAttributes) Then
        If Not sAttributes.Count = 0 Then
            sAttValues(8) = sAttributes.GetItem("Material Specification Production")
        End If
    Else
        sAttValues(8) = ""
    End If
End If

'*** Finish Code
If bRetrieveFromWebService = False Then
    On Error Resume Next
    Set oStrParam = Nothing
    Set oStrParam = oParameters.GetItem("Finish Code")
    If Not oStrParam Is Nothing Then
        sAttValues(9) = oStrParam.Value
    Else
        sAttValues(9) = ""
    End If
    On Error GoTo 0
Else
    If Not IsEmpty(sAttributes) Then
        If Not sAttributes.Count = 0 Then
            sAttValues(9) = sAttributes.GetItem("Finish Code")
        End If
    Else
        sAttValues(9) = ""
    End If
End If

'*** PCCN
If bRetrieveFromWebService = False Then
    On Error Resume Next
    Set oStrParam = Nothing
    Set oStrParam = oParameters.GetItem("PCCN")
    If Not oStrParam Is Nothing Then
        sAttValues(10) = oStrParam.Value
    Else
        sAttValues(10) = ""
    End If
    On Error GoTo 0
Else
    If Not IsEmpty(sAttributes) Then
        If Not sAttributes.Count = 0 Then
            sAttValues(10) = sAttributes.GetItem("PCCN")
        End If
    Else
        sAttValues(10) = ""
    End If
End If


'Add node to structure
Set oNode = oData.AddElement("Part", oData.RootNode)
Call oData.AddMultipleAttributes(oNode, sAttNames, sAttValues)
End Sub

Private Sub CheckInstance(ByVal oProduct1 As Product)

Dim sPartNumber As String

'Retrieve Part Number
sPartNumber = ""
On Error Resume Next
sPartNumber = oProduct1.PartNumber

If sPartNumber <> "" Then Exit Sub

'Part Number was not retrieved:
' - Part is not loaded
' - Part is in cgr

'Part in cgr
If oProduct1.HasAMasterShapeRepresentation Then
    oProduct1.ApplyWorkMode DEFAULT_MODE
End If
On Error GoTo 0

'Retrieve Part Number again
On Error Resume Next
sPartNumber = oProduct1.PartNumber
If sPartNumber = "" Then
    MsgBox ("Can't retrieve part number of " & oProduct1.Name)
    bStatus = False
End If

On Error GoTo 0

End Sub

Private Sub GetConsumableCageCode()

Dim iRow As Integer
Dim oLocalExcel As clsExcel

'Open Excel file
Set oLocalExcel = New clsExcel
Call oLocalExcel.GetNewExcel

Call oLocalExcel.OpenExcelFile(sAttributesTemplateFile)

iRow = 2
Do While oLocalExcel.App.Worksheets("Consumable-Compatibility").Cells(iRow, 1).Value <> ""
    
    If Not sConsumableCageCode.Exists(CStr(oLocalExcel.App.Worksheets("Consumable-Compatibility").Cells(iRow, 1).Value)) Then Call sConsumableCageCode.Add(CStr(oLocalExcel.App.Worksheets("Consumable-Compatibility").Cells(iRow, 1).Value), oLocalExcel.App.Worksheets("Consumable-Compatibility").Cells(iRow, 3).Value)
    
    iRow = iRow + 1
Loop

'Close Excel workbook and kill process
Call oLocalExcel.CloseAlreadyOpenWorkbook(sAttributesTemplateFile)
Set oLocalExcel = Nothing

End Sub



Private Sub DisplayInExcel(ByVal oData As clsXML, ByVal oNHA_RHS As Product)

Dim vArray() As Variant
Dim oList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim i As Integer

'Get Excel
Set oExcel = New clsExcel
Call oExcel.GetExcel

'Close already opened file
Call oExcel.CloseAlreadyOpenWorkbook("C:\Temp\DDPToolbar_PL.xls")

'Open Excel file
Call oExcel.OpenExcelFile(sPLTemplateFile)

Set oList = oData.SelectNodes("./Part", oData.RootNode)

'Redim
ReDim vArray(1 To oList.Length, 1 To 12)

'Scan all node
i = 0
For Each oNode In oList
    i = i + 1
    vArray(i, 2) = IIf(oData.GetAttributeValue(oNode, "QtyRHS") = "0", "", oData.GetAttributeValue(oNode, "QtyRHS"))
    vArray(i, 3) = IIf(oData.GetAttributeValue(oNode, "QtyLHS") = "0", "", oData.GetAttributeValue(oNode, "QtyLHS"))
    vArray(i, 6) = oData.GetAttributeValue(oNode, "CageCode")
    vArray(i, 7) = oData.GetAttributeValue(oNode, "PartNumber")
    vArray(i, 8) = oData.GetAttributeValue(oNode, "Description")
    vArray(i, 9) = oData.GetAttributeValue(oNode, "Size")
    vArray(i, 10) = oData.GetAttributeValue(oNode, "MatlSpec")
    vArray(i, 11) = oData.GetAttributeValue(oNode, "FinishCode")
    vArray(i, 12) = oData.GetAttributeValue(oNode, "PCCN")
Next

'Display in Excel
oExcel.App.Range(oExcel.App.Cells(2, 2), oExcel.App.Cells(oList.Length + 1, 13)).Value = vArray

'Run macro in Excel
oExcel.RunMacro ("DeleteTypeColumn")
oExcel.RunMacro ("ColumnsAutoFit")
If oNHA_RHS Is Nothing Then
    oExcel.RunMacro ("InsertBlankLineAfterAssy")
End If
oExcel.SaveWorkbook ("C:\Temp\DDPToolbar_PL.xls")


End Sub








