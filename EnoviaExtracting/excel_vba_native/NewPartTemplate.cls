Attribute VB_Name = "NewPartTemplate"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oExcel As clsExcel

Public Function InsertNewPartTemplate(ByVal oNHAProduct As Product, ByVal sPartNumber As String, _
                                      ByVal sTitle As String, ByVal sNewDocType As String, _
                                      ByVal iColumn As Integer, ByVal sExcelFileManagement As String) As Product
                                      
                                      
Dim oProductDoc As ProductDocument
Dim oAssyWindow As SpecsAndGeomWindow
Dim oNewProductWindow As SpecsAndGeomWindow
Dim oNewDocument As Document
Dim oNewProduct As Product
Dim iRow As Integer

'Define the selection of the active window
Set oProductDoc = CATIA.ActiveDocument
Set oSelection = oProductDoc.Selection

'Swap to Assembly wkb
Call SwapToAssemblyWorkbench

'On conserve la fenetre de l'assemblage
Set oAssyWindow = CATIA.ActiveWindow

'New document
If sNewDocType = "CATPart" Then
    Set oNewDocument = CATIA.Documents.NewFrom(sCATPartTemplateFile)
ElseIf sNewDocType = "CATProduct" Then
    Set oNewDocument = CATIA.Documents.NewFrom(sCATProductTemplateFile)
End If
Set oNewProduct = oNewDocument.Product
Set oNewProductWindow = CATIA.ActiveWindow
oNewProduct.PartNumber = sPartNumber

'On copie la part
Set oSelection = oNewDocument.Selection
oSelection.Clear
oSelection.Add oNewProduct
oSelection.Copy

'On la paste dans le NHA
Set oSelection = oProductDoc.Selection
oSelection.Clear
oSelection.Add oNHAProduct
oSelection.Paste

'On ferme la fenêtre et le document
oNewProductWindow.Close
oNewDocument.Close

'Change the Title attribute
Set oNewProduct = oNHAProduct.Products.Item(oNHAProduct.Products.Count)
On Error Resume Next
Set oParameters = oNewProduct.ReferenceProduct.UserRefProperties
Set oStrParam = oParameters.GetItem("Title")
oStrParam.Value = sTitle
On Error GoTo 0

'Open Excel file
If sExcelFileManagement Like "*Open*" Then
    Set oExcel = New clsExcel
    Call oExcel.GetNewExcel
    oExcel.App.ScreenUpdating = False
    oExcel.App.DisplayAlerts = False
    Call oExcel.OpenExcelFile(sAttributesTemplateFile)
End If

'Change other attributs (value from Excel file)
iRow = 3
Do While oExcel.App.Cells(iRow, 1).Value <> ""

    Call SetAttributValue(iRow, iColumn, oParameters)
    iRow = iRow + 1
Loop

'Set Defining Part Attribute
Call SetDefiningPart(oParameters, oNewProduct.PartNumber)

'Close Excel workbook and kill process
If sExcelFileManagement Like "*Close*" Then
    Call oExcel.CloseAlreadyOpenWorkbook(sAttributesTemplateFile)
    Set oExcel = Nothing
End If

Set InsertNewPartTemplate = oNewProduct

End Function
Public Sub PositionPartAtOriginNHA(ByVal oNewProduct As Product)

Dim oProdPosition 'As Position
Dim dPositionMatrix(11) As Variant

'Position Matrix
dPositionMatrix(0) = 1
dPositionMatrix(1) = 0
dPositionMatrix(2) = 0
dPositionMatrix(3) = 0
dPositionMatrix(4) = 1
dPositionMatrix(5) = 0
dPositionMatrix(6) = 0
dPositionMatrix(7) = 0
dPositionMatrix(8) = 1
dPositionMatrix(9) = 0
dPositionMatrix(10) = 0
dPositionMatrix(11) = 0

'Set Position
Set oProdPosition = oNewProduct.Position
oProdPosition.SetComponents dPositionMatrix

End Sub
Public Sub PositionPartToPart(ByVal oNewProduct As Product, ByVal oNewProductNHA As Product, ByVal oRefProduct As Product)
    Call SnapPartToPart(oNewProduct, oNewProductNHA, oRefProduct)
End Sub
Public Sub PositionPartWithGeometry(ByVal oNewProduct As Product, ByVal oZ0Face As PlanarFace, oZ0Prod As Product, _
                                    ByVal oZPositiveFace As PlanarFace, ByVal oZPositiveProd As Product, _
                                    ByVal oYAxisFace As Face, ByVal oYAxisProd As Product, _
                                    ByVal oOriginFace As Face, ByVal oOriginProd As Product)

Dim oNewPart As Part
Dim oNewHB As HybridBody
Dim oHybridShape As HybridShape
Dim oHybridShapeFactory As HybridShapeFactory
Dim oHybridShapeDirection As HybridShapeDirection
Dim oProdPosition 'As Position
Dim dOrigin(2)
Dim dX(2)
Dim dY(2)
Dim dZ(2)
Dim dPanelPosition(11) As Variant
Dim oReference As Reference

'Retrieve Part
Set oNewPart = oNewProduct.GetMasterShapeRepresentation(False).Part

'Add Geoset to New Panel Part
Set oNewHB = oNewPart.HybridBodies.Add
oNewHB.Name = "Temp_To_Delete"

'Create Extract Surfaces
Call ExtractFace(oZ0Prod, oNewPart, oZ0Face)
Call ExtractFace(oZPositiveProd, oNewPart, oZPositiveFace)
Call ExtractFace(oYAxisProd, oNewPart, oYAxisFace)
Call ExtractFace(oOriginProd, oNewPart, oOriginFace)

'Create planes on planar surfaces: oZ0Prod and oZPositiveProd
Set oHybridShapeFactory = oNewPart.HybridShapeFactory
For i = 1 To 2
    Set oHybridShape = oHybridShapeFactory.AddNewPlaneOffset(oNewHB.HybridShapes.Item(i), 0, True)
    oNewHB.AppendHybridShape oHybridShape
    oHybridShape.Name = "Plane" & i
Next
oNewPart.UpdateObject oNewHB

'Create planes tangent to surface: oYAxisProd and oOriginProd
For i = 3 To 4
    
    'Create point on surface
    Set oHybridShapeDirection = oHybridShapeFactory.AddNewDirectionByCoord(0, 0, 1)
    Set oHybridShape = oHybridShapeFactory.AddNewPointOnSurface(oNewHB.HybridShapes.Item(i), oHybridShapeDirection, 0)
    oNewHB.AppendHybridShape oHybridShape
    oHybridShape.Name = "Point" & i
    oNewPart.UpdateObject oNewHB
    
    'Create plane tangent to surface
    Set oHybridShape = oHybridShapeFactory.AddNewPlaneTangent(oNewHB.HybridShapes.Item(i), oHybridShape)
    oNewHB.AppendHybridShape oHybridShape
    oHybridShape.Name = "Plane" & i
    oNewPart.UpdateObject oNewHB
Next

'Intersect between Z0 plane and Yaxis plane
Set oHybridShape = oHybridShapeFactory.AddNewIntersection(oNewHB.HybridShapes.Item(5), oNewHB.HybridShapes.Item(8))
oNewHB.AppendHybridShape oHybridShape
oHybridShape.Name = "Intersect_Yaxis"

'Intersect between Z0 plane and Origin plane
Set oHybridShape = oHybridShapeFactory.AddNewIntersection(oNewHB.HybridShapes.Item(5), oNewHB.HybridShapes.Item(10))
oNewHB.AppendHybridShape oHybridShape
oHybridShape.Name = "Intersect_2"
oNewPart.UpdateObject oNewHB

'Origin
Set oHybridShape = oHybridShapeFactory.AddNewIntersection(oNewHB.HybridShapes.Item(11), oNewHB.HybridShapes.Item(12))
oNewHB.AppendHybridShape oHybridShape
oHybridShape.Name = "Intersect_Origin"
oNewPart.UpdateObject oNewHB
Set oReference = oNewPart.CreateReferenceFromObject(oHybridShape)
Set oHybridShape = oHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, oReference)
oNewHB.AppendHybridShape oHybridShape
oHybridShape.Name = "Point_Origin"
oNewPart.UpdateObject oNewHB
oHybridShape.GetCoordinates dOrigin  'GetCoordinates only works on point and not on intersect or translate

'Point on Z axis
Set oHybridShape = oHybridShapeFactory.AddNewProject(oNewHB.HybridShapes.Item(14), oNewHB.HybridShapes.Item(6))
oNewHB.AppendHybridShape oHybridShape
oHybridShape.Name = "Point_Project_on_Zaxis"
oNewPart.UpdateObject oNewHB
Set oReference = oNewPart.CreateReferenceFromObject(oHybridShape)
Set oHybridShape = oHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, oReference)
oHybridShape.Name = "Point_on_Zaxis"
oNewHB.AppendHybridShape oHybridShape
oNewPart.UpdateObject oNewHB
oHybridShape.GetCoordinates dZ    'GetCoordinates only works on point and not on intersect or translate

'Point on X axis
Set oDirection = oHybridShapeFactory.AddNewDirection(oNewHB.HybridShapes.Item(8))
Set oHybridShape = oHybridShapeFactory.AddNewTranslate(oNewHB.HybridShapes.Item(14), oDirection, 25.4)
oHybridShape.Name = "Point_Translate"
oNewHB.AppendHybridShape oHybridShape
oNewPart.UpdateObject oNewHB
Set oReference = oNewPart.CreateReferenceFromObject(oHybridShape)
Set oHybridShape = oHybridShapeFactory.AddNewProject(oReference, oNewHB.HybridShapes.Item(5))
oNewHB.AppendHybridShape oHybridShape
oHybridShape.Name = "Point_Project"
oNewPart.UpdateObject oNewHB
Set oReference = oNewPart.CreateReferenceFromObject(oHybridShape)
Set oHybridShape = oHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, oReference)
oNewHB.AppendHybridShape oHybridShape
oHybridShape.Name = "Point_on_Xaxis"
oNewPart.UpdateObject oNewHB
oHybridShape.GetCoordinates dX  'GetCoordinates only works on point and not on intersect or translate

'Delete Geoset in oNewPart
Set oReference = oNewPart.CreateReferenceFromGeometry(oNewHB)
oHybridShapeFactory.DeleteObjectForDatum oReference
oNewPart.Update

'Calculate position matrix of oNewPanelProd using the different points we just create
'X axis
dX(0) = dX(0) - dOrigin(0)
dX(1) = dX(1) - dOrigin(1)
dX(2) = dX(2) - dOrigin(2)

dLength = Sqr(dX(0) * dX(0) + dX(1) * dX(1) + dX(2) * dX(2))
dX(0) = dX(0) / dLength
dX(1) = dX(1) / dLength
dX(2) = dX(2) / dLength

'Z axis
dZ(0) = dZ(0) - dOrigin(0)
dZ(1) = dZ(1) - dOrigin(1)
dZ(2) = dZ(2) - dOrigin(2)

dLength = Sqr(dZ(0) * dZ(0) + dZ(1) * dZ(1) + dZ(2) * dZ(2))
dZ(0) = dZ(0) / dLength
dZ(1) = dZ(1) / dLength
dZ(2) = dZ(2) / dLength

'Y axis - Produit vectoriel (Z x X)
dY(0) = dZ(1) * dX(2) - dZ(2) * dX(1)
dY(1) = -(dZ(0) * dX(2) - dZ(2) * dX(0))
dY(2) = dZ(0) * dX(1) - dZ(1) * dX(0)

'Transfert to Matrix
dPanelPosition(0) = dX(0)
dPanelPosition(1) = dX(1)
dPanelPosition(2) = dX(2)
dPanelPosition(3) = dY(0)
dPanelPosition(4) = dY(1)
dPanelPosition(5) = dY(2)
dPanelPosition(6) = dZ(0)
dPanelPosition(7) = dZ(1)
dPanelPosition(8) = dZ(2)
dPanelPosition(9) = dOrigin(0)
dPanelPosition(10) = dOrigin(1)
dPanelPosition(11) = dOrigin(2)

'Panel Position
Set oProdPosition = oNewProduct.Position
oProdPosition.SetComponents dPanelPosition

End Sub

Private Sub SwapToAssemblyWorkbench()

Dim sWkb As String

sWkb = CATIA.GetWorkbenchId
If sWkb <> "Assembly" Then
    CATIA.StartWorkbench "Assembly"
End If

End Sub

Private Sub ExtractFace(ByVal oSourceProd As Product, ByVal oTargetPart As Part, ByVal oFace As Face)

Dim oSourcePart As Part
Dim oGeoset As HybridBody
Dim oHybridShapeFactory As HybridShapeFactory
Dim oHybridShape As HybridShape
Dim oReference As Reference

Set oSourcePart = oSourceProd.GetMasterShapeRepresentation(False).Part

'Add Geoset to oSourcePart
Set oGeoset = oSourcePart.HybridBodies.Add
oGeoset.Name = "Temp_To_Delete"

'Create Extract Face on oFace
Set oHybridShapeFactory = oSourcePart.HybridShapeFactory
Set oHybridShape = oHybridShapeFactory.AddNewExtract(oFace)
oHybridShape.PropagationType = 3
oHybridShape.ComplementaryExtract = False
oHybridShape.IsFederated = False
oGeoset.AppendHybridShape oHybridShape
oSourcePart.UpdateObject oGeoset

'Copy/Paste Extract Face to the Panel Part
oSelection.Clear
oSelection.Add oHybridShape
oSelection.Copy
oSelection.Clear
oSelection.Add oTargetPart.HybridBodies.Item("Temp_To_Delete")
oSelection.PasteSpecial "CATPrtResultWithOutLink"

oTargetPart.UpdateObject oTargetPart.HybridBodies.Item("Temp_To_Delete")

'Delete Geoset in oSourcePart
Set oReference = oSourcePart.CreateReferenceFromGeometry(oGeoset)
oHybridShapeFactory.DeleteObjectForDatum oReference

End Sub


Private Sub SetAttributValue(ByVal iRow As Integer, ByVal iColumn As Integer, ByVal oParameters As Parameters)

Dim oStrParam
Dim oRealParam

'All the conditions where we don't want to edit the attribut
If oExcel.App.Cells(iRow, 2).Value <> "CATIA" Then Exit Sub
If oExcel.App.Cells(iRow, iColumn).Value Like "NO_CHECK*" Then Exit Sub
If oExcel.App.Cells(iRow, iColumn).Value Like "NOT:*" Then Exit Sub


'Retrieve StrParam
On Error Resume Next
Set oStrParam = Nothing
Set oStrParam = oParameters.GetItem(oExcel.App.Cells(iRow, 1).Value)
If Not oStrParam Is Nothing Then
    oStrParam.Value = Split(oExcel.App.Cells(iRow, iColumn).Value, ";")(0)
    If oStrParam.Value = "Blank" Then oStrParam.Value = ""
    Exit Sub
End If

'StrParam not found. Look for oRealParam
Set oRealParam = Nothing
Set oRealParam = oParameters.GetItem(oExcel.App.Cells(iRow, 1).Value)
If Not oRealParam Is Nothing Then
    oRealParam.Value = CDbl(Split(oExcel.App.Cells(iRow, iColumn).Value, ";")(0))
    Exit Sub
End If


End Sub

Private Sub SetDefiningPart(ByVal oParameters As Parameters, ByVal sPartNumber As String)

Dim oDatasetType As StrParam
Dim oDefiningPart As StrParam
Dim sDefiningPart As String

'**Define Defining Part value
'We have a dash number in the Part Number
If UBound(Split(sPartNumber, "-")) >= 1 Then
    sDefiningPart = Split(sPartNumber, "-", 2)(0) & "-" & Left(Split(sPartNumber, "-", 2)(1), 3)

'No dash number in the Part Number. This case should not exit but I don't want the macro to crash if it happens
Else
    sDefiningPart = sPartNumber
End If


'Get Dataset Type
On Error Resume Next
Set oDatasetType = Nothing
Set oDatasetType = oParameters.GetItem("Dataset Type")
On Error GoTo 0

'Set Defining Part attribute when
' - Dataset Type = SKELETON
' - Dataset Type = CATALOG LIGHT VERSION
' - Dataset Type = FLEXIBLE REPRESENTATION
' - Dataset Type = SPLIT PART
' - Part Number like "*KIN*"
If oDatasetType.Value = "SKELETON" Or oDatasetType.Value = "CATALOG LIGHT VERSION" Or _
    oDatasetType.Value = "FLEXIBLE REPRESENTATION" Or oDatasetType.Value = "SPLIT PART" Or _
    sPartNumber Like "*KIN*" Then
    
    'Get Defining Part attribute
    On Error Resume Next
    Set oDefiningPart = Nothing
    Set oDefiningPart = oParameters.GetItem("Defining Part")
    On Error GoTo 0

    'Set Defining Part attribute to first 13 characters of Part Number
    If Not oDefiningPart Is Nothing Then

        oDefiningPart.ValuateFromString (sDefiningPart)
    End If
End If

End Sub


