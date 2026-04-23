Attribute VB_Name = "clsHardwareConstruction"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oProduct As Product
Private oPart As Part
Private oHybridBody As HybridBody
Private sParentList As String
Private oHybridShapeFactory As HybridShapeFactory
Private sInvertOrientation As String

'Product
Public Property Set Product(ByVal oValue As Product)
Set oProduct = oValue
End Property

Public Property Get Product() As Product
Set Product = oProduct
End Property

'Part
Public Property Get Part() As Part
Set Part = oPart
End Property

'HybridBody
Public Property Get HybridBody() As HybridBody
    Set HybridBody = oHybridBody
End Property

Public Property Set HybridBody(ByVal oValue As AnyObject)
    Set oHybridBody = oValue
End Property

Public Sub Initialize()

'Retrieve Part
Set oPart = oProduct.ReferenceProduct.Parent.Part

'Retrieve HybridShapeFactory
Set oHybridShapeFactory = oPart.HybridShapeFactory

'Create new Geoset if it doesn't already exist
On Error Resume Next
Err.Clear
Set oHybridBody = oPart.HybridBodies.Item("_DummyGeoset")
If Err.Number <> 0 Then
    Set oHybridBody = oPart.HybridBodies.Add
    oHybridBody.Name = "_DummyGeoset"

End If

'Set orientation
sInvertOrientation = ""

End Sub

'Parent list
Public Property Get ParentList() As String
     ParentList = sParentList
End Property

Public Property Let ParentList(ByVal oValue As String)
    sParentList = oValue
End Property

Public Property Get InvertOrientation() As String
    InvertOrientation = sInvertOrientation
End Property

Public Property Let InvertOrientation(ByVal oValue As String)
    sInvertOrientation = oValue
End Property

Public Sub CreateCenterLineFromCylinder(ByVal oSelectedFace As AnyObject)

Dim oArray(2) As Variant
Dim oPt As HybridShapePointCoord
Dim oDirection As HybridShapeDirection
Dim oLine As HybridShapeLinePtDir

'Create origin point
oSelectedFace.GetOrigin oArray
Set oPt = oHybridShapeFactory.AddNewPointCoord(oArray(0), oArray(1), oArray(2))
oHybridBody.AppendHybridShape oPt
oPt.Name = "_OriginPt"
oPart.Update

'Create centerline
oSelectedFace.GetDirection oArray
Set oDirection = oHybridShapeFactory.AddNewDirectionByCoord(oArray(0), oArray(1), oArray(2))
Set oLine = oHybridShapeFactory.AddNewLinePtDir(oPt, oDirection, -25.4, 25.4, False)
oHybridBody.AppendHybridShape oLine
oLine.Name = "_Centerline"
oPart.Update

End Sub

Public Sub CreateExtractSurface(ByVal oSelectedFace As AnyObject, ByVal sExtractName As String)

Dim oExtractFace As HybridShapeExtract

Set oExtractFace = oHybridShapeFactory.AddNewExtract(oSelectedFace)
oExtractFace.PropagationType = 3
oExtractFace.ComplementaryExtract = False
oExtractFace.IsFederated = False
oHybridBody.AppendHybridShape oExtractFace
oExtractFace.Name = sExtractName

End Sub


Public Sub CreatePlaneOnPlanarFace(ByVal oSelectedFace As AnyObject, ByVal sPlaneName As String)

Dim oPlaneOffset As HybridShapePlaneOffset

Set oPlaneOffset = oHybridShapeFactory.AddNewPlaneOffset(oSelectedFace, 0#, True)
oHybridBody.AppendHybridShape oPlaneOffset
oPlaneOffset.Name = sPlaneName

End Sub


Public Function CreateIntersectionPoint(ByVal sSurface As String) As Boolean

Dim oLine As HybridShape
Dim oSurface As HybridShape
Dim oIntersect As HybridShapeIntersection
Dim oPoint As HybridShapePointCoord
Dim oReference As Reference
Dim oFastenerHeadOriginPoint As HybridShapePointCoord
Dim bError As Boolean
Dim sResult As String

'Initial Set
Set oLine = oHybridBody.HybridShapes.Item("_Centerline")
Set oSurface = oHybridBody.HybridShapes.Item(sSurface)
bError = False

'Create intersect
Set oIntersect = oHybridShapeFactory.AddNewIntersection(oLine, oSurface)
oIntersect.PointType = 0
oHybridBody.AppendHybridShape oIntersect
oIntersect.Name = "_Intersect"

'Check if we have an intersect
Err.Clear
On Error Resume Next
oPart.UpdateObject oIntersect

If Err.Number <> 0 Then
    oHybridShapeFactory.DeleteObjectForDatum oIntersect
    sResult = MsgBox("The selected surface doesn't intersect the centerline", vbExclamation)
    bError = True
End If
On Error GoTo 0

CreateIntersectionPoint = bError

End Function
    
Public Sub CalculatePositionMatrix(ByRef oMatrix As Variant, ByVal sID As String)

Dim oIntersect As HybridShape
Dim oPlaneTangent As HybridShapePlaneTangent
Dim oSurface As HybridShape
Dim oNormalLine As HybridShape

Dim oHybridShape As HybridShape
Dim dOrigin(2)
Dim dX(2)
Dim dY(2)
Dim dZ(2)


'Initial Set
Set oIntersect = oHybridBody.HybridShapes.Item("_Intersect")
Set oSurface = oHybridBody.HybridShapes.Item("_Surface" & sID)

'Origin - I need to create a Point on top of the intersect to be able to use .GetCoordinates
Set oHybridShape = oHybridShapeFactory.AddNewPointCoordWithReference(0, 0, 0, oIntersect)
oHybridBody.AppendHybridShape oHybridShape
oHybridShape.Name = "_Origin"
oPart.UpdateObject oHybridShape
oHybridShape.GetCoordinates dOrigin

'Create Tangent Plane
Set oPlaneTangent = oHybridShapeFactory.AddNewPlaneTangent(oSurface, oIntersect)
oHybridBody.AppendHybridShape oPlaneTangent
oPlaneTangent.Name = "_TangentPlane"
oPart.Update

'Create Normal Line
Set oNormalLine = oHybridShapeFactory.AddNewLineNormal(oSurface, oIntersect, 0, 30, False)
oHybridBody.AppendHybridShape oNormalLine
oNormalLine.Name = "_NormalLine"
oPart.Update

'Point on Z axis
Set oHybridShape = oHybridShapeFactory.AddNewPointOnCurveWithReferenceFromDistance(oNormalLine, oIntersect, 25.4, False)
oHybridShape.DistanceType = 1
oHybridBody.AppendHybridShape oHybridShape
oHybridShape.Name = "_ZAxis"
oPart.UpdateObject oHybridShape
oHybridShape.GetCoordinates dZ

'Point on X axis
Set oHybridShape = oHybridShapeFactory.AddNewPointOnPlaneWithReference(oPlaneTangent, oIntersect, 25.4, 0#)
oHybridBody.AppendHybridShape oHybridShape
oHybridShape.Name = "_XAxis"
oPart.UpdateObject oHybridShape
oHybridShape.GetCoordinates dX
oPart.Update

'X vector
dX(0) = dX(0) - dOrigin(0)
dX(1) = dX(1) - dOrigin(1)
dX(2) = dX(2) - dOrigin(2)

dLength = Sqr(dX(0) * dX(0) + dX(1) * dX(1) + dX(2) * dX(2))
dX(0) = dX(0) / dLength
dX(1) = dX(1) / dLength
dX(2) = dX(2) / dLength

'Z Vector
dZ(0) = dZ(0) - dOrigin(0)
dZ(1) = dZ(1) - dOrigin(1)
dZ(2) = dZ(2) - dOrigin(2)

dLength = Sqr(dZ(0) * dZ(0) + dZ(1) * dZ(1) + dZ(2) * dZ(2))
dZ(0) = dZ(0) / dLength
dZ(1) = dZ(1) / dLength
dZ(2) = dZ(2) / dLength

'Y Vector - Produit vectoriel (Z x X)
dY(0) = dZ(1) * dX(2) - dZ(2) * dX(1)
dY(1) = -(dZ(0) * dX(2) - dZ(2) * dX(0))
dY(2) = dZ(0) * dX(1) - dZ(1) * dX(0)

'Transfer to Matrix
oMatrix(0) = dX(0)
oMatrix(1) = dX(1)
oMatrix(2) = dX(2)
oMatrix(3) = dY(0)
oMatrix(4) = dY(1)
oMatrix(5) = dY(2)
oMatrix(6) = dZ(0)
oMatrix(7) = dZ(1)
oMatrix(8) = dZ(2)
oMatrix(9) = dOrigin(0)
oMatrix(10) = dOrigin(1)
oMatrix(11) = dOrigin(2)

End Sub

