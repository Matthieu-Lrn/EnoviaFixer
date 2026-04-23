Attribute VB_Name = "CreateFlatPanel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oSketch As Sketch
Private oFactory2d As Factory2D
Private oPanelPart As Part
Private Const dError = 0.001

Public Sub LaunchCreationFlatPanel(ByVal oProductDocument As Document, ByVal oPanelProd As Product, _
                                   ByVal oSketchFace As PlanarFace, ByVal oSketchProd As Product, _
                                   ByVal oThkFace As PlanarFace, ByVal oThkProd As Product, _
                                   ByVal bCreateConstraints As Boolean)

Dim oSketchPart As Part
Dim oThkPart As Part
Dim oSPAWorkbench As Workbench

'Factories
Dim oShapeFactory As ShapeFactory
Dim oHybridShapeFactory As HybridShapeFactory

'Geoset
Dim oTempGeosets As HybridBodies
Dim oGeoset As HybridBody
Dim oDatumPlanesHB As HybridBody
Dim oDatumPointsHB As HybridBody

Dim oHybridShape As HybridShape
Dim oShape As Shape
Dim o2DGeometry As Geometry2D
Dim oPartBody As Body

Dim oReference As Reference
Dim oReference2 As Reference

Dim oVariant 'As Variant
Dim i As Integer
Dim sLog As String

Dim dRadius As Double
Dim dMaxLength As Double
Dim iOutsideContour As Integer

Dim oMeasurable 'As Measurable

Dim oElem As GeometricElement
Dim dElemCoord(3)
Dim oPreviousElem As GeometricElement
Dim dCOGSurface(2)

Dim oAllElemCollection As New Collection
Dim oOrderedCollection As New Collection
Dim oHolesDefinition() As String
Dim dHoleRadius As Double
Dim dHoleCenter(2)
Dim iItem As Integer

'First Elem
Dim oFirstElem As Geometry2D
Dim oFirstElemCoord(3)

'NextPoint
Dim oNextPoint 'As Point2D
Dim dNextPointCoord(1)

'Initial set
Set oSelection = oProductDocument.Selection
Call SwapToAssemblyWorkbench

'SPAWorkbench for Measurable object
Set oSPAWorkbench = oProductDocument.GetWorkbench("SPAWorkbench")

'Define Part object of oPanelProduct
Set oPanelPart = oPanelProd.GetMasterShapeRepresentation(False).Part

'Add temporary Geoset for construction.
'1st level
Set oGeoset = oPanelPart.HybridBodies.Add
oGeoset.Name = "Temp_To_Delete"
Set oTempGeosets = oGeoset.HybridBodies

'2nd level
Set oGeoset = oTempGeosets.Add
oGeoset.Name = "Extract"

Set oGeoset = oTempGeosets.Add
oGeoset.Name = "Boundary"

Set oGeoset = oTempGeosets.Add
oGeoset.Name = "Domains"

Set oGeoset = oTempGeosets.Add
oGeoset.Name = "Circle"

Set oGeoset = oTempGeosets.Add
oGeoset.Name = "Joins"

'Retrieve "Construction_Geometry" geoset or create if it doesn't exist
On Error Resume Next
Set oGeoset = Nothing
Set oGeoset = oPanelPart.HybridBodies.Item("Construction_Geometry")

If oGeoset Is Nothing Then
    Set oGeoset = oPanelPart.HybridBodies.Add
    oGeoset.Name = "Construction_Geometry"
End If
oPanelPart.Update
On Error GoTo 0

'Retrieve "Datum_Planes" geoset or create if it doesn't exist
On Error Resume Next
Set oDatumPlanesHB = Nothing
Set oDatumPlanesHB = oGeoset.HybridBodies.Item("Datum_Planes")

If oDatumPlanesHB Is Nothing Then
    Set oDatumPlanesHB = oGeoset.HybridBodies.Add
    oDatumPlanesHB.Name = "Datum_Planes"
End If
oPanelPart.Update
On Error GoTo 0

'Retrieve "Datum_Points" geoset or create if it doesn't exist
On Error Resume Next
Set oDatumPointsHB = Nothing
Set oDatumPointsHB = oGeoset.HybridBodies.Item("Datum_Points")

If oDatumPointsHB Is Nothing Then
    Set oDatumPointsHB = oGeoset.HybridBodies.Add
    oDatumPointsHB.Name = "Datum_Points"
End If
oPanelPart.Update
On Error GoTo 0

'************************SketchFace**************
Set oSketchPart = oSketchProd.GetMasterShapeRepresentation(False).Part

'Add Geoset to oSketchPart
Set oGeoset = oSketchPart.HybridBodies.Add
oGeoset.Name = "TempToDelete"

'Create Extract on oSketchFace
Set oHybridShapeFactory = oSketchPart.HybridShapeFactory
Set oHybridShape = oHybridShapeFactory.AddNewExtract(oSketchFace)
oHybridShape.PropagationType = 3
oHybridShape.ComplementaryExtract = False
oHybridShape.IsFederated = False
oGeoset.AppendHybridShape oHybridShape
oSketchPart.UpdateObject oGeoset

'Copy/Paste Extract Face to the Panel Part
oSelection.Clear
oSelection.Add oHybridShape
oSelection.Copy
oSelection.Clear
oSelection.Add oTempGeosets.Item(1)
oSelection.PasteSpecial "CATPrtResultWithOutLink"

oPanelPart.UpdateObject oTempGeosets.Item(1)

'Delete Geoset in oSketchPart
Set oReference = oSketchPart.CreateReferenceFromGeometry(oGeoset)
oHybridShapeFactory.DeleteObjectForDatum oReference

'************************ThkFace**************
Set oThkPart = oThkProd.GetMasterShapeRepresentation(False).Part

'Add Geoset to oThkPart
Set oGeoset = oThkPart.HybridBodies.Add
oGeoset.Name = "TempToDelete"

'Create Extract  on oThkFace
Set oHybridShapeFactory = oThkPart.HybridShapeFactory
Set oHybridShape = oHybridShapeFactory.AddNewExtract(oThkFace)
oHybridShape.PropagationType = 3
oHybridShape.ComplementaryExtract = False
oHybridShape.IsFederated = False
oGeoset.AppendHybridShape oHybridShape
oThkPart.UpdateObject oGeoset

'Copy/Paste Extract Face to the Panel Part
oSelection.Clear
oSelection.Add oHybridShape
oSelection.Copy
oSelection.Clear
oSelection.Add oTempGeosets.Item(1)
oSelection.PasteSpecial "CATPrtResultWithOutLink"

oPanelPart.UpdateObject oTempGeosets.Item(1)

'Delete Geoset in oThkPart
Set oReference = oThkPart.CreateReferenceFromGeometry(oGeoset)
oHybridShapeFactory.DeleteObjectForDatum oReference


'Create Sketch Plane
Set oHybridShapeFactory = oPanelPart.HybridShapeFactory
Set oReference = oPanelPart.CreateReferenceFromObject(oTempGeosets.Item(1).HybridShapes.Item(1))
Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
oMeasurable.GetCOG dCOGSurface
Set oHybridShape = oHybridShapeFactory.AddNewPlaneOffset(oPanelPart.OriginElements.PlaneXY, dCOGSurface(2), False)
oDatumPlanesHB.AppendHybridShape oHybridShape
oHybridShape.Name = "SketchPlane"
oPanelPart.UpdateObject oDatumPlanesHB

'Create Thk Plane
Set oHybridShapeFactory = oPanelPart.HybridShapeFactory
Set oReference = oPanelPart.CreateReferenceFromObject(oTempGeosets.Item(1).HybridShapes.Item(2))
Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
oMeasurable.GetCOG dCOGSurface
Set oHybridShape = oHybridShapeFactory.AddNewPlaneOffset(oPanelPart.OriginElements.PlaneXY, dCOGSurface(2), False)
oDatumPlanesHB.AppendHybridShape oHybridShape
oHybridShape.Name = "ThkPlane"
oPanelPart.UpdateObject oDatumPlanesHB

'Create Boundary Curve of the first Extract surface (the one in the sketch plane)
Set oReference = oPanelPart.CreateReferenceFromObject(oTempGeosets.Item(1).HybridShapes.Item(1))
Set oHybridShape = oHybridShapeFactory.AddNewBoundaryOfSurface(oReference)
oTempGeosets.Item(2).AppendHybridShape oHybridShape
oPanelPart.UpdateObject oTempGeosets.Item(2)

'Create 1 curve per domains
Set oReference = oPanelPart.CreateReferenceFromObject(oTempGeosets.Item(2).HybridShapes.Item(1))
oVariant = oHybridShapeFactory.AddNewDatums(oReference)
For i = 0 To UBound(oVariant)
      oTempGeosets.Item(3).AppendHybridShape oVariant(i)
Next
oPanelPart.UpdateObject oTempGeosets.Item(3)

'Check all domains for holes
ReDim oHolesDefinition(0)
For Each oHybridShape In oTempGeosets.Item(3).HybridShapes
    
    'Create Reference of the curve
    Set oReference = oPanelPart.CreateReferenceFromObject(oHybridShape)

    'Retrieve measurable object of the curve
    Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)

    On Error Resume Next
    dHoleRadius = oMeasurable.Radius
    
    'If Err.Number = 0 it means the curve is a circle. I concatenate the X and Y coordinates of the center and the
    'radius in one string
    If Err.Number = 0 Then
        oMeasurable.GetCenter dHoleCenter
        ReDim Preserve oHolesDefinition(UBound(oHolesDefinition) + 1)
        oHolesDefinition(UBound(oHolesDefinition)) = dHoleRadius & "|" & dHoleCenter(0) & "|" & dHoleCenter(1)
        oHybridShapeFactory.DeleteObjectForDatum oReference 'Remove the circle from the "Domains" geoset
    End If
    
    On Error GoTo 0
Next

'Measure length of each curve in "Domains" geoset. The longest one is the outside contour
dMaxLength = 0
For i = 1 To oTempGeosets.Item(3).HybridShapes.Count
    
    'Create Reference of the curve
    Set oReference = oPanelPart.CreateReferenceFromObject(oTempGeosets.Item(3).HybridShapes.Item(i))

    'Retrieve measurable object of the curve
    Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)

    If oMeasurable.Length > dMaxLength Then
        iOutsideContour = i
        dMaxLength = oMeasurable.Length
    End If
Next

'If a project a curve with 1 domain in a sketch the result is a "Used-Edge" which can't be modified.
'If a project a curve with 2 domains in a sketch the result is: lines, circles, points and curves which can be modified
'Let's create a circle that will be joined with each curve in the "Domains" geoset. When the Join object will be
'projected in the sketch, the Designer will be able to manualy modify the result.

'Create center point and dummy circle in Sketch plane
Set oHybridShape = oHybridShapeFactory.AddNewPointOnPlane(oDatumPlanesHB.HybridShapes.Item("SketchPlane"), -1.5 * 25.4, -1.5 * 25.4)
oTempGeosets.Item(4).AppendHybridShape oHybridShape
Set oReference = oPanelPart.CreateReferenceFromObject(oHybridShape)

'The radius of the dummy circle will be 33.7566 mm (1.329in). This will allow us to recognise it later
Set oHybridShape = oHybridShapeFactory.AddNewCircleCtrRad(oReference, oDatumPlanesHB.HybridShapes.Item("SketchPlane"), True, 33.7566)
oTempGeosets.Item(4).AppendHybridShape oHybridShape
oPanelPart.UpdateObject oTempGeosets.Item(4)

'Create Join object between the the Dummy Circle and each Curve of the "Domains" geoset
Set oReference = oPanelPart.CreateReferenceFromObject(oHybridShape) 'Dummy Circle
For i = 1 To oTempGeosets.Item(3).HybridShapes.Count

    Set oReference2 = oPanelPart.CreateReferenceFromObject(oTempGeosets.Item(3).HybridShapes.Item(i))
    Set oHybridShape = oHybridShapeFactory.AddNewJoin(oReference, oReference2)
    oTempGeosets.Item(5).AppendHybridShape oHybridShape
    oHybridShape.SetConnex (False)

Next
oPanelPart.UpdateObject oTempGeosets.Item(5)

'Loop All Joins
Set oPartBody = oPanelPart.Bodies.Item("PartBody")
For Each oHybridShape In oTempGeosets.Item(5).HybridShapes

    'Create Sketch and project the Join
    Set oReference = oPanelPart.CreateReferenceFromObject(oHybridShape)
    Set oSketch = oPartBody.Sketches.Add(oDatumPlanesHB.HybridShapes.Item("SketchPlane"))
    Set oFactory2d = oSketch.OpenEdition
    On Error Resume Next
    Set o2DGeometry = oFactory2d.CreateProjection(oReference)

    'Retrieve all geometric elements making up the closed contour and populate oAllElemCollection
    Set oAllElemCollection = Nothing
    For Each oElem In oSketch.GeometricElements
    
        'If element is a circle retrieve the diameter
        If TypeName(oElem) = "Circle2D" Then dRadius = oElem.Radius

        'If the radius of the circle = 33.7566 (Dummy Circle) then set it to construction geometry
        If TypeName(oElem) = "Circle2D" And Abs(dRadius - 33.7566) < dError Then
            oElem.Construction = True
        
        'Other Elements in the Collection except axis and points
        ElseIf TypeName(oElem) <> "Axis2D" And TypeName(oElem) <> "Point2D" Then
            oAllElemCollection.Add oElem
        End If
    Next

    'Let's start the closed contour sequence with the first element in the oAllElemCollection
    Set oElem = oAllElemCollection.Item(1)
    Set oOrderedCollection = Nothing
    oOrderedCollection.Add oElem
    oAllElemCollection.Remove (1)
    Set oNextPoint = oElem.EndPoint
    oNextPoint.GetCoordinates dNextPointCoord
        
    'Retrieve the right sequence of elements and connect them together
    'Populate oOrderedCollection
    Do While oAllElemCollection.Count <> 0
        iItem = 0

        For Each oElem In oAllElemCollection

            iItem = iItem + 1

            'Retrieve info from current element
            oElem.GetEndPoints dElemCoord

            'Start point of oElem = NextPoint
            If Abs(dElemCoord(0) - dNextPointCoord(0)) < dError And Abs(dElemCoord(1) - dNextPointCoord(1)) < dError Then
                oElem.StartPoint = oNextPoint
                Set oNextPoint = oElem.EndPoint
                oNextPoint.GetCoordinates dNextPointCoord
                oOrderedCollection.Add oElem
                oAllElemCollection.Remove (iItem)
                
                'oElem is the last element
                If oAllElemCollection.Count = 0 Then
                    Set oFirstElem = oOrderedCollection.Item(1)
                    Set oNextPoint = oFirstElem.StartPoint
                    oElem.EndPoint = oNextPoint
                End If
                
                Exit For
            
            'End point of oElem = NextPoint
            ElseIf Abs(dElemCoord(2) - dNextPointCoord(0)) < dError And Abs(dElemCoord(3) - dNextPointCoord(1)) < dError Then
                oElem.EndPoint = oNextPoint
                Set oNextPoint = oElem.StartPoint
                oNextPoint.GetCoordinates dNextPointCoord
                oOrderedCollection.Add oElem
                oAllElemCollection.Remove (iItem)
                
                'oElem is the last element
                If oAllElemCollection.Count = 0 Then
                    Set oFirstElem = oOrderedCollection.Item(1)
                    Set oNextPoint = oFirstElem.StartPoint
                    oElem.StartPoint = oNextPoint
                End If
                
                Exit For
            End If
        Next
    Loop

    'Add Constraints
    For i = 1 To oOrderedCollection.Count
    
        'Retrieve element
        Set oElem = oOrderedCollection.Item(i)
        
        'Previous element
        If oOrderedCollection.Count = 1 Then
            Set oPreviousElem = Nothing
        ElseIf i = 1 Then
            Set oPreviousElem = oOrderedCollection.Item(oOrderedCollection.Count)
        Else
            Set oPreviousElem = oOrderedCollection.Item(i - 1)
        End If
        
        If bCreateConstraints = True Then
            Call CreateConstraint(oElem, oPreviousElem)
        End If
        
    Next
    oPanelPart.UpdateObject oSketch
    oSketch.CloseEdition
    oSketch.Evaluate
    oPanelPart.Update

    On Error GoTo 0
Next


'Create pad
oPanelPart.InWorkObject = oPartBody
Set oShapeFactory = oPanelPart.ShapeFactory
Set oShape = oShapeFactory.AddNewPad(oPartBody.Sketches.Item(iOutsideContour), 25.4)

'Set Limits
Dim oLimit As Limit
Set oLimit = oShape.FirstLimit
Set oReference = oPanelPart.CreateReferenceFromObject(oDatumPlanesHB.HybridShapes.Item("ThkPlane"))
oLimit.LimitMode = catUpToPlaneLimit
oLimit.LimitingElement = oReference

Set oLimit = oShape.SecondLimit
Set oReference = oPanelPart.CreateReferenceFromObject(oDatumPlanesHB.HybridShapes.Item("SketchPlane"))
oLimit.LimitMode = catUpToPlaneLimit
oLimit.LimitingElement = oReference
oPanelPart.UpdateObject oShape

'Create pockets
sLog = ""
If oPartBody.Sketches.Count > 1 Then
    For i = 1 To oPartBody.Sketches.Count
        If i <> iOutsideContour Then
            Set oShape = oShapeFactory.AddNewPocket(oPartBody.Sketches.Item(i), 25.4)
            Set oLimit = oShape.FirstLimit
            oLimit.LimitMode = catUpThruNextLimit
            
            'Check if there is an error when creating the shape. It may be because the sketch is over-constraint
            On Error Resume Next
            oPanelPart.UpdateObject oShape
            
            'Delete shape
            If Err.Number <> 0 Then
            
                If sLog = "" Then
                    sLog = "Problem creating a pocket with " & oPartBody.Sketches.Item(i).Name
                Else
                    sLog = sLog & vbCrLf & "Problem creating a pocket with " & oPartBody.Sketches.Item(i).Name
                End If
                
                oSelection.Clear
                oSelection.Add oShape
                oSelection.Delete
            End If
            
            On Error GoTo 0
        End If
    Next
End If


'Create holes
For i = 1 To UBound(oHolesDefinition)

    'Create point
    Set oHybridShape = oHybridShapeFactory.AddNewPointOnPlane(oDatumPlanesHB.HybridShapes.Item("SketchPlane"), Split(oHolesDefinition(i), "|")(1), Split(oHolesDefinition(i), "|")(2))
    oDatumPointsHB.AppendHybridShape oHybridShape
    oPanelPart.UpdateObject oDatumPointsHB
    
    'Create Hole
    Set oHole = oShapeFactory.AddNewHoleFromRefPoint(oHybridShape, oDatumPlanesHB.HybridShapes.Item("SketchPlane"), 6.25)

    oHole.Type = catSimpleHole
    oHole.AnchorMode = catExtremPointHoleAnchor
    oHole.BottomType = catFlatHoleBottom
    oPanelPart.UpdateObject oHole
    Set HoleDia = oHole.Diameter
    HoleDia.Value = 2 * CDbl(Split(oHolesDefinition(i), "|")(0))

    Set oLimit = oHole.BottomLimit
    Set oReference = oPanelPart.CreateReferenceFromObject(oDatumPlanesHB.HybridShapes.Item("ThkPlane"))
    oLimit.LimitMode = catUpToPlaneLimit
    oLimit.LimitingElement = oReference

    Set oSketch = oHole.Sketch
    Set oReference = oPanelPart.CreateReferenceFromObject(oSketch)
    oHybridShapeFactory.GSMVisibility oReference, 0

    oPanelPart.UpdateObject oHole
Next

'Apply color to PartBody
'Call ApplyColorToBody(oProductDocument, oPartBody, CLng(Split(sMaterialType, "|")(1)), CLng(Split(sMaterialType, "|")(2)), CLng(Split(sMaterialType, "|")(3)))

'Delete temporary Geoset for construction.
Set oReference = oPanelPart.CreateReferenceFromGeometry(oPanelPart.HybridBodies.Item("Temp_To_Delete"))
oHybridShapeFactory.DeleteObjectForDatum oReference

'Delete "Datum_Points" if empty
Set oGeoset = oPanelPart.HybridBodies.Item("Construction_Geometry").HybridBodies.Item("Datum_Points")
If oGeoset.HybridShapes.Count = 0 Then
    Set oReference = oPanelPart.CreateReferenceFromGeometry(oGeoset)
    oHybridShapeFactory.DeleteObjectForDatum oReference
End If

'Update Part
On Error Resume Next
oPanelPart.Update
On Error GoTo 0

'Display log
If sLog <> "" Then
    MsgBox sLog
End If

End Sub


Private Sub SwapToAssemblyWorkbench()

Dim sWkb As String

sWkb = CATIA.GetWorkbenchId
If sWkb <> "Assembly" Then
    CATIA.StartWorkbench "Assembly"
End If

End Sub

Private Sub CreateConstraint(ByVal oElem As Geometry2D, ByVal oPreviousElem As Geometry2D)

Const dError = 0.001

Dim oCst As Constraint
Dim oReference As Reference
Dim oReference2 As Reference
Dim i As Integer

Dim dLineCoord(3)
Dim dCircleCoord(3)
Dim dCommonPointCoord(1)
Dim dCircleCenterCoord(1)
Dim oCircleCenter As Geometry2D
Dim dVectorLine(1)
Dim dVectorCircle(1)
Dim dScalarProduct As Double
Dim dLengthVectorLine As Double
Dim dLengthVectorCircle As Double

Dim bObjectTypeOK As Boolean

Dim dCircle1Coord(3)
Dim dCircle2Coord(3)
Dim dCircle1CenterCoord(1)
Dim dCircle2CenterCoord(1)
Dim dVectorCircle1(1)
Dim dVectorCircle2(1)
Dim dLengthVectorCircle1 As Double
Dim dLengthVectorCircle2 As Double



'Line2D objects : add H and V constraint
If TypeName(oElem) = "Line2D" Then

    oElem.GetEndPoints dLineCoord

    If Abs(dLineCoord(0) - dLineCoord(2)) < dError Then
        Set oReference = oPanelPart.CreateReferenceFromObject(oElem)
        Set oCst = oSketch.Constraints.AddMonoEltCst(catCstTypeVerticality, oReference)
    ElseIf Abs(dLineCoord(1) - dLineCoord(3)) < dError Then
        Set oReference = oPanelPart.CreateReferenceFromObject(oElem)
        Set oCst = oSketch.Constraints.AddMonoEltCst(catCstTypeHorizontality, oReference)
    End If
End If

'Circle2D objects : add Radius constraint
If TypeName(oElem) = "Circle2D" Then
    On Error GoTo 0
    
    'Add point at the center of the circle
    oElem.GetCenter dCircleCenterCoord
    Set oCircleCenter = oFactory2d.CreatePoint(dCircleCenterCoord(0), dCircleCenterCoord(1))
    oElem.CenterPoint = oCircleCenter
    
    'Add radius constraint
    Set oReference = oPanelPart.CreateReferenceFromObject(oElem)
    Set oCst = oSketch.Constraints.AddMonoEltCst(catCstTypeRadius, oReference)
End If

'Tangency between line and circle
bObjectTypeOK = False
If TypeName(oElem) = "Line2D" And TypeName(oPreviousElem) = "Circle2D" Then
    oElem.GetEndPoints dLineCoord
    oPreviousElem.GetEndPoints dCircleCoord
    oPreviousElem.GetCenter dCircleCenterCoord
    bObjectTypeOK = True
    
ElseIf TypeName(oPreviousElem) = "Line2D" And TypeName(oElem) = "Circle2D" Then
    oElem.GetEndPoints dCircleCoord
    oElem.GetCenter dCircleCenterCoord
    oPreviousElem.GetEndPoints dLineCoord
    bObjectTypeOK = True
End If

If bObjectTypeOK = True Then
    
    'Find common end point between the line and circle and calculate the two vectors:
    ' - dVectorLine : From common end point to other end point of line
    ' - dVectorCircle : From common end point to center of the circle
    For i = 0 To 1
    
        If Abs(dLineCoord(0) - dCircleCoord(2 * i)) < dError And Abs(dLineCoord(1) - dCircleCoord(2 * i + 1)) < dError Then
            dCommonPointCoord(0) = dLineCoord(0)
            dCommonPointCoord(1) = dLineCoord(1)
            
            dVectorLine(0) = dLineCoord(2) - dCommonPointCoord(0)
            dVectorLine(1) = dLineCoord(3) - dCommonPointCoord(1)
            
            dVectorCircle(0) = dCircleCenterCoord(0) - dCommonPointCoord(0)
            dVectorCircle(1) = dCircleCenterCoord(1) - dCommonPointCoord(1)
            
            Exit For
        ElseIf Abs(dLineCoord(2) - dCircleCoord(2 * i)) < dError And Abs(dLineCoord(3) - dCircleCoord(2 * i + 1)) < dError Then
            dCommonPointCoord(0) = dLineCoord(2)
            dCommonPointCoord(1) = dLineCoord(3)
            
            dVectorLine(0) = dLineCoord(0) - dCommonPointCoord(0)
            dVectorLine(1) = dLineCoord(1) - dCommonPointCoord(1)
            
            dVectorCircle(0) = dCircleCenterCoord(0) - dCommonPointCoord(0)
            dVectorCircle(1) = dCircleCenterCoord(1) - dCommonPointCoord(1)
           
            Exit For
        End If

    Next

    'Calculate angle between the two vectors using: Scalar Product / Length(dVectorLine) / Length(dVectorCircle)
    dScalarProduct = dVectorLine(0) * dVectorCircle(0) + dVectorLine(1) * dVectorCircle(1)
    dLengthVectorLine = Sqr(dVectorLine(0) * dVectorLine(0) + dVectorLine(1) * dVectorLine(1))
    dLengthVectorCircle = Sqr(dVectorCircle(0) * dVectorCircle(0) + dVectorCircle(1) * dVectorCircle(1))
    
    dScalarProduct = dScalarProduct / dLengthVectorLine
    dScalarProduct = dScalarProduct / dLengthVectorCircle

    'If the dScalarProduct = 0 than the angle is 90deg and the line and circle are tangent
    'Create tangency cst
    If Abs(dScalarProduct) < dError Then
        Set oReference = oPanelPart.CreateReferenceFromObject(oElem)
        Set oReference2 = oPanelPart.CreateReferenceFromObject(oPreviousElem)
        Set oCst = oSketch.Constraints.AddBiEltCst(catCstTypeTangency, oReference, oReference2)
    End If
       
End If


'Tangency between two circles
bObjectTypeOK = False
If TypeName(oElem) = "Circle2D" And TypeName(oPreviousElem) = "Circle2D" Then
    oElem.GetEndPoints dCircle1Coord
    oPreviousElem.GetEndPoints dCircle2Coord
    
    oElem.GetCenter dCircle1CenterCoord
    oPreviousElem.GetCenter dCircle2CenterCoord
    
    bObjectTypeOK = True
End If

If bObjectTypeOK = True Then

    'Find common end point between the 2 circles
    For i = 0 To 1
        If Abs(dCircle1Coord(0) - dCircle2Coord(2 * i)) < dError And Abs(dCircle1Coord(1) - dCircle2Coord(2 * i + 1)) < dError Then
            dCommonPointCoord(0) = dCircle1Coord(0)
            dCommonPointCoord(1) = dCircle1Coord(1)
            Exit For
        
        ElseIf Abs(dCircle1Coord(2) - dCircle2Coord(2 * i)) < dError And Abs(dCircle1Coord(3) - dCircle2Coord(2 * i + 1)) < dError Then
            dCommonPointCoord(0) = dCircle1Coord(2)
            dCommonPointCoord(1) = dCircle1Coord(3)
            Exit For
       End If
    Next
    
    'Calculate the two vectors:
    ' - dVectorCircle1 : From common end point to center of the circle1
    ' - dVectorCircle2 : From common end point to center of the circle2
    dVectorCircle1(0) = dCircle1CenterCoord(0) - dCommonPointCoord(0)
    dVectorCircle1(1) = dCircle1CenterCoord(1) - dCommonPointCoord(1)
    
    dVectorCircle2(0) = dCircle2CenterCoord(0) - dCommonPointCoord(0)
    dVectorCircle2(1) = dCircle2CenterCoord(1) - dCommonPointCoord(1)

    'Calculate angle between the two vectors using: Scalar Product / Length(dVectorLine) / Length(dVectorCircle)
    dScalarProduct = dVectorCircle1(0) * dVectorCircle2(0) + dVectorCircle1(1) * dVectorCircle2(1)
    dLengthVectorCircle1 = Sqr(dVectorCircle1(0) * dVectorCircle1(0) + dVectorCircle1(1) * dVectorCircle1(1))
    dLengthVectorCircle2 = Sqr(dVectorCircle2(0) * dVectorCircle2(0) + dVectorCircle2(1) * dVectorCircle2(1))
    
    dScalarProduct = dScalarProduct / dLengthVectorCircle1
    dScalarProduct = dScalarProduct / dLengthVectorCircle2

    'If the dScalarProduct =
    ' -1 than the angle is 180 deg and the two circles are tangent
    ' 1 than the angle is 0 deg and the two circles are tangent
    'Create tangency cst
    If Abs(dScalarProduct) - 1 < dError Then
        Set oReference = oPanelPart.CreateReferenceFromObject(oElem)
        Set oReference2 = oPanelPart.CreateReferenceFromObject(oPreviousElem)
        Set oCst = oSketch.Constraints.AddBiEltCst(catCstTypeTangency, oReference, oReference2)
    End If

End If


End Sub


Public Sub ApplyColorToBody(ByVal oRootDocument As Document, ByVal obody As Body, ByVal lR As Long, ByVal lG As Long, ByVal lB As Long)

Set oSelection = oRootDocument.Selection
oSelection.Clear

oSelection.Add obody
Set oVisProperties = oSelection.VisProperties
oVisProperties.SetRealColor lR, lG, lB, 0

End Sub










