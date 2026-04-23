Attribute VB_Name = "clsFinishTargetPart"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oProduct As Product
Private oPart As Part
Private sParentList As String
Private obody As Body
Private oRefGeoset As HybridBody
Private oHybridShapeFactory As HybridShapeFactory
Private oShapeFactory As ShapeFactory

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

Public Property Get RefGeoset() As HybridBody
Set RefGeoset = oRefGeoset
End Property
'Parent list
Public Property Get ParentList() As String
     ParentList = sParentList
End Property

Public Property Let ParentList(ByVal oValue As String)
    sParentList = oValue
End Property

Public Sub Initialize()

'Retrieve Part
Set oPart = oProduct.ReferenceProduct.Parent.Part

'Retrieve HybridShapeFactory
Set oHybridShapeFactory = oPart.HybridShapeFactory

'Retrieve ShapeFactory
Set oShapeFactory = oPart.ShapeFactory

End Sub

Public Sub AddNewFinishBody(ByVal iID As Integer)

'Add body and geoset
Set obody = oPart.Bodies.Add
obody.Name = "FinishBody_" & iID
Set oRefGeoset = obody.HybridBodies.Add
oRefGeoset.Name = "Reference"
oPart.Update

End Sub
    
Public Sub AddContructionGeoset()

Dim oConstructionGeoset As HybridBody
Dim oGeosetChild As HybridBody

'Add Construction Geoset
Set oConstructionGeoset = oRefGeoset.HybridBodies.Add
oConstructionGeoset.Name = "Construction_" & oRefGeoset.HybridBodies.Count
oPart.Update

'Add ExtractFace, Boundary, Domains, InsertHoleFace and Join geoset
Set oGeosetChild = oConstructionGeoset.HybridBodies.Add
oGeosetChild.Name = "ExtractedFace_" & oRefGeoset.HybridBodies.Count
Set oGeosetChild = oConstructionGeoset.HybridBodies.Add
oGeosetChild.Name = "Boundary_" & oRefGeoset.HybridBodies.Count
Set oGeosetChild = oConstructionGeoset.HybridBodies.Add
oGeosetChild.Name = "Domains_" & oRefGeoset.HybridBodies.Count
Set oGeosetChild = oConstructionGeoset.HybridBodies.Add
oGeosetChild.Name = "FillSurfaces_" & oRefGeoset.HybridBodies.Count
Set oGeosetChild = oConstructionGeoset.HybridBodies.Add
oGeosetChild.Name = "Join_" & oRefGeoset.HybridBodies.Count

oPart.Update

End Sub
    
Public Sub ManageHoles(ByVal oSPAWorkbench As Workbench, ByVal bDisableInsertHoles As Boolean, ByVal oInsertHole As Variant, _
                            ByVal bDisableSpecificHoles As Boolean, ByVal dLowerLimit As Double, ByVal dUpperLimit As Double)

Dim oConstructionGeoset As HybridBody
Dim oReference As Reference
Dim oReference2 As Reference
Dim oHybridShape As HybridShape
Dim oFill As HybridShape
Dim oVariant As Variant
Dim i As Integer
Dim oMeasurable 'As Measurable
Dim dHoleDiameter As Double
Dim bIsInsertHole As Boolean
Dim bIsSpecificDiaHole As Boolean
Dim oDirection As HybridShapeDirection
Dim oJoin As HybridShapeAssemble

For Each oConstructionGeoset In oRefGeoset.HybridBodies

    'Create Boundary Curve on the selected face
    Set oReference = oPart.CreateReferenceFromObject(oConstructionGeoset.HybridBodies.Item(1).HybridShapes.Item(1))
    Set oHybridShape = oHybridShapeFactory.AddNewBoundaryOfSurface(oReference)
    oConstructionGeoset.HybridBodies.Item(2).AppendHybridShape oHybridShape
    oPart.UpdateObject oConstructionGeoset.HybridBodies.Item(2)
    
    'Create une curve per domain
    Set oReference = oPart.CreateReferenceFromObject(oConstructionGeoset.HybridBodies.Item(2).HybridShapes.Item(1))
    oVariant = oHybridShapeFactory.AddNewDatums(oReference)
    For i = 0 To UBound(oVariant)
          oConstructionGeoset.HybridBodies.Item(3).AppendHybridShape oVariant(i)
    Next
    oPart.UpdateObject oConstructionGeoset.HybridBodies.Item(3)

    'Create a fill surface for each inner circle if diameter:
    ' - Is equal to an Insert Holes
    ' - Is between lower and upper limit
    'Option must be selected by user
    If bDisableInsertHoles = True Or bDisableSpecificHoles = True Then
    
        'Scan all curves in "Domains" geoset
        For Each oHybridShape In oConstructionGeoset.HybridBodies.Item(3).HybridShapes
        
            
            'Create Reference of the curve
            Set oReference = oPart.CreateReferenceFromObject(oHybridShape)
            
            'Retrieve measurable object of the curve
            Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
            
            'Measure hole diameter
            On Error Resume Next
            dHoleDiameter = oMeasurable.Radius * 2
            
            'If Err.Number = 0 then we have a circle
            If Err.Number = 0 Then
                
                'Initial set
                bIsInsertHole = False
                bIsSpecificDiaHole = False
                
                'Insert holes
                If bDisableInsertHoles = True Then
                
                    'Scan to see if dHoleDiameter is found in oInsertHole array
                    For i = LBound(oInsertHole, 1) To UBound(oInsertHole, 1)
                        If Round(dHoleDiameter, 6) = CDbl(oInsertHole(i, 2)) Then
                            bIsInsertHole = True
                            Exit For
                        End If
                    Next
                End If
                
                
                'Specific diameter hole
                If bDisableSpecificHoles = True Then
                
                    'Compare hole diameter with lower and upper limit
                    If Round(dHoleDiameter, 6) >= dLowerLimit And Round(dHoleDiameter, 6) <= dUpperLimit Then
                        bIsSpecificDiaHole = True
                    Else
                        bIsSpecificDiaHole = False
                    End If
                End If
                
                
                'Create the fill surface
                If bIsInsertHole = True Or bIsSpecificDiaHole = True Then
                    
                    'Create Fill surface
                    Set oFill = oHybridShapeFactory.AddNewFill()
                    oFill.AddBound oReference
                    oFill.Continuity = 0
                    
                    'Add to geoset
                    oConstructionGeoset.HybridBodies.Item(4).AppendHybridShape oFill
                    oPart.UpdateObject oConstructionGeoset.HybridBodies.Item(4)

                End If
            End If
            On Error GoTo 0
        Next
    End If
    
    'If we didn't create any Fill surface we just do a translate of 0 of the ExtractFace
    If oConstructionGeoset.HybridBodies.Item(4).HybridShapes.Count = 0 Then
    
        'Reference of Extract Face
        Set oReference = oPart.CreateReferenceFromObject(oConstructionGeoset.HybridBodies.Item(1).HybridShapes.Item(1))

        'Create translate direction
        Set oDirection = oHybridShapeFactory.AddNewDirectionByCoord(1, 0, 0)
        
        'Create Translate
        Set oHybridShape = oHybridShapeFactory.AddNewTranslate(oReference, oDirection, 0)
        oConstructionGeoset.HybridBodies.Item(5).AppendHybridShape oHybridShape
        oPart.UpdateObject oConstructionGeoset.HybridBodies.Item(5)
        
    'Join the extract face with all the Fill surface
    Else
    
        'Reference of Extract Face
        Set oReference = oPart.CreateReferenceFromObject(oConstructionGeoset.HybridBodies.Item(1).HybridShapes.Item(1))
        
        'Reference of the first Fill surface
        Set oReference2 = oPart.CreateReferenceFromObject(oConstructionGeoset.HybridBodies.Item(4).HybridShapes.Item(1))

        'Create Join
        Set oJoin = oHybridShapeFactory.AddNewJoin(oReference, oReference2)

        'Add other Fill surface to the join
        If oConstructionGeoset.HybridBodies.Item(4).HybridShapes.Count > 1 Then
            For i = 2 To oConstructionGeoset.HybridBodies.Item(4).HybridShapes.Count
                Set oReference = oPart.CreateReferenceFromObject(oConstructionGeoset.HybridBodies.Item(4).HybridShapes.Item(i))
                oJoin.AddElement oReference
            Next
            'Join Options
            oJoin.SetConnex 1
            oJoin.SetManifold 0
            oJoin.SetSimplify 1
            oJoin.SetSuppressMode 0
            oJoin.SetDeviation 0.001
            oJoin.SetAngularToleranceMode 0
            oJoin.SetAngularTolerance 0.5
            oJoin.SetFederationPropagation 0
        End If
        
        oConstructionGeoset.HybridBodies.Item(5).AppendHybridShape oJoin
        oPart.UpdateObject oConstructionGeoset.HybridBodies.Item(5)

    End If
        
        
    'Transform the Join in a datum surface
    Set oReference = oPart.CreateReferenceFromObject(oConstructionGeoset.HybridBodies.Item(5).HybridShapes.Item(1))
    oVariant = oHybridShapeFactory.AddNewDatums(oReference)
    oRefGeoset.AppendHybridShape oVariant(0)
    oRefGeoset.HybridShapes.Item(oRefGeoset.HybridShapes.Count).Name = "Face" & oRefGeoset.HybridShapes.Count
    oPart.UpdateObject oConstructionGeoset.HybridBodies.Item(3)

    'Delete construction geoset
    Set oReference = oPart.CreateReferenceFromObject(oConstructionGeoset)
    oHybridShapeFactory.DeleteObjectForDatum oReference
Next



End Sub


Public Sub CreateJoin()

Dim oJoin As HybridShapeAssemble
Dim oReference As Reference
Dim oReference2 As Reference
Dim i As Integer

'Create reference of the first two surfaces
Set oReference = oRefGeoset.HybridShapes.Item(1)
Set oReference2 = oRefGeoset.HybridShapes.Item(2)

Set oJoin = oHybridShapeFactory.AddNewJoin(oReference, oReference2)

If oRefGeoset.HybridShapes.Count > 2 Then
    For i = 3 To oRefGeoset.HybridShapes.Count
        Set oReference = oRefGeoset.HybridShapes.Item(i)
        oJoin.AddElement oReference
    Next

End If


'Options
oJoin.SetConnex 1
oJoin.SetManifold 0
oJoin.SetSimplify 1
oJoin.SetSuppressMode 0
oJoin.SetDeviation 0.001
oJoin.SetAngularToleranceMode 0
oJoin.SetAngularTolerance 0.5
oJoin.SetFederationPropagation 0


oRefGeoset.AppendHybridShape oJoin
oPart.UpdateObject oRefGeoset

End Sub

Public Sub CreateThickSurface(ByVal dThickness As Double)

Dim oThick As ThickSurface
Dim oReference As Reference

oPart.InWorkObject = obody
Set oReference = oPart.CreateReferenceFromObject(RefGeoset.HybridShapes.Item(oRefGeoset.HybridShapes.Count))

Set oThick = oShapeFactory.AddNewThickSurface(oReference, 0, -dThickness, 0)

oPart.Update
End Sub


Public Sub DeleteBody(ByVal oRootDocument As Document)

Set oSelection = oRootDocument.Selection
oSelection.Clear

oSelection.Add obody
oSelection.Delete
End Sub
