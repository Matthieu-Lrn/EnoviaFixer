Attribute VB_Name = "InstantiateHardware"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oDummy As New clsHardwareConstruction
Private oSurface1 As New clsHardwareConstruction
Private oSurface2 As New clsHardwareConstruction
Private oSurface3 As New clsHardwareConstruction
Private oCenterLine As New clsHardwareConstruction
Private oSPRootDocument As Document

Public Sub Launch(ByVal oSPOther As Product, ByVal oSPInsert As Product, _
                  ByVal oHardware1 As Product, ByVal oHardware2 As Product, ByVal oHardware3 As Product, _
                  ByVal oDoc As Document, ByVal sInsertsList As String)

Dim sType As String
Dim oSelectedGeometry
Dim oPosition(11) As Variant
Dim oNewHardware As Product
Dim bSurfaceSelection As Boolean
Dim sStatus As String
Dim sDeletionList As String

'Initial set
Set oSPRootDocument = oDoc

'Swap to Assembly Design
Call SwapToAssemblyWorkbench

'Add new dummy part under oSPOther
Set oDummy.Product = oSPOther.ReferenceProduct.Products.AddNewComponent("Part", "")
Call oDummy.Initialize
oDummy.ParentList = BuildParentList(oSPOther)    'Ca ne fonctionne pas si je fais: BuildParentList(oDummy.Product). Comprend pas pourquoi ??
oDummy.ParentList = oDummy.Product.Name & ";" & oDummy.ParentList
oDummy.Part.Update

'Refresh
CATIA.RefreshDisplay = True

Do
    
    '****************************************************************
    'Hardware Centerline

    'We ask the user to select the centerline
    sType = SelectCenterLine("Select the hardware centerline or ESC to end process")
    If sType = "" Then
        GoTo BackToUserForm
    End If
       
    Set oCenterLine.Product = oSelectedItem.LeafProduct
    Call oCenterLine.Initialize
    oCenterLine.ParentList = BuildParentList(oCenterLine.Product)
    
    'The selection is a Line
    If sType = "Line" Then
    
        'Copy/Paste the line to the dummy part
        Call CopyPasteElements(oCenterLine.ParentList, oDummy.ParentList, "_DummyGeoset")
        oDummy.Part.Update

        'Rename the copied line
        oDummy.HybridBody.HybridShapes.Item(oDummy.HybridBody.HybridShapes.Count).Name = "_Centerline"
        
    'The selection is a CylindricalFace
    Else
        
        Call HideObject(oCenterLine.HybridBody)

        'Create the centerline from the cylindrical face
        Call oCenterLine.CreateCenterLineFromCylinder(oSelectedItem.Value)

        'Copy/Paste the line to the dummy part
        Call CopyPasteElements(oCenterLine.ParentList, oDummy.ParentList, "_DummyGeoset", "_Centerline")
        oDummy.Part.Update

        'Rename the copied line
        oDummy.HybridBody.HybridShapes.Item(oDummy.HybridBody.HybridShapes.Count).Name = "_Centerline"

    
    End If
    
    Call SwapToAssemblyWorkbench
    
    'Delete "_DummyGeoset_" of oCenterLine
     If Not oCenterLine.HybridBody Is Nothing Then
        Call DeleteGeoset(oCenterLine.Part, oCenterLine.HybridBody)
        Set oCenterLine.HybridBody = Nothing
        oCenterLine.Part.Update
     End If

    Call SwapToAssemblyWorkbench
    
    
    '****************************************************************
    'Hardware 1
    
    'Add Hardware
    If Not oHardware1 Is Nothing Then
    
        'Selection of Surface 1
        If frmInstantiateHardware.chkHardware1Surface.Value = False Or oSurface1.Product Is Nothing Then
            bSurfaceSelection = True
        Else
            bSurfaceSelection = False
        End If
        
        'Instantiate Hardware in oSPOther. Inserts will be transfered to oSP Insert later
        sStatus = Instantiate("1", oSurface1, oSPOther, oHardware1, bSurfaceSelection)
        
        If sStatus = "BackToUserForm" Then
            GoTo BackToUserForm
        End If
        
        'Deletion list
        sDeletionList = "_Intersect|_Origin|_ZAxis|_TangentPlane|_NormalLine|_XAxis"

        If frmInstantiateHardware.chkHardware1Surface.Value = False Then
            sDeletionList = sDeletionList & "|_Surface1"
        End If
                        
        'Delete "_DummyGeoset_" of oSurface1
        If Not oSurface1.HybridBody Is Nothing Then
            Call DeleteGeoset(oSurface1.Part, oSurface1.HybridBody)
            Set oSurface1.HybridBody = Nothing
            oSurface1.Part.Update
        End If

        'Delete elements in "_DummyGeoset_" of oDummy
        Call DeleteHybridShapes(oDummy.Part, oDummy.HybridBody, sDeletionList)

        Call SwapToAssemblyWorkbench
        
        'Transfer inserts to oSPInsert
        Call TransferInsertsToSP(sInsertsList, oSPOther, oSPInsert)

    End If
    Call SwapToAssemblyWorkbench
            
    '****************************************************************
    'Hardware 2
    
    'Add Hardware
    If Not oHardware2 Is Nothing Then
    
        'Selection of Surface 2
        If frmInstantiateHardware.chkHardware2Surface.Value = False Or oSurface2.Product Is Nothing Then
            bSurfaceSelection = True
        Else
            bSurfaceSelection = False
        End If
        
        'Instantiate Hardware in oSPOther. Inserts will be transfered to oSP Insert later
        sStatus = Instantiate("2", oSurface2, oSPOther, oHardware2, bSurfaceSelection)
        
        If sStatus = "BackToUserForm" Then
            GoTo BackToUserForm
        End If
        
        'Deletion list
        sDeletionList = "_Intersect|_Origin|_ZAxis|_TangentPlane|_NormalLine|_XAxis"
        If frmInstantiateHardware.chkHardware2Surface.Value = False Then
            sDeletionList = sDeletionList & "|_Surface2"
        End If
        
        'Delete "_DummyGeoset_" of oSurface2
        If Not oSurface2.HybridBody Is Nothing Then
            Call DeleteGeoset(oSurface2.Part, oSurface2.HybridBody)
            Set oSurface2.HybridBody = Nothing
            oSurface2.Part.Update
        End If

        'Delete elements in "_DummyGeoset_" of oDummy
        Call DeleteHybridShapes(oDummy.Part, oDummy.HybridBody, sDeletionList)

        Call SwapToAssemblyWorkbench
        
        'Transfer inserts to oSPInsert
        Call TransferInsertsToSP(sInsertsList, oSPOther, oSPInsert)

    End If
    Call SwapToAssemblyWorkbench
    
    '****************************************************************
    'Hardware 3
    
    'Add Hardware
    If Not oHardware3 Is Nothing Then
    
        'Selection of Surface 3
        If frmInstantiateHardware.chkHardware3Surface.Value = False Or oSurface3.Product Is Nothing Then
            bSurfaceSelection = True
        Else
            bSurfaceSelection = False
        End If
        
        'Instantiate Hardware in oSPOther. Inserts will be transfered to oSP Insert later
        sStatus = Instantiate("3", oSurface3, oSPOther, oHardware3, bSurfaceSelection)
        
        If sStatus = "BackToUserForm" Then
            GoTo BackToUserForm
        End If
        
        'Deletion list
        sDeletionList = "_Intersect|_Origin|_ZAxis|_TangentPlane|_NormalLine|_XAxis"
        If frmInstantiateHardware.chkHardware3Surface.Value = False Then
            sDeletionList = sDeletionList & "|_Surface3"
        End If
                
        'Delete "_DummyGeoset_" of oSurface3
        If Not oSurface3.HybridBody Is Nothing Then
            Call DeleteGeoset(oSurface3.Part, oSurface3.HybridBody)
            Set oSurface3.HybridBody = Nothing
            oSurface3.Part.Update
        End If

        'Delete elements in "_DummyGeoset_" of oDummy
        Call DeleteHybridShapes(oDummy.Part, oDummy.HybridBody, sDeletionList)

        Call SwapToAssemblyWorkbench
        
        'Transfer inserts to oSPInsert
        Call TransferInsertsToSP(sInsertsList, oSPOther, oSPInsert)

    End If
    Call SwapToAssemblyWorkbench

    '****************************************************************
    'Delete
   
    'Delete elements in "_DummyGeoset_" of oDummy
    sDeletionList = "_Centerline"
    Call DeleteHybridShapes(oDummy.Part, oDummy.HybridBody, sDeletionList)
    Call SwapToAssemblyWorkbench
Loop

BackToUserForm:

'Delete "_DummyGeoset_" of oCenterLine
If Not oCenterLine.HybridBody Is Nothing Then
    Call DeleteGeoset(oCenterLine.Part, oCenterLine.HybridBody)
    Set oCenterLine.HybridBody = Nothing
    oCenterLine.Part.Update
End If

'Delete "_DummyGeoset_" of oSurface1
If Not oSurface1.HybridBody Is Nothing Then
    Call DeleteGeoset(oSurface1.Part, oSurface1.HybridBody)
    Set oSurface1.HybridBody = Nothing
    oSurface1.Part.Update
End If

'Delete "_DummyGeoset_" of oSurface2
If Not oSurface2.HybridBody Is Nothing Then
    Call DeleteGeoset(oSurface2.Part, oSurface2.HybridBody)
    Set oSurface2.HybridBody = Nothing
    oSurface2.Part.Update
End If

'Delete "_DummyGeoset_" of oSurface3
If Not oSurface3.HybridBody Is Nothing Then
    Call DeleteGeoset(oSurface3.Part, oSurface3.HybridBody)
    Set oSurface3.HybridBody = Nothing
    oSurface3.Part.Update
End If

'Clear objects
Set oSurface1 = Nothing
Set oSurface2 = Nothing
Set oSurface3 = Nothing

'Delete oDummy
If Not oDummy.Product Is Nothing Then
    Set oSelection = oSPRootDocument.Selection
    oSelection.Clear
    oSelection.Add oSPOther
    
    Call SwapToAssemblyWorkbench
    oSelection.Clear
    oSelection.Add oDummy.Product
    oSelection.Delete
End If

'Next section is required to activate (blue) an instance in the product structure
If Not oCenterLine.Product Is Nothing Then
    oSelection.Add oCenterLine.Product
    Set oSelection = KBESearch("(Name='xy plane' & CATGmoSearch.Plane),sel")
    CATIA.StartWorkbench "PrtCfg"
    CATIA.StartWorkbench "Assembly"
    oSelection.Clear
End If

End Sub

Private Function SelectCenterLine(ByVal sMessage As String) As String

Dim sType As String
Dim oPart As Part
Dim oReference As Reference
Dim oHybridShapeFactory As HybridShapeFactory

'Initial set
sType = ""

'Loop until a line or cylindrical face is selected
Do While sType = ""

    Set oSelectedItem = MultiWindowSelectElement(sMessage)
    If oSelectedItem Is Nothing Then Exit Do
    
    sString = oSelectedItem.LeafProduct.GetShapePathName(oSelectedItem.LeafProduct.GetDefaultShapeName)
    
    'Selection must be done in the same document as the SP
    If oSelectedItem.Document.Name <> oSPRootDocument.Name Then
        MsgBox ("Selection must be done in " & oSPRootDocument.Name & " window")
        GoTo NewSelection
    End If
    
    'Selection is a "CylindricalFace"
    If TypeName(oSelectedItem.Value) = "CylindricalFace" Then
        sType = "CylindricalFace"
        Exit Do
        
    'Selection is Hybridshape
    ElseIf TypeName(oSelectedItem.Value) Like "HybridShape*" Or TypeName(oSelectedItem.Value.Parent) Like "HybridShape*" Then
    
        Set oPart = oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part
        Set oHybridShapeFactory = oPart.HybridShapeFactory
   
        'Create the reference
        If TypeName(oSelectedItem.Value) Like "HybridShape*" Then
            Set oReference = oPart.CreateReferenceFromObject(oSelectedItem.Value)
        Else
            Set oReference = oPart.CreateReferenceFromObject(oSelectedItem.Value.Parent)
        End If
    
        'Retrieve geometry type
        If oHybridShapeFactory.GetGeometricalFeatureType(oReference) = 3 Then
            sType = "Line"
            Exit Do
        End If
    ElseIf TypeName(oSelectedItem.Value) = "StiDBItem" Then
        oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        sAnswer = MsgBox("Selected part was put in Design Mode. Please select again.", vbExclamation)
    ElseIf TypeName(oSelectedItem.Value) = "Product" Then
        sAnswer = MsgBox("You just selected an instance. Make sure the Product Selection icon is not active in the Assembly Design toolbar.", vbExclamation)
    Else
        sAnswer = MsgBox("Select a line or a cylindrical face.", vbExclamation)
    End If
    
NewSelection:
Loop
       
SelectCenterLine = sType
End Function

Private Function SelectFaceSurfacePlane(ByVal sMessage As String) As String

'This function checks that the type of the selected geometry is a: face, surface or plane

Dim oPart As Part
Dim oReference As Reference
Dim oHybridShapeFactory As HybridShapeFactory
Dim oHybridShape As HybridShape
Dim oSelectedGeometry As AnyObject
Dim sType As String

'Initial set
sType = ""

Do While sType = ""

    Set oSelectedItem = MultiWindowSelectElement(sMessage)
    If oSelectedItem Is Nothing Then Exit Do

    'Selection must be done in the same document as the SP
    If oSelectedItem.Document.Name <> oSPRootDocument.Name Then
        MsgBox ("Selection must be done in " & oSPRootDocument.Name & " window")
        GoTo NewSelection
    End If

    'Selection is a face on a solid
    If (TypeName(oSelectedItem.Value) = "Face" Or TypeName(oSelectedItem.Value) = "CylindricalFace") And _
       Not TypeName(oSelectedItem.Value.Parent) Like "HydridShape*" Then
       
        sType = "Face"
        Exit Do

    'Selection is a planarface on a solid
    ElseIf TypeName(oSelectedItem.Value) = "PlanarFace" And Not TypeName(oSelectedItem.Value.Parent) Like "HydridShape*" Then
                
        sType = "PlanarFace"
        Exit Do
        
    'Selection is Hybridshape
    ElseIf TypeName(oSelectedItem.Value) Like "HybridShape*" Or TypeName(oSelectedItem.Value.Parent) Like "HybridShape*" Then

        Set oPart = oSelectedItem.LeafProduct.ReferenceProduct.Parent.Part
        Set oHybridShapeFactory = oPart.HybridShapeFactory
   
        'Create the reference
        If TypeName(oSelectedItem.Value) Like "HybridShape*" Then
            Set oReference = oPart.CreateReferenceFromObject(oSelectedItem.Value)
        Else
            Set oReference = oPart.CreateReferenceFromObject(oSelectedItem.Value.Parent)
        End If
    
        'Retrieve geometry type
        If oHybridShapeFactory.GetGeometricalFeatureType(oReference) = 5 Then
            sType = "Surface"
            Exit Do
        ElseIf oHybridShapeFactory.GetGeometricalFeatureType(oReference) = 6 Then
            sType = "Plane"
            Exit Do
        Else
            MsgBox ("Select a plane, a surface or a face on a solid")
        End If
    ElseIf TypeName(oSelectedItem.Value) = "StiDBItem" Then
        oSelectedItem.LeafProduct.ApplyWorkMode DESIGN_MODE
        sAnswer = MsgBox("Selected part was put in Design Mode. Please select again.", vbExclamation)
    ElseIf TypeName(oSelectedItem.Value) = "Product" Then
        sAnswer = MsgBox("You just selected an instance. Make sure the Product Selection icon is not active in the Assembly Design toolbar.", vbExclamation)
    Else
        sAnswer = MsgBox("Select a plane, a surface or a face on a solid.", vbExclamation)
        
    End If

NewSelection:
Loop

SelectFaceSurfacePlane = sType

End Function

Private Sub HideObject(oObject As Variant)

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.Clear
oSelection.Add oObject

Dim oVisu As VisPropertySet

Set oVisu = oSelection.VisProperties
oVisu.SetShow 1

End Sub

Private Function Instantiate(ByVal sID As String, ByRef oSurface As clsHardwareConstruction, _
                             ByVal oSP As Product, ByVal oHardware As Product, ByVal bSurfaceSelection As Boolean) As String

Dim sType As String
Dim sStatus As String
Dim oSelectedGeometry
Dim oNewHardware As Product
Dim oMatrix(11) As Variant
Dim oPosition
Dim oControl As Control

'Initial Set
sStatus = "OK"

'Surface selection is required ?
If bSurfaceSelection Then

    'We ask the user to select the surface
    sType = SelectFaceSurfacePlane("Select the mating surface of  " & oHardware.PartNumber & " or ESC to end process")
    If sType = "" Then
        sStatus = "BackToUserForm"
        GoTo endsub
    End If
    
    
    Set oSurface.Product = oSelectedItem.LeafProduct
    Call oSurface.Initialize
    oSurface.ParentList = BuildParentList(oSurface.Product)
    Call HideObject(oSurface.HybridBody)
    
    'The selection is a face
    If sType = "Face" Then
    
        'Create extract surface - For intersection
        Call oSurface.CreateExtractSurface(oSelectedItem.Value, "_Surface" & sID)
        oSurface.Part.Update
        Call HideObject(oSurface.HybridBody)
    
        'Copy/Paste the extracted face in DummyPart
        Call CopyPasteElements(oSurface.ParentList, oDummy.ParentList, "_DummyGeoset", "_Surface" & sID)
        
        'Rename the copied element
        oDummy.HybridBody.HybridShapes.Item(oDummy.HybridBody.HybridShapes.Count).Name = "_Surface" & sID

       
    ElseIf sType = "PlanarFace" Then
               
        'Create a plane on a planar face
        Call SwapToAssemblyWorkbench
        Call oSurface.CreatePlaneOnPlanarFace(oSelectedItem.Value, "_Surface" & sID)
        oSurface.Part.Update
        
        'Copy/Paste the plane to the dummy part
        Call CopyPasteElements(oSurface.ParentList, oDummy.ParentList, "_DummyGeoset", "_Surface" & sID)
        oDummy.Part.Update
    
        'Rename the copied element
        oDummy.HybridBody.HybridShapes.Item(oDummy.HybridBody.HybridShapes.Count).Name = "_Surface" & sID
    
    'The selection is a Plane or Surface
    Else
    
        'Copy/Paste the plane or surface in DummyPart
        Call CopyPasteElements(oCenterLine.ParentList, oDummy.ParentList, "_DummyGeoset", oSelectedItem.Value.Name)
        oDummy.Part.Update

        'Rename the copied element
        oDummy.HybridBody.HybridShapes.Item(oDummy.HybridBody.HybridShapes.Count).Name = "_Surface" & sID

    End If
    
    Call SwapToAssemblyWorkbench
End If

'Create intersection point between "_Centerline" and the surface
Call SwapToAssemblyWorkbench
If oDummy.CreateIntersectionPoint("_Surface" & sID) Then
    sStatus = "BackToUserForm"
    GoTo endsub
End If
oDummy.Part.Update

'Retrieve Position Matrix
Call oDummy.CalculatePositionMatrix(oMatrix, sID)

'****************************************************************
'Add new instance of Hardware

'Create a new instance of the std part
oSP.Products.AddExternalComponent oHardware.ReferenceProduct.Parent
Set oNewHardware = oSP.Products.Item(oSP.Products.Count)

'Hardware Position
Set oPosition = oNewHardware.Position
oPosition.SetComponents oMatrix

'Refresh
CATIA.RefreshDisplay = True

'Retrieve chkHardwarexxOrientation control
For Each oControl In frmInstantiateHardware.Controls
    If oControl.Name = "chkHardware" & sID & "Orientation" Then Exit For
Next

'Invert position
If oSurface.InvertOrientation = "" Then
    sAnswer = MsgBox("Do you want to invert the orientation ?", vbYesNo)
    If oControl.Value = True Then
        oSurface.InvertOrientation = sAnswer
    End If
Else
    sAnswer = oSurface.InvertOrientation
End If

If sAnswer = vbYes Then
    oMatrix(3) = -oMatrix(3)
    oMatrix(4) = -oMatrix(4)
    oMatrix(5) = -oMatrix(5)
    oMatrix(6) = -oMatrix(6)
    oMatrix(7) = -oMatrix(7)
    oMatrix(8) = -oMatrix(8)
    oPosition.SetComponents oMatrix
End If
CATIA.RefreshDisplay = True

endsub:

Instantiate = sStatus

End Function

Private Sub CopyPasteElements(sSourceParentList As String, sTargetParentList As String, sTargetHybridBody As String, Optional sElementName As String = "")

'When there is more that one instance of the source or the target instance it is required to active the right part
'for the copy and paste operation


Dim sWkb As String
Dim i As Integer
Dim oProduct As Product

'Initial set
Set oSelection = CATIA.ActiveDocument.Selection

'We copy the active selection
If sElementName = "" Then

    'Copy
    CATIA.StartWorkbench "PrtCfg"
    oSelection.Copy
    CATIA.StartWorkbench "Assembly"

    'Retrieve parent instance of target element
    Set oProduct = RetrieveInstanceFromList(sTargetParentList)

    'Paste
    oSelection.Clear
    oSelection.Add oProduct
    oSelection.Search "Name=" & sTargetHybridBody & ",sel"
    CATIA.StartWorkbench "PrtCfg"
    oSelection.PasteSpecial "CATPrtResultWithOutLink"
    oSelection.Clear


'We copy the "sElementName"
Else
    'Retrieve parent instance of element to copy
    Set oProduct = RetrieveInstanceFromList(sSourceParentList)
    
    'Copy
    oSelection.Clear
    oSelection.Add oProduct
    oSelection.Search "Name=" & sElementName & ",sel"
    CATIA.StartWorkbench "PrtCfg"
    oSelection.Copy
    CATIA.StartWorkbench "Assembly"

    'Retrieve parent instance of target element
    Set oProduct = RetrieveInstanceFromList(sTargetParentList)

    'Paste
    oSelection.Add oProduct
    oSelection.Search "Name=" & sTargetHybridBody & ",sel"
    CATIA.StartWorkbench "PrtCfg"
    oSelection.PasteSpecial "CATPrtResultWithOutLink"
    oSelection.Clear

End If

End Sub

Private Function RetrieveInstanceFromList(sList As String) As Product

'From a parent list we can retrieve a specific instance in the tree

Dim oSelection
Dim i As Integer
Dim oProduct As Product
Dim oProdDocument As ProductDocument

'Initial set
Set oProdDocument = CATIA.ActiveDocument
Set oProduct = oProdDocument.Product
Set oSelection = oProdDocument.Selection

'Remove the last caracter which is a ";"
If Right(sList, 1) = ";" Then
    sList = Left(sList, Len(sList) - 1)
End If

'We select the instance in which the element to copy is found
For i = UBound(Split(sList, ";")) - 1 To 0 Step -1                  'We skip the first element in the list since it is the top product
    Set oProduct = oProduct.Products.Item(Split(sList, ";")(i))
Next

Set RetrieveInstanceFromList = oProduct

End Function

Private Sub DeleteGeoset(ByVal oPart As Part, ByVal oObject As AnyObject)

Dim oReference As Reference
Dim oHybridShapeFactory As HybridShapeFactory


Set oHybridShapeFactory = oPart.HybridShapeFactory

'Create a reference
Set oReference = oPart.CreateReferenceFromObject(oObject)

'Delete the object
oHybridShapeFactory.DeleteObjectForDatum oReference

End Sub

Public Sub DeleteHybridShapes(ByVal oPart As Part, ByVal oGeoset As HybridBody, ByVal sDeletionList As String)

Dim oReference As Reference
Dim oHybridShapeFactory As HybridShapeFactory
Dim i As Integer
Dim oHybridShape As HybridShape
Dim sName As String

'Initial set
Set oHybridShapeFactory = oPart.HybridShapeFactory
        
'Delete
For Each oHybridShape In oGeoset.HybridShapes
    For i = 0 To UBound(Split(sDeletionList, "|"))
        sName = Split(sDeletionList, "|")(i)
        If oHybridShape.Name = sName Then
            Set oReference = oPart.CreateReferenceFromObject(oHybridShape)
            oHybridShapeFactory.DeleteObjectForDatum oReference
            Exit For
        End If
    Next
Next

End Sub

Private Sub TransferInsertsToSP(ByVal sInsertsList As String, ByVal oSPOther As Product, ByVal oSPInsert As Product)

Dim oInsert As Product

'Retrieve last instantiated in oSPOther
Set oInsert = oSPOther.Products.Item(oSPOther.Products.Count)

'Cut/Paste if it is found in the list
If sInsertsList Like "*|" & oInsert.PartNumber & "|*" And Not oSPInsert Is Nothing Then
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.Clear
    oSelection.Add oInsert
    oSelection.Cut

    oSelection.Clear
    oSelection.Add oSPInsert
    oSelection.Paste
End If

End Sub
Private Sub SwapToAssemblyWorkbench()

Dim sWkb As String

sWkb = CATIA.GetWorkbenchId
If sWkb <> "Assembly" Then
    CATIA.StartWorkbench "Assembly"
End If

End Sub


