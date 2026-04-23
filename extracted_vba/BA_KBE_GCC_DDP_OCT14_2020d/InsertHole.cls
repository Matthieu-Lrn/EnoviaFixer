Attribute VB_Name = "InsertHole"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oLineCollection As New Collection
Private oExcel As clsExcel
Private Const MinimalDistanceBetweenPointAndPlane As Double = 0.0254    '0.001in
Private Const MaximumMisalignmentBetweenHoleAndInsert As Double = 0.0254 'value in mm
Private Const MinimalAngleBetweenLinetAndPlane As Double = 0.5
'********************************************************************************
'* Name: Insert Hole Collection
'* Purpose:
'*
'* Assumption:
'*
'* Author: Abhishek Kamboj
'* Language: VBA
'********************************************************************************
Public Sub LaunchCreation(ByVal oPanelObject As clsPanelCheck, _
                          ByVal oSPCollList As clsCollection, _
                          ByVal oProductDoc As Document, _
                          ByVal oDatumPlnGeo As HybridBody, _
                          ByVal oConstructnGeo As HybridBody, _
                          ByVal bDeleteInsertHoles As Boolean, _
                          Optional bCheckFlushness As Boolean = True, _
                          Optional bDeletUnusedHoles As Boolean = False)

    Dim oThkFaceSide1 As Plane
    Dim oThkFaceSide2 As Plane
    Dim oPanelProduct As Product
    Dim oClashes As Clashes
    Dim oClash As Clash
    Dim oConflicts As Conflicts
    Dim oConflict As Conflict
    
    Dim oPanelPart As Part
    Dim oHSF As HybridShapeFactory
    Dim oSF As ShapeFactory
    Dim oSubHB As HybridBody
    Dim oPartBody As Body
    
    Dim oPoint As HybridShapePointCoord
    Dim oLine 'As Line
    Dim oProject As HybridShapeProject
    
    Dim oSubBody As Body
    Dim oHole 'As Hole
    Dim oLimit As Limit
    Dim HoleDia As Length
    Dim dHoledia As Double
    Dim dHoleDepthOffset As Double
    Dim oHoleSupport As Plane
    Dim oHoleBottom As HybridShape
    Dim bHoleDefinitionFound As Boolean
    Dim dHoleDirection(2)
    Dim dLineDirection(2) As Variant
    Dim oDirection As HybridShapeDirection
    Dim oDirHole As HybridShapeDirection
    Dim dPanelThk As Double
    
    Dim oHS As HybridShape
    Dim oSketch As Sketch
    Dim oReference As Reference
    Dim oReference2 As Reference
    
    Dim oClashProdInsert As Product
    Dim oClashProdPanel As Product
    Dim i As Integer, j As Integer, k As Integer
    
    Dim dGlobalMatrixInsert(11) As Variant
    Dim dGlobalMatrixPanel(11) As Variant
    Dim dLocalMatrixInsert(11) As Variant
    Dim dProjectCoord(2) As Variant
    Dim sParentList As String
    
    Dim oSPAWorkbench As Workbench
    Dim oMeasurable 'As Measurable
    Dim dOffsetFromPlane As Double
    Dim dAngle As Double
    Dim oVisProperties As VisPropertySet
    
    Dim oParameters As Parameters
    Dim oParam As Parameter
    
    Dim oSymmetryPlane As Plane
    Dim oSymmetry As Symmetry
    
    Dim iLastRow As Integer
    Dim oArray() As Variant
    Dim oHoleDiaInfo As New clsCollection
    Dim oItem As clsCollection
    Dim bShowGeoset As Boolean
    
    bShowGeoset = False
    
    Set oThkFaceSide1 = oPanelObject.PanelPlane1
    Set oThkFaceSide2 = oPanelObject.PanelPlane2
    Set oPanelProduct = oPanelObject.PanelInstance
    
    'Swap workbench
    Call SwapToAssemblyWorkbench
    
    'Set oSelection
    Set oSelection = oProductDoc.Selection
    
    'SPAWorkbench for Measurable object
    Set oSPAWorkbench = oProductDoc.GetWorkbench("SPAWorkbench")
    
    'Retrieve the Clashes collection
    Set oClashes = oProductDoc.Product.GetTechnologicalObject("Clashes")
    
    'Open Excel file
    Set oExcel = New clsExcel
    Call oExcel.GetNewExcel
    Call oExcel.OpenExcelFile(sInsertHoleDefinition)
    
    'Transfer data to oArray
    iLastRow = 2
    Do While Trim(oExcel.App.Cells(iLastRow, 1).Value) <> ""
        iLastRow = iLastRow + 1
    Loop
    iLastRow = iLastRow - 1
    
    oArray = oExcel.App.Range(oExcel.App.Cells(2, 1), oExcel.App.Cells(iLastRow, 5)).Value
    
    'Close Excel workbook and kill process
    Call oExcel.CloseAlreadyOpenWorkbook(sInsertHoleDefinition)
    Set oExcel = Nothing
    
    Dim oHoleDiainfobyDia As clsCollection
    
    Set oHoleDiainfobyDia = New clsCollection '*** Key is Hole diameter instead of insert partnumber
    'Transfer data from oArray to oHoleDiaInfo
    For i = LBound(oArray, 1) To UBound(oArray, 1)
        Set oItem = New clsCollection
        Call oItem.Add("Hole Diameter", oArray(i, 2))
        Call oItem.Add("Hole Depth Offset", oArray(i, 3))
        Call oItem.Add("Symmetry", oArray(i, 4))
        Call oItem.Add("PanelThickness", oArray(i, 5))  'Abhishek
        Call oItem.Add("InsertNumber", oArray(i, 1))
        
        If Not oHoleDiaInfo.Exists(oArray(i, 1)) Then
            Call oHoleDiainfobyDia.Add(CStr(oArray(i, 2)), oItem)
            Call oHoleDiaInfo.Add(oArray(i, 1), oItem)
        Else
            sAnswer = MsgBox(oArray(i, 1) & " can be found twice in the Excel file." & vbCrLf & "Please contact Process & Tools group.", vbCritical)
            GoTo TheEnd
        End If
    Next
    
    'Panel Part
    Set oPanelPart = oPanelProduct.GetMasterShapeRepresentation(False).Part
    Set oHSF = oPanelPart.HybridShapeFactory
    Set oSF = oPanelPart.ShapeFactory
    
    'Retrieve PartBody
    On Error Resume Next
    Set oPartBody = Nothing
    Set oPartBody = oPanelPart.Bodies.Item("PartBody")
    If oPartBody Is Nothing Then
        sAnswer = MsgBox("PartBody can't be found." & vbCrLf & "Check the spelling and run the macro again.", vbCritical)
        GoTo TheEnd
    End If
    
    'Calculate Panel Thickness
    Set oReference = oPanelPart.CreateReferenceFromObject(oThkFaceSide1)
    Set oReference2 = oPanelPart.CreateReferenceFromObject(oThkFaceSide2)
    Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
    dPanelThk = oMeasurable.GetMinimumDistance(oReference2)
    'Progress bar
    Call frmProgress.progressBarInitialize("Create holes for inserts")
    
    oSelection.Clear
    'Calculate global position of panel
    Dim sParentListPanel As String
    sParentListPanel = BuildParentList(oPanelProduct)
    Call CalculateInstanceGlobalPositionMatrix(sParentListPanel, dGlobalMatrixPanel)
    'Calculate Global Position matrix of Plane1
    Dim U(2) As Double, V(2) As Double, W As Variant, p3(2) As Double
    Dim dGlobalMatrixPlane1(11) As Variant
    Dim dLocalMatrixPlane1(11) As Variant
    Dim dVector(9 To 11) As Variant, dVectorReturn(9 To 11) As Variant
    Dim Plane1Pt1(2) As Double, Plane1Pt2(2) As Double
    
    U(0) = oPanelObject.ArPlane1(3): U(1) = oPanelObject.ArPlane1(4): U(2) = oPanelObject.ArPlane1(5)
    V(0) = oPanelObject.ArPlane1(6): V(1) = oPanelObject.ArPlane1(7): V(2) = oPanelObject.ArPlane1(8)
    Call CrossProduct(U, V, W)
    
    dLocalMatrixPlane1(0) = U(0): dLocalMatrixPlane1(1) = U(1): dLocalMatrixPlane1(2) = U(2)
    dLocalMatrixPlane1(3) = V(0): dLocalMatrixPlane1(4) = V(1): dLocalMatrixPlane1(5) = V(2)
    dLocalMatrixPlane1(6) = W(0): dLocalMatrixPlane1(7) = W(1): dLocalMatrixPlane1(8) = W(2)
    dLocalMatrixPlane1(9) = oPanelObject.ArPlane1(0): dLocalMatrixPlane1(10) = oPanelObject.ArPlane1(1): dLocalMatrixPlane1(11) = oPanelObject.ArPlane1(2)
    
    Call MultiplyVector(dLocalMatrixPlane1, dGlobalMatrixPanel, dGlobalMatrixPlane1, True)
    Call MultiplyMatrix(dLocalMatrixPlane1, dGlobalMatrixPanel, dGlobalMatrixPlane1)
    dVector(9) = 0: dVector(10) = 0: dVector(11) = 1
    Call MultiplyVector(dVector, dGlobalMatrixPlane1, dVectorReturn, True)
    Plane1Pt1(0) = dGlobalMatrixPlane1(9): Plane1Pt1(1) = dGlobalMatrixPlane1(10): Plane1Pt1(2) = dGlobalMatrixPlane1(11)
    Plane1Pt2(0) = dVectorReturn(9): Plane1Pt2(1) = dVectorReturn(10): Plane1Pt2(2) = dVectorReturn(11)
    
    'Calculate Global Position matrix of Plane2
    
    Dim dGlobalMatrixPlane2(11) As Variant
    Dim dLocalMatrixPlane2(11) As Variant
    Dim Plane2Pt1(2) As Double, Plane2Pt2(2) As Double
    
    U(0) = oPanelObject.ArPlane2(3): U(1) = oPanelObject.ArPlane2(4): U(2) = oPanelObject.ArPlane2(5)
    V(0) = oPanelObject.ArPlane2(6): V(1) = oPanelObject.ArPlane2(7): V(2) = oPanelObject.ArPlane2(8)
    Call CrossProduct(U, V, W)
    
    dLocalMatrixPlane2(0) = U(0): dLocalMatrixPlane2(1) = U(1): dLocalMatrixPlane2(2) = U(2)
    dLocalMatrixPlane2(3) = V(0): dLocalMatrixPlane2(4) = V(1): dLocalMatrixPlane2(5) = V(2)
    dLocalMatrixPlane2(6) = W(0): dLocalMatrixPlane2(7) = W(1): dLocalMatrixPlane2(8) = W(2)
    dLocalMatrixPlane2(9) = oPanelObject.ArPlane2(0): dLocalMatrixPlane2(10) = oPanelObject.ArPlane2(1): dLocalMatrixPlane2(11) = oPanelObject.ArPlane2(2)
    
    Call MultiplyVector(dLocalMatrixPlane2, dGlobalMatrixPanel, dGlobalMatrixPlane2, True)
    Call MultiplyMatrix(dLocalMatrixPlane2, dGlobalMatrixPanel, dGlobalMatrixPlane2)
    dVector(9) = 0: dVector(10) = 0: dVector(11) = 1
    Call MultiplyVector(dVector, dGlobalMatrixPlane2, dVectorReturn, True)
    Plane2Pt1(0) = dGlobalMatrixPlane2(9): Plane2Pt1(1) = dGlobalMatrixPlane2(10): Plane2Pt1(2) = dGlobalMatrixPlane2(11)
    Plane2Pt2(0) = dVectorReturn(9): Plane2Pt2(1) = dVectorReturn(10): Plane2Pt2(2) = dVectorReturn(11)
    
    '***scan existing holes
    Call frmProgress.progressBarRepaint("Identifying Existing holes...")
    
    Dim oExistingHoles As Collection
    Set oExistingHoles = scanExistingHoles(oPanelObject, oHoleDiainfobyDia, Plane1Pt1, Plane1Pt2, Plane2Pt1, Plane2Pt2)
    On Error Resume Next
    oPanelObject.PanelPart.UpdateObject oPanelObject.PanelPart.MainBody
    On Error GoTo 0
    '****Delete Unused Lines and points in Holes HB
    Dim oColx As Collection
    Dim ObjtoDel
    For Each oSubHB In oConstructnGeo.HybridBodies
        If UCase(oSubHB.Name) Like "INSERTS_*" Then
            For k = 1 To 2
                oSelection.Clear
                oSelection.Add oSubHB
                oSelection.Search ("'Generative Shape Design'.'Geometric feature'.IsRoot=TRUE,sel")
                Set oColx = New Collection
                If oSelection.Count > 0 Then
                     For j = 1 To oSelection.Count
                        oColx.Add oSelection.Item(j).Value
                     Next
                End If
                For Each ObjtoDel In oColx
                    oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(ObjtoDel)
                Next
            Next
        End If
    Next
    '***Put std parts in design mode ans select them
    oSelection.Clear
    For j = 1 To oSPCollList.Count
        Set oSPNodeProduct = oSPCollList.GetItemByIndex(j)
        oSPNodeProduct.ApplyWorkMode DESIGN_MODE
        For i = 1 To oSPNodeProduct.Products.Count
            oSelection.Add oSPNodeProduct.Products.Item(i)
        Next
    Next
    oSelection.Add oPanelProduct 'Add Panel
    
    'Compute clash
    Set oClash = oClashes.AddFromSel
    oClash.Compute
    
    'Retrieve the conflicts
    Set oConflicts = oClash.Conflicts
    'Scan all conflicts
    For i = 1 To oConflicts.Count
        
        Call frmProgress.progressBarRepaint("Checking conflicts", oConflicts.Count, i)
    
        Set oConflict = oConflicts.Item(i)
        
        'Check Conflict Type
        If oConflict.Type <> catConflictTypeClash Then GoTo NextConflict
        
        'Retrieve products involved in the clash
        If oConflict.firstProduct.Name = oPanelProduct.Name Then
            Set oClashProdPanel = oConflict.firstProduct
            Set oClashProdInsert = oConflict.secondProduct
        ElseIf oConflict.secondProduct.Name = oPanelProduct.Name Then
            Set oClashProdInsert = oConflict.firstProduct
            Set oClashProdPanel = oConflict.secondProduct
        Else: GoTo NextConflict
        
        End If
        
        If oHoleDiaInfo.Exists(oClashProdInsert.PartNumber) Then
            'reset Global Matrix of Panel
            Call CalculateInstanceGlobalPositionMatrix(sParentListPanel, dGlobalMatrixPanel)
            'Calculate global position of insert
            sParentList = BuildParentList(oClashProdInsert)
            Call CalculateInstanceGlobalPositionMatrix(sParentList, dGlobalMatrixInsert)    '***Global StartPoint Insert
            'Vector of End point in Local matrix
            Dim dLocalEndPtInsert(9 To 11) As Double
            Dim dGlobalEndPtInsert(9 To 11) As Double
            
            '***End Point is always +ve direction of z axis of insert
            dLocalEndPtInsert(9) = 0: dLocalEndPtInsert(10) = 0: dLocalEndPtInsert(11) = oHoleDiaInfo.GetItem(oClashProdInsert.PartNumber).GetItem("PanelThickness")
            Call MultiplyVector(dLocalEndPtInsert, dGlobalMatrixInsert, dGlobalEndPtInsert, True)  '***Global EndPoint Insert
            
            '***Check if insert passes through either plane1 or plane2
            Dim IStart(2) As Double, IEnd(2) As Double
            IStart(0) = dGlobalMatrixInsert(9): IStart(1) = dGlobalMatrixInsert(10): IStart(2) = dGlobalMatrixInsert(11)
            IEnd(0) = dGlobalEndPtInsert(9): IEnd(1) = dGlobalEndPtInsert(10): IEnd(2) = dGlobalEndPtInsert(11)
            bCheck = False
            
            If iFindLineSegmentIntersectionToPlane(IStart, IEnd, Plane1Pt1, Plane1Pt2) = 1 Then
                bCheck = True
            ElseIf iFindLineSegmentIntersectionToPlane(IStart, IEnd, Plane2Pt1, Plane2Pt2) = 1 Then
                bCheck = True
            End If
            If Not bCheck Then GoTo NextConflict '***Insert doesn't intersect with either plane
            
            Dim dDistP1 As Double, dDistP2 As Double
            Dim bSide1 As Boolean
            
            bSide1 = False
            
            dDistP1 = dDistancePointToPlane(IStart, Plane1Pt1, Plane1Pt2)
            dDistP2 = dDistancePointToPlane(IStart, Plane2Pt1, Plane2Pt2)
            If Abs(dDistP1) < Abs(dDistP2) Then
                dOffsetFromPlane = Abs(dDistP1)
                Set oHoleSupport = oThkFaceSide1
                Set oHoleBottom = oThkFaceSide2
                bSide1 = True
            Else
                dOffsetFromPlane = Abs(dDistP2)
                Set oHoleSupport = oThkFaceSide2
                Set oHoleBottom = oThkFaceSide1
                bSide1 = False
            End If
            'Calculate position matrix of StartPoint Insert in the panel CATPart: dLocalMatrixInsert 'StartPoint Insert in PanelMatrix
            
            Call SubstractVector(dGlobalMatrixInsert, dGlobalMatrixPanel)
            Call InvertMatrix(dGlobalMatrixPanel)
            Call MultiplyVector(dGlobalMatrixInsert, dGlobalMatrixPanel, dLocalMatrixInsert)
            Call MultiplyMatrix(dGlobalMatrixInsert, dGlobalMatrixPanel, dLocalMatrixInsert)
            
            '*********Geoset**********
            Set oSubHB = RetrieveSubHB(oClashProdInsert.PartNumber, oConstructnGeo)
            '**********Point**********
            'Create point in the panel part using dLocalMatrixInsert
            Set oPoint = oHSF.AddNewPointCoord(dLocalMatrixInsert(9), dLocalMatrixInsert(10), dLocalMatrixInsert(11))
            oSubHB.AppendHybridShape oPoint
            oPanelPart.UpdateObject oSubHB

            '**************Line***********
            'Create line in the panel part using dLocalMatrixInsert. Line is parallel to Z axis of insert.
            Set oDirection = oHSF.AddNewDirectionByCoord(dLocalMatrixInsert(6), dLocalMatrixInsert(7), dLocalMatrixInsert(8))
            Set oLine = oHSF.AddNewLinePtDir(oPoint, oDirection, 0, 25.4, False)
            oSubHB.AppendHybridShape oLine
            oLine.Compute
            
            'Define a reference with the line
            Set oReference = oPanelPart.CreateReferenceFromObject(oLine)
            
            'Retreive measurable object of the line
            Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
            
            'Define a reference with the oHoleSupport
            Set oReference2 = oPanelPart.CreateReferenceFromObject(oHoleSupport)
            
            'Measure angle between oLine and oHoleSupport
            dAngle = oMeasurable.GetAngleBetween(oReference2)
            
            'Angle is smaller than MinimalAngleBetweenLinetAndPlane
            If Not (Abs(Abs(dAngle) - 90) < MinimalAngleBetweenLinetAndPlane Or Abs(Abs(dAngle) - 270) < MinimalAngleBetweenLinetAndPlane) Then
                'Angle is too big
                oSelection.Clear
                oSelection.Add oLine
                Set oVisProperties = oSelection.VisProperties
                oVisProperties.SetRealColor 255, 0, 0, 0
                bShowGeoset = True
                GoTo NextConflict
            End If
                
            'Delete oLine because it might not be normal to oHoleSupport
            oHSF.DeleteObjectForDatum oReference
        
            'Project oPoint on oHoleBottom and retrieve coordinates
            'Use oProject to calculate centerline direction
            Set oProject = oHSF.AddNewProject(oPoint, oHoleBottom)
            oSubHB.AppendHybridShape oProject
            oProject.Compute
            Set oReference = oPanelPart.CreateReferenceFromObject(oProject)
            Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
            oMeasurable.GetPoint dProjectCoord
            oHSF.DeleteObjectForDatum oReference
            '***Compare with existing hole
            Dim dInsertCood(2) As Double
            Dim dProjCood(2) As Double
            Dim dHoleCood(2) As Double
            Dim oHoleinfo As clsCollection
            Dim oExistingHole As Hole
            Dim cExistingHole As clsCollection
            
            Set oExistingHole = Nothing
            dInsertCood(0) = dLocalMatrixInsert(9): dInsertCood(1) = dLocalMatrixInsert(10): dInsertCood(2) = dLocalMatrixInsert(11)
            dProjCood(0) = dProjectCoord(0): dProjCood(1) = dProjectCoord(1): dProjCood(2) = dProjectCoord(2)
            For Each oHoleinfo In oExistingHoles
                If oHoleinfo.GetItemByKey("InsertNumber") = oClashProdInsert.PartNumber Then
                    dHoleCood(0) = oHoleinfo.GetItemByKey("HoleLocalOrigin")(0)
                    dHoleCood(1) = oHoleinfo.GetItemByKey("HoleLocalOrigin")(1)
                    dHoleCood(2) = oHoleinfo.GetItemByKey("HoleLocalOrigin")(2)
                    If dDistancePointToLineSeg(dHoleCood, dProjCood, dInsertCood) < MaximumMisalignmentBetweenHoleAndInsert Then
                        If oHoleinfo.GetItemByKey("HoleOnSide1") = bSide1 Then
                            Set oExistingHole = oHoleinfo.GetItemByKey("HoleObject")
                            Set cExistingHole = oHoleinfo
                            Exit For
                        End If
                    End If
                End If
            Next
            '**** if Insert not flush and check flush is ticked then
            '**** change colour of point
            '**** if a hole already exist then delete it
            If dOffsetFromPlane > MinimalDistanceBetweenPointAndPlane And bCheckFlushness Then
                oSelection.Clear
                oSelection.Add oPoint
                Set oVisProperties = oSelection.VisProperties
                oVisProperties.SetRealColor 255, 0, 0, 0
                bShowGeoset = True
                GoTo NextConflict
            End If
            If Not oExistingHole Is Nothing Then    '***if hole already exist the delete point
                oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(oPoint)
                cExistingHole.Add "KeepHole", "TRUE"
            Else    '*** if hole doesn't already exist
                'Create hole centerline. This centerline is normal to oHoleSupport
                Set oDirection = oHSF.AddNewDirectionByCoord(dProjectCoord(0) - dLocalMatrixInsert(9), _
                                                             dProjectCoord(1) - dLocalMatrixInsert(10), _
                                                             dProjectCoord(2) - dLocalMatrixInsert(11))
                Set oLine = oHSF.AddNewLinePtDir(oPoint, oDirection, -12.7, dPanelThk + 12.7, False)
                oSubHB.AppendHybridShape oLine
                oLine.Compute
                '***************Hole**************
                
                'Retrieve hole diameter and depth offset from oHoleDiaInfo
                If oHoleDiaInfo.Exists(oClashProdInsert.PartNumber) Then
                    dHoledia = oHoleDiaInfo.GetItemByKey(oClashProdInsert.PartNumber).GetItemByKey("Hole Diameter")
                    dHoleDepthOffset = oHoleDiaInfo.GetItemByKey(oClashProdInsert.PartNumber).GetItemByKey("Hole Depth Offset")
                
                'Hole diameter and depth offset not found in oHoleDiaInfo
                Else
                    oSelection.Clear
                    oSelection.Add oLine
                    oSelection.Add oPoint
                    Set oVisProperties = oSelection.VisProperties
                    oVisProperties.SetRealColor 0, 0, 255, 0
                    bShowGeoset = True
                    GoTo NextConflict
                End If
                
                '***Retrieve the Sub Body where the hole will be created
                Set oSubBody = RetrieveSubBody(oClashProdInsert.PartNumber, oPanelPart, oPartBody, oSF)
                
                '***Create hole
                oPanelPart.InWorkObject = oSubBody
                Set oHole = oSF.AddNewHoleFromRefPoint(oPoint, oHoleSupport, dPanelThk + dHoleDepthOffset)
                oHole.Type = catSimpleHole
                oHole.AnchorMode = catExtremPointHoleAnchor
                oHole.BottomType = catFlatHoleBottom
                
                Set oLimit = oHole.BottomLimit
                oLimit.LimitMode = catOffsetLimit
                
                oPanelPart.UpdateObject oHole
                Set HoleDia = oHole.Diameter
                HoleDia.Value = dHoledia
                oHole.SetDirection oLine
                
                'For some reason I don't understand the orientation of the hole may be inverted from the direction of oLine
                Call oHole.GetDirection(dHoleDirection) 'Retrieve hole direction
                oLine.GetDirection dLineDirection 'Retrieve line direction
                
                If (dHoleDirection(0) * dLineDirection(0) < 0) Or _
                   (dHoleDirection(1) * dLineDirection(1) < 0) Or _
                   (dHoleDirection(2) * dLineDirection(2) < 0) Then
                    oHole.Reverse
                End If
                'Hide sketch
                Set oSketch = oHole.Sketch
                Set oReference = oPanelPart.CreateReferenceFromObject(oSketch)
                oHSF.GSMVisibility oReference, 0
            End If
        End If
NextConflict:
    Next
    
    oPanelPart.InWorkObject = oPartBody
    On Error Resume Next
    oPanelPart.Update
    oClashes.Remove (1)
    On Error GoTo 0
    If oExistingHoles.Count > 0 Then
        For i = 1 To oExistingHoles.Count
            If bDeletUnusedHoles Then
                If oExistingHoles.Item(i).Exists("KeepHole") Then
                    If oPanelPart.IsInactive(oExistingHoles.Item(i).GetItemByKey("HoleObject")) Then oPanelPart.Activate oExistingHoles.Item(i).GetItemByKey("HoleObject")
                Else
                    Set oHole = oExistingHoles.Item(i).GetItemByKey("HoleObject")
                    Set oSketch = oHole.Sketch
                    oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(oSketch)
                    oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(oHole)
                End If
            Else
                If oPanelPart.IsInactive(oExistingHoles.Item(i).GetItemByKey("HoleObject")) Then oPanelPart.Activate oExistingHoles.Item(i).GetItemByKey("HoleObject")
            End If
        Next
    End If
    '***check for empty bodies and if found delete it
    Dim bEmpty As Boolean
    
    For Each obody In oPanelPart.Bodies
        bEmpty = True
        If obody.HybridBodies.Count <> 0 Then
            bEmpty = False
        End If
        If obody.Shapes.Count <> 0 Then
            bEmpty = False
        End If
        If obody.Sketches.Count <> 0 Then
            bEmpty = False
        End If
        If Not obody.HybridShapes Is Nothing Then
            bEmpty = False
        End If
        If obody.Name Like oPanelPart.MainBody.Name Then
            bEmpty = False
        End If
        
        If bEmpty Then
            oSelection.Clear
            oSelection.Add obody
            On Error Resume Next
            oSelection.Delete
            On Error GoTo 0
        End If
    Next
    '****Delete Unused Lines and points in Holes HB.... Only white coloured elements
    For Each oSubHB In oConstructnGeo.HybridBodies
        If UCase(oSubHB.Name) Like "INSERTS_*" Then
            For k = 1 To 2
                oSelection.Clear
                oSelection.Add oSubHB
                oSelection.Search ("'Generative Shape Design'.'Geometric feature'.Color=White,sel")
                oSelection.Search ("'Generative Shape Design'.'Geometric feature'.IsRoot=TRUE,sel")
                Set oColx = New Collection
                If oSelection.Count > 0 Then
                     For j = 1 To oSelection.Count
                        oColx.Add oSelection.Item(j).Value
                     Next
                End If
                For Each ObjtoDel In oColx
                    oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(ObjtoDel)
                Next
            Next
        End If
    Next
    'Create hole symmetry
    'Scan all bodies
    For Each oSubBody In oPanelPart.Bodies
        If oSubBody.Name Like "Inserts_*" Then
            
            'Insert's Part Number
            sPartNumber = Split(oSubBody.Name, "Inserts_")(1)
            
            'Create symmetry
            If Trim(UCase(oHoleDiaInfo.GetItemByKey(sPartNumber).GetItemByKey("Symmetry"))) = "YES" Then
                
                Set oSubHB = RetrieveSubHB(sPartNumber, oConstructnGeo)
                Set oSymmetryPlane = RetrieveSymmetryPlane(oSubHB, oHoleSupport, oHoleBottom, dPanelThk, oSPAWorkbench, oPanelPart)
                Set oReference = oPanelPart.CreateReferenceFromObject(oSymmetryPlane)
                oPanelPart.InWorkObject = oSubBody
                Set oSymmetry = oSF.AddNewSymmetry2(oReference)
                oPanelPart.InWorkObject = oPartBody
            End If
        End If
    Next
    On Error Resume Next
        oPanelPart.Update
    On Error GoTo 0
    
    If bShowGeoset Then
        For Each oSubHB In oConstructnGeo.HybridBodies
            If UCase(oSubHB.Name) Like "INSERTS_*" Then
               oHSF.GSMVisibility oPanelPart.CreateReferenceFromObject(oSubHB), 1
            End If
        Next
        oHSF.GSMVisibility oPanelPart.CreateReferenceFromObject(oConstructnGeo), 1
        
        Call MsgBox("Some of the hole(s) are not created" & vbCrLf & "Please analyse the result", vbExclamation)
        
    End If
    oSelection.Clear
TheEnd:
    frmProgress.Hide
End Sub

Private Sub SwapToAssemblyWorkbench()

Dim sWkb As String

sWkb = CATIA.GetWorkbenchId
If sWkb <> "Assembly" Then
    CATIA.StartWorkbench "Assembly"
End If

End Sub

Private Function RetrieveSubHB(ByVal sPartNumber As String, ByVal oParentHB As HybridBody) As HybridBody
'Checks if the Geoset exist. If not we creates it

Dim oHB As HybridBody
Dim bHBExist As Boolean

bHBExist = False

'Scan all geoset under the parent geoset
If oParentHB.HybridBodies.Count <> 0 Then
    For Each oHB In oParentHB.HybridBodies
        If oHB.Name = "Inserts_" & sPartNumber Then
            bHBExist = True
            Exit For
        End If
    Next
End If

'If the geoset doesn't exist we creates it
If Not bHBExist Then
    Set oHB = oParentHB.HybridBodies.Add
    oHB.Name = "Inserts_" & sPartNumber
End If

Set RetrieveSubHB = oHB

End Function

Private Function RetrieveSymmetryPlane(ByVal oSubHB As HybridBody, ByVal oPlane1 As HybridShape, ByVal oPlane2 As HybridShape, _
                                       ByVal dPanelThk As Double, ByVal oSPAWorkbench As Workbench, ByVal oPart As Part)
'Retreive the Symmetry plane in a specific Geoset or create it if it doesn't exist

Dim oHybridShape As HybridShape
Dim bHBExist As Boolean
Dim oHSF As HybridShapeFactory
Dim oReference As Reference
Dim oReference2 As Reference
Dim dDistance As Double

'Initial set
bHBExist = False
Set oHSF = oPart.HybridShapeFactory

'Retreive existing plane
For Each oHybridShape In oSubHB.HybridShapes
    If oHybridShape.Name = "SymmetryPlane" Then
        bHBExist = True
        Exit For
    End If
Next

'Create symmetry plane
If bHBExist = False Then
    
    Set oHybridShape = oHSF.AddNewPlaneOffset(oPlane1, dPanelThk / 2, False)
    oSubHB.AppendHybridShape oHybridShape
    oHybridShape.Name = "SymmetryPlane"
    oPart.UpdateObject oSubHB
    
    Set oReference = oPart.CreateReferenceFromObject(oPlane2)
    Set oReference2 = oPart.CreateReferenceFromObject(oHybridShape)
    Set oMeasurable = oSPAWorkbench.GetMeasurable(oReference)
    If oMeasurable.GetMinimumDistance(oReference2) <> dPanelThk / 2 Then
        oHSF.DeleteObjectForDatum oReference2
        Set oHybridShape = oHSF.AddNewPlaneOffset(oPlane1, -dPanelThk / 2, False)
        oSubHB.AppendHybridShape oHybridShape
        oHybridShape.Name = "SymmetryPlane"
        oPart.UpdateObject oSubHB
    End If
End If


Set RetrieveSymmetryPlane = oHybridShape
End Function

Private Function RetrieveSubBody(ByVal sPartNumber As String, ByVal oPart As Part, ByVal oPartBody As Body, oSF As ShapeFactory) As Body
'Check if the body exist. If not we creates it and assemble it to the PartBody

Dim obody As Body
Dim bExist As Boolean
Dim oAssemble As Assemble

bExist = False

For Each obody In oPart.Bodies
    If obody.Name = "Inserts_" & sPartNumber Then
        bExist = True
        Exit For
    End If
Next

If Not bExist Then
    Set obody = oPart.Bodies.Add
    obody.Name = "Inserts_" & sPartNumber
    oPart.InWorkObject = oPartBody
    Set oAssemble = oSF.AddNewAssemble(obody)
    oAssemble.Name = "Assemble_" & sPartNumber
    oPart.UpdateObject oAssemble
End If

Set RetrieveSubBody = obody

End Function

Private Function dAngleBetweenplanes(p1() As Double, p2() As Double, p3() As Double, p4() As Double) As Double
    Dim Vector1(2) As Double
    Dim Vector2(2) As Double
    
    '***Calculating Angle between Planes

    Vector1(0) = p2(0) - p1(0): Vector1(1) = p2(1) - p1(1): Vector1(2) = p2(2) - p1(3)
    Vector2(0) = p4(0) - p3(0): Vector2(1) = p4(1) - p3(1): Vector2(2) = p4(2) - p3(2)
    dAngleBetweenplanes = ArcCos(DotProd(Vector2, Vector1) / (LengthVector(Vector2) * LengthVector(Vector1)))
    dAngleBetweenplanes = Abs(Round(dAngle * 180 / PI, 9))
    
End Function
Private Function bPlanesAreSame(ByRef oPln1 As Plane, oPln2 As Plane, ByRef oPanelPart As Part) As Boolean
   
    Dim oMeasureable 'As Measurable
    Dim TOLERANCE As Double
    Dim TOLAngle As Double
    TOLERANCE = 0.001
    TOLAngle = 0.01
    
    Set oSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    Set oMeasureable = oSPAWorkbench.GetMeasurable(oPanelPart.CreateReferenceFromObject(oPln1))
    
    If Abs(oMeasureable.GetMinimumDistance(oPanelPart.CreateReferenceFromObject(oPln2))) < TOLERANCE Then
        If Abs(oMeasureable.GetAngleBetween(oPanelPart.CreateReferenceFromObject(oPln2))) < TOLAngle Then bPlanesAreSame = True
    End If
End Function
Private Function scanExistingHoles(ByVal oPanelObject As clsPanelCheck, _
                                   ByVal oHoleDiainfobyDia As clsCollection, _
                                   Plane1Pt1() As Double, _
                                   Plane1Pt2() As Double, _
                                   Plane2Pt1() As Double, _
                                   Plane2Pt2() As Double) As Collection
                                
    Dim oPart As Part
    Dim oHole As Hole
    Dim oShape 'As Shape
    Dim obody As Body
    Dim oSF As ShapeFactory
    Dim oHoleCol As Collection
    Dim oItem As clsCollection
    Dim dHoledia As Double
    Dim oHoleOrigin(2) As Variant
    Dim HoleGlobalstart(2) As Double
    Dim oHoleDir(2) As Variant
    Dim oHoleEnd(2) As Variant
    Dim sParentListPanel As String
    Dim PanelThickness As Double
    Dim dVector(9 To 11) As Variant
    Dim dVectorReturn(9 To 11) As Variant
    Dim dDistP1 As Double, dDistP2 As Double
    Dim MinDistanceBetweenHoleAndPlane As Double
    
   '****Ideally should be zero .... but just in case
    MinDistanceBetweenHoleAndPlane = 0.001
    
    Set oHoleCol = New Collection
   
    PanelThickness = oPanelObject.PanelThickness
    Set oPart = oPanelObject.PanelPart
    Dim dGlobalMatrixPanel(11) As Variant
    sParentListPanel = BuildParentList(oPanelObject.PanelInstance)
    Call CalculateInstanceGlobalPositionMatrix(sParentListPanel, dGlobalMatrixPanel)
    Set oHoleCol = New Collection
    For Each obody In oPart.Bodies
        For Each oShape In obody.Shapes
            If TypeName(oShape) = "Hole" Then
                Set oHole = oShape
                dHoledia = oHole.Diameter.Value
                If oHoleDiainfobyDia.Exists(CStr(Round(dHoledia, 6))) Then
                    oShape.GetOrigin oHoleOrigin
                    oShape.GetDirection oHoleDir
                    oHoleEnd(0) = oHoleDir(0) * PanelThickness + oHoleOrigin(0): oHoleEnd(1) = oHoleDir(1) * PanelThickness + oHoleOrigin(1): oHoleEnd(2) = oHoleDir(2) * PanelThickness + oHoleOrigin(2)
                    '***Calculate global start point
                    dVector(9) = oHoleOrigin(0): dVector(10) = oHoleOrigin(1): dVector(11) = oHoleOrigin(2)
                    Call MultiplyVector(dVector, dGlobalMatrixPanel, dVectorReturn, True)
                    HoleGlobalstart(0) = dVectorReturn(9): HoleGlobalstart(1) = dVectorReturn(10): HoleGlobalstart(2) = dVectorReturn(11)
                    dDistP1 = dDistancePointToPlane(HoleGlobalstart, Plane1Pt1, Plane1Pt2)
                    dDistP2 = dDistancePointToPlane(HoleGlobalstart, Plane2Pt1, Plane2Pt2)
                    If Abs(dDistP1) <= MinDistanceBetweenHoleAndPlane Or Abs(dDistP2) <= MinDistanceBetweenHoleAndPlane Then
                        If Abs(dDistP1) < Abs(dDistP2) Then
                            bSide1 = True
                        Else
                            bSide1 = False
                        End If
                        Set oItem = New clsCollection
                        oItem.Add "HoleDia", dHoledia
                        oItem.Add "InsertNumber", oHoleDiainfobyDia.GetItemByKey(CStr(Round(dHoledia, 6))).GetItemByKey("InsertNumber")
                        oItem.Add "HoleGlobalOrigin", HoleGlobalstart
                        oItem.Add "HoleLocalOrigin", oHoleOrigin
                        oItem.Add "HoleOnSide1", bSide1
                        oItem.Add "HoleObject", oHole
                        oHoleCol.Add oItem
                        If Not oPanelObject.PanelPart.IsInactive(oHole) Then oPanelObject.PanelPart.Inactivate oHole
                    End If
                End If
            End If
        Next
    Next
    
Set scanExistingHoles = oHoleCol
End Function


