Attribute VB_Name = "CheckPanelAssy"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'********************************************************************************
'* Name: Check Panel Assy
'* Purpose: Analyse inserts in a panel.
'*
'* Assumption:
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************


Private Const BM_CLICK As Integer = &HF5

'************************************************

Private Const dOffsetValue As Double = 0.9144   '*** 0.036 in for offset of intersection planes from Panel face
Private Const dMaxmimumMisalignmentTol As Double = 6.35    '*** if insert is more than tol from hole then ignore insert hole


Private oMeasurable 'As Measurable
Private oMeasurable2
Private oMeasurable3
Private oMeasureablePlane1 As Measurable
Private oMeasureablePlane2 As Measurable
Private oSPAWorkbench As Workbench
Private oHSF As HybridShapeFactory
Private oPanelPart As Part
Private oAutomationGeo As HybridBody
Private oGeoset As HybridBody
Private oTempGeosets As HybridBodies
Private oInsertHoleGeoset As HybridBody
Private oHS As HybridShapes
Private oRef1 As Reference
Private oRef2 As Reference
Private oPartBodyRef As Reference
Private oThicknessPlane1 As HybridShapePlaneOffset
Private oThicknessPlane2 As HybridShapePlaneOffset
Private oOffsetPlane1 As HybridShapePlaneOffset
Private oOffsetPlane2 As HybridShapePlaneOffset
Private oFill1 As HybridShapeFill
Private oFill2 As HybridShapeFill
Private oPanelBoundary1 As HybridShapeCurveExplicit
Private oPanelBoundary2 As HybridShapeCurveExplicit
Private oPartBody As Body
Private i As Integer
Private j As Integer
Private iHoleDia As Integer
Private iPNum As Integer
Private iPanelThk As Integer
Private iPanelCompatible As Integer
Private bIsInsertHole As Boolean
Private oIntersect As HybridShapeIntersection

Private ListOfTongueAssembles As Collection
Private MultiDomainArray As Variant
Private ListOfAllInserts As Collection
Private ListOfInsertLines As Collection
Private ListOfInsertObjects As Collection
Private ListOfInsertKeys As Collection
Private ListOfCurvesToCompare As Collection
Private InsertHoleCollection  As Collection
Private oSelectedAssy As AnyObject
Private oInstance As Product
Private oPoint As Point 'HybridShapePointCoord
Private oAssociatedPoint As HybridShapePointCoord
Private oLine As Line
Private oDirection As HybridShapeDirection
Private sParentList As String
Private sPanelPathInAssy As String

Private dGlobalMatrixInsert(11) As Variant
Private dGlobalMatrixInsertCopy As Variant
Private dGlobalMatrixPanel(11) As Variant
Private dLocalMatrixInsert(11) As Variant
Private dProjectCoord(2) As Variant
Private dHoleCenter(2)
Private dHoledia As Double
Private dMaxLength As Double

Private sString As String
Private sReport As String
Private sInsertReportCollection As Collection
Private sInsertkeyCollection As Collection
Private dGap As Double
Private iProgressmax As Integer
Private sLineKey As String

Private sWkb As String ' workbench ID

Public Sub LaunchAnalysis(ByRef oPanelCheck As clsPanelCheck, _
                          ByRef sSourceWindow As Window, iMax As Integer, ictr As Integer)
    Dim oSel As Selection
    Dim sWrongHole As String
    Dim bDontDelete As Boolean
    Dim dLengthBoundary1 As Double, dLengthBoundary2 As Double
    Dim dVector(9 To 11) As Double
    Dim dVectorReturn(9 To 11) As Double
    
    '*****Assigning values
    If sInsertdistanceCol.Count = 0 Then
       
    End If
    bDontDelete = False
    Set ListOfTongueAssembles = New Collection
    
    '***Accessing Column name from the table
    For i = 1 To UBound(sImportedInsertHoleTable, 2)
        If sImportedInsertHoleTable(1, i) = "Hole Diameter (mm)" Then iHoleDia = i
        If sImportedInsertHoleTable(1, i) = "Part Number" Then iPNum = i
        If sImportedInsertHoleTable(1, i) = "PanelThickness_mm" Then iPanelThk = i
        If sImportedInsertHoleTable(1, i) = "CompatiblePanels" Then iPanelCompatible = i
    Next
    '***Build Panelinstance parent list
    '*** This will be used to populate holes without inserts
    sPanelPathInAssy = BuildParentList(oPanelCheck.PanelInstance)
    
    '*** Populate List Of All Inserts available under the assy
    Set ListOfAllInserts = New Collection
    Call GenerateItemlist(oPanelCheck.PanelAssyInstance)
    '***if user presses cancel on progress toolbar
    If bCancelAction = True Then
        GoTo cleanExit:
    End If
    
    Call frmProgress.progressBarRepaint("Analysing Panel assies", iMax, ictr, "Creating Geometry inside panel for comparison", , , "Elapsed time: " & Format((DateTime.Timer - timeStart) / 86400, "hh:mm:ss"))
    
    '***Panel Part
    Set oPanelPart = oPanelCheck.PanelPart
    Set oHSF = oPanelPart.HybridShapeFactory
    Set oSF = oPanelPart.ShapeFactory
    Set oPartBody = oPanelPart.Bodies.Item("PartBody")
    Set oPartBodyRef = oPanelPart.CreateReferenceFromObject(oPartBody)
    Set oSPAWorkbench = oPanelCheck.PanelInstance.ReferenceProduct.Parent.GetWorkbench("SPAWorkbench")

    '*** Create Geometrical Set
    Set oAutomationGeo = oPanelPart.HybridBodies.Add
    oAutomationGeo.Name = "AutomationTemp_To_Delete"
    Call SwapToAssemblyWorkbench '*** to ensure that geoset is hidden & not instance
    Set oSel = oPanelCheck.PanelInstance.ReferenceProduct.Parent.Selection
    oSel.Clear
    oSel.Add oAutomationGeo
    oSel.VisProperties.SetShow catVisPropertyNoShowAttr
    
    '***Create geoset for insert holes
    Set oTempGeosets = oAutomationGeo.HybridBodies
    Set oInsertHoleGeoset = oTempGeosets.Add
    oInsertHoleGeoset.Name = "InsertHole"
    
    '*** Create geoset for first intersect geometery
    Set oTempGeosets = oAutomationGeo.HybridBodies
    Set oGeoset = oTempGeosets.Add
    oGeoset.Name = "FirstIntersect"
    
    '*** Create MeasurementPlane for face 1
    Set oThicknessPlane1 = oHSF.AddNewPlaneOffset(oPanelCheck.PanelPlane1, 0, False)
    oThicknessPlane1.Name = "Face1"
    oGeoset.AppendHybridShape oThicknessPlane1
    oThicknessPlane1.Compute
    Set oMeasureablePlane1 = oSPAWorkbench.GetMeasurable(oThicknessPlane1)
    
    '*** Create MeasurementPlane for face 2
    Set oThicknessPlane2 = oHSF.AddNewPlaneOffset(oPanelCheck.PanelPlane2, 0, False)
    oThicknessPlane1.Name = "Face2"
    oGeoset.AppendHybridShape oThicknessPlane2
    oThicknessPlane2.Compute
    Set oMeasureablePlane2 = oSPAWorkbench.GetMeasurable(oThicknessPlane2)
    
    '*** create offset plane for face 1
    Set oOffsetPlane1 = oHSF.AddNewPlaneOffset(oPanelCheck.PanelPlane1, dOffsetValue, False)
    oOffsetPlane1.Name = "OffsetPlane1"
    oGeoset.AppendHybridShape oOffsetPlane1
    oOffsetPlane1.Compute
    If oMeasureablePlane2.GetMinimumDistance(oPanelPart.CreateReferenceFromObject(oOffsetPlane1)) > oPanelCheck.PanelThickness Then
        Set oOffsetPlane1 = oHSF.AddNewPlaneOffset(oPanelCheck.PanelPlane1, -dOffsetValue, False)
        oPanelCheck.FlipPlane1Dir = True
        oOffsetPlane1.Name = "OffsetPlane1"
        oGeoset.AppendHybridShape oOffsetPlane1
        oOffsetPlane1.Compute
    End If
    
    '*** create intersect with part body & Plane offsetfrom Face 1
    Set oRef1 = oPanelPart.CreateReferenceFromGeometry(oOffsetPlane1)
    Set oIntersect = oHSF.AddNewIntersection(oRef1, oPartBodyRef)
    oGeoset.AppendHybridShape oIntersect
    oIntersect.Compute
    
    Set ListOfCurvesToCompare = New Collection
    '*** Break down multidomain intersect into individual domains
    '*** Populate geoset first intersect with all curves except circles which corresponds to insert holes
    '*** Get center of circle of insert holes & populate them in geoset Inserthole
    '*** Populate listOfCurvesTocompare with curves in FirstIntersect geoset
    dMaxLength = 0
    Set oRef1 = oPanelPart.CreateReferenceFromObject(oIntersect)
    Set InsertHoleCollection = New Collection
    MultiDomainArray = oHSF.AddNewDatums(oRef1)
    For i = 0 To UBound(MultiDomainArray)
        Call frmProgress.progressBarRepaint("Analysing Panel assies", iMax, ictr, "Identifying insert holes", , , "Elapsed time: " & Format((DateTime.Timer - timeStart) / 86400, "hh:mm:ss"))
        oGeoset.AppendHybridShape MultiDomainArray(i)
        ListOfCurvesToCompare.Add MultiDomainArray(i), MultiDomainArray(i).Name
        'Retrieve measurable object of the curve
        Set oRef2 = oPanelPart.CreateReferenceFromObject(MultiDomainArray(i))
        Set oMeasurable = oSPAWorkbench.GetMeasurable(oRef2)
        If oMeasurable.Length > dMaxLength Then
            dMaxLength = oMeasurable.Length
            Set oPanelBoundary1 = MultiDomainArray(i)   '*** find biggest curve
            dLengthBoundary1 = dMaxLength
        End If
        On Error Resume Next
            Err.Clear
            dHoledia = oMeasurable.Radius * 2
        'If Err.Number = 0 it means the curve is a circle.
        If Err.Number = 0 Then
            On Error GoTo 0
            For j = 2 To UBound(sImportedInsertHoleTable, 1)
                If CStr(Round(dHoledia, 6)) = CStr(sImportedInsertHoleTable(j, iHoleDia)) Then
                    oMeasurable.GetCenter dHoleCenter
                    'oHSF.DeleteObjectForDatum MultiDomainArray(i)
                    ListOfCurvesToCompare.Remove MultiDomainArray(i).Name
                    Set oPoint = oHSF.AddNewPointCoord(dHoleCenter(0), dHoleCenter(1), dHoleCenter(2))
                    oInsertHoleGeoset.AppendHybridShape oPoint
                    oPoint.Name = sImportedInsertHoleTable(j, iPNum) + "_" + CStr(i)
                    oPoint.Compute
                    InsertHoleCollection.Add oPoint, oPoint.Name
                End If
            Next
        End If
        On Error GoTo 0
    Next

    '***if user presses cancel on progress toolbar
    If bCancelAction = True Then
        GoTo cleanExit:
    End If
    
    '*** create fill
    Set oFill1 = oHSF.AddNewFill()
    oFill1.AddBound oPanelBoundary1
    oFill1.Continuity = 0
    oGeoset.AppendHybridShape oFill1
    oFill1.Compute
    
    '***Create geoset for second intersect geometry
    Set oGeoset = oTempGeosets.Add
    oGeoset.Name = "SecondIntersect"
    oPanelPart.InWorkObject = oGeoset
    
    '*** create offset plane for face 1
    Set oOffsetPlane2 = oHSF.AddNewPlaneOffset(oPanelCheck.PanelPlane2, dOffsetValue, False)
    oOffsetPlane2.Name = "OffsetPlane2"
    oGeoset.AppendHybridShape oOffsetPlane2
    oOffsetPlane2.Compute
    If oMeasureablePlane1.GetMinimumDistance(oPanelPart.CreateReferenceFromObject(oOffsetPlane2)) > oPanelCheck.PanelThickness Then
        Set oOffsetPlane2 = oHSF.AddNewPlaneOffset(oPanelCheck.PanelPlane2, -dOffsetValue, False)
        oPanelCheck.FlipPlane2Dir = True
        oOffsetPlane2.Name = "OffsetPlane2"
        oGeoset.AppendHybridShape oOffsetPlane2
        oOffsetPlane2.Compute
    End If
        
    '*** create intersect with part body & Plane offsetfrom Face 1
    Set oRef1 = oPanelPart.CreateReferenceFromGeometry(oOffsetPlane2)
    Set oIntersect = oHSF.AddNewIntersection(oRef1, oPartBodyRef)
    oGeoset.AppendHybridShape oIntersect
    oIntersect.Compute
    
    '*** break down multidomain intersect into individual domains and keep the biggest domain
    '*** Populate listOfCurvesTocompare with curves in SecondIntersect geoset
    Set oRef1 = oPanelPart.CreateReferenceFromObject(oIntersect)
    MultiDomainArray = oHSF.AddNewDatums(oRef1)
    dMaxLength = 0
    For i = 0 To UBound(MultiDomainArray)
        oGeoset.AppendHybridShape MultiDomainArray(i)
        Set oRef1 = oPanelPart.CreateReferenceFromObject(MultiDomainArray(i))
        Set oMeasurable = oSPAWorkbench.GetMeasurable(oRef1)
        If oMeasurable.Length > dMaxLength Then
            dMaxLength = oMeasurable.Length
            Set oPanelBoundary2 = MultiDomainArray(i)
            dLengthBoundary2 = dMaxLength
        End If
    Next
    ListOfCurvesToCompare.Add oPanelBoundary2, oPanelBoundary2.Name
    
    '*** create fill
    Set oFill2 = oHSF.AddNewFill()
    oFill2.AddBound oPanelBoundary2
    oFill2.Continuity = 0
    oGeoset.AppendHybridShape oFill2
    oFill2.Compute
    
    '*** Create Tongue boundary if checked in form
    Dim sDoNotKeep As String
    If dLengthBoundary1 > dLengthBoundary2 Then sDoNotKeep = oPanelBoundary2.Name Else sDoNotKeep = oPanelBoundary1.Name

    For i = 1 To ListOfCurvesToCompare.Count
        Set oRef1 = oPanelPart.CreateReferenceFromObject(ListOfCurvesToCompare.Item(i))
        Set oMeasurable = oSPAWorkbench.GetMeasurable(oRef1)
        If ListOfCurvesToCompare.Item(i).Name <> sDoNotKeep Then
            Call CreateTongueBoundary(ListOfCurvesToCompare.Item(i), oSel)
        End If
    Next

    '*** Create geoset for InsertGeometery
    Set oTempGeosets = oAutomationGeo.HybridBodies
    Set oGeoset = oTempGeosets.Add
    oGeoset.Name = "InsertDefinition"
    Set oMeasurable = oSPAWorkbench.GetMeasurable(oFill1)
    Set oMeasurable2 = oSPAWorkbench.GetMeasurable(oFill2)
    Set ListOfInsertLines = New Collection
    Set ListOfInsertObjects = New Collection
    Set sInsertkeyCollection = New Collection
    For i = 1 To ListOfAllInserts.Count
        Call frmProgress.progressBarRepaint("Analysing Panel assies", iMax, ictr, "Scanning inserts belonging to the panel", ListOfAllInserts.Count, i, "Elapsed time: " & Format((DateTime.Timer - timeStart) / 86400, "hh:mm:ss"))
        '***if user presses cancel on progress toolbar
        If bCancelAction = True Then
            GoTo cleanExit:
        End If
        
        Set oInstance = ListOfAllInserts.Item(i)
        '***Calculate global position of panel
        Call CalculateInstanceGlobalPositionMatrix(BuildParentList(oPanelCheck.PanelInstance), dGlobalMatrixPanel)
        '***Calculate global position of insert
        sParentList = BuildParentList(oInstance)
        Call CalculateInstanceGlobalPositionMatrix(sParentList, dGlobalMatrixInsert)
        dGlobalMatrixInsertCopy = dGlobalMatrixInsert
        '***Calculate position matrix of center point of insert in the panel CATPart: dLocalMatrixInsert
        Call SubstractVector(dGlobalMatrixInsert, dGlobalMatrixPanel)
        Call InvertMatrix(dGlobalMatrixPanel)
        Call MultiplyVector(dGlobalMatrixInsert, dGlobalMatrixPanel, dLocalMatrixInsert)
        Call MultiplyMatrix(dGlobalMatrixInsert, dGlobalMatrixPanel, dLocalMatrixInsert)
        '***Create point in the panel part using dLocalMatrixInsert
        Set oPoint = oHSF.AddNewPointCoord(dLocalMatrixInsert(9), dLocalMatrixInsert(10), dLocalMatrixInsert(11))
        oGeoset.AppendHybridShape oPoint
        oPoint.Compute
        '***Create line in the panel part using dLocalMatrixInsert. Line is parallel to Z axis of insert.
        Set oDirection = oHSF.AddNewDirectionByCoord(dLocalMatrixInsert(6), dLocalMatrixInsert(7), dLocalMatrixInsert(8))
        Set oLine = oHSF.AddNewLinePtDir(oPoint, oDirection, 0, oPanelCheck.PanelThickness, False)
        oGeoset.AppendHybridShape oLine
        oLine.Name = oInstance.ReferenceProduct.Name + "_" + sParentList
        oLine.Compute
        '*** verify that line is inside panel
        '*** distance from either of fills should be 0
        If oMeasurable.GetMinimumDistance(oLine) = 0 Or oMeasurable2.GetMinimumDistance(oLine) = 0 Then
            ListOfInsertLines.Add oLine, sParentList
            ListOfInsertObjects.Add ListOfAllInserts.Item(i), sParentList
            sParentList = AppendGlobalMatrix(sParentList, dGlobalMatrixInsertCopy)
            sInsertkeyCollection.Add sParentList, sParentList '*** to be used later for report generation
        End If
    Next
    
    '*** Analyse Inserts
    
    Set sInsertReportCollection = New Collection
    k = 1
    iProgressmax = ListOfInsertLines.Count
    
    For Each oLine In ListOfInsertLines
        Call frmProgress.progressBarRepaint("Analysing Panel assies", iMax, ictr, "Verifying inserts and holes...", iProgressmax, k, "Elapsed time: " & Format((DateTime.Timer - timeStart) / 86400, "hh:mm:ss"))
        k = k + 1
        '***if user presses cancel on progress toolbar
        If bCancelAction = True Then
            GoTo cleanExit:
        End If
        
        sReport = ""
        Set oRef1 = oPanelPart.CreateReferenceFromObject(oLine)
        Set oMeasurable = oSPAWorkbench.GetMeasurable(oRef1)
        sLineKey = Right(oLine.Name, Len(oLine.Name) - InStr(oLine.Name, "_")) '*** unique name to retrive objects from collections
        
        '*** check that insert is flush with the panel
        '*** -----------------------------------***
        
        Dim LineStartPt(2)
        Dim dDistanceLineStartToPlane1 As Double
        Dim dDistanceLineStartToPlane2 As Double
        Dim U(2) As Double, V(2) As Double, W As Variant, p3(2) As Double
        Dim bMisaligned As Boolean
        
        Set oMeasureable2 = oSPAWorkbench.GetMeasurable(oLine.Point)
        oMeasureable2.GetPoint LineStartPt  '**** array of coordinates
        
        '***Distance of Startpoint from Plane 1
        U(0) = oPanelCheck.ArPlane1(3): U(1) = oPanelCheck.ArPlane1(4): U(2) = oPanelCheck.ArPlane1(5)
        V(0) = oPanelCheck.ArPlane1(6): V(1) = oPanelCheck.ArPlane1(7): V(2) = oPanelCheck.ArPlane1(8)
        Call CrossProduct(U, V, W)         '*** w is a plane vector perpendicluar to it
        p3(0) = oPanelCheck.ArPlane1(0) - LineStartPt(0): p3(1) = oPanelCheck.ArPlane1(1) - LineStartPt(1): p3(2) = oPanelCheck.ArPlane1(2) - LineStartPt(2)
        sn = DotProd(W, p3)
        sd = LengthVector(W)
        dDistanceLineStartToPlane1 = Abs(sn / sd)
        
        '***Distance of Startpoint from Plane 2
        U(0) = oPanelCheck.ArPlane2(3): U(1) = oPanelCheck.ArPlane2(4): U(2) = oPanelCheck.ArPlane2(5)
        V(0) = oPanelCheck.ArPlane2(6): V(1) = oPanelCheck.ArPlane2(7): V(2) = oPanelCheck.ArPlane2(8)
        Call CrossProduct(U, V, W)         '*** w is a plane vector perpendicluar to it
        p3(0) = oPanelCheck.ArPlane2(0) - LineStartPt(0): p3(1) = oPanelCheck.ArPlane2(1) - LineStartPt(1): p3(2) = oPanelCheck.ArPlane2(2) - LineStartPt(2)
        sn = DotProd(W, p3)
        sd = LengthVector(W)
        dDistanceLineStartToPlane2 = Abs(sn / sd)
        
        bMisaligned = True
        If dDistanceLineStartToPlane1 < dFlushnessTol And (dDistanceLineStartToPlane2 - oPanelCheck.PanelThickness) < dFlushnessTol Then bMisaligned = False
        If dDistanceLineStartToPlane2 < dFlushnessTol And (dDistanceLineStartToPlane1 - oPanelCheck.PanelThickness) < dFlushnessTol Then bMisaligned = False
        
        If bMisaligned Then sReport = sReport + "---Insert not flush with Panel;"
        
        '***checking that insert is compatible with panel
        '*** -----------------------------------***
        
        For i = 2 To UBound(sImportedInsertHoleTable, 1)
            If Left(oLine.Name, InStr(oLine.Name, "_") - 1) = sImportedInsertHoleTable(i, iPNum) Then
                If oPanelCheck.PanelMaterial <> "" Then '*** Check compatibility as per stress  ': update 28Jan2018
                    If InStr(1, sImportedInsertHoleTable(i, iPanelCompatible), Replace(oPanelCheck.PanelMaterial, " ", "") & ";") = 0 Then
                       sReport = sReport + "---" + sMsgStressIncompatible + ";"
                    End If
                End If
                If CStr(sImportedInsertHoleTable(i, iPanelThk)) <> CStr(Round(oPanelCheck.PanelThickness, 3)) Then
                    sReport = sReport + "---" + CStr(sImportedInsertHoleTable(i, iPanelThk) / 25.4) + "inch insert in " + CStr(Round(oPanelCheck.PanelThickness, 3) / 25.4) + "inch panel" + ";"
                End If
            End If
        Next
        
        '***Checking edge distance with panels
        '*** -----------------------------------***
        Dim dEdgeDist As Double
        For i = 1 To ListOfCurvesToCompare.Count
            Set oRef2 = oPanelPart.CreateReferenceFromObject(ListOfCurvesToCompare.Item(i))
            dEdgeDist = oMeasurable.GetMinimumDistance(oRef2)
            If Round(dEdgeDist, 3) < dMiminumGapAllowed Then
                sReport = sReport + "---Insert is too close to Panel's edge [Gap " + CStr(Round(dEdgeDist / 25.4, 5)) + "in];---[Add 0.125in for edge with epocast];"
                Exit For
            End If
        Next
        '*** check edge distance with tongues
        '***-------------------------------*
        'If bChkTongue Then
            If ListOfTongueAssembles.Count > 0 Then
                For i = 1 To ListOfTongueAssembles.Count
                    Set oRef2 = oPanelPart.CreateReferenceFromObject(ListOfTongueAssembles.Item(i))
                    dEdgeDist = oMeasurable.GetMinimumDistance(oRef2)
                    If Round(dEdgeDist, 3) < dMinGapTongue Then
                        sReport = sReport + "---Insert is too close to tongue [Gap " + CStr(Round(dEdgeDist / 25.4, 5)) + "in];"
                    End If
                Next
            End If
        'End If
        '***checking distance between inserts
        '*** -----------------------------------***
        Dim dMeasuredInsertGap As Double
        For i = 1 To ListOfInsertLines.Count
            If ListOfInsertLines.Item(i).Name <> oLine.Name Then
                Set oRef2 = oPanelPart.CreateReferenceFromObject(ListOfInsertLines.Item(i))
                dMeasuredInsertGap = oMeasurable.GetMinimumDistance(oRef2)
                For j = 1 To sInsertdistanceCol.Count
                    If Round(dMeasuredInsertGap, 3) < CDbl(sInsertdistanceCol.Item(j)) Then
                        sString = ListOfInsertLines.Item(i).Name
                        sString = ListOfAllInserts.Item(Right(sString, Len(sString) - InStr(sString, "_"))).Name
                        sReport = sReport + "---Insert is too close to " + sString + "[Gap " + CStr(Round(dMeasuredInsertGap / 25.4, 6)) + "in]" & "|" & sInsertdistanceCol.Item(j) & ";"
                    End If
                Next
            End If
        Next
        '*** check Panel hole center with inserts
        '*** -----------------------------------***
        
        dGap = 20 '** arbitrary value to begin
        '*** find closest hole center
        Set oAssociatedPoint = Nothing
        For Each oPoint In InsertHoleCollection
            Set oRef2 = oPanelPart.CreateReferenceFromObject(oPoint)
            If oMeasurable.GetMinimumDistance(oRef2) < dGap Then
                dGap = oMeasurable.GetMinimumDistance(oRef2)
                Set oAssociatedPoint = oPoint
            End If
        Next
        
        '*** check that insert & hole are same
        '*** -----------------------------------***
        
        sWrongHole = ""
        If Not oAssociatedPoint Is Nothing Then
            If Left(oAssociatedPoint.Name, InStr(oAssociatedPoint.Name, "_") - 1) <> Left(oLine.Name, InStr(oLine.Name, "_") - 1) And dGap < dMaxmimumMisalignmentTol Then
                sWrongHole = "---Insert hole diameter doesn't match Insert Type;"
            End If
        End If
        sReport = sReport + sWrongHole  '*** GSUP-1957
        '*** if closest hole center is not colinear with line generate report
        If dGap > dMinMisalignmentTol And dGap < dMaxmimumMisalignmentTol Then
            sReport = sReport + "---Insert Hole is misaligned;"                 '*** GSUP-1957
            '*** since associated insert is found for hole then remove it from insert hole collection
            InsertHoleCollection.Remove oAssociatedPoint.Name
            sReport = sReport + sWrongHole
        ElseIf dGap > dMaxmimumMisalignmentTol Then
            sReport = sReport + "---Insert hole not found;"
        ElseIf dGap < dMinMisalignmentTol Then
            '*** since associated insert is found for hole then remove it from insert hole collection
            InsertHoleCollection.Remove oAssociatedPoint.Name
        End If
        
        '***Remove useless ";" from start & end
        While Right(sReport, 1) = ";"
            sReport = Mid(sReport, 1, Len(sReport) - 1)
        Wend
        
        While Left(sReport, 1) = ";"
            sReport = Mid(sReport, 2, Len(sReport))
        Wend
        '*** add one ";" at the end to give line break between reports
        If sReport <> "" Then sReport = sReport + ";"
        
        sInsertReportCollection.Add sReport, sLineKey '**** add report to report collection
        ListOfInsertLines.Remove sLineKey
    Next
    '***Add holes without inserts in report list

    If InsertHoleCollection.Count > 0 Then
        For Each oPoint In InsertHoleCollection
            sReport = "---Hole for " + Left(oPoint.Name, InStr(oPoint.Name, "_") - 1) + " without insert;"
            sInsertReportCollection.Add sReport, oPoint.Name
            ListOfInsertObjects.Add oPoint, oPoint.Name
            sInsertkeyCollection.Add "Hole_for " & Left(oPoint.Name, InStr(oPoint.Name, "_") - 1) & ";||" & sPointwithoutholematrix(oPoint, oPanelCheck)
        Next
    End If
    Set oPanelCheck.InsertReportCollection = sInsertReportCollection
    Set oPanelCheck.InsertkeyCollection = sInsertkeyCollection
    '*** make sure that automation geoset is not deleted
cleanExit:
    If Not oAutomationGeo Is Nothing And Not bDontDelete Then oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(oAutomationGeo)
    oPanelCheck.DeleteThicknessGeo
    If bCancelAction = True Then Call frmKBEMain.resetToolbar
End Sub
Private Sub GenerateItemlist(ByVal oAssembly As Product)
    Dim sParentList As String
    
    For Each oInstance In oAssembly.Products
        If oInstance.Products.Count > 0 Then
            Call GenerateItemlist(oInstance)  '*** recursion to scan all sub branches
        Else
            bCheck = False
            On Error Resume Next
            Err.Clear
            Set oTest = oInstance.ReferenceProduct '*** test to check if instance is cgr mode
            If Err.Number <> 0 Then oInstance.ApplyWorkMode DEFAULT_MODE
            Err.Clear
            Set oTest = oInstance.ReferenceProduct  '*** if instance cannot be loaded
            If Err.Number = 0 Then
                On Error GoTo 0
                For i = 2 To UBound(sImportedInsertHoleTable, 1)
                    If oInstance.ReferenceProduct.Name Like sImportedInsertHoleTable(i, iPNum) Then
                        bCheck = True
                    End If
                Next
                If bCheck = True Then
                    sParentList = BuildParentList(oInstance)
                    ListOfAllInserts.Add oInstance, sParentList
                End If
            End If
            On Error GoTo 0
        End If
    Next
End Sub
Private Sub SwapToAssemblyWorkbench()
    Dim sWkb As String
    
    sWkb = CATIA.GetWorkbenchId
    If sWkb <> "Assembly" Then
        CATIA.StartWorkbench "Assembly"
    End If
End Sub
'******* Tongue Boundary Creation using WINAPI***********
Private Sub CreateTongueBoundary(ByRef oSelectedBoundary As HybridShapeCurveExplicit, ByRef oSel As Selection)
    Dim oTongueCurveGeoset As HybridBody
    Dim oTongueAssemble As HybridShapeAssemble
    Dim Texto As String
    Dim OKhwnd As Long
    Dim oHSextract1 As HybridShapeExtract
    Dim oRefCurv1 As Reference
    Dim oRefCurv2 As Reference
    Dim iCounter As Integer
    Dim DisassembleForm_hwnd As Long
    
    Dim Dir(2)
    Dim POC(8)
    Dim PV(2) As Double
    Dim dLengthV  As Double
    '*** Create Geometrical Set
    Set oTempGeosets = oAutomationGeo.HybridBodies
    Set oTongueCurveGeoset = oTempGeosets.Add
    oTongueCurveGeoset.Name = "TongueBoundaries"
    '***swap workbenches
    CATIA.ActiveDocument.Selection.Clear
    CATIA.ActiveDocument.Selection.Add oTongueCurveGeoset
    'oSel.Clear
    'oSel.Add oTongueCurveGeoset
    sWkb = CATIA.GetWorkbenchId
    If sWkb <> "PrtCfg" Then
        CATIA.StartWorkbench "PrtCfg"
    End If
    '********
    
    oPanelPart.InWorkObject = oTongueCurveGeoset
    CATIA.ActiveDocument.Selection.Clear
    CATIA.ActiveDocument.Selection.Add oSelectedBoundary
    
    CATIA.StartCommand ("Disassemble")
    '***WinAPI Commands
    DisassembleForm_hwnd = FindWindowLike("Disassemble")
    button_hwnd = lChildWindow(DisassembleForm_hwnd, "Button", "OK")
    SendMessage button_hwnd, BM_CLICK, 0, 0
    '************
    
    oSel.Clear
    oSel.Add oTongueCurveGeoset
    oSel.Search "'Generative Shape Design'.Curve,sel"
    
    Set oTongueAssemble = oHSF.AddNewJoin(oRefCurv1, oRefCurv2)
    oTongueAssemble.SetConnex 0

    iCountCurve = 0
    For iCounter = 1 To oSel.Count
       Set oRef1 = oSel.Item2(iCounter).Reference
       Set oMeasurable = oSPAWorkbench.GetMeasurable(oRef1)
       oMeasurable.GetPointsOnCurve POC
       PV(0) = POC(6) - POC(0): PV(1) = POC(7) - POC(1): PV(2) = POC(8) - POC(2)
       dLengthV = LengthVector(PV)
       If Round(oMeasurable.Length, 3) = Round(dLengthV, 3) Then '****ensuring a straight line
            If (dLengthV > dLTonguewidthmin And dLengthV < dLTonguewidthmax) Or (dLengthV > dTTonguewidthmin And dLengthV < dTTonguewidthmax) Then
                'Set oRefCurv1 = oSel.Item2(iCounter).Reference
                Set oHSextract1 = oHSF.AddNewExtract(oRef1)
                oTongueCurveGeoset.AppendHybridShape oHSextract1
                oHSextract1.PropagationType = 3
                oHSextract1.IsFederated = False
                oHSextract1.ComplementaryExtract = False
                oHSextract1.Compute
                Set oRefCurv1 = oPanelPart.CreateReferenceFromObject(oHSextract1)
                oTongueAssemble.AddElement oRefCurv1
            End If
        End If
    Next
    
    oTongueAssemble.Compute
     
    If oTongueAssemble.GetElementsSize > 1 Then
        ListOfTongueAssembles.Add oTongueAssemble
        oAutomationGeo.AppendHybridShape oTongueAssemble
    Else
        oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(oTongueAssemble)
    End If

End Sub


'********************************************************
Public Function AppendGlobalMatrix(ByRef sInputString As String, ByVal dGlobalMatrix As Variant) As String

    AppendGlobalMatrix = sInputString & "||" & _
                        CStr(dGlobalMatrix(0)) & ";" & _
                        CStr(dGlobalMatrix(1)) & ";" & _
                        CStr(dGlobalMatrix(2)) & ";" & _
                        CStr(dGlobalMatrix(3)) & ";" & _
                        CStr(dGlobalMatrix(4)) & ";" & _
                        CStr(dGlobalMatrix(5)) & ";" & _
                        CStr(dGlobalMatrix(6)) & ";" & _
                        CStr(dGlobalMatrix(7)) & ";" & _
                        CStr(dGlobalMatrix(8)) & ";" & _
                        CStr(dGlobalMatrix(9)) & ";" & _
                        CStr(dGlobalMatrix(10)) & ";" & _
                        CStr(dGlobalMatrix(11))
End Function
Private Function sPointwithoutholematrix(ByRef oPoint As Point, ByRef oPanelCheck As clsPanelCheck) As String
    Dim dVector(9 To 11) As Double
    Dim dLocalMatrix(11) As Variant
    Dim dVectorReturn(9 To 11) As Double
    Dim U(2) As Double, V(2) As Double, W As Variant
    Dim oPlane(8)
    Dim bFlip As Boolean
    Dim oMeasurePl 'As Measurable

    If oMeasureablePlane1.GetMinimumDistance(oPanelPart.CreateReferenceFromObject(oPoint)) < oMeasureablePlane2.GetMinimumDistance(oPanelPart.CreateReferenceFromObject(oPoint)) Then
        Set oMeasurePl = oSPAWorkbench.GetMeasurable(oPanelCheck.PanelPlane1)
        oMeasurePl.GetPlane oPlane
        bFlip = oPanelCheck.FlipPlane1Dir
    Else
        Set oMeasurePl = oSPAWorkbench.GetMeasurable(oPanelCheck.PanelPlane2)
        oMeasurePl.GetPlane oPlane
        bFlip = oPanelCheck.FlipPlane2Dir
    End If
    U(0) = oPlane(3): U(1) = oPlane(4): U(2) = oPlane(5)
    V(0) = oPlane(6): V(1) = oPlane(7): V(2) = oPlane(8)
    If Not bFlip Then
        Call CrossProduct(U, V, W)
    Else
        Call CrossProduct(V, U, W)
    End If
    dLocalMatrix(0) = W(0): dLocalMatrix(1) = W(1): dLocalMatrix(2) = W(2)  ' x vector =z vector
    If Not bFlip Then
        dLocalMatrix(3) = V(0): dLocalMatrix(4) = V(1): dLocalMatrix(5) = V(2)
        dLocalMatrix(6) = -U(0): dLocalMatrix(7) = -U(1): dLocalMatrix(8) = -U(2)
    Else
        dLocalMatrix(3) = U(0): dLocalMatrix(4) = U(1): dLocalMatrix(5) = U(2)
        dLocalMatrix(6) = -V(0): dLocalMatrix(7) = -V(1): dLocalMatrix(8) = -V(2)
    End If
        
    dLocalMatrix(9) = oPoint.X.Value: dLocalMatrix(10) = oPoint.Y.Value: dLocalMatrix(11) = oPoint.Z.Value
    Call CalculateInstanceGlobalPositionMatrix(BuildParentList(oPanelCheck.PanelInstance), dGlobalMatrixPanel)
    
    'Caculate Vector Position > Return result in Global Matrix
    Call MultiplyVector(dLocalMatrix, dGlobalMatrixPanel, dGlobalMatrixPanel, True)
    'Caculate Rotation Matrix (dLocalMatrix X dGlobalMatrix) > Return result in Global Matrix
    Call MultiplyMatrix(dLocalMatrix, dGlobalMatrixPanel, dGlobalMatrixPanel)
    
    sPointwithoutholematrix = CStr(dGlobalMatrixPanel(0)) & ";" & _
                                CStr(dGlobalMatrixPanel(1)) & ";" & _
                                CStr(dGlobalMatrixPanel(2)) & ";" & _
                                CStr(dGlobalMatrixPanel(3)) & ";" & _
                                CStr(dGlobalMatrixPanel(4)) & ";" & _
                                CStr(dGlobalMatrixPanel(5)) & ";" & _
                                CStr(dGlobalMatrixPanel(6)) & ";" & _
                                CStr(dGlobalMatrixPanel(7)) & ";" & _
                                CStr(dGlobalMatrixPanel(8)) & ";" & _
                                CStr(dGlobalMatrixPanel(9)) & ";" & _
                                CStr(dGlobalMatrixPanel(10)) & ";" & _
                                CStr(dGlobalMatrixPanel(11))
End Function

