Attribute VB_Name = "frmInsertHole"
Attribute VB_Base = "0{4B268193-3BB2-418D-915D-B4E6B6293F22}{1039ED6A-68FA-4E68-9051-F8F63C8AF54C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oSPCollList As clsCollection
Private oSPAWorkbench As INFITF.Workbench
Private oPanelProd As Product
Private oFirstProd As Product
Private oSecondProd As Product
Private oPanelPart As Part
Private oDatumPlnGeo As HybridBody
Private oConstructnGeo As HybridBody
Private oPanelCheckObject As clsPanelCheck
Private oThkFaceSide1 As PlanarFace
Private oThkFaceSide2 As PlanarFace
Private oRootDocument As Document
Private oInsertHole As New InsertHole

Public Sub InitializeForm()
    With Me
        .Height = dTitleBarHeight + 296
        .Width = 190
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        .Lst_SPCollectors.AddItem sNoSelection
        '.Lbl_ListBox.Caption = "RightClick to highlight" & vbCrLf & "DoubleClick to remove"
        Call Lst_SPCollectors_Change
        .cmdPanelPart.SetFocus
        .Show vbModless
    End With
        Set oSPCollList = New clsCollection
End Sub

Private Sub cmdCancel_Click()
    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    'Log File
    Call AddToLogFile("Insert Hole", oPanelPart.Parent.Name)
    Call oInsertHole.LaunchCreation(oPanelCheckObject, _
                                    oSPCollList, _
                                    oRootDocument, _
                                    oDatumPlnGeo, _
                                    oConstructnGeo, _
                                    True, _
                                    chkBx_Panelflushness.Value, _
                                    chkBx_DeleteUnusedHoles.Value)

    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub

Private Sub resetUserform()
    With Me
        
        .txtPanelPart.Text = sNoSelection
        .txtThkSurfaceSide1.Text = sNoSelection
        .txtThkSurfaceSide2.Text = sNoSelection
    End With
End Sub

Private Sub cmdOKActivation()

    
    On Error Resume Next
    If Not (Me.txtPanelPart.Text = sNoSelection Or _
            Me.txtThkSurfaceSide1.Text = sNoSelection Or _
            Me.txtThkSurfaceSide2.Text = sNoSelection) Then
            
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
    '    Me.cmdOK.SetFocus
    Else
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
        Me.cmdCancel.Default = True
    '    Me.cmdCancel.SetFocus
    End If
    On Error GoTo 0
End Sub

Private Sub cmdThkSurfaceActivation()
    On Error Resume Next
        If Not (Me.txtPanelPart.Text = sNoSelection) Then
            Me.cmdThkSurfaceSide1.ForeColor = -2147483630
            Me.cmdThkSurfaceSide1.Locked = False
            Me.cmdThkSurfaceSide2.ForeColor = -2147483630
            Me.cmdThkSurfaceSide2.Locked = False
        
        Else
            Me.cmdThkSurfaceSide1.ForeColor = -2147483631
            Me.cmdThkSurfaceSide1.Locked = True
            Me.cmdThkSurfaceSide2.ForeColor = -2147483631
            Me.cmdThkSurfaceSide2.Locked = True
        End If
    On Error GoTo 0
End Sub

Private Sub cmdPanelPart_Click()
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtPanelPart.Text = sNoSelection
    
    Call selectPart(oPanelPart, oPanelProd, oRootDocument, , "Please select the panel part!   ")
        
    '****If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oPanelPart Is Nothing Then
            MsgBox "Please select a panel part!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    '****Get Existing Panel Thicknesses
    Set oPanelCheckObject = New clsPanelCheck
    Set oPanelCheckObject.PanelInstance = oPanelProd
    Call GetPanelThickness(oPanelProd.ReferenceProduct, oPanelPart, oPanelCheckObject)
    oPanelCheckObject.PanelThicknessStatus = CheckPanelThickness(oPanelCheckObject.PanelThickness)
    If oPanelCheckObject.PanelThicknessStatus = EnumPanelThicknessStatus.OK Then
        '*****Find or create Construction_Geometry and Datum_Planes  Geosets
        '**Plane 1
        Dim oPln As Plane
        Set oPln = FindPlaneInDatumGeoset(oPanelCheckObject.ArPlane1)
        If oPln Is Nothing Then
            Set oPanelCheckObject.PanelPlane1 = CreatePlaneInGeosetfromMeasuredata(oPanelCheckObject.ArPlane1, oDatumPlnGeo)
        Else
            Set oPanelCheckObject.PanelPlane1 = oPln
        End If
        '**Plane 2
        Set oPln = FindPlaneInDatumGeoset(oPanelCheckObject.ArPlane2)
        If oPln Is Nothing Then
            Set oPanelCheckObject.PanelPlane2 = CreatePlaneInGeosetfromMeasuredata(oPanelCheckObject.ArPlane2, oDatumPlnGeo)
        Else
            Set oPanelCheckObject.PanelPlane2 = oPln
        End If
        Me.txtThkSurfaceSide1.Text = oPanelCheckObject.PanelPlane1.Name
        Me.txtThkSurfaceSide2.Text = oPanelCheckObject.PanelPlane2.Name
        Set oFirstProd = oPanelCheckObject.PanelInstance
        Set oSecondProd = oPanelCheckObject.PanelInstance
    End If
    Call oPanelCheckObject.DeleteThicknessGeo
    
    '***Get Default SPcollector and populate the form
    Dim oPanelassyInstance As Product
    Dim oSPColInstance As Product
    
    Set oPanelCheckObject.PanelAssyInstance = oPanelCheckObject.PanelInstance.Parent.Parent
    oPanelCheckObject.PanelAssyStatus = CheckPanelAssyStatus(oPanelCheckObject.PanelAssyInstance, oPanelCheckObject.PanelInstance, oSPColInstance)
    If oPanelCheckObject.PanelAssyStatus = EnumPanelAssyStatus.OK And Not oSPColInstance Is Nothing Then
        sParentList = BuildParentList(oSPColInstance)
        Call Lst_SPCollectorsAddItem(sParentList, oSPColInstance)
    End If
    
    'Update form with selected object
    If Not oPanelPart Is Nothing Then
        Me.txtPanelPart.Text = oPanelPart.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtPanelPart.Text = sNoSelection
        Set oPanelPart = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
        Call cmdThkSurfaceActivation
    End If
End Sub

Private Sub cmdStdCollector_Click()
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide

    Dim j As Integer
    Dim iObject As Object
    Dim sParentList As String
    Dim oSPCollProduct As Product

    Call selectAssy(oSPCollProduct, oRootDocument, , "Please select Std Part Collector with inserts!   ")
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSPCollProduct Is Nothing Then    ' If user did not choose anything (i.e. pressed esc.)
            MsgBox "Please select an SP Collector product!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        Else
            Set iObject = oSPCollProduct
            Do  'Find SPCollector instance in parents
                'We don't care about the "Products" or "ProductDocuments" objects
                If TypeName(iObject) = "Product" And iObject.Name Like "*#SP##*" Then Exit Do
                Set iObject = iObject.Parent
            Loop Until TypeName(iObject) = "Application"
            If TypeName(iObject) = "Application" Then
                MsgBox "Please select an SP Collector product!  ", vbCritical, "Wrong Selection..."
                GoTo noSelection
            Else
                Set oSPCollProduct = iObject
            End If
        End If
    Else
        GoTo noSelection
    End If
    'Update form with selected object
    sParentList = BuildParentList(oSPCollProduct)
    '***Add item to SP collector list
    Call Lst_SPCollectorsAddItem(sParentList, oSPCollProduct)
    Set sTargetWindow = CATIA.ActiveWindow
noSelection:
    Set oSPCollProduct = Nothing
    Set sTargetWindow = Nothing
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If
End Sub

Private Sub cmdThkSurfaceSide1_Click()
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtThkSurfaceSide1.Text = sNoSelection
    
    On Error Resume Next
    Call selectPlanarFace(oThkFaceSide1, oFirstProd, oRootDocument, "Please select the planar face of the panel!   ")
    On Error GoTo 0
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oThkFaceSide1 Is Nothing Then
            MsgBox "Please select a planar face in an assembly!  ", vbCritical, "Wrong Selection..."
            Set oFirstProd = Nothing
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    '*** both faces belong to same panel
    If sCheckBothFacesBelongToSamePart(oFirstProd, oSecondProd) = "KO" Then
        Set oThkFaceSide1 = Nothing
        Set oFirstProd = Nothing
        MsgBox "Please make all selections from same CATPart", vbExclamation, "Wrong Selection"
        GoTo noSelection
    End If
    '***Create plane from selected Face
    Dim AR(8)
    Dim oMeasurable
    Dim oPlnx As HybridShapePlaneEquation
    If oSPAWorkbench Is Nothing Then Set oSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    
    Set oMeasurable = oSPAWorkbench.GetMeasurable(oThkFaceSide1)
    oMeasurable.GetPlane AR
    Dim oPln As Plane
    Set oPln = FindPlaneInDatumGeoset(AR)
    If oPln Is Nothing Then
        Set oPanelCheckObject.PanelPlane1 = CreatePlaneInGeosetfromMeasuredata(AR, oDatumPlnGeo)
    Else
        Set oPanelCheckObject.PanelPlane1 = oPln
    End If
    oPanelCheckObject.SetArPlane1 AR
    'Update form with selected object
    If Not oThkFaceSide1 Is Nothing Then
        Me.txtThkSurfaceSide1.Text = oPanelCheckObject.PanelPlane1.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtThkSurfaceSide1.Text = sNoSelection
        Set oThkFaceSide1 = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, except if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub cmdThkSurfaceSide2_Click()
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtThkSurfaceSide2.Text = sNoSelection
    
    On Error Resume Next
    Call selectPlanarFace(oThkFaceSide2, oSecondProd, oRootDocument, "Please select the planar face of the panel!   ")
    On Error GoTo 0
        
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oThkFaceSide2 Is Nothing Then
            MsgBox "Please select a planar face in an assembly!  ", vbCritical, "Wrong Selection..."
            Set oSecondProd = Nothing
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    '*** both faces belong to same panel
    If sCheckBothFacesBelongToSamePart(oSecondProd, oFirstProd) = "KO" Then
        Set oThkFaceSide2 = Nothing
        Set oSecondProd = Nothing
        MsgBox "Please make all selections from same CATPart", vbExclamation, "Wrong Selection"
        GoTo noSelection
    End If
    '***Create plane from selected Face
    Dim AR(8)
    Dim oMeasurable
    Dim oPlnx As HybridShapePlaneEquation
    If Not oSPAWorkbench Is Nothing Then Set oSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    
    Set oMeasurable = oSPAWorkbench.GetMeasurable(oThkFaceSide2)
    oMeasurable.GetPlane AR
    Dim oPln As Plane
    Set oPln = FindPlaneInDatumGeoset(AR)
    If oPln Is Nothing Then
        Set oPanelCheckObject.PanelPlane2 = CreatePlaneInGeosetfromMeasuredata(AR, oDatumPlnGeo)
    Else
        Set oPanelCheckObject.PanelPlane2 = oPln
    End If
    oPanelCheckObject.SetArPlane2 AR
    'Update form with selected object
    If Not oThkFaceSide2 Is Nothing Then
        Me.txtThkSurfaceSide2.Text = oPanelCheckObject.PanelPlane2.Name
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtThkSurfaceSide2.Text = sNoSelection
        Set oThkFaceSide2 = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub

Private Sub Lst_SPCollectors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Lst_SPCollectors.ListCount > 0 Then
        If Lst_SPCollectors.List(0, 0) <> sNoSelection Then
            If oSPCollList.Exists(Lst_SPCollectors.Value) Then
                oSPCollList.RemoveByKey Lst_SPCollectors.Value
            End If
            Lst_SPCollectors.RemoveItem Lst_SPCollectors.ListIndex
        Else
            Call Lst_SPCollectors_Change
        End If
    End If
    If Lst_SPCollectors.ListCount = 0 Then
        Lst_SPCollectors.AddItem sNoSelection
        Call Lst_SPCollectors_Change
    End If
End Sub

Private Sub Lst_SPCollectors_Change()
    If Lst_SPCollectors.List(0, 0) = sNoSelection Then
        Lst_SPCollectors.ForeColor = &H80000011
        Lst_SPCollectors.ListIndex = -1
    Else
        Lst_SPCollectors.ForeColor = &H80000008
    End If

End Sub
Private Sub Lst_SPCollectorsAddItem(ByVal sItem As String, ByRef oInstance As Product)
    If Not oSPCollList.Exists(sItem) Then
        Lst_SPCollectors.AddItem sItem
        oSPCollList.Add sItem, oInstance
        Lst_SPCollectors.ListIndex = Lst_SPCollectors.ListCount - 1
        If Lst_SPCollectors.List(0, 0) = sNoSelection Then
            Lst_SPCollectors.RemoveItem 0
        End If
    Else
        For j = 0 To Lst_SPCollectors.ListCount - 1
            If Lst_SPCollectors.List(j) = sItem Then
                Lst_SPCollectors.ListIndex = j
            End If
        Next
    End If
End Sub

Private Sub Lst_SPCollectors_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        CATIA.ActiveDocument.Selection.Clear
        On Error Resume Next
           CATIA.ActiveDocument.Selection.Add oSPCollList.GetItem(Lst_SPCollectors.Value)
           CATIA.StartCommand ("Reframe on")
        On Error GoTo 0
    End If
End Sub

'Private Sub txtDistanceCtrToCtr_Change()
'    Call cmdOKActivation
'End Sub
'
'Private Sub txtDistanceCtrToCtr_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    '0 to 9
'    If KeyAscii >= 48 And KeyAscii <= 57 Then
'        KeyAscii = KeyAscii
'
'    'Only one "."
'    ElseIf KeyAscii = 46 And Not Me.txtDistanceCtrToCtr.Text Like "*.*" Then
'        KeyAscii = 46
'    Else
'        KeyAscii = 0
'    End If
'End Sub
Private Sub txtPanelPart_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
       CATIA.ActiveDocument.Selection.Clear
       On Error Resume Next
          CATIA.ActiveDocument.Selection.Add oPanelCheckObject.PanelInstance
          CATIA.StartCommand ("Reframe on")
       On Error GoTo 0
    End If
End Sub
Private Sub txtThkSurfaceSide1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
       CATIA.ActiveDocument.Selection.Clear
       On Error Resume Next
          CATIA.ActiveDocument.Selection.Add oPanelCheckObject.PanelPlane1
          CATIA.StartCommand ("Center graph")
       On Error GoTo 0
    End If
End Sub

Private Sub txtThkSurfaceSide2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
       CATIA.ActiveDocument.Selection.Clear
       On Error Resume Next
          CATIA.ActiveDocument.Selection.Add oPanelCheckObject.PanelPlane2
          CATIA.StartCommand ("Center graph")
       On Error GoTo 0
    End If
End Sub

'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub
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
Private Function CreatePlaneInGeosetfromMeasuredata(dPlane As Variant, ByRef oGeoset As HybridBody, Optional ByVal sName As String = "") As Plane
    Dim oHSF As HybridShapeFactory
    Dim oPlane As HybridShapePlaneEquation
    
    Dim dNormal As Variant
    Dim dLengthNormal As Double
    Dim dOrigin(2) As Double
    Dim dV1(2) As Double, dV2(2) As Double
    Dim a As Double, b As Double, c As Double, d As Double
    
    dOrigin(0) = dPlane(0): dOrigin(1) = dPlane(1): dOrigin(2) = dPlane(2)
    dV1(0) = dPlane(3): dV1(1) = dPlane(4): dV1(2) = dPlane(5)
    dV2(0) = dPlane(6): dV2(1) = dPlane(7): dV2(2) = dPlane(8)
    
    Call CrossProduct(dV1, dV2, dNormal)
    
    dLengthNormal = LengthVector(dNormal)
    
    dNormal(0) = dNormal(0) / dLengthNormal: dNormal(1) = dNormal(1) / dLengthNormal: dNormal(2) = dNormal(2) / dLengthNormal
    
    a = dNormal(0): b = dNormal(1): c = dNormal(2)
    d = (DotProd(dNormal, dOrigin))
    
    Set oHSF = oPanelPart.HybridShapeFactory
    Set oPlane = oHSF.AddNewPlaneEquation(a, b, c, d)
    oGeoset.AppendHybridShape oPlane
    If sName <> "" Then oPlane.Name = sName
    oPlane.Compute
    Set CreatePlaneInGeosetfromMeasuredata = oPlane
End Function
Private Function sCheckBothFacesBelongToSamePart(oFAssy As Product, oSAssy As Product) As String
    '*** This check verifies that both faces belong to same part
    '*** if only one face is chosen yet then it outputs OK
    '*** if both faces are chosen befoe triggering the check then it verifies that both facses belong to same part
    Dim sCheck As String
    Dim bNotNothing As Boolean
    
    sCheck = "KO"
    
    If oFAssy Is Nothing Or oSAssy Is Nothing Then bNotNothing = False Else bNotNothing = True
    
    If bNotNothing Then
        If oFAssy.Name = oSAssy.Name And oFAssy.ReferenceProduct.Name = oSAssy.ReferenceProduct.Name Then
            sCheck = "OK"
        End If
    Else
        sCheck = "OK"
    End If
    sCheckBothFacesBelongToSamePart = sCheck
End Function
Private Function FindPlaneInDatumGeoset(dArPlane As Variant) As Plane
        '*****Find plane in Construction_Geometry and Datum_Planes  Geosets and
        
        Dim oGeo As HybridBody
        Dim oHSF As HybridShapeFactory
        For Each oGeo In oPanelPart.HybridBodies
            If UCase(oGeo.Name) Like "CONSTRUCTION*GEOMETRY" Then
                Set oConstructnGeo = oGeo
                Exit For
            End If
        Next
        If oConstructnGeo Is Nothing Then
            Set oConstructnGeo = oPanelPart.HybridBodies.Add
            oConstructnGeo.Name = "Construction_Geometry"
        End If
        For Each oGeo In oConstructnGeo.HybridBodies
            If UCase(oGeo.Name) Like "DATUM*PLANES" Then
                Set oDatumPlnGeo = oGeo
                Exit For
            End If
        Next
        If oDatumPlnGeo Is Nothing Then
            Set oDatumPlnGeo = oConstructnGeo.HybridBodies.Add
            oDatumPlnGeo.Name = "Datum_Planes"
        End If
        
        Dim oPln As Plane
        Dim oInIPln As Plane
        Dim oTemGeo As HybridBody
        Set oTemGeo = oPanelPart.HybridBodies.Add
        Set oInIPln = CreatePlaneInGeosetfromMeasuredata(dArPlane, oTemGeo, "PlaneToCheck")
        
        For Each oPln In oDatumPlnGeo.HybridShapes
            If bPlanesAreSame(oPln, oInIPln, oPanelPart) Then
                Set FindPlaneInDatumGeoset = oPln
                Exit For
            End If
        Next
        Set oHSF = oPanelPart.HybridShapeFactory
        oHSF.DeleteObjectForDatum oPanelPart.CreateReferenceFromObject(oTemGeo)
End Function
