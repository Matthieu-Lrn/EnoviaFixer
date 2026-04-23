Attribute VB_Name = "frmDisplayInCatia"
Attribute VB_Base = "0{7AD5D151-A030-4623-B246-E64D4FEA4FA5}{732BE36E-DBD3-4873-B276-72BD9EB382AC}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oAutomationGeo As HybridBody
Private oHSF As HybridShapeFactory
Private cSections As Sections
Private oMasterSection 'As Section
Private sWindow As Window
Private oTopAssy As Product
Private oPointProd As Product
Private oPointPart As Part
Private oPoint As HybridShapePointCoord
Private oSphere As HybridShapeSphere
Private oExcel As clsExcel
Public Sub LaunchDisplayInCatia(Optional ByVal sSwind As Window, Optional oExcl As clsExcel)
    Dim oAxis As AxisSystem
    Dim oSel As Selection
    
    Set sWindow = sSwind
    If oExcl Is Nothing Then '***Search for Runnning Excel File
        Set oExcel = New clsExcel
        oExcel.GetExistingExcel
        If oExcel.App Is Nothing Then
            MsgBox "Please open the relevant panel assy analysis report" & vbCrLf & _
                    "before launching Display in Catia."
            Exit Sub
        End If
    Else
        Set oExcel = oExcl
    End If

    Call SwapToAssemblyWorkbench
    Set oTopAssy = CATIA.ActiveDocument.Product
    '************************************
    CATIA.DisplayFileAlerts = False
    
    Set oPointProd = oTopAssy.Products.AddNewComponent("Part", "")
    Set oPointPart = oPointProd.ReferenceProduct.Parent.Part
    Set oAutomationGeo = oPointPart.HybridBodies.Add
    Set oSel = oPointProd.ReferenceProduct.Parent.Selection
    oSel.Clear
    oSel.Add oAutomationGeo
    oSel.VisProperties.SetShow catVisPropertyNoShowAttr
    
    Set oHSF = oPointPart.HybridShapeFactory
    Set oPoint = oHSF.AddNewPointCoord(0, 0, 0)
    oAutomationGeo.AppendHybridShape oPoint
    oPoint.Compute
    Set oAxis = oPointPart.AxisSystems.Add
    oPointPart.UpdateObject oAxis
    Set oSphere = oHSF.AddNewSphere(oPointPart.CreateReferenceFromObject(oPoint), _
                                    oPointPart.CreateReferenceFromObject(oAxis), _
                                    12.7, -45#, 45#, 0#, 180#)
    oSphere.Limitation = 1
    oAutomationGeo.AppendHybridShape oSphere
    oSphere.Compute
    
    CATIA.DisplayFileAlerts = True
    '************************************
    With Me
        .Height = dTitleBarHeight + 58
        .Width = 120
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdSelect.Top = .Height - dTitleBarHeight - .cmdCancel.Height - 32
        .cmdSelect.Left = .Width - .cmdSelect.Width - 12
        .cmdSelect.ControlTipText = "Select an insert from excel report and" + vbNewLine + " press Select To Highlight it in Catia"
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        
        .cmdSelect.SetFocus
        .Show vbModless
    End With
    
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    cSections.Remove "Automation_001Section" '*** remove any previously created c Section
    CATIA.DisplayFileAlerts = False
    oTopAssy.Products.Remove oPointProd.Name
    CATIA.DisplayFileAlerts = True
    On Error GoTo 0
    Me.Hide

    Call frmKBEMain.resetToolbar
End Sub
'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub
Private Sub cmdSelect_Click()
    Dim sAssyName As String
    Dim sPathlist As Variant
    Dim oProd As Product
    Dim oProduct As Product
    Dim bProductFound As Boolean
    Dim sVar As Variant
    sAssyName = ""
    '**Getting value of coloum1 and selected row
       
    On Error Resume Next
        sAssyName = oExcel.App.Cells(oExcel.App.ActiveCell.Row, 1).Value
    On Error GoTo 0
    
    Set oProd = CATIA.ActiveDocument.Product
    If sAssyName = "" Or InStr(1, sAssyName, "||") = 0 Then
        MsgBox "Please select the row with insert name in Excel" & vbCrLf & _
               "Also verify if active Sheet is an analisys report", vbExclamation
    ElseIf sAssyName Like "Hole_for*" Then
        Call SelectAssemblyInCatia(sAssyName)
    Else
        sPathlist = Split(sAssyName, "||")
        sPathlist = Split(sPathlist(0), ";")
        If UBound(sPathlist) > 2 Then
            For i = UBound(sPathlist) - 2 To 0 Step -1
                If Not sPathlist(i) = "" Then
                    For Each oProduct In oProd.Products
                        If oProduct.Name = sPathlist(i) Then
                            Set oProd = oProduct
                            bProductFound = True
                            Exit For
                        End If
                    Next
                    If Not bProductFound Then
                        MsgBox "Cannot find the Selection in Catia!"
                        Exit Sub
                    End If
                End If
            Next
            If Not bProductFound Then
                MsgBox "Cannot find the Selection in Catia!"
                Exit Sub
            Else
                Call SelectAssemblyInCatia(sAssyName, oProd)
                bProductFound = False
            End If
        Else
            MsgBox ("Please select the row with insert name in Excel")
        End If
    End If

End Sub
Private Sub SelectAssemblyInCatia(ByVal sAssyName As String, Optional oProdtoSelect As Product)
    Dim dGlobalMatrix(11) As Variant
    Dim xGlobalMatrix(11) As Variant
    Dim iObject As AnyObject
    Dim yVector(3 To 5) As Double
    Dim zVector(6 To 8) As Double
    
    If Not oProdtoSelect Is Nothing Then Set iObject = oProdtoSelect
    '*** make sure that active window is the one in which tested assembly was selected
    If Not sWindow Is Nothing And Not CATIA.ActiveWindow Is sWindow Then sWindow.Activate
    '*** creating global matrix for section view
    Call oParseGlobalPositionMatrix(sAssyName, dGlobalMatrix)
    If UBound(dGlobalMatrix) = 11 Then
        Set cSections = CATIA.ActiveDocument.Product.GetTechnologicalObject("Sections")
        On Error Resume Next
        cSections.Remove "Automation_001Section" '*** remove any previously created c Section
        On Error GoTo 0
        Set oMasterSection = cSections.Add
        oMasterSection.Type = catSectionTypePlane
        oMasterSection.CutMode = 1
        oMasterSection.Name = "Automation_001Section"
        oMasterSection.Width = 1
        oMasterSection.Height = 1
        oMasterSection.Behavior = catSectionBehaviorAutomatic
      
        If sAssyName Like "Hole_for*" Then
            oPoint.X.Value = dGlobalMatrix(9): oPoint.Y.Value = dGlobalMatrix(10): oPoint.Z.Value = dGlobalMatrix(11)
            oPoint.Compute
            oSphere.Compute
            Set iObject = oSphere
            oMasterSection.Width = 5000
            oMasterSection.Height = 5000
        End If
        With CATIA.ActiveWindow.ActiveViewer
            .Viewpoint3D.PutOrigin Array(dGlobalMatrix(9), dGlobalMatrix(10), dGlobalMatrix(11))
            .Viewpoint3D.PutSightDirection Array(dGlobalMatrix(6), dGlobalMatrix(7), dGlobalMatrix(8))
            .Viewpoint3D.PutUpDirection Array(dGlobalMatrix(3), dGlobalMatrix(4), dGlobalMatrix(5))
            .Update
        End With
        On Error Resume Next
        oMasterSection.SetPosition dGlobalMatrix
        If Err.Number <> 0 Then
             MsgBox ("Assembly cannot be found" & vbExclamation)
             cSections.Remove "Automation_001Section"
             On Error GoTo 0
             Exit Sub
        End If
        On Error GoTo 0
    
        CATIA.ActiveDocument.Selection.Clear
        CATIA.ActiveDocument.Selection.Add iObject
        CATIA.StartCommand ("Reframe on")
    Else
        MsgBox "Non Compatible Excel Report" & vbCrLf & "Please Verify that Active Document is a Panel Assy Check report", vbExclamation
    End If
    'CATIA.ActiveDocument.Selection.Clear
    
End Sub
Public Sub oParseGlobalPositionMatrix(ByVal sAssyName As String, ByRef GlobalMatrix As Variant)
    Dim sGlobalMatrix(11) As Variant
    Dim sdummy As String
    Dim sVar As Variant
    If InStr(sAssyName, "||") > 0 Then sdummy = Split(sAssyName, "||")(1)
    If Right(sdummy, 1) = ";" Then sdummy = Left(sdummy, Len(sdummy) - 1)
    
    sVar = Split(sdummy, ";")
    ReDim dGlobalMatrix(UBound(sVar))
    For i = 0 To UBound(sVar)
        GlobalMatrix(i) = CDbl(sVar(i))
    Next
End Sub
Private Sub SwapToAssemblyWorkbench()
    Dim sWkb As String
    
    sWkb = CATIA.GetWorkbenchId
    If sWkb <> "Assembly" Then
        CATIA.StartWorkbench "Assembly"
    End If
End Sub
