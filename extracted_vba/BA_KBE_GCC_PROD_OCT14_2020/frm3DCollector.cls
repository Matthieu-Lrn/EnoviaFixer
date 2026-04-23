Attribute VB_Name = "frm3DCollector"
Attribute VB_Base = "0{2FC44FD1-3FDD-411C-9942-5326FA144F76}{4F029AE5-990E-4700-9566-8F4D1489A509}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private o3DCollector As New Collector3D

'********************************************************************************
'* Name: Create/Sync. 3D Collector
'* Purpose: Create and synchronize 3D Collectors (see documentation)
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************

Public Sub initialize3DCollectorForm()
    
    Call enableRelationUpdate
    
    Set oSeedAssy = Nothing
    Set oSeedPartCollection = New Collection
    
    '***** Window size according to number of shifters
    With Me
        .Caption = "3D Collector Definition"
        .fra3DCollector.Top = .fraMode.Height + .fraMode.Top
        .fraNHA.Top = .fraMode.Height + .fraMode.Top
        .fraSeedAssy.Top = .fraNHA.Height + .fraNHA.Top
        .fraOptions.Top = .fraSeedAssy.Height + .fraSeedAssy.Top
        .fraDatum.Top = .fraOptions.Height + .fraOptions.Top
        .fraOperations.Top = .fraDatum.Height + .fraDatum.Top
        .Height = dTitleBarHeight + .fraOperations.Top + .fraOperations.Height + 33
        .Width = 182
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        
        Call resetUserform
        
        .optModifyExisting.Value = True
        Call SetFormMode
        
        Call SetOperationMode("AllSPs", True)
        Call cmdOKActivation
        .Show vbModeless
    End With
    
End Sub

'********************************************************************************
'* Name: OK Button
'* Purpose: When user clicks OK, create/sync. 3D collector and close form
'********************************************************************************
Private Sub cmdOK_Click()
    
    Me.Hide
    bCancelAction = False
    CATIA.DisplayFileAlerts = False
'    CATIA.Interactive = False 'Do not allow user to click something
'    frmKBEMain.Enabled = False 'Do not allow user to click something
    
    Call o3DCollector.launchColl3D
    
    frmProgress.Hide
    CATIA.DisplayFileAlerts = True
    CATIA.Interactive = True
    CATIA.RefreshDisplay = True
    
'    'Needed to "reactivate" Catia after CATIA.Interactive = True
'    On Error Resume Next
'    CATIA.ActiveWindow.ActivatePrevious
'    CATIA.ActiveWindow.ActivateNext
'    On Error GoTo 0
    
    If Me.Visible Then Exit Sub
    Call cmdCancel_Click

End Sub

'********************************************************************************
'* Name: Cancel Button
'* Purpose: When user clicks Cancel, close form
'********************************************************************************
Private Sub cmdCancel_Click()
    '*****Save settings (ATBS & Publis)
    Call SetTagValueToFile(sSettingsPath, sActiveToolbarModule & ".3DCOLLECTOR.PUBLISH", Str(Me.cbPublish.Value))
    Call SetTagValueToFile(sSettingsPath, sActiveToolbarModule & ".3DCOLLECTOR.ATB", Str(Me.cbATBParams.Value))
    
    Call resetUserform
    Call frmKBEMain.resetToolbar
End Sub


'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

'********************************************************************************
'* Name: Reset Userform
'* Purpose: Put default values, reset collections
'********************************************************************************
Private Sub resetUserform()
    
    Dim oSettingValue As Variant
    
    With Me
        .txtSelection3DCollector.Text = sNoSelection
        .txtSelectionNHA.Text = sNoSelection
        .txtSelectionSA.Text = sNoSelection
        
        On Error Resume Next
        '***** Uncheck all check boxes
        For Each oControl In .Controls
            If TypeName(oControl) = "CheckBox" Then oControl.Value = False
        Next
        
        .cbPositionCollector.Value = True
        
        .cbPublish.Value = False    'Default value if not found in saved settings
        .cbPublish.Value = CBool(GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".3DCOLLECTOR.PUBLISH"))
        
        .cbATBParams.Value = False  'Default value if not found in saved settings
        .cbATBParams.Value = CBool(GetTagValueFromFile(sSettingsPath, sActiveToolbarModule & ".3DCOLLECTOR.ATB"))
        
        On Error GoTo 0
        
    End With
    
    Set oSeedLingPart = Nothing
    Set oNHAAssy = Nothing
    Set oSeedAssy = Nothing
    Set oSeedPartCollection = New Collection
    
End Sub

'********************************************************************************
'* Name: Set default / focus
'* Purpose: Set the focus on Cancel Button if there's no seed part selected,
'*          and OK button is deactivated.
'*          Otherwise, set the focus on OK button
'********************************************************************************
Private Sub cmdOKActivation()
    
    On Error Resume Next
    If Me.imgModifyTrue.Visible = True And Not (Me.txtSelection3DCollector.Text = sNoSelection Or Me.txtSelectionSA.Text = sNoSelection) Then
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
        Me.cmdOK.SetFocus
    ElseIf Me.imgCreateTrue.Visible = True And Not (Me.txtSelectionNHA.Text = sNoSelection Or Me.txtSelectionSA.Text = sNoSelection) Then
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.Default = True
        Me.cmdOK.SetFocus
    Else
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
        Me.cmdCancel.Default = True
        Me.cmdCancel.SetFocus
    End If
    On Error GoTo 0
    
End Sub

Private Sub SetFormMode()
    
    With Me
        .imgCreateTrue.Visible = .optCreateNew.Value
        .imgCreateFalse.Visible = Not (.imgCreateTrue.Visible)
        .imgModifyTrue.Visible = .optModifyExisting.Value
        .imgModifyFalse.Visible = Not (.imgModifyTrue.Visible)
    
        If .optModifyExisting.Value Then
            .optModifyExisting.Value = True
            .fra3DCollector.Visible = True: .fraNHA.Visible = False
        ElseIf .optCreateNew.Value Then
            .optCreateNew.Value = True
            .fra3DCollector.Visible = False: fraNHA.Visible = True
            .fraNHA.Left = .fra3DCollector.Left
        End If
    End With
    
End Sub

Private Sub SetOperationMode(ByVal sMode As String, ByVal bValue As Boolean)
    
    If sMode = "AllSPs" Then
        
        ImgAllSPTrue.Visible = bValue
        ImgAllSPFalse.Visible = Not bValue
        If bValue = True Then
            ImgSelectedSPTrue.Visible = Not bValue
            ImgSelectedSPFalse.Visible = bValue
        End If
        
    ElseIf sMode = "SelectedSPs" Then
        
        ImgSelectedSPTrue.Visible = bValue
        ImgSelectedSPFalse.Visible = Not bValue
        If bValue = True Then
            ImgAllSPTrue.Visible = Not bValue
            ImgAllSPFalse.Visible = bValue
        End If
        
    End If

End Sub

Private Sub ChkAllSeedParts_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetOperationMode "AllSPs", ImgAllSPFalse.Visible
End Sub

Private Sub chkSelectedSP_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetOperationMode "SelectedSPs", ImgSelectedSPFalse.Visible
End Sub

Private Sub cmdSelectSA_Click()
    
    Me.Hide
    Me.txtSelectionSA.Text = sNoSelection
    Set oSeedPartCollection = New Collection
    On Error Resume Next
    If oSeedAssy Is Nothing Then
        Call selectRootAssy(oSeedAssy, oSeedPartRootDocument, "*_SEED_ASSY_*", "Please select a Seed Assembly!  ")
    Else
        If Not sSourceWindow Is Nothing And Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    End If
    
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If (oSeedAssy Is Nothing Or oSeedAssy.Products.Count = 0) Then
            MsgBox "Please select a valid Seed Assy!  ", vbCritical, "No Seed Parts found..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    On Error GoTo 0
    
    Call frmSeedPartsList.ListBoxInitialize(oSeedAssy, Me.Left + 15, Me.Top + Me.cmdSelectSA.Top)

    If sSeedPartsList <> "" Then
        Me.txtSelectionSA.Text = oSeedAssy.PartNumber
        Set sSourceWindow = CATIA.ActiveWindow
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
        Set oSeedPartCollection = frmSeedPartsList.PopulateSeedPartCollection
    Else
noSelection:
        Me.txtSelectionSA.Text = sNoSelection
        Set oSeedAssy = Nothing
        Set oSelectionFromSeedPart = Nothing
        Set sSourceWindow = Nothing
        Set oSeedPartCollection = New Collection
    End If
    
    If bToolbarCommandSelected = False Then Me.Show vbModeless
    
End Sub

Private Sub cmdSelect3DCollector_Click()
        
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtSelection3DCollector.Text = sNoSelection
    
    On Error Resume Next
    Call selectPart(oSeedLingPart, oSeedLingInstance, oSeedLingRootDocument, "*-*COL*", "Please select a 3D Collector!   ")
    
    ' Find NHA from part instance
    If TypeName(oSeedLingInstance.Parent.Parent) = "Product" Then
        Set oNHAAssy = oSeedLingInstance.Parent.Parent
    Else
        Set oNHAAssy = Nothing
    End If
    
    ' If user did not choose anything (i.e. pressed esc.)
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oSeedLingPart Is Nothing Then
            MsgBox "Please select a valid 3D Collector!  ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    
    If Not oSeedLingPart Is Nothing Then
        Me.txtSelection3DCollector.Text = oSeedLingPart.Name
        Set oSelectionFromSeedLing = oSeedLingRootDocument.Selection
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtSelection3DCollector.Text = sNoSelection
        Set oSeedLingPart = Nothing
        Set oNHAAssy = Nothing
        Set oSelectionFromSeedLing = Nothing
        Set sTargetWindow = Nothing
    End If
    On Error GoTo 0
     
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If

End Sub


Private Sub cmdSelectNHA_Click()
    
    Me.Hide
    Me.txtSelectionNHA.Text = sNoSelection
    On Error Resume Next
    Call selectAssy(oNHAAssy, oSeedLingRootDocument, , "Please select a Next Higher Assembly!   ")
    
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If oNHAAssy Is Nothing Then
            MsgBox "Please select a valid and loaded Next Higher Assembly!   ", vbCritical, "Wrong Selection..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    On Error GoTo 0

    If ProductNameCheck(oNHAAssy.PartNumber) = True Then
        Me.txtSelectionNHA.Text = oNHAAssy.Name
        Set oSelectionFromSeedLing = oSeedLingRootDocument.Selection
        Set sTargetWindow = CATIA.ActiveWindow
    Else
noSelection:
        Me.txtSelectionNHA.Text = sNoSelection
        Set oNHAAssy = Nothing
        Set oSeedLingRootDocument = Nothing
        Set oSelectionFromSeedLing = Nothing
        Set sTargetWindow = Nothing
    End If
    
    ' Show form, expcept if user ended command clicking another command on toolbar
    If bToolbarCommandSelected = False Then
        Me.Show vbModeless
        Call cmdOKActivation
    End If
    
End Sub

Private Sub ImgAllSPFalse_Click()
    SetOperationMode "AllSPs", ImgAllSPFalse.Visible
End Sub

Private Sub ImgAllSpTrue_Click()
    SetOperationMode "AllSPs", ImgAllSPFalse.Visible
End Sub

Private Sub ImgCreateFalse_Click()
    Me.optCreateNew.Value = True
End Sub

Private Sub ImgModifyFalse_Click()
    Me.optModifyExisting.Value = True
End Sub

Private Sub ImgSelectedSPFalse_Click()
    SetOperationMode "SelectedSPs", ImgSelectedSPFalse.Visible
End Sub

Private Sub ImgSelectedSPTrue_Click()
    SetOperationMode "SelectedSPs", ImgSelectedSPFalse.Visible
End Sub

Private Sub optCreateNew_Change()
    Call SetFormMode
    Call cmdOKActivation
End Sub

Private Sub optModifyExisting_Change()
    Call SetFormMode
    Call cmdOKActivation
End Sub

Private Sub txtSelection3DCollector_Change()
    Call cmdOKActivation
End Sub

Private Sub txtSelectionNHA_Change()
    Call cmdOKActivation
End Sub

Private Sub txtSelectionSA_Change()
    Call cmdOKActivation
End Sub

'********************************************************************************
'* Name: DS Globe Click
'* Purpose: Open a new product with right name to create a collector, tile vertically
'********************************************************************************
Private Sub imgDSGlobe_Click()
    
    Dim i As Integer
    Dim oProductDoc As ProductDocument
    
    On Error Resume Next
    If Me.imgCreateTrue.Visible Then
        Set oProductDoc = CATIA.Documents.Add("Product")
        
        i = 0
        Do
            Err.Clear
            i = i + 1
            oProductDoc.Product.ReferenceProduct.PartNumber = "7K00AAA001-" & Format(i, "00#") '"GXRS0000-0000-00-" & Format(i, "00#")
        Loop Until Err.Number = 0 Or i >= 500
        
        CATIA.Windows.Arrange catArrangeTiledVertical
    End If
    On Error GoTo 0
    
End Sub

