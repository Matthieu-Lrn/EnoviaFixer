Attribute VB_Name = "frmProject"
Attribute VB_Base = "0{2400BC35-C3B6-48FE-8738-CDC28D574381}{2A2D2417-14CA-4EE2-9D9E-38C20A19BC97}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oCabinetProject As New CabinetProject

Public Sub initializeProjectForm()
    
    '***** Window size according to number of shifters
    With Me
        .StartUpPosition = Manual
        .fraNHA.Top = 3.75
        .fraSeedAssy.Top = .fraNHA.Height + .fraNHA.Top
        .Height = dTitleBarHeight + .fraSeedAssy.Top + .fraSeedAssy.Height + 33
        .Width = 182
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        .Caption = "Cabinet Project"
        .cmdCancel.SetFocus
        
        '***** Default values
        Call resetUserform
        Call cmdOKActivation
        .Show vbModeless
    End With

End Sub

'********************************************************************************
'* Name: Cancel Button
'* Purpose: When user clicks Cancel, close form
'********************************************************************************
Private Sub cmdCancel_Click()
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
'* Purpose: Empty all listboxes, textboxes and collections
'*          Deactivate buttons in frames
'********************************************************************************
Private Sub resetUserform()
    Me.txtSelectionNHA.Text = sNoSelection
    Me.txtSelectionSA.Text = sNoSelection
End Sub


'********************************************************************************
'* Name: Set default / focus
'* Purpose: Set the focus on Cancel Button if there's no seed part selected
'*          Otherwise, set the focus on OK button
'********************************************************************************
Private Sub cmdOKActivation()
    
    On Error Resume Next
    If Not (Me.txtSelectionNHA.Text = sNoSelection Or Me.txtSelectionSA.Text = sNoSelection) Then
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

    If ProductNameCheck(oNHAAssy.PartNumber, False, True) = True Then
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

Private Sub txtSelectionNHA_Change()
    Call cmdOKActivation
End Sub

Private Sub cmdSelectSA_Click()

    Me.Hide
    Me.txtSelectionSA.Text = sNoSelection¸
    Set oSeedPartCollection = New Collection
    
    On Error Resume Next
    Call selectRootAssy(oSeedAssy, oSeedPartRootDocument, "*_SEED_ASSY_*", "Please select a Seed Assembly!  ")
    
    If bStopMultiSelection = False And bToolbarCommandSelected = False Then
        If (oSeedAssy Is Nothing Or oSeedAssy.Products.Count = 0) Then
            MsgBox "Please select a valid Seed Assy!  ", vbCritical, "No Seed Parts found..."
            GoTo noSelection
        End If
    Else
        GoTo noSelection
    End If
    On Error GoTo 0
    
    If Not oSeedAssy Is Nothing Then
        Me.txtSelectionSA.Text = oSeedAssy.PartNumber
        Set sSourceWindow = CATIA.ActiveWindow
        Set oSelectionFromSeedPart = oSeedPartRootDocument.Selection
    Else
noSelection:
        Me.txtSelectionSA.Text = sNoSelection
        Set oSeedAssy = Nothing
        Set sSourceWindow = Nothing
        Set oSelectionFromSeedPart = Nothing
    End If
    
    If bToolbarCommandSelected = False Then Me.Show vbModeless
    
End Sub

Private Sub txtSelectionSA_Change()
    Call cmdOKActivation
End Sub

Private Sub cmdOK_Click()
    
    Me.Hide
    bCancelAction = False
    CATIA.DisplayFileAlerts = False
'    CATIA.Interactive = False 'Do not allow user to click something
'    frmKBEMain.Enabled = False 'Do not allow user to click something
    
    Call oCabinetProject.launchCabinetProject
    
    frmProgress.Hide
    CATIA.DisplayFileAlerts = True
    CATIA.Interactive = True
    CATIA.RefreshDisplay = True
    
'    'Needed to "reactivate" Catia after CATIA.Interactive = True
'    On Error Resume Next
'    CATIA.ActiveWindow.ActivatePrevious
'    CATIA.ActiveWindow.ActivateNext
'    On Error GoTo 0
    
    Call cmdCancel_Click
    
End Sub
