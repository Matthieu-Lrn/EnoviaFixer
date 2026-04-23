Attribute VB_Name = "frmFlatTree"
Attribute VB_Base = "0{4FE3865E-63A9-402D-8826-FEF48C219AA2}{A24141A7-384C-42B6-B6D4-2995D957D87C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private i As Long
Private iNoOfBodies As Long
Private sCageCode As String
Private sSupplierCode As String
Private oFlatTree As New FlatTree

'********************************************************************************
'* Name: Create/Sync. Flat Tree
'* Purpose: Create and synchronize Flat Tree (see documentation)
'*
'* Assumption:
'*
'* Author:
'* Updated by: Julien Bigaouette
'* Language: VBA
'********************************************************************************

Public Sub initializeFlatTreeForm()
    
    Call enableRelationUpdate
    
    '***** Window size according to number of shifters
    With Me
        .Caption = "Flat Tree Definition"
        
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        .optCreate.Value = True
        
        Call fillComboBoxes
        Call resetUserform
        Call SetFormMode
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
    
    Call oFlatTree.launchFlatTree
    
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
    Me.Hide
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
    
    iBdyCntIndex = 1
    iKBEBodyCntIndex = 0
    
    With Me
        .txtSelection3DCollector.Text = sNoSelection
        .lblRangeKBEBodies.Caption = "Range"
        .txtRangeKBEBodies.Value = ""
        .txtRangeKBEBodies.Enabled = False
        .txtRangeKBEBodies.BackColor = &H80000016
        
        .cmbSupplierCode.ListWidth = 0
        .cmbSupplierCode.Value = .cmbSupplierCode.List(39) '35 corresponds to "BI", 39 corresponds to "BM"
        sSupplierCode = .cmbSupplierCode.Value
        
        .cmbCAGECode.ListWidth = 250
        .cmbCAGECode.Value = .cmbCAGECode.List(49) '49 corresponds to "BOMBARDIER COMPLETION CENTER INC. 3AJ74",
        sCageCode = .cmbCAGECode.Value
        
    End With
    
    Set oSeedLingPart = Nothing
    Set oNHAAssy = Nothing
    
End Sub

'********************************************************************************
'* Name: Set default / focus
'* Purpose: Set the focus on Cancel Button if there's no seed part selected,
'*          and OK button is deactivated.
'*          Otherwise, set the focus on OK button
'********************************************************************************

Private Sub cmdOKActivation()
    
    On Error Resume Next
    If Not Me.txtSelection3DCollector.Text = sNoSelection Then
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
        
        .Width = 182
        .imgCreateTrue.Visible = .optCreate.Value
        .imgCreateFalse.Visible = Not (.imgCreateTrue.Visible)
        .imgCleanTrue.Visible = .optClean.Value
        .imgCleanFalse.Visible = Not (.imgCleanTrue.Visible)
        .imgReleaseTrue.Visible = .optRelease.Value
        .imgReleaseFalse.Visible = Not (.imgReleaseTrue.Visible)
        
        .fraMode.Top = 3.75
        
        If .optCreate.Value Then
            .fra3DCollector.Visible = True: .fraOptions.Visible = True: .fraAttributes.Visible = True
            
            .fra3DCollector.Top = fraMode.Top + fraMode.Height
            .fraOptions.Top = fra3DCollector.Top + fra3DCollector.Height
            .fraAttributes.Top = .fraOptions.Top + .fraOptions.Height
            .Height = dTitleBarHeight + .fraAttributes.Top + .fraAttributes.Height + 33
            
        Else 'If .optClean.Value or .optRelease.Value Then
            .fra3DCollector.Visible = True: .fraOptions.Visible = False: .fraAttributes.Visible = False
            
            .fra3DCollector.Top = fraMode.Top + fraMode.Height
            .Height = dTitleBarHeight + .fra3DCollector.Top + .fra3DCollector.Height + 33
        End If
        
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        
        If .Top > dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height Then .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 10
        If .Top < 0 Then .Top = 10
        
    End With
    
End Sub

Private Sub cmdSelect3DCollector_Click()
    
    Me.cmdCancel.Default = True
    Me.cmdCancel.SetFocus
    Me.Hide
    Me.txtSelection3DCollector.Text = sNoSelection
    
    On Error Resume Next
    Call selectPart(oSeedLingPart, oSeedLingInstance, oSeedLingRootDocument, "*-*COL*", "Please select a 3D Collector!   ")
    If oSeedLingInstance Is Nothing Then GoTo noSelection
    
    ' Find NHA from part instance
    Set oNHAAssy = Nothing
    Set oNHAAssy = oSeedLingInstance.Parent.Parent
    
    If oNHAAssy Is Nothing Then
        MsgBox "NHA Cannot be found!" & vbCrLf & "Please select a 3D Collector in product context!   ", vbCritical, "Wrong Selection..."
        GoTo noSelection
    ElseIf Me.optCreate.Value = True Then 'Validate NHA
        If ProductNameCheck(oNHAAssy.PartNumber, True, False) = False Then GoTo noSelection
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
        
        'Find at least one BodyKBE (or BodyDBE in the case of Collector Template)
        Set oBody = Nothing
        For Each iObject In oSeedLingPart.Bodies
            If iObject.Name Like "Body?BE*" Then Set oBody = iObject: Exit For
        Next
         
        If oBody Is Nothing Then
            MsgBox "No KBE Bodies found in selected part.     " & vbCrLf & "Please select a valid 3D Collector!     ", vbCritical, "Warning"
            GoTo noSelection
        End If
        
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

Private Sub optClean_Change()
    Call SetFormMode
End Sub

Private Sub imgCleanFalse_Click()
    Me.optClean.Value = True
End Sub

Private Sub optCreate_Change()
    Call SetFormMode
End Sub

Private Sub ImgCreateFalse_Click()
    Me.optCreate.Value = True
End Sub

Private Sub optRelease_Change()
    Call SetFormMode
End Sub

Private Sub imgReleaseFalse_Click()
    Me.optRelease.Value = True
End Sub

Private Sub txtRangeKBEBodies_Change()
    If Me.txtRangeKBEBodies.Value <> "" And Val(Me.txtRangeKBEBodies.Text) > 0 Then iNoOfBodies = Val(Me.txtRangeKBEBodies.Text)
End Sub

Private Sub txtRangeKBEBodies_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii > 47 And KeyAscii < 58) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRangeKBEBodies_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Val(Me.txtRangeKBEBodies.Text) < 1 Then Me.txtRangeKBEBodies.Text = 1
    If Me.txtRangeKBEBodies.Text = "" Then Me.txtRangeKBEBodies.Text = iNoOfBodies
    If Val(Me.txtRangeKBEBodies.Text) > iMaxRangeKBEBodies Then Me.txtRangeKBEBodies.Text = iMaxRangeKBEBodies
End Sub

Private Sub fraOptions_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     
    If Me.txtRangeKBEBodies.Enabled = True Then
        If Val(Me.txtRangeKBEBodies.Text) < 1 Then Me.txtRangeKBEBodies.Text = 1
        If Me.txtRangeKBEBodies.Text = "" Then Me.txtRangeKBEBodies.Text = iNoOfBodies
        If Val(Me.txtRangeKBEBodies.Text) > iMaxRangeKBEBodies Then Me.txtRangeKBEBodies.Text = iMaxRangeKBEBodies
    End If
    
End Sub

Private Sub cmbCAGECode_Change()
    If Me.cmbCAGECode.Value <> "" Then sCageCode = Me.cmbCAGECode.Value
End Sub

Private Sub cmbSupplierCode_Change()
    If Me.cmbSupplierCode.Value <> "" Then sSupplierCode = Me.cmbSupplierCode.Value
End Sub

Private Sub txtSelection3DCollector_Change()
    
    If Not txtSelection3DCollector = sNoSelection And Not oSeedLingPart Is Nothing Then
        
        iMaxRangeKBEBodies = 0
        For i = 1 To oSeedLingPart.Bodies.Count
            If oSeedLingPart.Bodies.Item(i).Name Like "*_KBE_*" Or oSeedLingPart.Bodies.Item(i).Name Like "*_BodyDBE*" Then iMaxRangeKBEBodies = iMaxRangeKBEBodies + 1
        Next i
        
        If iMaxRangeKBEBodies > 0 Then
            Me.lblRangeKBEBodies.Caption = "Range [Max: " & iMaxRangeKBEBodies & "]"
            If Val(Me.txtRangeKBEBodies.Text) > iMaxRangeKBEBodies Or Me.txtRangeKBEBodies.Text = "" Then Me.txtRangeKBEBodies.Text = iMaxRangeKBEBodies
        End If
        Me.txtRangeKBEBodies.Enabled = True
        Me.txtRangeKBEBodies.BackColor = &H80000005
    Else
        Me.txtRangeKBEBodies.Value = ""
        Me.txtRangeKBEBodies.Enabled = False
        Me.txtRangeKBEBodies.BackColor = &H80000016
    End If
    
    Call cmdOKActivation
    
End Sub

Private Sub fillComboBoxes()
        
    Dim sPropertyFilePath As String
    Dim iTxtFileLine As Long
    Dim sParameterName As String
    Dim sParameterValues As String
    Dim iValueIndex As Long
    
    'Empty Combo Boxes
    Me.cmbCAGECode.Clear
    Me.cmbSupplierCode.Clear
        
    'Get property file
    sPropertyFilePath = GetTagValueFromFile(sKBEPathFile, "Help Values")        '***** Path for Property file
    
    On Error Resume Next
    iTxtFileLine = 0
    Do
        iTxtFileLine = iTxtFileLine + 1
        sParameterName = ReadTextFile(sPropertyFilePath, iTxtFileLine)
        If sParameterName = "" Then Exit Do
        
        iTxtFileLine = iTxtFileLine + 1
        sParameterValues = ReadTextFile(sPropertyFilePath, iTxtFileLine)
        
        If sParameterName Like "*CAGE Code*" Then
            For iValueIndex = 0 To UBound(Split(sParameterValues, ":"))
                If Split(sParameterValues, ":")(iValueIndex) <> "" Then Me.cmbCAGECode.AddItem Split(sParameterValues, ":")(iValueIndex)
            Next iValueIndex
        ElseIf sParameterName Like "*Major Supplier Code*" Then
            For iValueIndex = 0 To UBound(Split(sParameterValues, ":"))
                If Split(sParameterValues, ":")(iValueIndex) <> "" Then Me.cmbSupplierCode.AddItem Split(sParameterValues, ":")(iValueIndex)
            Next iValueIndex
        End If
        
        If Me.cmbCAGECode.ListCount > 0 And Me.cmbSupplierCode.ListCount > 0 Then Exit Do
        
    Loop Until sParameterName = "" Or sParameterName = "END"
    
End Sub
