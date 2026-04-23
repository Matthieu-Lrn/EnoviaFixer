Attribute VB_Name = "frmStdPartsInstantiation"
Attribute VB_Base = "0{6EA70C0E-9D9B-4EE9-9AED-23BF999FF1C8}{9FF14A7C-C272-4B65-9529-10917B69EC09}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private oInstantiateStdParts As New InstantiateSTDParts
Private ActivityRuleSelected As Boolean

'********************************************************************************
'* Name: Instantiate Standard parts
'* Purpose: Create a new instance for selected standard parts, create a constraint
'*          to put it in position, and if needed create an activity rule at product level
'*
'* Assumption: See below
'*
'* Author:
'* Updated by: François Charette, Julien Bigaouette
'* Language: VBA
'********************************************************************************
'*
'* A constraint name has the following nomenclature:
'* -  Std Part instance name | Axis Name | Seed Part Number
'* Example: "OVALV-1PLC1.23 | SP_M77_Door_BA02_DatumAxisLatch | KBE_M77_SEED_PART_BA-V06"
'*
'******************************************************************************************
'*
'* A Std Part boolean parameter has the following nomenclature:
'*
'* - SP_zzzzzzzzzzz_yyyyyyyyy_Boolean
'* or
'* - SA_zzzzzzzzzzz_yyyyyyyyy_Boolean
'*
'* The corresponding axis system must be named:
'* - SP_zzzzzzzzzzz or SA_zzzzzzzzzzz   Note that the string "zzzzzzzzzz" can contain "_"
'*
'* The corresponding Part Number will be:
'* - "yyyyyyyy"  Note that a "-" found in the Part Number must be replaced by a "_" in the string "zzzzzzzzzzzz"
'*
'******************************************************************************************

Public Sub initializeStdPartsForm()
    
    Call enableRelationUpdate
    
    With Me
        '***** Window size according to number of shifters
        .Height = dTitleBarHeight + 162
        .Width = 182
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdCancel.Top = .Height - dTitleBarHeight - 24
        .cmdCancel.Left = .Width - .cmdCancel.Width - 12
        .cmdCancel.SetFocus
        .cmdOK.Top = .Height - dTitleBarHeight - 24
        .cmdOK.Left = .Width - .cmdOK.Width - .cmdCancel.Width - 18
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        .Caption = "Standard Parts Instantiation"
        
        
        '***** Default values
        .txtSelectionSA.Text = sNoSelection
        .optBoolParamTrue.Value = True
        .chkDeleteStdPart.Value = False
        .chkActivityRule.Value = False
        .Show vbModeless
    End With
    
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
        Set sSourceWindow = Nothing
        Set oSeedPartCollection = New Collection
    End If
    
    If bToolbarCommandSelected = False Then Me.Show vbModeless
    
End Sub

Private Sub txtSelectionSA_Change()
    If Not (Me.txtSelectionSA.Text = sNoSelection) Then
        Me.cmdOK.ForeColor = -2147483630
        Me.cmdOK.Locked = False
        Me.cmdOK.SetFocus
        Me.cmdOK.Default = True
    Else
        Me.cmdOK.ForeColor = -2147483631
        Me.cmdOK.Locked = True
        Me.cmdCancel.SetFocus
        Me.cmdCancel.Default = True
    End If
End Sub

Private Sub optBoolParamTrueFalse_Change()
    If Me.optBoolParamTrueFalse.Value Then
        ActivityRuleSelected = Me.chkActivityRule.Value
        Me.chkActivityRule.Value = True
        Me.chkActivityRule.Enabled = False
    Else
        Me.chkActivityRule.Value = ActivityRuleSelected
        Me.chkActivityRule.Enabled = True
    End If
End Sub

Private Sub cmdOK_Click()
    
    Me.Hide
    CATIA.DisplayFileAlerts = False
'    CATIA.Interactive = False 'Do not allow user to click something
'    frmKBEMain.Enabled = False 'Do not allow user to click something
    
    Call oInstantiateStdParts.launchInstantiation
    
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
